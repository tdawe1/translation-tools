#!/usr/bin/env python3
"""
translate_pptx_inplace.py

JA -> EN PowerPoint translator that replaces text in the original file while preserving layout.
- Parses PPTX XML directly (no extra libs required).
- Batches strings to the OpenAI API with a strict JSON response format.
- Caches translations (JSON sidecar) to avoid rework/re-costs.
- Emits a bilingual CSV for QA and a JSON audit report (remaining JP counts, etc.).

Usage:
  python translate_pptx_inplace.py --in input.pptx --out output_en.pptx \
    --model gpt-4o-2024-08-06

Production Presets:
  Conservative (rock-solid):  --model gpt-4o-2024-08-06 (auto batch 8-12, max retries)
  Balanced (recommended):     --model gpt-4o-2024-08-06 (auto batch 10-14) 
  Cost-lean (good quality):   --model gpt-4o-mini (auto batch 12-16)

Batch sizes auto-calculated based on model and token estimates.
Override with --batch N (8-24 recommended range).

Env:
  OPENAI_API_KEY must be set.
"""
import asyncio
import argparse, json, os, re, shutil, sys, time, zipfile, logging
from xml.etree import ElementTree as ET
from pathlib import Path
from datetime import datetime

def get_timestamped_filename(filepath):
    """Create a timestamped backup filename if the file exists."""
    if os.path.exists(filepath):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path_obj = Path(filepath)
        backup_name = f"{path_obj.stem}_{timestamp}{path_obj.suffix}"
        return backup_name
    return filepath

def backup_existing_files(cache_file, bilingual_csv, audit_json, log_file):
    """Backup existing output files with timestamps."""
    files_backed_up = []
    
    for filepath in [cache_file, bilingual_csv, audit_json, log_file]:
        if os.path.exists(filepath):
            backup_name = get_timestamped_filename(filepath)
            shutil.move(filepath, backup_name)
            files_backed_up.append(f"{filepath} -> {backup_name}")
    
    if files_backed_up:
        print("Backed up existing files:")
        for backup in files_backed_up:
            print(f"  {backup}")
        print()
    
    return files_backed_up

# Logging will be set up in main() after parsing args

# Import style consistency modules
try:
    from style_normalize import normalize_block, get_style_guide, apply_style_guide_to_prompt, detect_content_type as detect_content_type_from_text
    from style_checker import model_style_check, apply_style_fixes, run_style_check
    from pptx_format import apply_deck_formatting_profile
    from style_mechanics_normalize import normalize_punct, bullet_fragment
    STYLE_MODULES_AVAILABLE = True
except ImportError:
    print("Warning: Style modules not found. Running without style consistency features.")
    STYLE_MODULES_AVAILABLE = False

# ---- OpenAI client (official library) ----
try:
    from openai import OpenAI, AsyncOpenAI
from utils.gpt_adapter import GPT5Adapter
except Exception:
    print("ERROR: The 'openai' package is required. Install via: pip install openai", file=sys.stderr)
    raise

# ---- Regex helpers ----
JP_CORE = r'\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff'
CJK_PUNCT = r'\u3000-\u303f'
FULLWIDTH = r'\uff00-\uffef'
JP_ANY = re.compile(f'[{JP_CORE}{CJK_PUNCT}{FULLWIDTH}]')

# Masking patterns for fragile content
RX_NUM = re.compile(r"\d[\d,.\-\u2013%]*")
RX_URL = re.compile(r"https?://\S+|www\.\S+")
RX_CODE= re.compile(r"[A-Z]{2,}\d[\w\-]*")

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

# Global storage for notes content during processing
_slide_notes_content = {}

# Global storage for slides needing layout tightening
_slides_need_tightening = set()

def count_jp_chars(s: str) -> int:
    return len(JP_ANY.findall(s))

def mask_fragile(s):
    i, maps = 1, {}
    def do(rx, tag, s):
        nonlocal i
        def repl(m):
            nonlocal i
            k = f"⟦{tag}_{i}⟧"
            maps[k] = m.group(0)
            i += 1
            return k
        return rx.sub(repl, s)
    s = do(RX_URL,"URL",s); s = do(RX_NUM,"NUM",s); s = do(RX_CODE,"CODE",s)
    return s, maps

def unmask_fragile(s, maps):
    for k, v in maps.items():
        s = s.replace(k, v)
    return s

def normalize_para_text(p_el):
    """Extract full visible text for a paragraph (concatenate runs, insert '\n' for a:br)."""
    br_tag = A_NS + "br"
    t_tag = A_NS + "t"
    r_tag = A_NS + "r"

    parts = []
    for node in p_el:
        if node.tag == r_tag:
            t = node.find(t_tag)
            parts.append("" if t is None or t.text is None else t.text)
        elif node.tag == br_tag:
            parts.append("\n")
        else:
            t = node.find(f".//{t_tag}")
            if t is not None and t.text:
                parts.append(t.text)

    return "".join(parts)

def set_para_text(p_el, new_text: str):
    """Word-aware replacement. Preserves word boundaries and turns '\n' into <a:br/>."""
    t_tag = A_NS + "t"
    r_tag = A_NS + "r"
    br_tag = A_NS + "br"
    import re

    # Collect runs (preserve overall styling distribution), clear <a:br/> and run text
    runs = [child for child in p_el if child.tag == r_tag]
    if not runs:
        r = ET.Element(r_tag)
        ET.SubElement(r, t_tag).text = ""
        p_el.insert(0, r)
        runs = [r]

    for child in list(p_el):
        if child.tag == br_tag:
            p_el.remove(child)
    for r in runs:
        t = r.find(t_tag)
        if t is None:
            t = ET.SubElement(r, t_tag)
        t.text = ""

    # Tokenize: keep whitespace; use None sentinel for newline
    def tokenize(s): return re.findall(r"\S+|\s+", s)
    tokens = []
    lines = new_text.split("\n")
    for i, line in enumerate(lines):
        tokens.extend(tokenize(line))
        if i < len(lines) - 1:
            tokens.append(None)  # newline marker

    # Single run: dump text, insert <a:br/> at markers
    if len(runs) == 1:
        t = runs[0].find(t_tag)
        buf = []
        br_count = 0
        for tok in tokens:
            if tok is None:
                # Insert <a:br/> after the run
                br = ET.Element(br_tag)
                run_idx = list(p_el).index(runs[0])
                p_el.insert(run_idx + 1 + br_count, br)
                br_count += 1
            else:
                buf.append(tok)
        t.text = "".join(buf).strip()
        return

    # Multi-run: distribute on word boundaries proportional to original text lengths
    orig_lens = [len((r.find(t_tag).text or "")) for r in runs]
    total_words = sum(len(x) for x in tokens if isinstance(x, str))
    total_base = sum(orig_lens) or total_words or 1
    targets = []
    acc = 0
    for L in orig_lens:
        share = round(total_words * (L / total_base))
        targets.append(share)
        acc += share
    if targets:
        targets[-1] += (total_words - acc)  # fix rounding drift

    def consume(n_chars):
        taken = []
        count = 0
        while tokens:
            tok = tokens[0]
            if tok is None:  # stop before newline; caller will insert <a:br/>
                break
            need = len(tok)
            # respect word boundaries
            if count > 0 and not tok.isspace() and count + need > n_chars:
                break
            taken.append(tokens.pop(0))
            count += need
            if tokens and tokens[0] is None:
                break
        return "".join(taken)

    # Fill each run, inserting <a:br/> exactly where newlines occur
    for r, n in zip(runs, targets):
        t = r.find(t_tag)
        t.text = consume(n)
        while tokens and tokens[0] is None:
            tokens.pop(0)
            br = ET.Element(br_tag)
            run_idx = list(p_el).index(r)
            p_el.insert(run_idx + 1, br)

    # Any leftovers go into the last run
    if tokens:
        tail = "".join(tok for tok in tokens if isinstance(tok, str))
        last_t = runs[-1].find(t_tag)
        last_t.text = (last_t.text or "") + tail

def extract_all_paragraphs(z: zipfile.ZipFile, slide_range: set | None = None):
    """Return a flat list of (slide_name, paragraph_index, text)."""
    paras = []
    slide_files = sorted([n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")])

    if slide_range:
        filtered_slides = []
        for sf in slide_files:
            match = re.search(r'slide(\d+)\.xml', sf)
            if match and int(match.group(1)) in slide_range:
                filtered_slides.append(sf)
        slide_files = filtered_slides

    for sf in slide_files:
        root = ET.fromstring(z.read(sf))
        for idx, p_el in enumerate(root.iter(A_NS + "p")):
            text = normalize_para_text(p_el)
            if text.strip():
                paras.append((sf, idx, text))
    return paras, slide_files

def _ensure_autofit(root):
    # For every txBody, ensure <a:bodyPr><a:normAutofit/></a:bodyPr>
    for tx in root.iter(A_NS + "txBody"):
        bodyPr = tx.find(A_NS + "bodyPr")
        if bodyPr is None:
            bodyPr = ET.SubElement(tx, A_NS + "bodyPr")
        if bodyPr.find(A_NS + "normAutofit") is None and bodyPr.find(A_NS + "spAutoFit") is None:
            ET.SubElement(bodyPr, A_NS + "normAutofit")

def _ensure_autofit_on_tree(root, mode: str, font_scale_min: int, line_spacing_pct: int, tight_margins: bool):
    """Enable slide-wide text autofit & spacing. Safe, mechanical; does not change wording."""
    bodyPr_tag = A_NS + "bodyPr"
    norm_tag   = A_NS + "normAutofit"
    shape_tag  = A_NS + "spAutoFit"
    no_tag     = A_NS + "noAutofit"
    p_tag      = A_NS + "p"
    pPr_tag    = A_NS + "pPr"
    lnSpc_tag  = A_NS + "lnSpc"
    spcPct_tag = A_NS + "spcPct"

    # 1) Per text frame: set autofit and tighten insets if requested
    for bp in list(root.iter(bodyPr_tag)):
        # remove conflicting children
        for ch in list(bp):
            if ch.tag in (norm_tag, shape_tag, no_tag):
                bp.remove(ch)
        if mode == "none":
            bp.append(ET.Element(no_tag))
        elif mode == "shape":
            bp.append(ET.Element(shape_tag))
        else:
            na = ET.Element(norm_tag)
            # allow PowerPoint to shrink fonts and line spacing to fit
            na.set("fontScale", str(font_scale_min))           # e.g., 90000 = 90%
            na.set("lnSpcReduction", "12000")                  # ~12% line-space reduction headroom
            bp.append(na)
        if tight_margins:
            # Insets: values are EMUs; these are conservative, non-zero margins
            bp.set("lIns", "45720")   # ~0.05"
            bp.set("rIns", "45720")
            bp.set("tIns", "22860")   # ~0.025"
            bp.set("bIns", "22860")

    # 2) Normalize line spacing to a fixed percentage (e.g., 100%)
    for p in list(root.iter(p_tag)):
        pPr = p.find(pPr_tag)
        if pPr is None:
            pPr = ET.SubElement(p, pPr_tag)
        ln = pPr.find(lnSpc_tag)
        if ln is None:
            ln = ET.SubElement(pPr, lnSpc_tag)
        sp = ln.find(spcPct_tag)
        if sp is None:
            sp = ET.SubElement(ln, spcPct_tag)
        sp.set("val", str(line_spacing_pct))

def _use_responses_api(model: str) -> bool:
    m = (model or "").lower()
    # Prefer Responses API for latest models like gpt-5 family
    return m.startswith("gpt-5") or os.getenv("OPENAI_USE_RESPONSES") == "1"

def make_array_schema(expected_len: int | None):
    """Build a strict JSON Schema for string arrays."""
    return {
        "name": "BatchArrayOfStrings",
        "schema": {
            "type": "array",
            "items": {"type": "string"},
            "minItems": 1
        },
        "strict": True
    }

def _responses_create(client, model: str, sys_prompt: str, user_payload: dict, temperature: float):
    # Use adapter for responses
    try:
        input_data = [
            {"role": "system", "content": [{"type": "input_text", "text": sys_prompt}]},
            {"role": "user", "content": [{"type": "input_text", "text": json.dumps(user_payload, ensure_ascii=False)}]}
        ]
        resp = client.responses_create(
            model=model,
            input=input_data,
            temperature=temperature,
            # reasoning and text stripped by adapter
        )
        return resp.strip() if resp else ""
    except Exception:
        raise

def _chat_create(client, model: str, sys_prompt: str, user_payload: dict, temperature: float):
    """Sync version with response_format fallback."""
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
            ],
            temperature=temperature,
            response_format={"type": "json_object"},
        )
    except Exception:
        # Fallback: schema in prompt
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": sys_prompt + "\nReturn ONLY a JSON array."},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
            ],
            temperature=temperature,
        )
    return resp.choices[0].message.content.strip()

async def _chat_create_async(client, model: str, sys_prompt: str, user_payload: dict, temperature: float):
    """Async version with response_format fallback."""
    try:
        resp = await client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
            ],
            temperature=temperature,
            response_format={"type": "json_object"},
        )
    except Exception:
        # Fallback: schema in prompt
        resp = await client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": sys_prompt + "\nReturn ONLY a JSON array."},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
            ],
            temperature=temperature,
        )
    return resp.choices[0].message.content.strip()

async def _responses_create_compat_async(aclient, *, model, input, temperature, json_schema, max_output_tokens):
    """Async Responses API wrapper with JSON schema fallback."""
    try:
        resp = await aclient.responses.create(
            model=model,
            input=input,
            temperature=temperature,
            max_output_tokens=max_output_tokens,
            response_format={"type": "json_schema", "json_schema": json_schema, "strict": True},
        )
    except TypeError as e:
        if "response_format" in str(e):
            # Fallback: inline schema in prompt
            schema_text = f"Return ONLY a valid JSON value matching this JSON Schema:\n{json.dumps(json_schema, indent=2)}"
            fallback_input = input.copy()
            if fallback_input and len(fallback_input) > 0:
                fallback_input[0]["content"] = schema_text + "\n\n" + fallback_input[0]["content"]
            
            resp = await aclient.responses.create(
                model=model,
                input=fallback_input,
                temperature=temperature,
                max_output_tokens=max_output_tokens,
            )
        else:
            raise
    
    # Extract content from response
    content = getattr(resp, "output_text", None)
    if not content and getattr(resp, "output", None):
        try:
            content = resp.output[0].content[0].text
        except Exception:
            content = None
    if not content and getattr(resp, "choices", None):
        content = resp.choices[0].message.content
    
    return content.strip() if content else ""

def _extract_json_array(s: str, expected_len: int):
    import json, re
    s = re.sub(r"^```(?:json)?|```$", "", s.strip(), flags=re.M)
    dec = json.JSONDecoder()
    in_str = esc = False; i = 0; n = len(s)
    while i < n:
        ch = s[i]
        if esc: esc = False
        elif ch == '\\' and in_str: esc = True
        elif ch == '"' and in_str: in_str = not in_str
        elif not in_str and ch == '[':
            try:
                obj, end = dec.raw_decode(s, i)
            except json.JSONDecodeError:
                i += 1; continue
            if isinstance(obj, list) and (expected_len == 0 or len(obj) >= expected_len):
                return obj[:expected_len] if expected_len else obj
            i = end; continue
        i += 1
    return None

def build_style_guide_text(style_preset: str, style_file: str | None) -> str:
    """Return style guide text used in prompts."""
    if style_file:
        try:
            return Path(style_file).read_text(encoding="utf-8")
        except Exception:
            pass
    if style_preset in ("gengo", "", None):
        # Default to project STYLE_GUIDE.md
        for candidate in ("STYLE_GUIDE.md", "./STYLE_GUIDE.md"):
            p = Path(candidate)
            if p.exists():
                return p.read_text(encoding="utf-8")
        # Fallback minimal if file not present
        return (
            "Mirror tone from Japanese; neutral–professional if ambiguous. "
            "Quotes use double marks; commas inside quotes. Serial comma for clarity. "
            "Dates Month Day, Year. Thousands separators. Tilde ranges → en dashes. "
            "Convert JP punctuation to EN. Bullet fragments, no terminal period. "
            "If too long: condense ~15% → Notes spill → shrink-to-fit."
        )
    if style_preset == "minimal":
        return "Translate naturally, keep numbers/URLs exact, preserve list structure."
    return ""

def make_producer_prompt(items, style_guide: str, glossary: dict) -> str:
    tone_inference = (
        "Tone & register: Infer and mirror tone from the Japanese. "
        "If ambiguous, default to neutral–professional. Do not add hype or weaken claims. "
        "Translate the entire block naturally (not word-by-word)."
    )
    return (
        f"{tone_inference}\n\n"
        "Preserve tags/placeholders exactly. Keep list structure.\n\n"
        f"STYLE_GUIDE (Gengo-aligned):\n{style_guide}\n\n"
        f"GLOSSARY:\n{json.dumps(glossary, ensure_ascii=False)}\n\n"
        f"ITEMS:\n{json.dumps(items, ensure_ascii=False)}"
    )

REVIEWER_INSTRUCTIONS = """Given JP source and the tagged EN candidate, check fidelity:
omissions, additions, number/url mismatches, glossary violations, tag integrity.
Then check mechanics (Gengo-aligned) and tone drift. Return JSON only.
Schema:
{
  "omissions": [], "additions": [],
  "number_mismatches": [], "url_mismatches": [],
  "glossary_violations": [],
  "tag_integrity": {"ok": true, "details": []},
  "mechanics": {
    "quotes_rule": true,
    "periods_commas_inside_quotes": true,
    "serial_comma_missed": [],
    "date_style_violations": [],
    "thousands_separator_missed": [],
    "range_dash_needed": []
  },
  "structure": {
    "bullet_terminal_punct": [],
    "parallelism_mismatch": []
  },
  "tone": {
    "over_formalized": false,
    "over_casual": false,
    "added_hype_terms": []
  }
}"""

def make_reviewer_prompt(jp_source, en_candidate, glossary, style_guide):
    return (
        f"{REVIEWER_INSTRUCTIONS}\n"
        f"STYLE_GUIDE (Gengo-aligned):\n{style_guide}\n"
        f"JP:\n{jp_source}\nEN:\n{en_candidate}"
    )

def calculate_expansion_ratio(original_jp: str, translated_en: str) -> float:
    """Calculate expansion ratio between Japanese and English text."""
    jp_len = len(original_jp.strip())
    en_len = len(translated_en.strip())
    return en_len / jp_len if jp_len > 0 else 1.0

def condense_text_block(client, model: str, text: str, target_ratio: float = 0.85) -> str:
    """Stage 1: Compress text by removing filler while preserving meaning."""
    if not text or len(text) < 50:  # Skip very short text
        return text
        
    reduction_pct = int((1 - target_ratio) * 100)
    prompt = f"""Shorten this English text by ~{reduction_pct}% while preserving all meaning.

REQUIREMENTS:
- Keep all numbers, URLs, and technical terms exactly as-is
- Preserve any markup tags or placeholders ⟦…⟧
- Use concise fragments for bullets, not full sentences
- Remove filler: "in order to"→"to", "utilize"→"use", "as well as"→"and"
- Drop unnecessary articles ("the", "a") and instances of "that"
- One verb per bullet; cut adverbs where possible
- Maintain professional tone and parallel structure
- Do NOT change meaning or remove actual content

Text to shorten:
{text}"""

    try:
        if _use_responses_api(model):
            resp = client.responses.create(
                model=model,
                reasoning_effort="high",
                text={"verbosity": "low"}, 
                input=[{"role": "user", "content": prompt}],
                response_format={"type": "text"},
                temperature=0.2,
            )
            content = getattr(resp, "output_text", None)
            if not content and getattr(resp, "output", None):
                try:
                    content = resp.output[0].content[0].text
                except Exception:
                    pass
            return content.strip() if content else text
        else:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2,
            )
            return resp.choices[0].message.content.strip()
    except Exception:
        return text  # Fallback to original if compression fails

def spill_to_notes(text_block: str, content_type: str = "bullet") -> tuple[str, str]:
    """Stage 2: Move overflow content to Notes with reference stub."""
    import re
    
    if content_type == "title":
        # For titles, just truncate at reasonable length and add ellipsis
        if len(text_block) > 80:  # Conservative title length
            words = text_block.split()
            truncated = []
            char_count = 0
            for word in words:
                if char_count + len(word) + 1 > 75:  # Leave room for ellipsis
                    break
                truncated.append(word)
                char_count += len(word) + 1
            
            stub_text = " ".join(truncated) + "..."
            spilled_content = f"Full title: {text_block}"
            return stub_text, spilled_content
    
    elif content_type == "bullet":
        # Split bullets at sentence boundaries or logical breaks
        sentences = re.split(r'(?<=[.!?;])\s+', text_block)
        if len(sentences) <= 1:
            # Single sentence - try to split at conjunctions or commas
            parts = re.split(r'\s*(?:,\s*(?:and|but|or)|;\s*)\s*', text_block)
            if len(parts) > 1:
                stub_text = parts[0] + " (detail → Notes)"
                spilled_content = f"Additional details: {" ".join(parts[1:])}"
                return stub_text, spilled_content
            else:
                # Last resort: split at halfway point on word boundary
                words = text_block.split()
                split_point = len(words) // 2
                stub_text = " ".join(words[:split_point]) + " (more → Notes)"
                spilled_content = f"Continued: {" ".join(words[split_point:])}"
                return stub_text, spilled_content
        else:
            # Multiple sentences - keep first, spill rest
            stub_text = sentences[0] + " (detail → Notes)"
            spilled_content = f"Additional details: {" ".join(sentences[1:])}"
            return stub_text, spilled_content
    
    elif content_type == "table":
        # For table cells, aggressive abbreviation + Notes reference
        words = text_block.split()
        if len(words) > 5:
            stub_text = " ".join(words[:3]) + "... (Notes)"
            spilled_content = f"Full content: {text_block}"
            return stub_text, spilled_content
    
    # Default fallback
    words = text_block.split()
    if len(words) > 8:
        stub_text = " ".join(words[:6]) + " (→Notes)"
        spilled_content = f"Complete text: {text_block}"
        return stub_text, spilled_content
    
    return text_block, ""  # No spill needed

def verify_content_integrity(original_jp: str, stub_en: str, notes_en: str, glossary: dict) -> bool:
    """Reviewer function: verify no numbers/URLs/glossary terms lost in split."""
    combined_en = stub_en + " " + notes_en
    
    # Check for numbers (including Japanese numerals and percentages)
    import re
    jp_numbers = re.findall(r'\d+(?:[,.]?\d+)*[%％]?', original_jp)
    en_numbers = re.findall(r'\d+(?:[,.]?\d+)*[%％]?', combined_en)
    
    if len(jp_numbers) != len(en_numbers):
        return False
    
    # Check URLs
    jp_urls = re.findall(r'https?://\S+|www\.\S+', original_jp)  
    en_urls = re.findall(r'https?://\S+|www\.\S+', combined_en)
    
    if len(jp_urls) != len(en_urls):
        return False
    
    # Check glossary terms are preserved
    for jp_term, en_term in glossary.items():
        if jp_term in original_jp and en_term not in combined_en:
            return False
    
    return True

def add_notes_to_slide(zout: zipfile.ZipFile, slide_name: str, notes_content: list[str]) -> None:
    """Add or update slide notes with spilled content."""
    if not any(notes_content):  # No notes to add
        return
        
    # Generate notes slide XML filename 
    slide_num = slide_name.split("slide")[1].split(".xml")[0]
    notes_name = f"ppt/notesSlides/notesSlide{slide_num}.xml"
    
    # Combine all non-empty notes content
    combined_notes = "\n\n".join(note for note in notes_content if note.strip())
    if not combined_notes.strip():
        return
    
    # Create basic notes slide XML structure
    notes_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="0" cy="0"/>
                    <a:chOff x="0" y="0"/>
                    <a:chExt cx="0" cy="0"/>
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Notes Placeholder"/>
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1"/>
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" idx="1"/>
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr/>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US"/>
                            <a:t>{combined_notes}</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
</p:notes>"""
    
    try:
        # Add notes slide to zip
        zout.writestr(notes_name, notes_xml.encode('utf-8'))
    except Exception:
        # If notes creation fails, continue without notes
        pass

def apply_layout_tightening(root, is_aggressive: bool = False):
    """Stage 3: Apply layout optimizations to buy space."""
    import xml.etree.ElementTree as ET
    
    # Find all text bodies and apply tightening
    for txBody in root.iter(A_NS + "txBody"):
        # Ensure autofit is enabled (shrink-to-fit)
        bodyPr = txBody.find(A_NS + "bodyPr")
        if bodyPr is None:
            bodyPr = ET.SubElement(txBody, A_NS + "bodyPr")
        
        # Set autofit with minimum font size guards
        if bodyPr.find(A_NS + "normAutofit") is None and bodyPr.find(A_NS + "spAutoFit") is None:
            normAutofit = ET.SubElement(bodyPr, A_NS + "normAutofit")
            # Set font scale limits to prevent text from becoming unreadable
            normAutofit.set("fontScale", "85000")  # Minimum 85% font scaling
            normAutofit.set("lnSpcReduction", "15000")  # Maximum 15% line spacing reduction
        
        # Tighten margins
        bodyPr.set("lIns", "36000")   # Left margin: 2pt (was default ~7pt)
        bodyPr.set("rIns", "36000")   # Right margin: 2pt  
        bodyPr.set("tIns", "18000")   # Top margin: 1pt (was default ~5pt)
        bodyPr.set("bIns", "18000")   # Bottom margin: 1pt
        bodyPr.set("wrap", "square")  # Ensure text wrapping
        
        # Apply paragraph-level optimizations
        for p in txBody.iter(A_NS + "p"):
            pPr = p.find(A_NS + "pPr")
            if pPr is None:
                pPr = ET.SubElement(p, A_NS + "pPr")
            
            # Tighten line spacing
            lnSpc = pPr.find(A_NS + "lnSpc")
            if lnSpc is None:
                lnSpc = ET.SubElement(pPr, A_NS + "lnSpc")
            spcPct = lnSpc.find(A_NS + "spcPct")
            if spcPct is None:
                spcPct = ET.SubElement(lnSpc, A_NS + "spcPct")
            spcPct.set("val", "110000")  # 110% line spacing (was default ~120%)
            
            # Remove extra spacing before/after paragraphs
            spcBef = pPr.find(A_NS + "spcBef")
            if spcBef is not None:
                pPr.remove(spcBef)
            spcAft = pPr.find(A_NS + "spcAft")  
            if spcAft is not None:
                pPr.remove(spcAft)
            
            # Optimize bullet indents
            lvl = int(pPr.get("lvl", "0"))
            if lvl > 0:
                # Tighten bullet indentation
                if lvl == 1:
                    pPr.set("marL", "228600")    # 0.32" left margin (was ~0.5")
                    pPr.set("indent", "-228600") # Hanging indent to align text
                elif lvl == 2:
                    pPr.set("marL", "457200")    # 0.64" left margin
                    pPr.set("indent", "-228600")
                else:
                    pPr.set("marL", str(228600 * (lvl + 1)))
                    pPr.set("indent", "-228600")
            
            # Apply font size guards to prevent unreadable text
            for r in p.iter(A_NS + "r"):
                rPr = r.find(A_NS + "rPr")
                if rPr is not None:
                    # Check if font size is specified
                    sz = rPr.get("sz")
                    if sz:
                        font_size = int(sz)
                        # Determine if this is likely a title based on context or size
                        is_title = font_size > 2800 or "title" in (p.get("class", "")).lower()
                        
                        # Set minimum font sizes
                        min_size = 1800 if is_title else 1100  # 18pt for titles, 11pt for body
                        if font_size < min_size:
                            rPr.set("sz", str(min_size))

def detect_content_type(para_element) -> str:
    """Detect if paragraph is title, bullet, or table content."""
    # Check parent elements and attributes for context
    parent = para_element.getparent() if hasattr(para_element, 'getparent') else None
    
    # Look for title indicators in parent shape properties  
    current = para_element
    while current is not None:
        if current.tag and "title" in current.tag.lower():
            return "title"
        if hasattr(current, 'getparent'):
            current = current.getparent()
        else:
            break
    
    # Check for bullet/list indicators
    pPr = para_element.find(A_NS + "pPr")
    if pPr is not None:
        if pPr.find(A_NS + "buChar") is not None or pPr.find(A_NS + "buAutoNum") is not None:
            return "bullet"
        if pPr.get("lvl") is not None and int(pPr.get("lvl", "0")) > 0:
            return "bullet"
    
    # Check for table context (simplified detection)
    if any("table" in str(elem.tag).lower() for elem in para_element.iter()):
        return "table"
    
    return "bullet"  # Default assumption

def apply_style_consistency_workflow(client, translations, original_items, glossary, deck_tone, offline_mode=False):
    """
    Apply comprehensive style consistency workflow:
    1. Style normalization (deterministic)
    2. Style checking with model (JSON diagnostics)
    3. Authority fixes (deterministic)
    
    Args:
        client: OpenAI client
        translations: List of translated strings
        original_items: Original Japanese strings for context
        glossary: Glossary dict for terminology consistency
        deck_tone: Deck tone profile
        
    Returns:
        Style-consistent translations
    """
    if not STYLE_MODULES_AVAILABLE:
        return translations
    
    # Stage 1: Deterministic style normalization
    normalized_translations = []
    for translation in translations:
        # Detect content type for appropriate normalization
        content_type = detect_content_type_from_text(translation)
        if content_type == 'title':
            normalized = normalize_punct(translation)
        else:
            normalized = bullet_fragment(normalize_punct(translation))
        normalized_translations.append(normalized)
    
    # Stage 2: Model-based style checking (if enabled and not in offline mode)
    enable_style_checking = os.getenv("ENABLE_STYLE_CHECKING", "1") == "1"
    if not offline_mode and enable_style_checking and _use_responses_api(os.getenv("OPENAI_MODEL", "gpt-5")):
        try:
            # Run style diagnostics
            diagnostics = model_style_check(client, normalized_translations, glossary, deck_tone)
            
            # Apply authority fixes based on diagnostics
            fixed_translations = apply_style_fixes(normalized_translations, diagnostics)
            
            return fixed_translations
            
        except Exception as e:
            import traceback
            print(f"Style checking failed, using normalized translations: {e}")
            print(f"Full traceback: {traceback.format_exc()}")
            return normalized_translations
    elif not offline_mode:
        # Fallback to local-only style checking for consistency (skip in offline mode)
        local_diagnostics = run_style_check(client, normalized_translations, glossary, deck_tone)
        fixed_translations = apply_style_fixes(normalized_translations, local_diagnostics)
        return fixed_translations
    else:
        # In offline mode, just return the normalized translations
        logging.info("Offline mode: skipping style checking")
        return normalized_translations

def mock_translate(items):
    """Generate mock translations for offline testing."""
    mock_translations = []
    for i, item in enumerate(items):
        if not item or not item.strip():
            mock_translations.append("")
            continue
        
        # Generate predictable mock translation
        mock = f"Mock EN {i+1}: {item[:20]}..." if len(item) > 20 else f"Mock EN {i+1}: {item}"
        mock_translations.append(mock)
    
    return mock_translations

def batch_translate(client, model: str, items, glossary, offline_mode=False):
    """Translate list of strings JA->EN. Returns list of translations in order.
    Uses GPT-5 reasoning model with deep thinking for best fidelity.
    Falls back to Chat Completions for non-GPT-5 models.
    Expects a strict JSON array output.
    """
    if offline_mode:
        logging.info(f"Running in offline mode - using mock translations for {len(items)} items")
        return mock_translate(items)
    
    global _slide_notes_content
    logging.debug(f"Starting batch translation of {len(items)} items with model {model}")
    # Apply masking to protect fragile content
    items_masked, maps = zip(*[mask_fragile(x) for x in items]) if items else ([], [])
    
    # Compose system prompt with style guide integration
    style_guide = build_style_guide_text(
        os.getenv("STYLE_PRESET", "gengo"), os.getenv("STYLE_GUIDE_FILE")
    )
    
    sys_prompt = make_producer_prompt(items, style_guide, glossary)

    user_payload = {
        "glossary": glossary or {},
        "strings": list(items_masked),
        "instructions": [
            "Return ONLY a JSON array of translated strings in the same order.",
            "No code fences, no commentary."
        ],
    }

    use_responses = _use_responses_api(model)
    # Allow temperature override
    try:
        temperature = float(os.getenv("OPENAI_TEMPERATURE", "0.6"))
    except Exception:
        temperature = 0.6

    for attempt in range(3):
        try:
            if use_responses:
                content = _responses_create(client, model, sys_prompt, user_payload, temperature)
            else:
                content = _chat_create(client, model, sys_prompt, user_payload, temperature)
        except Exception:
            # Backoff and retry on transient errors
            time.sleep(1 + attempt)
            continue

        # Try robust JSON parsing first
        data = _extract_json_array(content, len(items))
        if data:
            # Unmask fragile content in results
            out = [unmask_fragile(str(y), maps[i]) for i, y in enumerate(data)]
            
            # Apply expansion policy if text is too long
            if _use_responses_api(model) and os.getenv("ENABLE_EXPANSION_POLICY", "1") == "1":
                processed_out = []
                notes_content = []
                
                for i, (original, translated) in enumerate(zip(items, out)):
                    expansion_ratio = calculate_expansion_ratio(original, translated)
                    content_type = "bullet"  # Default; could be enhanced to detect titles/tables
                    
                    # Define thresholds by content type
                    threshold = 1.8 if "title" in original.lower() else (1.2 if "table" in original.lower() else 1.4)
                    
                    if expansion_ratio > threshold:
                        # Stage 1: Try compression first
                        condensed = condense_text_block(client, model, translated, target_ratio=0.85)
                        new_ratio = calculate_expansion_ratio(original, condensed)
                        
                        if new_ratio > threshold:
                            # Stage 2: Spill to Notes
                            stub_text, spilled_content = spill_to_notes(condensed, content_type)
                            
                            # Verify content integrity
                            if verify_content_integrity(original, stub_text, spilled_content, glossary or {}):
                                processed_out.append(stub_text)
                                notes_content.append(spilled_content)
                                # Still might need tightening
                                final_ratio = calculate_expansion_ratio(original, stub_text)
                                if final_ratio > (threshold * 0.9):  # Still close to threshold
                                    _slides_need_tightening.add(original)
                            else:
                                # Integrity check failed, use condensed version without spill
                                processed_out.append(condensed)
                                notes_content.append("")
                                # Definitely need tightening since spill failed
                                _slides_need_tightening.add(original)
                        else:
                            # Compression worked, check if still needs tightening
                            processed_out.append(condensed)
                            notes_content.append("")
                            if new_ratio > (threshold * 0.85):  # Still somewhat long
                                _slides_need_tightening.add(original)
                    else:
                        # Check if borderline case that could benefit from tightening
                        if expansion_ratio > (threshold * 0.8):  # Within 20% of threshold
                            _slides_need_tightening.add(original)
                        processed_out.append(translated)
                        notes_content.append("")
                
                # Store notes content globally for PPTX write-back
                # Map original text to notes content for lookup during processing
                for original, notes in zip(items, notes_content):
                    if notes.strip():
                        _slide_notes_content[original] = notes
                
                # Load deck tone profile
                deck_tone = None
                deck_tone_path = "deck_tone.json"
                if os.path.exists(deck_tone_path):
                    with open(deck_tone_path, "r", encoding="utf-8") as f:
                        deck_tone = json.load(f)

                # Apply style consistency workflow
                final_out = apply_style_consistency_workflow(client, processed_out, items, glossary, deck_tone, offline_mode)
                        
                return final_out
            else:
                # Load deck tone profile
                deck_tone = None
                deck_tone_path = "deck_tone.json"
                if os.path.exists(deck_tone_path):
                    with open(deck_tone_path, "r", encoding="utf-8") as f:
                        deck_tone = json.load(f)

                # Apply style consistency to simple path too
                final_out = apply_style_consistency_workflow(client, out, items, glossary, deck_tone, offline_mode)
                return final_out
            
        # Fallback to simple JSON parsing
        try:
            data = json.loads(content)
            if isinstance(data, list) and len(data) == len(items):
                out = [unmask_fragile(str(y), maps[i]) for i, y in enumerate(data)]
                
                # Apply expansion policy for fallback path too
                if _use_responses_api(model) and os.getenv("ENABLE_EXPANSION_POLICY", "1") == "1":
                    processed_out = []
                    notes_content = []
                    
                    for i, (original, translated) in enumerate(zip(items, out)):
                        expansion_ratio = calculate_expansion_ratio(original, translated)
                        content_type = "bullet"
                        threshold = 1.8 if "title" in original.lower() else (1.2 if "table" in original.lower() else 1.4)
                        
                        if expansion_ratio > threshold:
                            condensed = condense_text_block(client, model, translated, target_ratio=0.85)
                            new_ratio = calculate_expansion_ratio(original, condensed)
                            
                            if new_ratio > threshold:
                                stub_text, spilled_content = spill_to_notes(condensed, content_type)
                                if verify_content_integrity(original, stub_text, spilled_content, glossary or {}):
                                    processed_out.append(stub_text)
                                    notes_content.append(spilled_content)
                                    final_ratio = calculate_expansion_ratio(original, stub_text)
                                    if final_ratio > (threshold * 0.9):
                                        _slides_need_tightening.add(original)
                                else:
                                    processed_out.append(condensed)
                                    notes_content.append("")
                                    _slides_need_tightening.add(original)
                            else:
                                processed_out.append(condensed)
                                notes_content.append("")
                                if new_ratio > (threshold * 0.85):
                                    _slides_need_tightening.add(original)
                        else:
                            if expansion_ratio > (threshold * 0.8):
                                _slides_need_tightening.add(original)
                            processed_out.append(translated)
                            notes_content.append("")
                    
                    # Store notes content globally
                    for original, notes in zip(items, notes_content):
                        if notes.strip():
                            _slide_notes_content[original] = notes
                    
                    # Apply style consistency workflow
                    final_out = apply_style_consistency_workflow(client, processed_out, items, glossary, None, offline_mode)
                    
                    return final_out
                else:
                    # Apply style consistency to fallback path
                    final_out = apply_style_consistency_workflow(client, out, items, glossary, None, offline_mode)
                    return final_out
        except Exception:
            # Not valid JSON array; retry
            time.sleep(1 + attempt)
            continue

    return items

async def translate_batch(items, attempt=1, args=None, client=None, model=None, glossary=None, idx=None, json_debug_dir=None):
    """Translate a single batch with retry and split logic."""
    items_masked, maps = zip(*[mask_fragile(x) for x in items]) if items else ([], [])
    
    style_guide = build_style_guide_text(
        os.getenv("STYLE_PRESET", "gengo"), os.getenv("STYLE_GUIDE_FILE")
    )
    sys_prompt = make_producer_prompt(items, style_guide, glossary)
    user_payload = {
        "glossary": glossary or {},
        "strings": list(items_masked),
        "instructions": ["Return ONLY a JSON array of translated strings in the same order.", "No code fences, no commentary."],
    }
    
    temperature = float(os.getenv("OPENAI_TEMPERATURE", "0.6"))
    max_output_tokens = args.max_output_tokens or min(4096, 80 * len(items))
    
    # Build JSON schema
    json_schema = make_array_schema(len(items))
    
    # Prepare input for responses API
    input_messages = [
        {"role": "system", "content": [{"type": "input_text", "text": sys_prompt}]},
        {"role": "user", "content": [{"type": "input_text", "text": json.dumps(user_payload, ensure_ascii=False)}]}
    ]
    
    try:
        # Try responses API first
        if _use_responses_api(model):
            content = await _responses_create_compat_async(
                client, model=model, input=input_messages, 
                temperature=temperature, json_schema=json_schema, 
                max_output_tokens=max_output_tokens
            )
        else:
            # Fallback to chat completions
            content = await _chat_create_async(client, model, sys_prompt, user_payload, temperature)
        
        # Extract JSON array
        data = _extract_json_array(content, len(items))
        if data:
            out = [unmask_fragile(str(y), maps[i]) for i, y in enumerate(data)]
            logging.info(f"[Batch {idx}] Completed {len(out)} items")
            return {"idx": idx, "items": items, "translations": out}
        
        # If we get here, JSON parsing failed
        raise ValueError("Failed to extract valid JSON array")
        
    except Exception as e:
        # Write debug artifacts
        if json_debug_dir:
            os.makedirs(json_debug_dir, exist_ok=True)
            debug_file = os.path.join(json_debug_dir, f"batch_{idx}_attempt_{attempt}_raw.txt")
            schema_file = os.path.join(json_debug_dir, f"batch_{idx}_attempt_{attempt}_schema.json")
            prompt_file = os.path.join(json_debug_dir, f"batch_{idx}_attempt_{attempt}_prompt.txt")
            
            with open(debug_file, "w", encoding="utf-8") as f:
                f.write(content if 'content' in locals() else str(e))
            with open(schema_file, "w", encoding="utf-8") as f:
                json.dump(json_schema, f, indent=2)
            with open(prompt_file, "w", encoding="utf-8") as f:
                f.write(sys_prompt)
        
        # Retry logic
        if attempt <= args.max_retries:
            logging.warning(f"[Batch {idx}] Attempt {attempt} failed, retrying: {e}")
            await asyncio.sleep(0.5 * attempt)  # Brief backoff
            return await translate_batch(items, attempt + 1, args, client, model, glossary, idx, json_debug_dir)
        
        # Split logic
        elif args.on_batch_fail == "split" and len(items) > 1:
            logging.info(f"[Batch {idx}] Splitting batch of {len(items)} items after {args.max_retries} failed attempts")
            mid = len(items) // 2
            left_items = items[:mid]
            right_items = items[mid:]
            
            # Process both halves
            left_result = await translate_batch(left_items, 1, args, client, model, glossary, f"{idx}_L", json_debug_dir)
            right_result = await translate_batch(right_items, 1, args, client, model, glossary, f"{idx}_R", json_debug_dir)
            
            # Combine results
            combined_items = left_result["items"] + right_result["items"]
            combined_translations = left_result["translations"] + right_result["translations"]
            
            return {"idx": idx, "items": combined_items, "translations": combined_translations}
        
        else:
            raise ValueError(f"Failed to get valid JSON for batch {idx} after {args.max_retries} attempts")

async def run_async_translation(client, model, missing, glossary, batch_size, concurrency, args):
    """Run concurrent batch translations with robust error handling."""
    # Set up debug directory
    if args.json_debug_dir is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_debug_dir = f"run-{timestamp}/json_failures"
    else:
        json_debug_dir = args.json_debug_dir
    
    os.makedirs(json_debug_dir, exist_ok=True)
    
    sem = asyncio.Semaphore(concurrency)
    
    async def limited_batch(items, idx):
        async with sem:
            return await translate_batch(items, 1, args, client, model, glossary, idx, json_debug_dir)
    
    tasks = []
    for i in range(0, len(missing), batch_size):
        batch = missing[i:i+batch_size]
        task = limited_batch(batch, i // batch_size)
        tasks.append(task)
    
    results = await asyncio.gather(*tasks)
    
    # Reassemble in order
    cache_updates = {}
    for result in sorted(results, key=lambda x: x["idx"]):
        for item, trans in zip(result["items"], result["translations"]):
            cache_updates[item] = trans
    
    return cache_updates

def estimate_batch_size(model: str, avg_len: float, max_array_items: int = 20) -> int:
    """Auto-size batch based on model and item length."""
    if "mini" in model.lower():
        target_tokens, fallback = 8000, 14
    else:
        target_tokens, fallback = 10000, 12
    
    if avg_len > 0:
        tokens_per_item = avg_len * 2.5 + 50
        batch = int(target_tokens / tokens_per_item)
    else:
        batch = fallback
    
    return max(8, min(max_array_items, batch))

def warm_pass(client, warm_model, missing, glossary, batch_size):
    """Prefill cache with cheaper model."""
    print(f"Warm pass with {warm_model}: {len(missing)} items")
    cache_updates = {}
    
    i = 0
    while i < len(missing):
        batch = missing[i:i+batch_size]
        out = batch_translate(client, warm_model, batch, glossary)
        for s, t in zip(batch, out):
            cache_updates[s] = t
        i += batch_size
        print(f"Warm: {i}/{len(missing)}")
    
    return cache_updates

def get_flagged_items(cache, reviewer_results=None):
    """Get items flagged by reviewer for upgrade."""
    # Simplified: would parse reviewer JSON for real flagged items
    # For now, return empty set
    return set()

def upgrade_pass(client, model, flagged, glossary, batch_size):
    """Retranslate only flagged items."""
    if not flagged:
        return {}
    
    print(f"Upgrade pass: {len(flagged)} flagged items")
    flagged_list = list(flagged)
    cache_updates = {}
    
    for i in range(0, len(flagged_list), batch_size):
        batch = flagged_list[i:i+batch_size]
        out = batch_translate(client, model, batch, glossary)
        for s, t in zip(batch, out):
            cache_updates[s] = t
    
    return cache_updates

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True, help="Input PPTX")
    ap.add_argument("--out", dest="outp", required=True, help="Output PPTX path")
    ap.add_argument("--cache", default="translation_cache.json", help="Path to JSON cache")
    ap.add_argument("--bilingual_csv", default="bilingual.csv", help="Output bilingual CSV")
    ap.add_argument("--audit_json", default="audit.json", help="Audit report JSON")
    ap.add_argument("--glossary", default=None, help="Optional glossary JSON {JA: EN}")
    ap.add_argument("--model", default=os.getenv("OPENAI_MODEL", "gpt-5"))
    ap.add_argument("--batch", type=int, default=40, help="Batch size for API calls")
    ap.add_argument("--slides", default=None, help="Slide range, e.g., '1-6'")
    ap.add_argument("--style-preset", default="gengo", choices=["gengo","minimal"], help="Style preset to load into prompts (default: gengo)")
    ap.add_argument("--style-file", default=None, help="Path to custom style guide file")
    ap.add_argument("--concurrency", type=int, default=1, help="Number of concurrent API requests")
    ap.add_argument("--warm-with", default=None, help="Model for warm pass (e.g., gpt-4o-mini)")
    ap.add_argument("--upgrade-flagged", action="store_true", help="Retranslate only reviewer-flagged items")
    ap.add_argument("--auto-batch", action="store_true", help="Auto-size batches based on model")
    ap.add_argument("--fresh", action="store_true", help="Backup existing output files with timestamps before creating new ones")
    ap.add_argument("--offline", action="store_true", help="Run in offline mode using mock translations for testing")
    ap.add_argument("--cache-only", action="store_true", help="Do not call any API and do not mock; require all translations to exist in cache")
    ap.add_argument("--max-retries", type=int, default=2, help="Number of retry attempts before splitting batch (default: 2)")
    ap.add_argument("--on-batch-fail", choices=["split", "abort"], default="split", help="Action on batch failure: split or abort (default: split)")
    ap.add_argument("--json-debug-dir", default=None, help="Directory for JSON failure debug artifacts (default: run-<timestamp>/json_failures)")
    ap.add_argument("--max-array-items", type=int, default=20, help="Maximum array items for auto-batch (default: 20)")
    ap.add_argument("--max-output-tokens", type=int, default=None, help="Maximum output tokens (default: auto-calculated)")
    ap.add_argument("--autofit-mode", choices=["norm","shape","none"], default="norm",
        help="norm = shrink text to fit shape; shape = expand shape to fit text; none = disable autofit")
    ap.add_argument("--font-scale-min", type=int, default=90000,
        help="Minimum font scale (percent * 1000) for norm autofit; 90000 = 90%")
    ap.add_argument("--line-spacing-pct", type=int, default=100000,
        help="Paragraph line spacing percentage (100000 = 100%)")
    ap.add_argument("--tight-margins", action="store_true",
        help="Reduce text insets (left/right ≈0.5em, top/bottom ≈0.25em) to gain space")
    args = ap.parse_args()
    
    # Backup existing files if --fresh flag is used  
    if args.fresh:
        backup_existing_files(args.cache, args.bilingual_csv, args.audit_json, "translation.log")
    
    # Set up logging after potential backup
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('translation.log'),
            logging.StreamHandler(sys.stdout)
        ]
    )

    slide_range = set()
    if args.slides:
        parts = args.slides.split('-')
        if len(parts) == 2:
            start, end = int(parts[0]), int(parts[1])
            slide_range = set(range(start, end + 1))

    args.concurrency = max(1, args.concurrency)

    # Skip API setup in offline mode or cache-only mode
    if args.offline or args.cache_only:
        logging.info("Running in OFFLINE MODE - using mock translations")
        if args.cache_only:
            logging.info("Cache-only mode: will not translate any missing strings; cache must be complete")
        client = None  # No API client needed
    else:
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            logging.error("OPENAI_API_KEY not set in environment")
            sys.exit(2)
        
        base_url = os.getenv("OPENAI_BASE_URL", "").strip()
        if base_url:
            client = OpenAI(api_key=api_key, base_url=base_url)
        else:
            client = OpenAI(api_key=api_key)
    
    logging.info(f"Starting translation: {args.inp} -> {args.outp}")
    logging.info(f"Model: {args.model}, Batch size: {args.batch}")
    if args.offline:
        logging.info("Offline mode: generating mock translations for testing")

    glossary = {}
    if args.glossary and os.path.exists(args.glossary):
        with open(args.glossary, "r", encoding="utf-8") as f:
            glossary = json.load(f)

    cache = {}
    if os.path.exists(args.cache):
        with open(args.cache, "r", encoding="utf-8") as f:
            cache = json.load(f)

    with zipfile.ZipFile(args.inp, "r") as zin:
        paras, slide_files = extract_all_paragraphs(zin, slide_range)
    
    logging.info(f"Extracted {len(paras)} paragraphs from {len(slide_files)} slides")

    src_strings = [t for _, _, t in paras if JP_ANY.search(t)]
    uniq = list(dict.fromkeys(src_strings))
    # Treat identity-mapped entries as missing to avoid caching failures where source == target
    missing = [s for s in uniq if s not in cache or cache.get(s) == s]
    
    logging.info(f"Found {len(src_strings)} Japanese strings, {len(uniq)} unique")
    logging.info(f"Cache has {len(cache)} entries, {len(missing)} strings need translation")

    # In cache-only mode, require that there are no missing items
    if args.cache_only:
        if missing:
            sample = missing[:10]
            logging.error(
                "cache-only mode: %d strings missing from cache. Export them with 'python scripts/export_missing_jp.py --in %s --out missing_jp.json --cache %s', translate externally, then merge with 'python scripts/merge_into_cache.py --updates translated.json --cache %s'",
                len(missing), args.inp, args.cache, args.cache
            )
            # Also write a convenience file listing missing items
            try:
                miss_path = "missing_jp.cache_only.json"
                with open(miss_path, "w", encoding="utf-8") as f:
                    json.dump({s: "" for s in missing}, f, ensure_ascii=False, indent=2)
                print(f"Wrote template of missing items: {miss_path}")
            except Exception:
                pass
            sys.exit(3)
        else:
            logging.info("cache-only mode: cache covers all strings; proceeding without any API/mocks")

    batch_size = args.batch
    if args.auto_batch and missing:
        avg_len = sum(len(s) for s in missing) / len(missing)
        batch_size = estimate_batch_size(args.model, avg_len, args.max_array_items)
        print(f"Auto-batch: size={batch_size} (avg_len={avg_len:.1f})")

    # Warm pass
    if args.warm_with and args.warm_with != args.model:
        if args.cache_only:
            logging.info("cache-only mode: skipping warm pass")
        else:
            uncached = [s for s in missing if s not in cache]
            if uncached:
                warm_batch = batch_size
                if args.auto_batch:
                    avg_len = sum(len(s) for s in uncached) / len(uncached)
                    warm_batch = estimate_batch_size(args.warm_with, avg_len, args.max_array_items)
                updates = warm_pass(client, args.warm_with, uncached, glossary, warm_batch)
                cache.update(updates)
                with open(args.cache, "w", encoding="utf-8") as f:
                    json.dump(cache, f, ensure_ascii=False, indent=2)

    # Upgrade pass  
    if args.upgrade_flagged and not args.cache_only:
        flagged = get_flagged_items(cache)
        if flagged:
            updates = upgrade_pass(client, args.model, flagged, glossary, batch_size)
            cache.update(updates)
            with open(args.cache, "w", encoding="utf-8") as f:
                json.dump(cache, f, ensure_ascii=False, indent=2)
        # Refresh missing list
        missing = [s for s in uniq if s not in cache or cache.get(s) == s]

    # Main translation
    calls = 0  # Initialize calls counter
    if missing and not args.cache_only:
        if args.concurrency > 1:
            # Async path
            print(f"Async translation: {len(missing)} items, concurrency={args.concurrency}")
            if base_url:
                async_client = AsyncOpenAI(api_key=api_key, base_url=base_url)
            else:
                async_client = AsyncOpenAI(api_key=api_key)
            
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                updates = loop.run_until_complete(
                    run_async_translation(async_client, args.model, missing, glossary, batch_size, args.concurrency, args)
                )
                cache.update(updates)
                calls = (len(missing) + batch_size - 1) // batch_size  # Estimate calls for async
            finally:
                loop.close()
        else:
            # Sync path (existing)
            i = 0
            while i < len(missing):
                batch = missing[i:i+batch_size]
                out = batch_translate(client, args.model, batch, glossary, args.offline)
                calls += 1
                for s, t in zip(batch, out):
                    cache[s] = t
                i += batch_size
                print(f"Progress: {min(i, len(missing))}/{len(missing)}")
    elif args.cache_only:
        logging.info("cache-only mode: no translation performed; using cache as-is")

    with open(args.cache, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

    # Build bilingual CSV
    import csv
    with open(args.bilingual_csv, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["slide_xml","paragraph_idx","Japanese","English"])
        for sf, idx, jp in paras:
            en = cache.get(jp, jp)
            w.writerow([sf, idx, jp, en])

    # Write output PPTX
    tmp = args.outp + ".tmp"
    shutil.copyfile(args.inp, tmp)

    before_total = 0
    after_total = 0
    per_before = {}
    per_after = {}

    with zipfile.ZipFile(args.inp, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name in slide_files:
                root = ET.fromstring(data)
                texts = []
                for p in root.iter(A_NS + "p"):
                    t = normalize_para_text(p)
                    texts.append(t)
                per_before[name] = sum(count_jp_chars(t) for t in texts)
                before_total += per_before[name]

                changed = False
                for p in root.iter(A_NS + "p"):
                    src_text = normalize_para_text(p)
                    if src_text.strip() and JP_ANY.search(src_text):
                        tgt = cache.get(src_text)
                        if tgt:
                            set_para_text(p, tgt)
                            changed = True
                if changed:
                    # Apply Stage 3: Layout tightening for slides marked as needing it
                    if name in _slides_need_tightening:
                        apply_layout_tightening(root)
                    
                    # Apply consistent PPTX formatting profile
                    if STYLE_MODULES_AVAILABLE and os.getenv("ENABLE_FORMATTING_PROFILE", "1") == "1":
                        apply_deck_formatting_profile(root)
                    
                    _ensure_autofit(root)
                    
                    # After text replacement, make layout robust against EN overflow
                    _ensure_autofit_on_tree(
                        root,
                        args.autofit_mode,
                        args.font_scale_min,
                        args.line_spacing_pct,
                        args.tight_margins,
                    )
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    
                    # Process notes content for this slide
                    slide_notes = []
                    for p in root.iter(A_NS + "p"):
                        orig_text = normalize_para_text(p)
                        if orig_text in _slide_notes_content:
                            slide_notes.append(_slide_notes_content[orig_text])
                    
                    # Add notes to slide if any content was spilled
                    if slide_notes:
                        add_notes_to_slide(zout, name, slide_notes)

                # Recalc after
                root2 = ET.fromstring(data)
                txt2 = []
                for p in root2.iter(A_NS + "p"):
                    t = normalize_para_text(p)
                    txt2.append(t)
                per_after[name] = sum(count_jp_chars(t) for t in txt2)
                after_total += per_after[name]

            zout.writestr(name, data)

    os.replace(tmp, args.outp)

    with open(args.audit_json, "w", encoding="utf-8") as f:
        json.dump({
            "unique_strings": len(uniq),
            "api_calls": calls,
            "jp_chars_before": before_total,
            "jp_chars_after": after_total,
            "per_slide_before": per_before,
            "per_slide_after": per_after
        }, f, ensure_ascii=False, indent=2)

    # Run style consistency audit
    if STYLE_MODULES_AVAILABLE and os.getenv("ENABLE_STYLE_AUDIT", "1") == "1":
        try:
            from audit_style import run_full_audit, generate_audit_report, should_fail_ci
            
            # Load glossary for audit
            audit_glossary = {}
            if args.glossary and os.path.exists(args.glossary):
                with open(args.glossary, "r", encoding="utf-8") as f:
                    audit_glossary = json.load(f)
            
            # Run comprehensive style audit
            audit_results = run_full_audit(args.bilingual_csv, audit_glossary)
            
            # Generate report
            report_path = args.bilingual_csv.replace('.csv', '_STYLE_REPORT.csv')
            issue_count = generate_audit_report(audit_results, report_path)
            
            if issue_count > 0:
                print(f"Style issues found: {issue_count}")
                print(f"Style report: {report_path}")
                
                # Check if should fail (for CI integration)
                should_fail, reason = should_fail_ci(audit_results)
                if should_fail:
                    print(f"WARNING: {reason}")
            else:
                print("Style audit: PASSED")
                
        except Exception as e:
            print(f"Style audit failed: {e}")

    print("DONE")
    print("Output:", args.outp)
    print("Bilingual CSV:", args.bilingual_csv)
    print("Audit JSON:", args.audit_json)
    print("Remaining JP chars:", after_total)

if __name__ == "__main__":
    main()
