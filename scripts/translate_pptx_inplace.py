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
import argparse, json, os, re, shutil, sys, time, zipfile, logging
from xml.etree import ElementTree as ET
from pathlib import Path

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('translation.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

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
    from openai import OpenAI
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
        t = r.find(t_tag) or ET.SubElement(r, t_tag)
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

def _use_responses_api(model: str) -> bool:
    m = (model or "").lower()
    # Prefer Responses API for latest models like gpt-5 family
    return m.startswith("gpt-5") or os.getenv("OPENAI_USE_RESPONSES") == "1"

def _responses_create(client, model: str, sys_prompt: str, user_payload: dict, temperature: float):
    # OpenAI Responses API with GPT-5 reasoning model
    try:
        # Configure reasoning effort based on model - high for main translation, minimal for reviews
        if model.startswith("gpt-5-mini"):
            effort = "minimal"  # Fast reviewer
        else:
            effort = os.getenv("OPENAI_REASONING_EFFORT", "high")  # Deep thinking for translation
        
        resp = client.responses.create(
            model=model,
            input=[
                {"role": "system", "content": [{"type": "input_text", "text": sys_prompt}]},
                {"role": "user", "content": [{"type": "input_text", "text": json.dumps(user_payload, ensure_ascii=False)}]}
            ],
            reasoning={"effort": effort},
            text={"verbosity": "low"},  # Concise responses, avoid chatty prose
            temperature=temperature,
            response_format={"type": "json"},
        )
        # New SDKs expose output_text; fall back if absent
        content = getattr(resp, "output_text", None)
        if not content:
            # Fallback to choices/message style if present
            if getattr(resp, "choices", None):
                content = resp.choices[0].message.content
        if not content and getattr(resp, "output", None):
            try:
                # Attempt to read the first text content
                content = resp.output[0].content[0].text
            except Exception:
                content = None
        return content.strip() if content else ""
    except Exception:
        raise

def _chat_create(client, model: str, sys_prompt: str, user_payload: dict, temperature: float):
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
        ],
        temperature=temperature,
    )
    return resp.choices[0].message.content.strip()

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

def apply_style_consistency_workflow(client, translations, original_items, glossary, deck_tone):
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
    
    # Stage 2: Model-based style checking (if enabled)
    enable_style_checking = os.getenv("ENABLE_STYLE_CHECKING", "1") == "1"
    if enable_style_checking and _use_responses_api(os.getenv("OPENAI_MODEL", "gpt-5")):
        try:
            # Run style diagnostics
            diagnostics = model_style_check(client, normalized_translations, glossary, deck_tone)
            
            # Apply authority fixes based on diagnostics
            fixed_translations = apply_style_fixes(normalized_translations, diagnostics)
            
            return fixed_translations
            
        except Exception as e:
            print(f"Style checking failed, using normalized translations: {e}")
            return normalized_translations
    else:
        # Fallback to local-only style checking for consistency
        local_diagnostics = run_style_check(client, normalized_translations, glossary, deck_tone)
        fixed_translations = apply_style_fixes(normalized_translations, local_diagnostics)
        return fixed_translations

def batch_translate(client, model: str, items, glossary):
    """Translate list of strings JA->EN. Returns list of translations in order.
    Uses GPT-5 reasoning model with deep thinking for best fidelity.
    Falls back to Chat Completions for non-GPT-5 models.
    Expects a strict JSON array output.
    """
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
                final_out = apply_style_consistency_workflow(client, processed_out, items, glossary, deck_tone)
                        
                return final_out
            else:
                # Load deck tone profile
                deck_tone = None
                deck_tone_path = "deck_tone.json"
                if os.path.exists(deck_tone_path):
                    with open(deck_tone_path, "r", encoding="utf-8") as f:
                        deck_tone = json.load(f)

                # Apply style consistency to simple path too
                final_out = apply_style_consistency_workflow(client, out, items, glossary, deck_tone)
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
                    final_out = apply_style_consistency_workflow(client, processed_out, items, glossary)
                    
                    return final_out
                else:
                    # Apply style consistency to fallback path
                    final_out = apply_style_consistency_workflow(client, out, items, glossary)
                    return final_out
        except Exception:
            # Not valid JSON array; retry
            time.sleep(1 + attempt)
            continue

    return items

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
    args = ap.parse_args()

    slide_range = set()
    if args.slides:
        parts = args.slides.split('-')
        if len(parts) == 2:
            start, end = int(parts[0]), int(parts[1])
            slide_range = set(range(start, end + 1))

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        logging.error("OPENAI_API_KEY not set in environment")
        sys.exit(2)
    
    logging.info(f"Starting translation: {args.inp} -> {args.outp}")
    logging.info(f"Model: {args.model}, Batch size: {args.batch}")

    base_url = os.getenv("OPENAI_BASE_URL", "").strip()
    if base_url:
        client = OpenAI(api_key=api_key, base_url=base_url)
    else:
        client = OpenAI(api_key=api_key)

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

    # Auto-sizing batch logic based on model and content
    def calculate_optimal_batch_size(model: str, avg_tokens_per_item: int = 50) -> int:
        """Calculate optimal batch size based on model and estimated token usage."""
        # Model-specific token targets
        if model == "gpt-4o-mini":
            TARGET_REQ_TOKENS = 8_000
        elif model.startswith("gpt-4o-"):
            TARGET_REQ_TOKENS = 10_000
        else:
            TARGET_REQ_TOKENS = 8_000  # Conservative default
        
        MAX_ARRAY_ITEMS = 24  # Keep JSON outputs manageable
        
        # Estimate tokens (rough calculation)
        prefix_tokens = 1500  # System prompt + schema + glossary (often cached)
        tokens_per_block = avg_tokens_per_item * 2  # Input + expected output
        
        raw = int((TARGET_REQ_TOKENS - prefix_tokens) / tokens_per_block)
        batch_size = max(8, min(MAX_ARRAY_ITEMS, raw))  # Clamp to 8..24
        
        return batch_size
    
    # Use auto-sizing if no explicit batch size given, or validate the provided size
    if args.batch == 40:  # Default value, use auto-sizing
        optimal_batch = calculate_optimal_batch_size(args.model)
        logging.info(f"Auto-calculated optimal batch size: {optimal_batch} for model {args.model}")
        args.batch = optimal_batch
    else:
        # Validate user-provided batch size
        optimal_batch = calculate_optimal_batch_size(args.model)
        if args.batch > 24:
            logging.warning(f"Batch size {args.batch} > 24 may cause JSON truncation. Recommended: {optimal_batch}")
        elif args.batch < 8:
            logging.warning(f"Batch size {args.batch} < 8 may be inefficient. Recommended: {optimal_batch}")
    
    # ETA estimation setup
    from eta import ETAEstimator, fmt_hms
    eta = ETAEstimator(alpha=0.25)
    start = time.time()
    total_batches = (len(missing) + args.batch - 1) // args.batch
    total_items = len(missing)
    processed_items = 0
    retry_count = 0

    i = 0
    calls = 0
    while i < len(missing):
        batch = missing[i:i+args.batch]
        
        # Time the batch translation with retry handling
        t0 = time.time()
        batch_success = False
        attempt = 0
        
        while not batch_success and attempt < 3:
            try:
                out = batch_translate(client, args.model, batch, glossary)
                batch_success = True
            except Exception as e:
                attempt += 1
                retry_count += 1
                logging.warning(f"Batch failed (attempt {attempt}/3): {e}")
                
                # Check retry rate and adjust batch size if needed
                retry_rate = retry_count / max(1, calls + 1)
                if retry_rate > 0.05 and args.batch > 8:  # >5% retries
                    new_batch_size = max(8, int(args.batch * 0.75))  # Drop by 25%
                    logging.warning(f"High retry rate ({retry_rate:.1%}), reducing batch size: {args.batch} → {new_batch_size}")
                    args.batch = new_batch_size
                    total_batches = (len(missing) + args.batch - 1) // args.batch
                
                if attempt < 3:
                    time.sleep(1 + attempt)  # Backoff
                else:
                    logging.error(f"Batch failed after 3 attempts, skipping batch")
                    out = ["[TRANSLATION_FAILED]"] * len(batch)
        
        t1 = time.time()
        
        # Update ETA estimator
        batch_size = len(batch)
        eta.update((t1 - t0) / max(1, batch_size))
        processed_items += batch_size
        remaining_items = total_items - processed_items
        est_seconds = (remaining_items * (eta.sec_per_item or 0))
        
        calls += 1
        for s, t in zip(batch, out):
            cache[s] = t
        
        # Display progress with retry info
        elapsed = time.time() - start
        retry_info = f" | Retries: {retry_count}" if retry_count > 0 else ""
        logging.info(f"[{calls}/{total_batches}] Processed {processed_items}/{total_items} items | ETA {fmt_hms(est_seconds)} | Elapsed {fmt_hms(elapsed)}{retry_info}")
        
        i += args.batch

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