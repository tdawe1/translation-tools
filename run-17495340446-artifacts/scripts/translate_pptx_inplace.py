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
    --model gpt-4o --batch 40 --glossary glossary.json

Env:
  OPENAI_API_KEY must be set.
"""
import argparse, json, os, re, shutil, sys, time, zipfile
from xml.etree import ElementTree as ET

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

def count_jp_chars(s: str) -> int:
    return len(JP_ANY.findall(s))

def mask_fragile(s):
    i, maps = 1, {}
    def do(rx, tag, s):
        nonlocal i
        def repl(m):
            nonlocal i
            k = f"⟦{tag}_{i}⟧"; maps[k] = m.group(0); i += 1; return k
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
    """Word-aware replacement that preserves line breaks as <a:br/> and avoids mid-word splits."""
    t_tag = A_NS + "t"
    r_tag = A_NS + "r"
    br_tag = A_NS + "br"

    # Get existing runs (to preserve basic styling distribution if present)
    runs = [child for child in p_el if child.tag == r_tag]
    if not runs:
        r = ET.Element(r_tag)
        ET.SubElement(r, t_tag).text = ""
        p_el.insert(0, r)
        runs = [r]

    # Clear all run text and remove existing <a:br/>
    for child in list(p_el):
        if child.tag == br_tag:
            p_el.remove(child)
    for r in runs:
        t = r.find(t_tag) or ET.SubElement(r, t_tag)
        t.text = ""

    # Split translated text into "lines" (we'll re-insert <a:br/> nodes)
    lines = new_text.split("\n")

    # Tokenize by words but keep whitespace separators so we can reassemble cleanly
    def tokenize(s):
        return re.findall(r'\S+|\s+', s)

    # Build a sequence of text tokens and line break markers
    elements = []
    for li, line in enumerate(lines):
        elements.extend(tokenize(line))
        if li < len(lines) - 1:
            elements.append(None)  # sentinel = newline -> <a:br/>

    # Simple case: one run - put all text in it and insert <a:br/> elements properly
    if len(runs) == 1:
        run = runs[0]
        t_elem = run.find(t_tag)
        text_parts = []
        insert_pos = 0
        
        for elem in elements:
            if elem is None:
                # Insert accumulated text first
                if text_parts:
                    t_elem.text = "".join(text_parts)
                    text_parts = []
                
                # Insert <a:br/> after the run
                br = ET.Element(br_tag)
                parent = p_el
                run_idx = list(parent).index(run)
                parent.insert(run_idx + 1 + insert_pos, br)
                insert_pos += 1
            else:
                text_parts.append(elem)
        
        # Set any remaining text
        if text_parts:
            t_elem.text = (t_elem.text or "") + "".join(text_parts)
        
        return

    # Multiple runs: distribute text across runs while preserving word boundaries
    text_only = [e for e in elements if e is not None]
    total_chars = sum(len(e) for e in text_only)
    
    # Calculate target lengths based on original distribution
    orig_lens = [len((r.find(t_tag) or ET.Element(t_tag)).text or "") for r in runs]
    orig_total = sum(orig_lens) or total_chars or 1
    
    targets = []
    for orig_len in orig_lens:
        target = int(total_chars * (orig_len / orig_total)) if orig_total > 0 else total_chars // len(runs)
        targets.append(max(1, target))  # At least 1 char per run
    
    # Adjust for rounding errors
    diff = total_chars - sum(targets)
    if diff != 0 and targets:
        targets[-1] += diff

    # Distribute text tokens across runs
    text_idx = 0
    br_positions = []  # Track where to insert <a:br/> elements
    
    for run_idx, (run, target_len) in enumerate(zip(runs, targets)):
        t_elem = run.find(t_tag)
        current_len = 0
        run_text = []
        
        # Fill this run up to target length, respecting word boundaries
        while text_idx < len(text_only) and current_len < target_len:
            token = text_only[text_idx]
            token_len = len(token)
            
            # If adding this token would exceed target and we have some text already,
            # and it's not whitespace, stop here
            if current_len > 0 and current_len + token_len > target_len and not token.isspace():
                break
                
            run_text.append(token)
            current_len += token_len
            text_idx += 1
        
        t_elem.text = "".join(run_text)
    
    # Put any remaining text in the last run
    if text_idx < len(text_only):
        last_run = runs[-1]
        last_t = last_run.find(t_tag)
        remaining = "".join(text_only[text_idx:])
        last_t.text = (last_t.text or "") + remaining
    
    # Handle line breaks by inserting <a:br/> elements between runs where needed
    # This is a simplified approach - insert breaks based on original newline positions
    elem_idx = 0
    for i, elem in enumerate(elements):
        if elem is None:  # This was a newline
            # Find appropriate position to insert <a:br/>
            # Insert after first run as a simple approach
            if len(runs) > 1:
                br = ET.Element(br_tag)
                parent = p_el
                first_run_idx = list(parent).index(runs[0])
                parent.insert(first_run_idx + 1, br)

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
    # OpenAI Responses API
    try:
        # Support "high thinking" baseline via reasoning.effort
        effort = os.getenv("OPENAI_REASONING_EFFORT", "high")
        resp = client.responses.create(
            model=model,
            input=[
                {"role": "system", "content": [{"type": "output_text", "text": sys_prompt}]},
                {"role": "user", "content": [{"type": "input_text", "text": json.dumps(user_payload, ensure_ascii=False)}]},
            ],
            reasoning={"effort": effort},
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
    
    # Try to find JSON array patterns, being more flexible
    patterns = [
        r"\[\s*(?:\"|{|\d)",  # Original pattern
        r"\[\s*\{",           # Array of objects
        r"\[\s*\"",           # Array of strings
    ]
    
    for pattern in patterns:
        m = re.search(pattern, s, flags=re.S)
        if m:
            start = m.start()
            # Use JSONDecoder for more robust parsing
            decoder = json.JSONDecoder()
            try:
                # Try to decode from the array start position
                arr, idx = decoder.raw_decode(s, start)
                if isinstance(arr, list) and (expected_len == 0 or len(arr) >= expected_len):
                    # Accept arrays with >= expected length to handle chatty responses
                    return arr[:expected_len] if expected_len > 0 else arr
            except (json.JSONDecodeError, ValueError):
                # Fall back to manual bracket matching
                depth = 0
                in_string = False
                escape_next = False
                
                for i, ch in enumerate(s[start:], start):
                    if escape_next:
                        escape_next = False
                        continue
                    if ch == '\\' and in_string:
                        escape_next = True
                        continue
                    if ch == '"' and not escape_next:
                        in_string = not in_string
                    elif not in_string:
                        if ch == "[": 
                            depth += 1
                        elif ch == "]":
                            depth -= 1
                            if depth == 0:
                                frag = s[start:i+1]
                                try:
                                    arr = json.loads(frag)
                                    if isinstance(arr, list) and (expected_len == 0 or len(arr) >= expected_len):
                                        return arr[:expected_len] if expected_len > 0 else arr
                                except Exception:
                                    pass
                                break
    return None

def build_style_guide_text(style_preset: str, style_file: str | None) -> str:
    if style_file and os.path.exists(style_file):
        try:
            with open(style_file, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            pass

    preset = (style_preset or "").strip().lower()
    if preset in {"gengo", "gengo-ja-en", "gengo_ja_en"}:
        return (
            "Follow these JP→EN style rules (Gengo-inspired):\n"
            "- Tone: Natural, clear business English; avoid overly literal phrasing.\n"
            "- Honorifics: Omit honorifics unless required for meaning.\n"
            "- Formatting: Preserve line breaks and bullet structure. Do not add new bullets.\n"
            "- Punctuation: Use ASCII punctuation; convert full-width to half-width.\n"
            "- Capitalization: Sentence case for sentences; Title Case for slide/section titles.\n"
            "- Numerals: Keep numbers, percentages (%), units (e.g., GB), and URLs as-is.\n"
            "- Dates: Use target-locale English (e.g., January 5, 2025) where explicit.\n"
            "- Proper nouns: Keep brand/product capitalization; do not translate names.\n"
            "- Acronyms: Expand on first use if unclear, then use acronym.\n"
            "- Currency: Do not convert values; if symbol ambiguous, append ISO code (e.g., JPY).\n"
            "- Register: Prefer active voice; concise and persuasive B2B tone.\n"
            "- No additions: Do not summarize, omit, or invent content."
        )
    return ""

def batch_translate(client, model: str, items, glossary):
    """Translate list of strings JA->EN. Returns list of translations in order.
    Uses Responses API for gpt-5 models; falls back to Chat Completions otherwise.
    Expects a strict JSON array output.
    """
    # Apply masking to protect fragile content
    items_masked, maps = zip(*[mask_fragile(x) for x in items]) if items else ([], [])
    
    # Compose system prompt with optional style guide
    style_guide = build_style_guide_text(
        os.getenv("STYLE_PRESET", ""), os.getenv("STYLE_GUIDE_FILE")
    )
    sys_prompt = (
        "You are a professional Japanese-to-English translator for B2B marketing decks. "
        "Translate faithfully and naturally; keep the meaning and tone persuasive yet neutral. "
        "Do NOT summarize or add content. Preserve line breaks. "
        "Keep numbers, URLs, and variable-like tokens intact. "
        "Use sentence case for sentences; Title Case for slide titles where appropriate. "
        "Respect the glossary exactly when terms occur. "
        "Never return Japanese text in the output. If a term is untranslatable (product name, brand), retain it but translate surrounding text. "
        + ("\n" + style_guide if style_guide else "")
    )

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
        temperature = float(os.getenv("OPENAI_TEMPERATURE", "0.2"))
    except Exception:
        temperature = 0.2

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
            return out
            
        # Fallback to simple JSON parsing
        try:
            data = json.loads(content)
            if isinstance(data, list) and len(data) == len(items):
                out = [unmask_fragile(str(y), maps[i]) for i, y in enumerate(data)]
                return out
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
    args = ap.parse_args()

    slide_range = set()
    if args.slides:
        parts = args.slides.split('-')
        if len(parts) == 2:
            start, end = int(parts[0]), int(parts[1])
            slide_range = set(range(start, end + 1))

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        print("ERROR: Set OPENAI_API_KEY in environment.", file=sys.stderr)
        sys.exit(2)

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

    src_strings = [t for _, _, t in paras if JP_ANY.search(t)]
    uniq = list(dict.fromkeys(src_strings))
    # Treat identity-mapped entries as missing to avoid caching failures where source == target
    missing = [s for s in uniq if s not in cache or cache.get(s) == s]

    i = 0
    calls = 0
    while i < len(missing):
        batch = missing[i:i+args.batch]
        out = batch_translate(client, args.model, batch, glossary)
        calls += 1
        for s, t in zip(batch, out):
            cache[s] = t
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
                    _ensure_autofit(root)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)

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

    print("DONE")
    print("Output:", args.outp)
    print("Bilingual CSV:", args.bilingual_csv)
    print("Audit JSON:", args.audit_json)
    print("Remaining JP chars:", after_total)

if __name__ == "__main__":
    main()
