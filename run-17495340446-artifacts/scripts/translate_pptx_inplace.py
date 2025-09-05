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
    """Word-aware replacement. Preserves word boundaries and turns '\n' into <a:br/>."""
    t_tag = A_NS + "t"; r_tag = A_NS + "r"; br_tag = A_NS + "br"
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
        targets.append(share); acc += share
    if targets:
        targets[-1] += (total_words - acc)  # fix rounding drift

    def consume(n_chars):
        taken, count = [], 0
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
    dec = json.JSONDecoder()
    in_str = esc = False; i = 0; n = len(s)
    while i < n:
        ch = s[i]
        if esc: esc = False
        elif ch == '\\' and in_str: esc = True
        elif ch == '"': in_str = not in_str
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
