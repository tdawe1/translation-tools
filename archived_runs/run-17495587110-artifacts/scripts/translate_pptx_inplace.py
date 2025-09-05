#!/usr/bin/env python3
"""
translate_pptx_inplace.py

JA -> EN PowerPoint translator that replaces text in the original file while preserving layout.
- Parses PPTX XML directly (no extra libs required).
- Batches strings to the OpenAI API with a strict JSON response format.
- Caches translations (JSON sidecar) to avoid rework/re-costs.
- Emits a bilingual CSV for QA and a JSON audit report (remaining JP counts, etc.).

Enhancements (opt-in via env):
- USE_TAGS=1 — preserve inline formatting via lightweight tags [b] [i] [u] [sup] [sub].
- USE_PLACEHOLDERS=1 — lock numbers/URLs/IDs with ⟦…⟧ placeholders.
- ENABLE_AUTOFIT=1 — enable shrink-to-fit on text bodies to reduce overlaps.

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

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

def count_jp_chars(s: str) -> int:
    return len(JP_ANY.findall(s))

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

# ---- Inline formatting tags support ----
def extract_run_style(r_el):
    rpr = r_el.find(A_NS + "rPr")
    if rpr is None:
        return {}
    sty = {}
    if rpr.get("b") in ("1", "true"):
        sty["b"] = True
    if rpr.get("i") in ("1", "true"):
        sty["i"] = True
    if rpr.get("u") and rpr.get("u") != "none":
        sty["u"] = True
    base = rpr.get("baseline")
    if base:
        try:
            val = int(base)
            if val > 0:
                sty["sup"] = True
            elif val < 0:
                sty["sub"] = True
        except Exception:
            pass
    return sty

def style_open_tags(sty):
    s = ""
    if sty.get("b"): s += "[b]"
    if sty.get("i"): s += "[i]"
    if sty.get("u"): s += "[u]"
    if sty.get("sup"): s += "[sup]"
    if sty.get("sub"): s += "[sub]"
    return s

def style_close_tags(sty):
    s = ""
    if sty.get("sub"): s += "[/sub]"
    if sty.get("sup"): s += "[/sup]"
    if sty.get("u"): s += "[/u]"
    if sty.get("i"): s += "[/i]"
    if sty.get("b"): s += "[/b]"
    return s

def tagged_para_text(p_el):
    parts = []
    for node in p_el:
        if node.tag == A_NS + "r":
            t = node.find(A_NS + "t")
            txt = "" if t is None or t.text is None else t.text
            if not txt:
                continue
            sty = extract_run_style(node)
            parts.append(style_open_tags(sty) + txt + style_close_tags(sty))
        elif node.tag == A_NS + "br":
            parts.append("\n")
        else:
            t = node.find(f".//{A_NS}t")
            if t is not None and t.text:
                parts.append(t.text)
    return "".join(parts)

# ---- Placeholder masking ----
NUM_RE = re.compile(r"(?<!\w)(?:\d[\d,\.\-–%]*|\d{4})")
URL_RE = re.compile(r"https?://[^\s]+|\b[\w.%-]+@[\w.-]+\.[A-Za-z]{2,}\b")
CODE_RE = re.compile(r"\b[A-Z]{2,}[A-Z0-9\-_.]*\d+[A-Z0-9\-_.]*\b")

def mask_placeholders(s: str):
    mapping = {}
    ctr = {"NUM":0, "URL":0, "CODE":0}
    def repl(pattern, kind, text):
        def _r(m):
            ctr[kind]+=1
            key=f"⟦{kind}_{ctr[kind]}⟧"
            mapping[key]=m.group(0)
            return key
        return pattern.sub(_r, text)
    out = s
    out = repl(URL_RE, "URL", out)
    out = repl(NUM_RE, "NUM", out)
    out = repl(CODE_RE, "CODE", out)
    return out, mapping

def unmask_placeholders(s: str, mapping: dict):
    for k,v in mapping.items():
        s = s.replace(k, v)
    return s

def set_para_text(p_el, new_text: str):
    """Replace paragraph text while preserving number of runs (rough distribution).
    """
    t_tag = A_NS + "t"
    r_tag = A_NS + "r"

    runs = [child for child in p_el if child.tag == r_tag]
    if not runs:
        r = ET.Element(r_tag)
        t = ET.SubElement(r, t_tag)
        t.text = ""
        p_el.insert(0, r)
        runs = [r]

    N = len(runs)
    L = len(new_text)
    if N == 1:
        chunks = [new_text]
    else:
        base = L // N
        rem = L % N
        chunks = []
        start = 0
        for i in range(N):
            size = base + (1 if i < rem else 0)
            chunks.append(new_text[start:start+size])
            start += size

    for r, chunk in zip(runs, chunks):
        t = r.find(t_tag)
        if t is None:
            t = ET.SubElement(r, t_tag)
        t.text = chunk

    for r in runs[len(chunks):]:
        t = r.find(t_tag)
        if t is not None:
            t.text = ""

def set_para_text_tagged(p_el, tagged_text: str):
    """Replace paragraph content from tagged text, reconstructing runs for b/i/u/sup/sub.
    Tags: [b] [/b] [i] [/i] [u] [/u] [sup] [/sup] [sub] [/sub]
    """
    # Remove existing runs and line breaks
    for child in list(p_el):
        if child.tag in (A_NS+"r", A_NS+"br"):
            p_el.remove(child)

    # Parse simple tags
    segments = []  # list of (set(styles), text)
    stack = []
    buf = []
    s = tagged_text
    i = 0
    def flush():
        if buf:
            segments.append((set(stack), ''.join(buf)))
            buf.clear()
    while i < len(s):
        if s[i] == '[':
            j = s.find(']', i)
            if j != -1:
                tag = s[i+1:j]
                if tag in ("b","i","u","sup","sub"):
                    flush(); stack.append(tag); i = j+1; continue
                if tag in ("/b","/i","/u","/sup","/sub"):
                    flush(); tname = tag[1:]
                    for k in range(len(stack)-1, -1, -1):
                        if stack[k]==tname:
                            del stack[k]; break
                    i = j+1; continue
        buf.append(s[i]); i+=1
    flush()

    # Rebuild runs according to segments
    for styles, text in segments:
        if not text:
            continue
        r = ET.SubElement(p_el, A_NS+"r")
        rpr = ET.SubElement(r, A_NS+"rPr")
        if "b" in styles: rpr.set("b","1")
        if "i" in styles: rpr.set("i","1")
        if "u" in styles: rpr.set("u","sng")
        if "sup" in styles: rpr.set("baseline","30000")
        if "sub" in styles: rpr.set("baseline","-25000")
        t = ET.SubElement(r, A_NS+"t")
        if text and (text[0].isspace() or text.endswith(' ')):
            t.set("xml:space","preserve")
        t.text = text

def extract_all_paragraphs(z: zipfile.ZipFile, slide_range: set = None):
    """Return a flat list of (slide_name, paragraph_index, text).
    If USE_TAGS=1, paragraph text includes inline tags.
    """
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
            text = tagged_para_text(p_el) if os.getenv("USE_TAGS") == "1" else normalize_para_text(p_el)
            if text.strip():
                paras.append((sf, idx, text))
    return paras, slide_files

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
    # Compose system prompt with optional style guide
    style_guide = build_style_guide_text(
        os.getenv("STYLE_PRESET", ""), os.getenv("STYLE_GUIDE_FILE")
    )
    sys_prompt = (
        "You are translating Japanese slide content to clear US-English for business decks. "
        "Translate faithfully and naturally. Do NOT add or remove content. Preserve line breaks. "
        "If inline tags like [b] [i] [u] [sup] [sub] or [li-lN]…[/li] are present, preserve them exactly. "
        "If placeholders like ⟦NUM_1⟧ ⟦URL_2⟧ ⟦CODE_3⟧ are present, preserve them exactly. "
        "Keep numbers, URLs, and variable-like tokens intact. "
        "Use sentence case for body; Title Case for slide titles when appropriate. "
        "Respect the glossary exactly when terms occur. "
        + ("\n" + style_guide if style_guide else "")
    )

    user_payload = {
        "glossary": glossary or {},
        "strings": items,
        "instructions": [
            "Return ONLY a JSON array of translated strings in the same order.",
            "Do not alter tag markers or placeholders; keep them exactly as provided.",
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
        except Exception as e:
            # Backoff and retry on transient errors
            time.sleep(1 + attempt)
            continue

        try:
            data = json.loads(content)
            if isinstance(data, list) and len(data) == len(items):
                return [str(x) for x in data]
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

    use_placeholders = os.getenv("USE_PLACEHOLDERS") == "1"

    # Prepare list of source strings (possibly masked)
    masks = {}
    src_strings = []
    for _, _, t in paras:
        if JP_ANY.search(t):
            if use_placeholders:
                masked, mp = mask_placeholders(t)
                masks[masked] = mp
                src_strings.append(masked)
            else:
                src_strings.append(t)

    uniq = list(dict.fromkeys(src_strings))
    # Treat identity-mapped or JP-like cached entries as missing to avoid caching failures
    missing = [
        s for s in uniq
        if s not in cache or cache.get(s) == s or (cache.get(s) and count_jp_chars(cache.get(s)) > 0)
    ]

    i = 0
    calls = 0
    while i < len(missing):
        batch = missing[i:i+args.batch]
        out = batch_translate(client, args.model, batch, glossary)
        calls += 1
        for s, t in zip(batch, out):
            if use_placeholders and s in masks:
                t = unmask_placeholders(t, masks[s])
            cache[s] = t
        i += args.batch

    with open(args.cache, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

    # Build bilingual CSV
    import csv
    with open(args.bilingual_csv, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Japanese", "English"])
        for s in uniq:
            src_disp = s
            if use_placeholders and s in masks:
                src_disp = unmask_placeholders(s, masks[s])
            w.writerow([src_disp, cache.get(s, src_disp)])

    # Write output PPTX
    tmp = args.outp + ".tmp"
    shutil.copyfile(args.inp, tmp)

    before_total = 0
    after_total = 0
    per_before = {}
    per_after = {}

    enable_autofit = os.getenv("ENABLE_AUTOFIT") == "1"
    use_tags = os.getenv("USE_TAGS") == "1"
    with zipfile.ZipFile(args.inp, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name in slide_files:
                root = ET.fromstring(data)
                texts = []
                for p in root.iter(A_NS + "p"):
                    t = tagged_para_text(p) if use_tags else normalize_para_text(p)
                    texts.append(t)
                per_before[name] = sum(count_jp_chars(t) for t in texts)
                before_total += per_before[name]

                changed = False
                for p in root.iter(A_NS + "p"):
                    src_text = tagged_para_text(p) if use_tags else normalize_para_text(p)
                    if src_text.strip() and JP_ANY.search(src_text):
                        key = src_text
                        if use_placeholders:
                            key, _ = mask_placeholders(src_text)
                        tgt = cache.get(key)
                        if tgt:
                            if use_tags:
                                set_para_text_tagged(p, tgt)
                            else:
                                set_para_text(p, tgt)
                            changed = True
                # Optionally enable shrink-to-fit on text bodies
                if enable_autofit:
                    for tx in root.findall(f".//{A_NS}txBody"):
                        bodyPr = tx.find(A_NS+"bodyPr")
                        if bodyPr is None:
                            bodyPr = ET.SubElement(tx, A_NS+"bodyPr")
                        if bodyPr.find(A_NS+"normAutofit") is None:
                            ET.SubElement(bodyPr, A_NS+"normAutofit")
                if changed:
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)

                # Recalc after
                root2 = ET.fromstring(data)
                txt2 = []
                for p in root2.iter(A_NS + "p"):
                    t = tagged_para_text(p) if use_tags else normalize_para_text(p)
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
