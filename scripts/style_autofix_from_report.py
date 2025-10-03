#!/usr/bin/env python3
import re, sys, json, csv
from pathlib import Path

# --- lightweight normalizers (Gengo-aligned mechanics, no voice changes) ---
JP_TO_ASCII = str.maketrans({
    "、": ", ", "。": ". ", "「": '"', "」": '"', "（": "(", "）": ")", "［": "[", "］": "]",
    "！": "!", "？": "?", "：": ":", "；": ";", "／": "/", "　": " ", "～": "~", "％": "%", "￥": "¥",
})

STOPWORDS = set([
    "a","an","and","as","at","but","by","for","from","in","into","nor","of","on",
    "or","over","per","the","to","via","with"
])

def ascii_punct(s:str)->str:
    return (s or "").translate(JP_TO_ASCII)

def collapse_space(s:str)->str:
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"\s+([,.;:!?%)])", r"\1", s)  # no space before closing punct
    s = re.sub(r"([(])\s+", r"\1", s)         # no space right after opening (
    s = re.sub(r"\s+([%])", r"\1", s)         # 50 % -> 50%
    return s.strip()

def normalize_dashes(s:str)->str:
    # number ranges: 5-10 -> 5–10 (en dash)
    return re.sub(r"(?<=\d)\s*-\s*(?=\d)", "–", s)

def currency_percent_units(s:str)->str:
    # ¥ 120,000 -> ¥120,000
    s = re.sub(r"¥\s+(\d)", r"¥\1", s)
    # 10GB -> 10 GB (for common units)
    s = re.sub(r"(?i)(\d)(?=(kb|mb|gb|tb|pb|%)\b)", r"\1 ", s)
    # 50 % -> 50%
    s = s.replace(" %", "%")
    return s

def quotes_ellipses(s:str)->str:
    s = s.replace("…", "...")  # safer for fonts
    return s

def kill_bullet_trailer(s:str)->str:
    # For short fragments (typical bullets), drop trailing . ; :
    t = s.strip()
    if len(t) <= 90 and not re.search(r"[.!?].+\w", t):  # no full sentence inside
        t = re.sub(r"[.;:]\s*$", "", t)
    return t

def sentence_case_fragment(s:str)->str:
    # Keep first letter as-is if ALLCAPS/Proper nouns; otherwise minimal tweak
    if re.match(r"^[A-Z0-9\W]+$", s.strip()):
        return s
    # If looks like a title ALL words capitalized, downcase mid-words except proper nouns is risky -> skip
    return s

def title_case_simple(s:str)->str:
    # Simple Title Case (best-effort); use only when report flags "title_case"
    words = re.split(r"(\s+)", s.strip())
    out=[]
    for i,w in enumerate(words):
        if not w.strip(): out.append(w); continue
        core = re.sub(r"^[\"'(\[]|[)\"'\]]$", "", w)  # naive strip outer quotes/brackets
        low = core.lower()
        if i==0 or i==len(words)-1 or low not in STOPWORDS:
            tc = core[:1].upper()+core[1:]
        else:
            tc = low
        out.append(w.replace(core, tc))
    return "".join(out)

def apply_mechanics(s:str, *, drop_bullet_trailer=False, dash=True, quotes=True, currency=True):
    if not isinstance(s,str): return s
    t = s
    t = ascii_punct(t)
    t = quotes_ellipses(t) if quotes else t
    t = normalize_dashes(t) if dash else t
    t = currency_percent_units(t) if currency else t
    t = collapse_space(t)
    if drop_bullet_trailer:
        t = kill_bullet_trailer(t)
    return t

# --- main ---
def main():
    if len(sys.argv) < 3:
        print("usage: python scripts/style_autofix_from_report.py bilingual_STYLE_REPORT.csv translation_cache.json [--aggressive]")
        sys.exit(2)
    report_path = Path(sys.argv[1])
    cache_path = Path(sys.argv[2])
    aggressive = ("--aggressive" in sys.argv)

    cache = json.loads(cache_path.read_text(encoding="utf-8"))
    # Collect keys touched by style report (if present). If not, we still run mechanics over all.
    keys_touched = set()

    # The report is project-specific; we accept any columns and rely only on `key` and `rule` if available.
    rules_by_key = {}
    if report_path.exists():
        with report_path.open(encoding="utf-8") as f:
            rdr = csv.DictReader(f)
            for row in rdr:
                k = row.get("key") or row.get("cache_key") or ""
                r = (row.get("rule") or row.get("category") or "").lower()
                if not k: continue
                keys_touched.add(k)
                rules_by_key.setdefault(k, set()).add(r)

    changed = 0
    for k, v in list(cache.items()):
        if not isinstance(v, str): continue

        # Decide if this looks like a bullet-ish fragment by key heuristic
        is_fragment = ("|para" in k) or ("|shape" in k)
        drop_trailer = is_fragment

        new = apply_mechanics(v, drop_bullet_trailer=drop_trailer)

        # Optional: apply simple Title Case only if report flagged it
        if aggressive and k in rules_by_key and any("title" in r for r in rules_by_key[k]):
            new = title_case_simple(new)

        if new != v:
            cache[k] = new
            changed += 1

    cache_path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Autofixed mechanics in cache: {changed} entries updated")

if __name__ == "__main__":
    main()
