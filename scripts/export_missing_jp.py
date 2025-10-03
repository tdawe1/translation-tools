#!/usr/bin/env python3
"""
export_missing_jp.py

Extract unique Japanese strings from a PPTX and optionally filter out items
already present in the translation cache. Produces a JSON file that you can
translate manually (e.g., in this chat) and then merge back into the cache.

Usage:
  python scripts/export_missing_jp.py --in input.pptx --out missing_jp.json \
    --cache translation_cache.json

Notes:
  - By default, only JP strings that are missing from the cache OR whose value
    equals the key (identity mapping) are exported.
  - Use --include-all to export all JP strings regardless of cache.
"""
import argparse, json, os, re, zipfile
from xml.etree import ElementTree as ET

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# Japanese and related ranges
JP_CORE = r"\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff"
CJK_PUNCT = r"\u3000-\u303f"
FULLWIDTH = r"\uff00-\uffef"
JP_ANY_RX = re.compile(f"[{JP_CORE}{CJK_PUNCT}{FULLWIDTH}]")


def normalize_para_text(p_el):
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


def extract_jp_strings(pptx_path: str) -> list[str]:
    strings = []
    with zipfile.ZipFile(pptx_path, "r") as z:
        slide_names = sorted(
            [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
        )
        for name in slide_names:
            root = ET.fromstring(z.read(name))
            for p in root.iter(A_NS + "p"):
                text = normalize_para_text(p)
                if text and JP_ANY_RX.search(text):
                    strings.append(text)
    # de-duplicate preserving order
    uniq = list(dict.fromkeys(strings))
    return uniq


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True, help="Input PPTX path")
    ap.add_argument("--out", dest="outp", default="missing_jp.json", help="Output JSON path")
    ap.add_argument("--cache", dest="cache", default="translation_cache.json", help="Translation cache JSON {JP: EN}")
    ap.add_argument("--include-all", action="store_true", help="Export all JP strings regardless of cache")
    args = ap.parse_args()

    if not os.path.exists(args.inp):
        raise SystemExit(f"Input not found: {args.inp}")

    uniq = extract_jp_strings(args.inp)

    cache = {}
    if os.path.exists(args.cache):
        with open(args.cache, "r", encoding="utf-8") as f:
            try:
                cache = json.load(f)
            except Exception:
                cache = {}

    if args.include_all:
        missing = uniq
    else:
        missing = [s for s in uniq if s not in cache or cache.get(s) == s]

    # Write as a dict template for easy fill: {JP: ""}
    out_map = {s: "" for s in missing}
    with open(args.outp, "w", encoding="utf-8") as f:
        json.dump(out_map, f, ensure_ascii=False, indent=2)

    print(f"Found {len(uniq)} unique JP strings; {len(missing)} to translate.")
    print(f"Wrote template: {args.outp}")


if __name__ == "__main__":
    main()

