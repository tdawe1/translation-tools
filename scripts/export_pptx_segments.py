#!/usr/bin/env python3
"""
export_pptx_segments.py

Extract unique Japanese text segments from a PPTX and write a translation
template JSON (array of {"jp": ..., "en": ""}). Useful for fully local
workflows where translations are provided manually or by a local agent.

Usage:
  python scripts/export_pptx_segments.py --in inputs/source.pptx \
    --out translations/source_template.json
"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# XML namespace
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# Japanese character ranges (match translate_pptx_inplace.py)
JP_CORE = r"\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff"
CJK_PUNCT = r"\u3000-\u303f"
FULLWIDTH = r"\uff00-\uffef"
JP_ANY = re.compile(f"[{JP_CORE}{CJK_PUNCT}{FULLWIDTH}]")


def normalize_para_text(p_el: ET.Element) -> str:
    """Extract full visible text for a paragraph (respect <a:br/>)."""
    br_tag = A_NS + "br"
    t_tag = A_NS + "t"
    r_tag = A_NS + "r"

    parts: list[str] = []
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


def collect_unique_japanese_strings(pptx_path: Path) -> list[str]:
    """Return a de-duplicated list of Japanese-containing strings from PPTX."""
    if not pptx_path.exists():
        raise FileNotFoundError(pptx_path)

    all_strings: list[str] = []
    with zipfile.ZipFile(pptx_path, "r") as z:
        slide_files = sorted(
            n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        )
        for sf in slide_files:
            root = ET.fromstring(z.read(sf))
            for p in root.iter(A_NS + "p"):
                text = normalize_para_text(p)
                if text and JP_ANY.search(text):
                    all_strings.append(text)

    # Preserve first occurrence order
    return list(dict.fromkeys(all_strings))


def main() -> None:
    ap = argparse.ArgumentParser(description="Export unique JP segments from PPTX to a JSON template")
    ap.add_argument("--in", dest="inp", required=True, help="Input PPTX path")
    ap.add_argument("--out", dest="outp", required=True, help="Output template JSON path")
    args = ap.parse_args()

    inp = Path(args.inp)
    outp = Path(args.outp)
    outp.parent.mkdir(parents=True, exist_ok=True)

    segments = collect_unique_japanese_strings(inp)
    template = [{"jp": s, "en": ""} for s in segments]
    outp.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {len(template)} unique JP segments to {outp}")


if __name__ == "__main__":
    main()

