#!/usr/bin/env python3
"""
audit_pptx_jp_count.py
Counts Japanese characters (Kanji/Kana + CJK punct + Fullwidth forms) in a PPTX.
Usage:
  python audit_pptx_jp_count.py file.pptx
"""
import sys, re, zipfile
from xml.etree import ElementTree as ET

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
JP_CORE = r'\u3040-\u309F\u30A0-\u30FF\u31F0-\u31FF\u3400-\u4DBF\u4E00-\u9FFF'
CJK_PUNCT = r'\u3000-\u303F'
FULLWIDTH = r'\uFF00-\uFFEF'
JP_ANY = re.compile(f'[{JP_CORE}{CJK_PUNCT}{FULLWIDTH}]')

def count_file(path):
    total = 0
    per_slide = {}
    with zipfile.ZipFile(path, 'r') as z:
        slides = sorted([n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")])
        for sf in slides:
            root = ET.fromstring(z.read(sf))
            s = ""
            for t in root.iter(A_NS + "t"):
                if t.text:
                    s += t.text
            cnt = len(JP_ANY.findall(s))
            per_slide[sf] = cnt
            total += cnt
    print("Total JP-like chars:", total)
    for k, v in sorted(per_slide.items()):
        print(f"{k}: {v}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python audit_pptx_jp_count.py file.pptx")
        sys.exit(2)
    count_file(sys.argv[1])
