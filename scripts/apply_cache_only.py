#!/usr/bin/env python3
"""
apply_cache_only.py

Applies translations from a cache (JSON mapping {JP: EN}) to a PPTX without
calling any API. It replaces paragraph texts exactly matching cache keys while
preserving runs and layout. Useful for residual sweeps and offline fixes.

Usage:
  python scripts/apply_cache_only.py --in input.pptx --out output.pptx \
    --cache translation_cache.json
"""
import argparse, json, os, zipfile
from xml.etree import ElementTree as ET

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

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

def set_para_text(p_el, new_text: str):
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

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--in', dest='inp', required=True)
    ap.add_argument('--out', dest='outp', required=True)
    ap.add_argument('--cache', dest='cache', default='translation_cache.json')
    args = ap.parse_args()

    # Load cache
    cache = {}
    if os.path.exists(args.cache):
        with open(args.cache, 'r', encoding='utf-8') as f:
            cache = json.load(f)

    # Process PPTX
    tmp = args.outp + '.tmp'
    with zipfile.ZipFile(args.inp, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name.startswith('ppt/slides/slide') and name.endswith('.xml'):
                root = ET.fromstring(data)
                changed = False
                for p in root.iter(A_NS + 'p'):
                    src = normalize_para_text(p)
                    if src and src in cache:
                        set_para_text(p, cache[src])
                        changed = True
                if changed:
                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            zout.writestr(name, data)

    os.replace(tmp, args.outp)
    print('Wrote', args.outp)

if __name__ == '__main__':
    main()

