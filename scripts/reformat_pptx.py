#!/usr/bin/env python3
"""
reformat_pptx.py

Apply the deck formatting profile to all slides in a PPTX without changing text.
Useful for fixing footer sizes (e.g., "Confidential" wrapping) after translation.

Usage:
  python scripts/reformat_pptx.py --in input.pptx --out output.pptx
"""
import argparse, shutil, zipfile
from xml.etree import ElementTree as ET

from pptx_format import apply_deck_formatting_profile

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True, help="Input PPTX path")
    ap.add_argument("--out", dest="outp", required=True, help="Output PPTX path")
    args = ap.parse_args()

    tmp = args.outp + ".tmp"
    with zipfile.ZipFile(args.inp, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name.startswith("ppt/slides/") and name.endswith(".xml"):
                try:
                    root = ET.fromstring(data)
                    apply_deck_formatting_profile(root)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                except Exception:
                    # On parse errors, just copy original
                    pass
            zout.writestr(name, data)
    shutil.move(tmp, args.outp)
    print("Wrote:", args.outp)

if __name__ == "__main__":
    main()

