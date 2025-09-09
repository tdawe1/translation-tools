printf '%s\n' \
'#!/usr/bin/env python3
import argparse, zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

def scale_xml(xml_bytes: bytes, scale: float, min_pt_hundred=600):
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return xml_bytes  # non-XML or unexpected, leave as-is

    for el in root.iter():
        if el.tag.endswith("rPr") or el.tag.endswith("defRPr") or el.tag.endswith("endParaRPr"):
            sz = el.get("sz")
            if sz and sz.isdigit():
                v = int(sz)  # hundredths of a point (18pt => 1800)
                new_v = max(min_pt_hundred, int(round(v * scale)))
                if new_v != v:
                    el.set("sz", str(new_v))
    return ET.tostring(root, encoding="utf-8")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True)
    ap.add_argument("--out", dest="outp", required=True)
    ap.add_argument("--scale", type=float, required=True, help="e.g., 0.8 for 20%% shrink")
    ap.add_argument("--include-masters", action="store_true", help="also scale slideMasters/layouts/notes")
    args = ap.parse_args()

    targets = ["ppt/slides/"]
    if args.include_masters:
        targets += ["ppt/slideMasters/", "ppt/slideLayouts/", "ppt/notesSlides/"]

    with zipfile.ZipFile(args.inp, "r") as zin, zipfile.ZipFile(args.outp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.endswith(".xml") and any(info.filename.startswith(t) for t in targets):
                data = scale_xml(data, args.scale)
            zout.writestr(info, data)

if __name__ == "__main__":
    main()
' > scripts/scale_pptx_fonts.py
chmod +x scripts/scale_pptx_fonts.py
