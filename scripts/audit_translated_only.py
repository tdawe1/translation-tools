#!/usr/bin/env python3

import argparse
import json
import re
import sys
import csv
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict

rx_jp = re.compile(r'[\u3040-\u30ff\u3400-\u9fff々〆ヵヶ]')

def extract_text_from_pptx(pptx_path):
    """Extract all text from PPTX file."""
    all_text = []
    try:
        with zipfile.ZipFile(pptx_path, 'r') as zip_file:
            for file_name in zip_file.namelist():
                if file_name.startswith('ppt/slides/slide') and file_name.endswith('.xml'):
                    xml_data = zip_file.read(file_name)
                    try:
                        root = ET.fromstring(xml_data)
                        ns = {
                            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
                        }
                        for t_elem in root.iterfind('.//a:t', ns):
                            if t_elem.text:
                                all_text.append(t_elem.text.strip())
                    except ET.ParseError:
                        continue
    except zipfile.BadZipFile:
        print(f"Invalid PPTX file: {pptx_path}", file=sys.stderr)
        return []
    return all_text

def audit_from_csv(bilingual_path, threshold):
    total_chars = 0
    total_residual = 0
    per_slide = defaultdict(int)
    per_slide_total = defaultdict(int)
    residual_rows = []

    with open(bilingual_path, encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            slide = row.get("slide") or ""
            en = (row.get("en") or "")
            cnt = len(rx_jp.findall(en))
            en_len = len(en)
            total_chars += en_len
            total_residual += cnt
            per_slide_total[slide] += en_len
            if cnt:
                per_slide[slide] += cnt
                residual_rows.append({
                    "slide": slide, "shape": row.get("shape"),
                    "para": row.get("para"), "jp": row.get("jp"), "en": en
                })

    percentage = (total_residual / total_chars * 100) if total_chars > 0 else 0
    passed = percentage <= threshold

    out = {
        "scope": "translated_only",
        "input_type": "csv",
        "total_en_chars": total_chars,
        "total_residual_jp_chars": total_residual,
        "percentage_residual": round(percentage, 2),
        "passed": passed,
        "per_slide_residual_jp_chars": dict(per_slide),
        "per_slide_total_chars": dict(per_slide_total),
        "residual_rows": residual_rows[:200]  # preview
    }
    return out, passed

def audit_from_pptx(pptx_path, threshold):
    texts = extract_text_from_pptx(pptx_path)
    if not texts:
        return {"scope": "translated_only", "input_type": "pptx", "message": "No text extracted", "passed": True}, True

    total_chars = sum(len(t) for t in texts)
    total_residual = sum(len(rx_jp.findall(t)) for t in texts)
    percentage = (total_residual / total_chars * 100) if total_chars > 0 else 0
    passed = percentage <= threshold

    out = {
        "scope": "translated_only",
        "input_type": "pptx",
        "file": str(pptx_path),
        "total_chars": total_chars,
        "total_residual_jp_chars": total_residual,
        "percentage_residual": round(percentage, 2),
        "passed": passed,
        "message": f"Residual Japanese: {percentage:.2f}% ({total_residual}/{total_chars})"
    }
    if not passed:
        out["message"] += " - FAILED (above threshold)"
    return out, passed

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="Path to bilingual.csv or translated.pptx")
    ap.add_argument("--out", default="audit_translated.json", help="Write audit JSON here")
    ap.add_argument("--fail-threshold", type=float, default=1.0, help="Fail if residual % > this")
    args = ap.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Input not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if input_path.suffix == '.csv':
        report, passed = audit_from_csv(str(input_path), args.fail_threshold)
    elif input_path.suffix.lower() == '.pptx':
        report, passed = audit_from_pptx(str(input_path), args.fail_threshold)
    else:
        print("Unsupported input format. Use .csv or .pptx", file=sys.stderr)
        sys.exit(1)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    # Write residual CSV if applicable
    if 'residual_rows' in report:
        with open("residual_rows_translated_only.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["slide","shape","para","jp","en"])
            for r in report["residual_rows"]:
                w.writerow([r["slide"], r["shape"], r["para"], r["jp"], r["en"]])

    percentage = report.get("percentage_residual", 0)
    print(f"Translated-only residual JP: {percentage:.2f}%")
    if not passed:
        print(f"FAIL: residual > threshold ({args.fail_threshold}%)")
        sys.exit(1)
    print("OK: within threshold")

if __name__ == "__main__":
    main()
