#!/usr/bin/env python3
"""
export_review_text.py

Generate reviewer-friendly text files from bilingual.csv.
Produces:
- review_en_only.txt      (English-only, slide by slide)
- review_bilingual.txt    (JP -> EN pairs grouped by slide)

Usage:
  python scripts/export_review_text.py --bilingual bilingual.csv [--out-dir outputs]
"""
import argparse, csv, os, re
from pathlib import Path

JP_CORE = r"\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff"
JP_ANY = re.compile(f"[{JP_CORE}]")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--bilingual", default="bilingual.csv", help="Path to bilingual CSV")
    ap.add_argument("--out-dir", default="outputs", help="Directory to write review files")
    ap.add_argument("--include-nonjp", action="store_true", help="Include rows with no JP chars (default: only JP->EN rows)")
    args = ap.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    rows = []
    with open(args.bilingual, "r", encoding="utf-8") as f:
        rdr = csv.DictReader(f)
        for row in rdr:
            slide_xml = row.get("slide_xml") or ""
            idx = row.get("paragraph_idx") or "0"
            jp = row.get("Japanese") or ""
            en = row.get("English") or ""
            m = re.search(r"slide(\d+)\.xml", slide_xml)
            slide_no = int(m.group(1)) if m else 0
            try:
                pidx = int(idx)
            except Exception:
                pidx = 0
            if not args.include_nonjp and not JP_ANY.search(jp or ""):
                continue
            rows.append((slide_no, pidx, jp, en))

    rows.sort(key=lambda r: (r[0], r[1]))

    # English-only file
    en_path = out_dir / "review_en_only.txt"
    with en_path.open("w", encoding="utf-8") as f:
        cur = None
        for slide_no, pidx, jp, en in rows:
            if slide_no != cur:
                cur = slide_no
                f.write(f"\n=== Slide {slide_no} ===\n")
            f.write(en.strip() + "\n")

    # Bilingual file
    bi_path = out_dir / "review_bilingual.txt"
    with bi_path.open("w", encoding="utf-8") as f:
        cur = None
        for slide_no, pidx, jp, en in rows:
            if slide_no != cur:
                cur = slide_no
                f.write(f"\n=== Slide {slide_no} ===\n")
            f.write("JP: " + jp.strip() + "\n")
            f.write("EN: " + en.strip() + "\n\n")

    print("Wrote:", en_path)
    print("Wrote:", bi_path)

if __name__ == "__main__":
    main()

