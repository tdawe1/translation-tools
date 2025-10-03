#!/usr/bin/env python3
import csv, json, re, argparse, sys
rx_jp = re.compile(r'[\u3040-\u30ff\u3400-\u9fff々〆ヵヶ]')

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--bilingual", default="bilingual.csv", help="Pipeline bilingual export (source JP + translated EN)")
    ap.add_argument("--out", default="audit_translated.json", help="Write translated-only audit JSON here")
    ap.add_argument("--fail-threshold", type=int, default=0, help="Exit non-zero if residual JP chars > this")
    args = ap.parse_args()

    total = 0
    per_slide = {}
    residual_rows = []

    with open(args.bilingual, encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            slide = row.get("slide") or ""
            en = (row.get("en") or "")
            cnt = len(rx_jp.findall(en))
            if cnt:
                total += cnt
                per_slide[slide] = per_slide.get(slide, 0) + cnt
                residual_rows.append({
                    "slide": slide, "shape": row.get("shape"),
                    "para": row.get("para"), "jp": row.get("jp"), "en": en
                })

    out = {
        "scope": "translated_only",
        "total_residual_jp_chars": total,
        "per_slide_residual_jp_chars": per_slide,
        "residual_rows": residual_rows[:200]  # preview (full detail lives in CSV below)
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    # Also write a flat CSV you can open quickly
    import csv as _csv
    with open("residual_rows_translated_only.csv", "w", newline="", encoding="utf-8") as f:
        w = _csv.writer(f)
        w.writerow(["slide","shape","para","jp","en"])
        for r in residual_rows:
            w.writerow([r["slide"], r["shape"], r["para"], r["jp"], r["en"]])

    print(f"Translated-only residual JP chars: {total}")
    if total > args.fail_threshold:
        print(f"FAIL: residual > threshold ({args.fail_threshold})")
        sys.exit(1)
    print("OK: within threshold")

if __name__ == "__main__":
    main()
