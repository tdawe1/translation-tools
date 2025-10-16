#!/usr/bin/env python3
"""
translations_json_to_cache.py

Convert a filled template (array of {"jp": ..., "en": ...}) into a cache
JSON mapping {jp: en} suitable for scripts/apply_cache_only.py.

Usage:
  python scripts/translations_json_to_cache.py \
    --in translations/source_translations.json \
    --out outputs/source_cache.json
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path


def main() -> None:
    ap = argparse.ArgumentParser(description="Convert translations JSON (jp/en pairs) to cache mapping {jp: en}")
    ap.add_argument("--in", dest="inp", required=True, help="Input translations JSON (array of {jp,en})")
    ap.add_argument("--out", dest="outp", required=True, help="Output cache JSON {jp: en}")
    args = ap.parse_args()

    inp = Path(args.inp)
    outp = Path(args.outp)
    outp.parent.mkdir(parents=True, exist_ok=True)

    data = json.loads(inp.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise SystemExit("Input must be a JSON array of objects with 'jp' and 'en' fields")

    mapping: dict[str, str] = {}
    missing = []
    for row in data:
        if not isinstance(row, dict):
            continue
        # Preserve jp exactly (including newlines/spaces) to match PPTX text
        jp = (row.get("jp") or "")
        en = (row.get("en") or "").strip()
        if not jp.strip():
            continue
        if not en:
            missing.append(jp)
            continue
        mapping[jp] = en

    if missing:
        print(f"WARNING: {len(missing)} entries missing 'en' – they will be skipped")

    outp.write_text(json.dumps(mapping, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote cache with {len(mapping)} entries to {outp}")


if __name__ == "__main__":
    main()
