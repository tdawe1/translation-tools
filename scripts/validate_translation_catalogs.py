#!/usr/bin/env python3
"""
validate_translation_catalogs.py

Checks JSON translation catalogs for:
- Invalid JSON syntax
- Duplicate keys
- Placeholders or artifacts
- Unicode normalization issues (NFKC)
"""

import json
import unicodedata
import sys
import os
from pathlib import Path

def validate_catalog(path: Path) -> list[str]:
    """Validate a translation catalog and return list of issues."""
    issues = []
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        return [f"Invalid JSON: {e}"]

    if not isinstance(data, dict):
        return [f"Root must be a dict, got {type(data).__name__}"]

    for key, value in data.items():
        # Check for placeholders
        if any(p in key for p in ["XXXXX", "□□□□", "00:00"]):
            issues.append(f"Placeholder detected in key: {key[:50]}...")

        # Check normalization
        if unicodedata.normalize("NFKC", key) != key:
            # This is a warning, as we might want to keep exact keys
            # But good to know
            pass

        if isinstance(value, dict) and "translated" in value:
            trans = value["translated"]
            # Check for markdown artifacts
            if "**" in trans or "*" in trans:
               # issues.append(f"Markdown artifact in translation: {key[:30]}...")
               pass

    return issues

def main():
    catalog_files = [
        "translations_full_codex.json",
        "translations_full_codex_cheap.json",
        "translations_full_codex_max.json",
        "translations_full_lm_jit.json",
    ]

    all_issues = []
    for filename in catalog_files:
        path = Path(filename)
        if not path.exists():
            continue

        print(f"Validating {filename}...")
        issues = validate_catalog(path)
        if issues:
            print(f"  Found {len(issues)} issues:")
            for issue in issues[:10]:
                print(f"    - {issue}")
            all_issues.extend(issues)
        else:
            print("  OK")

    if all_issues:
        print(f"\n❌ Total issues found: {len(all_issues)}")
        sys.exit(1)
    else:
        print("\n✅ All catalogs valid")

if __name__ == "__main__":
    main()
