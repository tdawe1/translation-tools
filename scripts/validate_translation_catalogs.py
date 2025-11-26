#!/usr/bin/env python3
"""Validate and normalize translation catalog JSON files."""
import json
import unicodedata
import sys
from pathlib import Path


def validate_catalog(path: Path) -> list[str]:
    """Validate a translation catalog and return issues."""
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
        if any(placeholder in key for placeholder in ["XXXXX", "□□□□", "00:00"]):
            issues.append(f"Placeholder detected in key: {key[:50]}...")

        # Check for Markdown artifacts
        if isinstance(value, dict) and "translated" in value:
            trans = value["translated"]
            if "**" in trans or "*" in trans:
                issues.append(f"Markdown artifact in: {key[:30]}... → {trans[:30]}...")

        # Check normalization
        normalized_key = unicodedata.normalize("NFKC", key)
        if normalized_key != key:
            issues.append(f"Key not NFKC normalized: {key[:50]}...")

    return issues


def main():
    """Validate all translation catalogs."""
    catalog_files = [
        "translations_full_codex.json",
        "translations_full_codex_cheap.json",
        "translations_full_codex_max.json",
        "translations_full_lm_jit.json",
        "temp_translations_debug.json",
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

    if all_issues:
        print(f"\n❌ Total issues found: {len(all_issues)}")
        sys.exit(1)

    print("\n✅ All catalogs valid")


if __name__ == "__main__":
    main()
