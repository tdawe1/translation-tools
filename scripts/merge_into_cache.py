#!/usr/bin/env python3
"""
merge_into_cache.py

Merge a JSON mapping of updates {JP: EN} into the main translation cache.
Empty-string values are ignored. Existing entries are overwritten.

Usage:
  python scripts/merge_into_cache.py --updates translated.json \
    --cache translation_cache.json [--backup]
"""
import argparse, json, os, shutil


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--updates", required=True, help="JSON file with {JP: EN} updates")
    ap.add_argument("--cache", default="translation_cache.json", help="Target cache JSON")
    ap.add_argument("--backup", action="store_true", help="Create timestamped backup of the cache")
    args = ap.parse_args()

    if not os.path.exists(args.updates):
        raise SystemExit(f"Updates file not found: {args.updates}")

    updates = json.loads(open(args.updates, "r", encoding="utf-8").read())
    if not isinstance(updates, dict):
        raise SystemExit("Updates must be a JSON object {JP: EN}")

    cache = {}
    if os.path.exists(args.cache):
        cache = json.loads(open(args.cache, "r", encoding="utf-8").read())
        if not isinstance(cache, dict):
            cache = {}

    if args.backup and os.path.exists(args.cache):
        base, ext = os.path.splitext(args.cache)
        i = 1
        while True:
            candidate = f"{base}.bak{i}{ext}"
            if not os.path.exists(candidate):
                shutil.copyfile(args.cache, candidate)
                print(f"Backup: {candidate}")
                break
            i += 1

    changed = 0
    for k, v in updates.items():
        if isinstance(k, str) and isinstance(v, str) and v.strip() != "":
            if cache.get(k) != v:
                cache[k] = v
                changed += 1

    with open(args.cache, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

    print(f"Merged {changed} entries into {args.cache}")


if __name__ == "__main__":
    main()

