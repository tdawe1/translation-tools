#!/usr/bin/env python3
"""
Scrub translation cache to remove JP→JP identity mappings.
These occur when Japanese text wasn't translated and remains identical.
"""
import json
import re
import sys
import os

JP_CORE = r'\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff'
CJK_PUNCT = r'\u3000-\u303f'
FULLWIDTH = r'\uff00-\uffef'
JP_ANY = re.compile(f'[{JP_CORE}{CJK_PUNCT}{FULLWIDTH}]')

def has_japanese(text):
    """Check if text contains Japanese characters."""
    return bool(JP_ANY.search(str(text)))

def main():
    cache_file = sys.argv[1] if len(sys.argv) > 1 else "../translation_cache.json"
    
    if not os.path.exists(cache_file):
        print(f"Cache file not found: {cache_file}")
        sys.exit(1)
    
    # Load cache
    with open(cache_file, "r", encoding="utf-8") as f:
        cache = json.load(f)
    
    print(f"Original cache entries: {len(cache)}")
    
    # Find problematic entries
    to_remove = []
    for key, value in cache.items():
        # Remove if key == value (identity mapping)
        if key == value:
            to_remove.append(key)
        # Remove if value still contains Japanese (failed translation)  
        elif has_japanese(value):
            to_remove.append(key)
    
    print(f"Found {len(to_remove)} problematic entries:")
    for key in to_remove[:10]:  # Show first 10
        print(f"  '{key}' -> '{cache[key]}'")
    if len(to_remove) > 10:
        print(f"  ... and {len(to_remove) - 10} more")
    
    # Remove problematic entries
    for key in to_remove:
        del cache[key]
    
    print(f"Cleaned cache entries: {len(cache)}")
    
    # Create backup
    backup_file = cache_file + ".backup"
    if os.path.exists(backup_file):
        os.remove(backup_file)
    os.rename(cache_file, backup_file)
    print(f"Backed up original to: {backup_file}")
    
    # Write cleaned cache
    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    
    print(f"Wrote cleaned cache to: {cache_file}")
    print(f"Removed {len(to_remove)} entries that need re-translation")

if __name__ == "__main__":
    main()