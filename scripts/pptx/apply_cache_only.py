#!/usr/bin/env python3
"""Shim: apply cache to PPTX without API (package path)."""
from ..apply_cache_only import main as _main  # type: ignore

def main():
    _main()

if __name__ == "__main__":
    main()

