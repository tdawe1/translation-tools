#!/usr/bin/env python3
"""Shim: PPTX JP character audit (package path)."""
from ..audit_pptx_jp_count import *  # type: ignore

def main():
    import sys
    from ..audit_pptx_jp_count import count_file  # type: ignore
    if len(sys.argv) < 2:
        print("Usage: python -m scripts.audit.pptx_jp_count file.pptx")
        return
    count_file(sys.argv[1])

if __name__ == "__main__":
    main()

