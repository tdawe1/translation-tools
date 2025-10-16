#!/usr/bin/env python3
"""Shim: manual/local DOCX translation CLI (package path)."""
from ..manual_docx_translation import main as _main  # type: ignore

def main():
    _main()

if __name__ == "__main__":
    main()

