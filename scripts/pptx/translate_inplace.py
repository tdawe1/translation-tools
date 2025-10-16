#!/usr/bin/env python3
"""Shim: PPTX in-place translator (package path).

Exports main/translate_batch from the original CLI to provide a stable
import path `scripts.pptx.translate_inplace` during the restructure.
"""
import sys
from pathlib import Path

# Ensure legacy script module imports (e.g., 'style_normalize') resolve
_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from ..translate_pptx_inplace import main as _main, translate_batch  # type: ignore

def main():
    _main()

if __name__ == "__main__":
    main()
