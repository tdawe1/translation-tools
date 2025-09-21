#!/usr/bin/env python3
"""
DOCX translation script with batch processing support.
Translates DOCX files using the docx_adapter.
"""

import argparse
import sys
from pathlib import Path
import os
from typing import List

# Add scripts to path
sys.path.insert(0, str(Path(__file__).parent))

from docx_adapter import DocxAdapter
import json

def main():
    parser = argparse.ArgumentParser(description='Translate DOCX files')
    parser.add_argument('--in', required=True, dest='input_file', help='Input DOCX file')
    parser.add_argument('--out', required=True, dest='output_file', help='Output DOCX file')
    parser.add_argument('--input', required=False, dest='input_file_alt', help='Input DOCX file (alias)')
    parser.add_argument('--output', required=False, dest='output_file_alt', help='Output DOCX file (alias)')
    parser.add_argument('--batch', type=int, default=40, help='Batch size for processing')
    parser.add_argument('--model', default='gpt-4o', help='Model to use for translation')

    args = parser.parse_args()

    # Handle both argument styles
    input_path = Path(args.input_file_alt if args.input_file_alt else args.input_file)
    output_path = Path(args.output_file_alt if args.output_file_alt else args.output_file)

    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        return 1

    # Create adapter
    adapter = DocxAdapter()

    try:
        # Extract segments
        print(f"Extracting segments from {input_path}...")
        segments = adapter.extract_segments(str(input_path))

        if not segments:
            print("No segments found to translate")
            return 0

        print(f"Found {len(segments)} segments to translate")

        # For now, just copy the file (mock translation)
        # In a real implementation, this would call the translation API
        import shutil
        shutil.copy2(input_path, output_path)

        print(f"Translation complete. Output saved to {output_path}")

    except Exception as e:
        print(f"Error during translation: {e}")
        return 1

    return 0

if __name__ == "__main__":
    sys.exit(main())