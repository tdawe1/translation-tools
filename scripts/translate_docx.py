#!/usr/bin/env python3
"""
DOCX translation script.
"""

import argparse
import asyncio
import sys
from pathlib import Path
import tempfile
import shutil

# Add backend to path
sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from translation_orchestrator import orchestrator


async def main_async():
    """Async main entry point."""
    parser = argparse.ArgumentParser(description="Translate DOCX documents")

    # Support both --in/--out and --input/--output
    parser.add_argument('--in', required=True, dest='input_file', help='Input DOCX file')
    parser.add_argument('--input', required=False, dest='input_file_alt', help='Input DOCX file (alias)')
    parser.add_argument('--out', required=True, dest='output_file', help='Output DOCX file')
    parser.add_argument('--output', required=False, dest='output_file_alt', help='Output DOCX file (alias)')
    parser.add_argument('--model', default='gpt-4', help='Translation model')
    parser.add_argument('--source-lang', default='auto', help='Source language')
    parser.add_argument('--target-lang', default='en', help='Target language')
    parser.add_argument('--glossary-id', help='Glossary ID')
    parser.add_argument('--batch-size', type=int, default=1, help='Batch size')
    parser.add_argument('--no-backup', action='store_true', help='Do not create backup')
    parser.add_argument('--no-cache', action='store_true', help='Do not use cache')
    parser.add_argument('--bilingual-csv', action='store_true', help='Generate bilingual CSV')
    parser.add_argument('--json-audit', action='store_true', help='Generate JSON audit')

    args = parser.parse_args()

    # Handle both argument styles
    input_file = args.input_file or args.input_file_alt
    output_file = args.output_file or args.output_file_alt

    if not input_file or not output_file:
        parser.error("Both input and output files are required")

    input_path = Path(input_file)
    output_path = Path(output_file)

    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    # Run translation
    try:
        result = await orchestrator.translate_document(
            input_path=input_path,
            output_path=output_path,
            model=args.model,
            source_lang=args.source_lang,
            target_lang=args.target_lang,
            glossary_id=args.glossary_id,
            batch_size=args.batch_size,
            backup=not args.no_backup,
            cache=not args.no_cache,
            bilingual_csv=args.bilingual_csv,
            json_audit=args.json_audit,
            no_backup=args.no_backup,
            no_cache=args.no_cache,
        )

        print(f"Translation completed successfully")
        print(f"Segments translated: {result.segments_translated}/{result.total_segments}")
        print(f"Words translated: {result.words_translated}/{result.total_words}")
        print(f"Processing time: {result.processing_time:.2f}s")

        if result.warnings:
            print("Warnings:")
            for warning in result.warnings:
                print(f"  - {warning}")

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    """Main entry point."""
    asyncio.run(main_async())


if __name__ == "__main__":
    main()