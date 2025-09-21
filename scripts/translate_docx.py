#!/usr/bin/env python3
"""
translate_docx.py

Japanese-to-English DOCX translator using the adapter pattern.
Integrates with existing batch translation, cache, and glossary systems.

Usage:
  python translate_docx.py --in input.docx --out output_en.docx \
    --model gpt-4o-2024-08-06 --batch 10
"""

import argparse
import asyncio
import json
import logging
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))  # For backend

# Import from existing translation system
from backend.document_adapter import Segment
from backend.translation_orchestrator import TranslationResult, orchestrator
from scripts.docx_adapter import DocxAdapter
from scripts.translate_pptx_inplace import (
    backup_existing_files,
)

# Import style modules if available
try:
    from scripts.style_checker import apply_style_fixes, model_style_check
    from scripts.style_normalize import get_style_guide, normalize_block
    STYLE_MODULES_AVAILABLE = True
except ImportError:
    print("Warning: Style modules not found. Running without style consistency features.")
    STYLE_MODULES_AVAILABLE = False

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def load_cache(cache_file):
    """Load translation cache from JSON file."""
    if cache_file and os.path.exists(cache_file):
        with open(cache_file, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache, cache_file):
    """Save translation cache to JSON file."""
    if cache_file:
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)


def load_glossary(glossary_file):
    """Load glossary from JSON file."""
    if glossary_file and os.path.exists(glossary_file):
        with open(glossary_file, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def setup_argument_parser() -> argparse.ArgumentParser:
    """Set up command line argument parser."""
    parser = argparse.ArgumentParser(
        description="Translate Japanese DOCX documents to English",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Production Presets:
  Conservative (rock-solid):  --model gpt-4o-2024-08-06 --batch 8
  Balanced (recommended):     --model gpt-4o-2024-08-06 --batch 12
  Cost-lean (good quality):   --model gpt-4o-mini --batch 16

Environment Variables:
  OPENAI_API_KEY           Required: OpenAI API key
  STYLE_PRESET             Style preset (default: "gengo")
  STYLE_GUIDE_FILE         Path to custom style guide
  GLOSSARY_FILE            Path to terminology glossary
        """
    )

    # Required arguments
    parser.add_argument('--in', dest='input_file', required=True,
                        help='Input DOCX file to translate')
    parser.add_argument('--out', dest='output_file', required=True,
                        help='Output path for translated DOCX')

    # Translation options
    parser.add_argument('--model', default='gpt-4o-2024-08-06',
                        help='OpenAI model to use (default: gpt-4o-2024-08-06)')
    parser.add_argument('--batch', type=int, default=12,
                        help='Batch size for translation (default: 12)')
    parser.add_argument('--temperature', type=float, default=0.6,
                        help='Translation temperature (default: 0.6)')
    parser.add_argument('--max-retries', type=int, default=3,
                        help='Maximum retry attempts (default: 3)')

    # I/O options
    parser.add_argument('--no-backup', action='store_true',
                        help='Skip creating backup files')
    parser.add_argument('--no-cache', action='store_true',
                        help='Disable translation cache')
    parser.add_argument('--cache-file',
                        help='Custom cache file path')
    parser.add_argument('--glossary-file',
                        help='Path to glossary JSON file')

    # Output options
    parser.add_argument('--bilingual-csv', action='store_true',
                        help='Export bilingual CSV for QA')
    parser.add_argument('--json-audit', action='store_true',
                        help='Export detailed JSON audit report')

    # Debug options
    parser.add_argument('--dry-run', action='store_true',
                        help='Extract text but don\'t translate')
    parser.add_argument('--verbose', '-v', action='store_true',
                        help='Verbose output')

    return parser


# Remove extract_segments_for_translation as orchestrator handles it


# Remove translate_segments as orchestrator handles it


# Remove update_segments_with_translations as not needed


def generate_bilingual_csv(
    segments: List[Segment],
    original_texts: List[str],
    translated_texts: List[str],
    output_path: str
):
    """Generate bilingual CSV for QA."""
    import csv

    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Segment ID', 'Original Japanese', 'Translated English', 'Context'])

        for segment, original, translated in zip(segments, original_texts, translated_texts):
            writer.writerow([
                segment.id,
                original,
                translated,
                segment.context or ''
            ])

    logger.info(f"Bilingual CSV saved to {output_path}")


def generate_audit_report(
    input_file: str,
    output_file: str,
    segments: List[Segment],
    metadata,  # DocumentMetadata from adapter
    processing_time: float,
    cache_stats: Dict[str, int],
    output_path: str
):
    """Generate detailed audit report."""
    # Normalise metadata to a dictionary regardless of source type.
    metadata_dict: Dict[str, Any] = {}
    if metadata:
        if isinstance(metadata, dict):
            metadata_dict = metadata
        else:
            try:
                from dataclasses import asdict

                metadata_dict = asdict(metadata)
            except TypeError:
                # Fallback to attribute inspection for plain objects
                metadata_dict = getattr(metadata, "__dict__", {}) or {}

    custom_properties = metadata_dict.get("custom_properties") or {}

    report = {
        "translation_info": {
            "input_file": input_file,
            "output_file": output_file,
            "timestamp": datetime.now().isoformat(),
            "processing_time_seconds": processing_time,
            "model": os.getenv("OPENAI_MODEL", "unknown")
        },
        "document_metadata": {
            "title": custom_properties.get("title"),
            "author": custom_properties.get("author"),
            "paragraph_count": custom_properties.get("paragraph_count")
            or metadata_dict.get("paragraph_count")
            or metadata_dict.get("segment_count"),
            "table_count": custom_properties.get("table_count")
            or metadata_dict.get("table_count", 0),
            "has_headers": metadata_dict.get("has_headers_footers", False),
            "has_footers": metadata_dict.get("has_headers_footers", False),
            "has_footnotes": metadata_dict.get("has_footnotes", False)
        },
        "translation_stats": {
            "total_segments": len(segments),
            "segments_translated": len(segments),
            "cache_hits": cache_stats.get("hits", 0),
            "cache_misses": cache_stats.get("misses", 0)
        }
    }

    segments_payload = []
    for seg in segments:
        seg_metadata = {}
        if isinstance(seg.metadata, dict):
            seg_metadata = seg.metadata
        elif seg.metadata:
            seg_metadata = getattr(seg.metadata, "__dict__", {}) or {}

        segments_payload.append({
            "id": seg.id,
            "file_path": seg.file_path,
            "paragraph_index": seg.paragraph_index,
            "run_index": seg.run_index,
            "context": seg.context,
            "metadata": {
                "bold": seg_metadata.get('bold', False),
                "italic": seg_metadata.get('italic', False),
                "underline": seg_metadata.get('underline', False),
                "color": seg_metadata.get('color'),
                "size": seg_metadata.get('size'),
                "font": seg_metadata.get('font')
            }
        })

    report["segments"] = segments_payload

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=2, ensure_ascii=False)

    logger.info(f"Audit report saved to {output_path}")


async def main():
    """Main entry point."""
    parser = setup_argument_parser()
    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    # Check OpenAI API key
    if not os.getenv("OPENAI_API_KEY"):
        print("ERROR: OPENAI_API_KEY environment variable is required", file=sys.stderr)
        sys.exit(1)

    # Validate input file
    input_path = Path(args.input_file)
    if not input_path.exists():
        print(f"ERROR: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if input_path.suffix.lower() != '.docx':
        print(f"ERROR: Input file must be a DOCX file: {input_path}", file=sys.stderr)
        sys.exit(1)

    # Prepare output paths
    output_path = Path(args.output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Set up cache and glossary
    cache = {} if args.no_cache else load_cache(args.cache_file)
    glossary = load_glossary(args.glossary_file) if args.glossary_file else None

    # Backup files if needed
    if not args.no_backup:
        cache_file = args.cache_file or f"{input_path.stem}_cache.json"
        bilingual_file = f"{output_path.stem}_bilingual.csv" if args.bilingual_csv else None
        audit_file = f"{output_path.stem}_audit.json" if args.json_audit else None
        log_file = f"{input_path.stem}_translation.log"
        backup_existing_files(cache_file, bilingual_file, audit_file, log_file)

    start_time = datetime.now()

    try:
        # Use orchestrator for full pipeline
        logger.info("Translating document using orchestrator...")
        result: TranslationResult = await orchestrator.translate_document(
            input_path=str(input_path),
            output_path=str(output_path),
            model=args.model,
            glossary=glossary,
            cache=cache,
            batch_size=args.batch,
            temperature=args.temperature
        )

        # Save cache if updated (orchestrator handles internal updates)
        if not args.no_cache:
            cache_file = args.cache_file or f"{input_path.stem}_cache.json"
            save_cache(cache, cache_file)

        # For bilingual and audit, extract from result or re-extract
        if args.bilingual_csv or args.json_audit:
            # Re-extract original for reporting
            adapter = DocxAdapter(str(input_path))
            original_segments = adapter.extract_segments(str(input_path))
            metadata = adapter.collect_metadata(str(input_path))
            # Re-extract from output for real translated texts
            translated_adapter = DocxAdapter(str(output_path))
            all_translated_segments = translated_adapter.extract_segments(str(output_path))
            # Match by ID for bilingual (only JP segments)
            trans_map = {s.id: s.text for s in all_translated_segments}
            original_texts = [s.text for s in original_segments if s.has_japanese]
            translated_texts = [trans_map.get(s.id, '[Unmatched]') for s in original_segments if s.has_japanese]

            if args.bilingual_csv:
                csv_path = f"{output_path.stem}_bilingual.csv"
                generate_bilingual_csv(original_segments, original_texts, translated_texts, csv_path)

        if args.json_audit:
            audit_path = f"{output_path.stem}_audit.json"
            processing_time = (datetime.now() - start_time).total_seconds()
            cache_stats = {"hits": result.cache_hits, "misses": len(original_segments) - result.cache_hits}
            generate_audit_report(
                str(input_path),
                str(output_path),
                original_segments,
                metadata,
                processing_time,
                cache_stats,
                audit_path
            )

        # Apply style fixes if available
        if STYLE_MODULES_AVAILABLE and output_path.exists():
            logger.info("Applying style fixes...")
            try:
                apply_style_fixes(str(output_path))
                logger.info("Style fixes applied successfully to output file")
            except Exception as e:
                logger.warning(f"Style fixes failed: {e}")

        processing_time = (datetime.now() - start_time).total_seconds()
        logger.info(f"Translation completed in {processing_time:.1f} seconds")
        logger.info(f"Output saved to: {output_path}")

    except Exception as e:
        logger.error(f"Translation failed: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    asyncio.run(main())
