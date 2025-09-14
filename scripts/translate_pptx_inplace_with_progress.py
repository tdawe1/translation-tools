#!/usr/bin/env python3
"""
translate_pptx_inplace_with_progress.py

Enhanced version of translate_pptx_inplace.py with real-time progress tracking.
Emits progress updates via WebSocket during translation.
"""

import asyncio
import argparse, json, os, re, shutil, sys, time, zipfile, logging
from xml.etree import ElementTree as ET
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional

# Import the original translation functions
from translate_pptx_inplace import (
    get_timestamped_filename,
    backup_existing_files,
    batch_translate,
    process_pptx_file,
    count_jp_chars,
    JP_ANY,
    extract_jp_from_pptx,
    apply_translations_to_pptx,
    build_style_guide_text,
    apply_deck_formatting_profile,
    _slides_need_tightening,
    scale_slide_fonts,
    create_bilingual_csv,
    create_audit_report
)

# Import progress tracker
from progress_tracker import get_tracker, sync_update_progress

def calculate_estimated_tokens(text_blocks: List[str]) -> int:
    """Estimate total tokens needed for translation"""
    # Rough estimation: 1 token ≈ 4 characters for English
    # Japanese is more compact, so use 3 characters per token
    total_chars = sum(len(block) for block in text_blocks)
    return int(total_chars / 3)

async def translate_with_progress(args):
    """Main translation function with progress tracking"""
    # Generate job ID
    job_id = f"job_{int(time.time())}_{hash(args.input_file) % 10000}"

    # Get file info
    file_path = Path(args.input_file)
    file_size = file_path.stat().st_size
    file_name = file_path.name

    logger.info(f"Starting translation job {job_id} for {file_name}")

    # Start progress tracking
    await get_tracker().start_job(
        job_id=job_id,
        file_name=file_name,
        file_size=file_size,
        estimated_tokens=0,  # Will update after extraction
        estimated_cost=0.0    # Will update after extraction
    )

    try:
        # Step 1: Extract text
        await get_tracker().update_stage("extracting", "processing")
        logger.info("Extracting Japanese text...")

        jp_blocks = extract_jp_from_pptx(args.input_file)
        total_tokens = calculate_estimated_tokens(jp_blocks)

        # Update estimates
        await get_tracker().update_progress(
            total_tokens=total_tokens,
            estimated_cost=total_tokens * 0.00002  # Rough estimate
        )

        logger.info(f"Found {len(jp_blocks)} text blocks, ~{total_tokens} tokens")

        # Step 2: Translate
        await get_tracker().update_stage("translating", "processing")
        logger.info("Translating text...")

        # Load glossary if provided
        glossary = {}
        if args.glossary and os.path.exists(args.glossary):
            with open(args.glossary, 'r', encoding='utf-8') as f:
                glossary = json.load(f)

        # Initialize OpenAI client
        from openai import OpenAI
        client = OpenAI()

        # Process in batches
        batch_size = args.batch or 10
        translated_blocks = []
        total_blocks = len(jp_blocks)

        for i in range(0, total_blocks, batch_size):
            batch = jp_blocks[i:i + batch_size]
            batch_num = i // batch_size + 1
            total_batches = (total_blocks + batch_size - 1) // batch_size

            await get_tracker().update_batch_progress(batch_num, total_batches)

            logger.info(f"Translating batch {batch_num}/{total_batches}")

            if args.offline:
                # Use mock translations for offline mode
                translations = [f"[Mock translation: {block}]" for block in batch]
            else:
                translations = batch_translate(client, args.model, batch, glossary, args.offline)

            translated_blocks.extend(translations)

            # Update progress
            tokens_this_batch = sum(len(block) for block in batch)
            await get_tracker().update_tokens(
                tokens_processed=len(translated_blocks),
                cost_increment=tokens_this_batch * 0.00002
            )

        # Step 3: Apply translations
        await get_tracker().update_stage("applying", "processing")
        logger.info("Applying translations to PPTX...")

        # Backup existing files
        cache_file = f"{args.output_file}.cache.json"
        bilingual_csv = f"{args.output_file}.bilingual.csv"
        audit_json = f"{args.output_file}.audit.json"
        log_file = f"{args.output_file}.log"

        backup_existing_files(cache_file, bilingual_csv, audit_json, log_file)

        # Create translation mapping
        translation_map = {jp: en for jp, en in zip(jp_blocks, translated_blocks)}

        # Apply translations
        apply_translations_to_pptx(
            args.input_file,
            args.output_file,
            translation_map,
            cache_file,
            args.model,
            args.temperature,
            args.autofit,
            glossary
        )

        # Step 4: Post-processing
        await get_tracker().update_stage("finalizing", "processing")
        logger.info("Finalizing document...")

        # Apply deck formatting if enabled
        if args.formatting_profile and os.getenv("ENABLE_FORMATTING_PROFILE") == "1":
            logger.info("Applying deck formatting profile...")
            apply_deck_formatting_profile(args.output_file)

        # Scale fonts if needed
        if _slides_need_tightening and args.autofit != "none":
            logger.info("Scaling fonts for overflow prevention...")
            scale_slide_fonts(args.output_file, min_font_size=6)

        # Create outputs
        create_bilingual_csv(jp_blocks, translated_blocks, bilingual_csv)
        create_audit_report(args.output_file, audit_json)

        # Calculate quality score (simplified)
        quality_score = 0.95  # Default high score
        if any("error" in t.lower() for t in translated_blocks):
            quality_score = 0.7

        await get_tracker().set_quality_score(quality_score)

        # Complete job
        await get_tracker().complete_job(success=True)
        logger.info(f"Translation completed successfully: {args.output_file}")

    except Exception as e:
        logger.error(f"Translation failed: {e}")
        await get_tracker().complete_job(success=False, error_message=str(e))
        raise

async def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(description="Translate PPTX with real-time progress")
    parser.add_argument("--in", dest="input_file", required=True, help="Input PPTX file")
    parser.add_argument("--out", dest="output_file", required=True, help="Output PPTX file")
    parser.add_argument("--model", default="gpt-4o-2024-08-06", help="OpenAI model")
    parser.add_argument("--batch", type=int, help="Batch size for translation")
    parser.add_argument("--temperature", type=float, default=0.6, help="Temperature")
    parser.add_argument("--autofit", choices=["norm", "shape", "none"], default="norm", help="Autofit mode")
    parser.add_argument("--glossary", help="Glossary file path")
    parser.add_argument("--offline", action="store_true", help="Offline mode (no API calls)")
    args = parser.parse_args()

    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(f"{args.output_file}.log"),
            logging.StreamHandler(sys.stdout)
        ]
    )

    # Connect to WebSocket server
    await get_tracker().connect()

    try:
        await translate_with_progress(args)
    finally:
        await get_tracker().disconnect()

if __name__ == "__main__":
    asyncio.run(main())