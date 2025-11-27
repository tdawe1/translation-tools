#!/usr/bin/env python3
"""
translate_pdf.py

Main PDF orchestrator for Japanese-to-English translation pipeline.
Integrates extraction, translation, layout optimization, back-projection, and audit components
while reusing the existing PPTX translation cache and glossary systems.

Usage:
  python translate_pdf.py --in input.pdf --out output_en.pdf --model gpt-4o
  
Production Presets:
  Conservative (rock-solid):  --model gpt-4o-2024-08-06
  Balanced (recommended):     --model gpt-4o-2024-08-06  
  Cost-lean (good quality):   --model gpt-4o-mini

Features:
- Complete end-to-end PDF translation pipeline
- Unified cache system with PPTX pipeline (translation_cache.json)
- Shared glossary support
- CLI parity with PPTX translator
- Layout optimization for text expansion
- Comprehensive audit reporting
- Bilingual CSV output
"""

import argparse
import json
import logging
import os
import sys
import csv
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed

# Import PDF translation components
try:
    from extract_pdf import PDFExtractor, ExtractionResult
    from pdf_layout_engine import PDFLayoutEngine, TextBlock, ContentType, LayoutConstraint
    from apply_pdf_translation import PDFBackProjector
    from audit_pdf import PDFAuditor
    PDF_COMPONENTS_AVAILABLE = True
except ImportError as e:
    print(f"ERROR: PDF translation components not found: {e}")
    print("Make sure all PDF translation scripts are in the scripts/ directory")
    PDF_COMPONENTS_AVAILABLE = False

# Import PPTX translation system components
try:
    from translate_pptx_inplace import (
        backup_existing_files, 
        get_timestamped_filename,
        batch_translate,
        get_llm_client
    )
    PPTX_SYSTEM_AVAILABLE = True
except ImportError:
    print("WARNING: PPTX translation system not available - some features disabled")
    PPTX_SYSTEM_AVAILABLE = False

# OpenAI client (if available)
try:
    from openai import OpenAI, AsyncOpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    print("WARNING: OpenAI package not available - translation will be disabled")
    OPENAI_AVAILABLE = False

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pdf_translation.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

class PDFTranslationOrchestrator:
    """Main orchestrator for PDF translation pipeline."""
    
    def __init__(self, 
                 input_path: str,
                 output_path: str,
                 model: str = "gpt-4o-2024-08-06",
                 cache_file: str = "translation_cache.json",
                 glossary_file: Optional[str] = None,
                 batch_size: int = 10,
                 temperature: float = 0.6,
                 pages: Optional[str] = None,
                 offline: bool = False,
                 cache_only: bool = False,
                 verbose: bool = False,
                 concurrency: int = 1,
                 do_audit: bool = True,
                 do_csv: bool = True):
        """
        Initialize PDF translation orchestrator.
        
        Args:
            input_path: Path to input PDF file
            output_path: Path to output PDF file  
            model: OpenAI model to use for translation
            cache_file: Path to translation cache file
            glossary_file: Optional path to glossary file
            batch_size: Batch size for API calls
            temperature: Temperature for translation
            pages: Page range to process (e.g., "1-10")
            offline: Run in offline mode
            cache_only: Use only cache, no API calls
            verbose: Enable verbose logging
            concurrency: Number of concurrent API requests
            do_audit: Generate audit report after translation
            do_csv: Generate bilingual CSV after translation
        """
        self.input_path = input_path
        self.output_path = output_path
        self.model = model
        self.cache_file = cache_file
        self.glossary_file = glossary_file
        self.batch_size = batch_size
        self.temperature = temperature
        self.pages = pages
        self.offline = offline
        self.cache_only = cache_only
        self.verbose = verbose
        self.concurrency = concurrency
        self.do_audit = do_audit
        self.do_csv = do_csv
        
        # Initialize components
        self.extractor = PDFExtractor() if PDF_COMPONENTS_AVAILABLE else None
        self.layout_engine = PDFLayoutEngine() if PDF_COMPONENTS_AVAILABLE else None
        self.auditor = PDFAuditor() if PDF_COMPONENTS_AVAILABLE else None
        
        # Statistics
        self.stats = {
            'total_blocks': 0,
            'translated_blocks': 0,
            'layout_optimizations': 0,
            'cache_hits': 0,
            'api_calls': 0,
            'start_time': None,
            'end_time': None
        }
        
        # Load cache and glossary
        self.cache = self._load_cache()
        self.glossary = self._load_glossary()
        
        if verbose:
            logging.getLogger().setLevel(logging.DEBUG)
        
        logger.info(f"Initialized PDF translation orchestrator")
        logger.info(f"  Input: {input_path}")
        logger.info(f"  Output: {output_path}")
        logger.info(f"  Model: {model}")
        logger.info(f"  Cache: {cache_file} ({len(self.cache)} entries)")
        logger.info(f"  Glossary: {glossary_file if glossary_file else 'None'}")
        logger.info(f"  Concurrency: {concurrency}")
    
    def _load_cache(self) -> Dict[str, str]:
        """Load translation cache from file."""
        cache = {}
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, "r", encoding="utf-8") as f:
                    cache = json.load(f)
                logger.info(f"Loaded {len(cache)} entries from cache")
            except Exception as e:
                logger.warning(f"Failed to load cache: {e}")
        return cache
    
    def _load_glossary(self) -> Dict[str, str]:
        """Load glossary from file."""
        glossary = {}
        if self.glossary_file and os.path.exists(self.glossary_file):
            try:
                with open(self.glossary_file, "r", encoding="utf-8") as f:
                    glossary_data = json.load(f)
                    # Handle both list and dict formats
                    if isinstance(glossary_data, list):
                        for item in glossary_data:
                            if isinstance(item, dict) and 'original' in item and 'translated' in item:
                                glossary[item['original']] = item['translated']
                    elif isinstance(glossary_data, dict):
                        glossary = glossary_data
                logger.info(f"Loaded {len(glossary)} entries from glossary")
            except Exception as e:
                logger.warning(f"Failed to load glossary: {e}")
        return glossary
    
    def _save_cache(self) -> None:
        """Save translation cache to file."""
        try:
            with open(self.cache_file, "w", encoding="utf-8") as f:
                json.dump(self.cache, f, ensure_ascii=False, indent=2)
            logger.debug(f"Saved {len(self.cache)} entries to cache")
        except Exception as e:
            logger.error(f"Failed to save cache: {e}")
    
    def _parse_page_range(self, pages_str: str) -> Tuple[int, int]:
        """Parse page range string like '1-10' or '5'."""
        if not pages_str:
            return 0, -1  # All pages
        
        if '-' in pages_str:
            start, end = pages_str.split('-', 1)
            return int(start.strip()), int(end.strip())
        else:
            page_num = int(pages_str.strip())
            return page_num, page_num
    
    def _filter_pages_by_range(self, extraction_result) -> 'ExtractionResult':
        """Filter extraction result to include only specified pages."""
        if not self.pages:
            return extraction_result
        
        start_page, end_page = self._parse_page_range(self.pages)
        
        # Convert to 0-based indexing
        start_idx = start_page - 1
        end_idx = end_page - 1 if end_page > 0 else -1
        
        filtered_pages = []
        for page in extraction_result.pages:
            if (page.page_num >= start_idx) and (end_idx == -1 or page.page_num <= end_idx):
                filtered_pages.append(page)
        
        # Recalculate statistics
        total_blocks = sum(len(p.text_blocks) for p in filtered_pages)
        japanese_blocks = sum(len([b for b in p.text_blocks if self._contains_japanese(b.text)]) for p in filtered_pages)
        
        return ExtractionResult(
            filename=extraction_result.filename,
            pages=filtered_pages,
            total_blocks=total_blocks,
            total_japanese_blocks=japanese_blocks,
            extraction_time=extraction_result.extraction_time,
            extraction_methods=extraction_result.extraction_methods,
            metadata=extraction_result.metadata
        )
    
    def _contains_japanese(self, text: str) -> bool:
        """Check if text contains Japanese characters."""
        # Using the same pattern as PPTX system
        jp_pattern = re.compile(r'[\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]')
        return bool(jp_pattern.search(text))
    
    def _extract_unique_japanese_text(self, extraction_result) -> List[str]:
        """Extract unique Japanese text strings for translation."""
        japanese_texts = []
        
        for page in extraction_result.pages:
            for block in page.text_blocks:
                if self._contains_japanese(block.text):
                    japanese_texts.append(block.text)
        
        # Remove duplicates while preserving order
        return list(dict.fromkeys(japanese_texts))
    
    def _translate_with_cache(self, text_list: List[str]) -> Dict[str, str]:
        """Translate text list using cache and API calls with concurrency."""
        if not text_list:
            return {}
        
        # Find uncached texts
        uncached = []
        translations = {}
        
        for text in text_list:
            if text in self.cache:
                translations[text] = self.cache[text]
                self.stats['cache_hits'] += 1
                logger.debug(f"Cache hit: '{text[:50]}...'")
            else:
                uncached.append(text)
        
        if uncached and not self.offline and not self.cache_only:
            if not OPENAI_AVAILABLE or not PPTX_SYSTEM_AVAILABLE:
                logger.error("Translation requires OpenAI client and PPTX system")
                return {}
            
            try:
                # Initialize LLM client (OpenAI or Gemini)
                client = get_llm_client(self.model)
                
                logger.info(f"Translating {len(uncached)} uncached items with {self.model} (concurrency={self.concurrency})")
                
                total_batches = (len(uncached) + self.batch_size - 1) // self.batch_size
                future_to_batch = {}
                
                with ThreadPoolExecutor(max_workers=self.concurrency) as executor:
                    for i in range(0, len(uncached), self.batch_size):
                        batch_idx = i // self.batch_size + 1
                        batch_items = uncached[i : i + self.batch_size]
                        
                        future = executor.submit(batch_translate, client, self.model, batch_items, self.glossary, self.offline)
                        future_to_batch[future] = (batch_idx, batch_items)
                        logger.info(f"Submitted batch {batch_idx}/{total_batches} ({len(batch_items)} items)")
                    
                    for future in as_completed(future_to_batch):
                        batch_idx, batch_items = future_to_batch[future]
                        try:
                            batch_results = future.result()
                            logger.info(f"Finished batch {batch_idx}")
                            
                            # Check for result mismatch
                            if len(batch_results) != len(batch_items):
                                logger.warning(f"Batch {batch_idx} result count mismatch ({len(batch_results)} vs {len(batch_items)}). Using fallbacks.")
                                # Pad with original text or truncate
                                if len(batch_results) < len(batch_items):
                                    batch_results.extend(batch_items[len(batch_results):])
                                else:
                                    batch_results = batch_results[:len(batch_items)]
                            
                            # Add to translations dict and cache
                            for text, translation in zip(batch_items, batch_results):
                                translations[text] = translation
                                self.cache[text] = translation
                            
                            self.stats['api_calls'] += 1
                            # Periodically save cache to prevent data loss on crash
                            if self.stats['api_calls'] % 5 == 0:
                                self._save_cache()
                                
                        except Exception as exc:
                            logger.error(f"Batch {batch_idx} generated an exception: {exc}")
                            # Fallback to original text
                            for text in batch_items:
                                translations[text] = text
                                self.cache[text] = text
                
                self._save_cache()
                logger.info(f"Added {len(uncached)} new translations to cache")
                
            except Exception as e:
                logger.error(f"Translation process failed: {e}")
                # Fallback: use original text as translation for all uncached
                for text in uncached:
                    if text not in translations:
                        translations[text] = text
        
        elif self.cache_only and uncached:
            logger.warning(f"Cache-only mode: {len(uncached)} items not found in cache")
            for text in uncached:
                translations[text] = text
        
        # Return all translations (cached + newly translated)
        return translations
    
    def _convert_to_layout_blocks(self, extraction_result, translations: Dict[str, str]) -> List[TextBlock]:
        """Convert extraction result to layout engine blocks with translations."""
        layout_blocks = []
        
        for page in extraction_result.pages:
            for block in page.text_blocks:
                if self._contains_japanese(block.text):
                    # Map block type from extractor to layout engine
                    content_type_map = {
                        'title': ContentType.TITLE,
                        'header': ContentType.HEADER,
                        'footer': ContentType.FOOTER,
                        'caption': ContentType.CAPTION,
                        'table': ContentType.TABLE,
                        'body': ContentType.BODY,
                        'unknown': ContentType.BODY
                    }
                    
                    content_type = content_type_map.get(block.block_type, ContentType.BODY)
                    
                    # Determine layout constraint
                    constraint_map = {
                        'table': LayoutConstraint.FIXED,
                        'title': LayoutConstraint.FLEXIBLE,
                        'header': LayoutConstraint.FLEXIBLE,
                        'footer': LayoutConstraint.FLEXIBLE,
                        'caption': LayoutConstraint.FLEXIBLE,
                        'body': LayoutConstraint.FLEXIBLE,
                        'unknown': LayoutConstraint.FLEXIBLE
                    }
                    
                    constraint = constraint_map.get(block.block_type, LayoutConstraint.FLEXIBLE)
                    
                    layout_block = TextBlock(
                        id=block.id,
                        text=block.text,  # Add this required field
                        jp_text=block.text,
                        en_text=translations.get(block.text, block.text),
                        content_type=content_type,
                        constraint=constraint,
                        x=block.x0,
                        y=block.y0,
                        width=block.x1 - block.x0,
                        height=block.y1 - block.y0,
                        font_size=block.font_size,
                        font_name=block.font_name,
                        line_spacing=block.line_height,
                        char_spacing=block.char_spacing,
                        min_font_size=8.0,
                        max_font_size=72.0
                    )
                    
                    layout_blocks.append(layout_block)
        
        self.stats['total_blocks'] = len(layout_blocks)
        return layout_blocks
    
    def _prepare_translations_for_backprojection(self, layout_blocks: List[TextBlock]) -> List[Dict[str, Any]]:
        """Prepare translation data for back-projection."""
        translations = []
        
        for block in layout_blocks:
            translation_item = {
                "original": block.jp_text,
                "translated": block.en_text,
                "position": {
                    "page": int(block.id.split('_')[1]),  # Extract page number from block ID
                    "bbox": [block.x, block.y, block.x + block.width, block.y + block.height]
                },
                "font_scaling": block.optimized_font_size / block.font_size if hasattr(block, 'optimized_font_size') else 1.0,
                "layout_adjustments": {
                    "content_type": block.content_type.value,
                    "constraint": block.constraint.value
                }
            }
            translations.append(translation_item)
        
        return translations
    
    def _generate_bilingual_csv(self, extraction_result, translations: Dict[str, str], output_path: str) -> None:
        """Generate bilingual CSV file."""
        try:
            with open(output_path, "w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["page", "block_id", "japanese", "english", "block_type", "font_size"])
                
                for page in extraction_result.pages:
                    for block in page.text_blocks:
                        if self._contains_japanese(block.text):
                            en_text = translations.get(block.text, block.text)
                            writer.writerow([page.page_num + 1, block.id, block.text, en_text, block.block_type, block.font_size])
            
            logger.info(f"Bilingual CSV saved to: {output_path}")
        except Exception as e:
            logger.error(f"Failed to generate bilingual CSV: {e}")
    
    def _copy_pdf_for_translation(self) -> None:
        """Copy input PDF to output location before applying translations."""
        try:
            import shutil
            shutil.copy2(self.input_path, self.output_path)
            logger.debug(f"Copied {self.input_path} to {self.output_path}")
        except Exception as e:
            logger.error(f"Failed to copy PDF: {e}")
            raise
    
    def translate_pdf(self) -> bool:
        """Execute the complete PDF translation pipeline."""
        if not PDF_COMPONENTS_AVAILABLE:
            logger.error("PDF translation components not available")
            return False
        
        self.stats['start_time'] = datetime.now()
        logger.info(f"Starting PDF translation pipeline")
        
        try:
            # Step 1: Extract text and layout from PDF
            logger.info("Step 1: Extracting text and layout from PDF")
            extraction_result = self.extractor.extract_text_blocks(self.input_path, detailed=True)
            
            # Filter by page range if specified
            extraction_result = self._filter_pages_by_range(extraction_result)
            
            logger.info(f"Extracted {extraction_result.total_blocks} blocks ({extraction_result.total_japanese_blocks} Japanese)")
            
            if extraction_result.total_japanese_blocks == 0:
                logger.warning("No Japanese text found in the PDF")
                return True
            
            # Step 2: Extract unique Japanese text for translation
            logger.info("Step 2: Preparing unique Japanese text for translation")
            unique_japanese = self._extract_unique_japanese_text(extraction_result)
            logger.info(f"Found {len(unique_japanese)} unique Japanese text segments")
            
            # Step 3: Translate using unified cache system
            logger.info("Step 3: Translating text using unified cache system")
            translations = self._translate_with_cache(unique_japanese)
            self.stats['translated_blocks'] = len(translations)
            logger.info(f"Translated {len(translations)} text segments")
            
            # Step 4: Convert to layout blocks and optimize
            logger.info("Step 4: Optimizing layout for text expansion")
            layout_blocks = self._convert_to_layout_blocks(extraction_result, translations)
            
            # Apply layout optimization
            optimized_blocks = self.layout_engine.optimize_font_sizes(layout_blocks)
            
            # Count optimizations
            optimizations = sum(1 for block in optimized_blocks if block.adjustment_applied)
            self.stats['layout_optimizations'] = optimizations
            logger.info(f"Applied {optimizations} layout optimizations")
            
            # Step 5: Prepare translations for back-projection
            logger.info("Step 5: Preparing for back-projection")
            translation_data = self._prepare_translations_for_backprojection(optimized_blocks)
            
            # Save translations to temporary JSON file
            translations_json = f"temp_translations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(translations_json, "w", encoding="utf-8") as f:
                json.dump(translation_data, f, ensure_ascii=False, indent=2)
            
            # Step 6: Copy PDF and back-project translations
            logger.info("Step 6: Copying PDF and back-projecting translations")
            self._copy_pdf_for_translation()
            
            # Apply translations
            back_projector = PDFBackProjector(
                self.input_path,
                self.output_path,
                translations_json,
                # Loosen matching for PDF fragments (line breaks, ellipses)
                fuzzy_threshold=0.7,
                token_threshold=0.55
            )
            back_projector.load_translations()
            back_projector.extract_text_blocks()
            back_projector.process_document()
            
            # Clean up temporary file
            if os.path.exists(translations_json):
                os.remove(translations_json)
            
            logger.info(f"Back-projection completed: {self.output_path}")
            
            if self.do_audit:
                logger.info("Step 7: Generating audit report")
                audit_report = self.auditor.generate_audit_report(self.output_path, self.input_path)

                audit_json = self.output_path.replace('.pdf', '_audit.json')
                try:
                    from audit_pdf import save_report_json
                    save_report_json(audit_report, audit_json)
                    logger.info(f"Audit report saved to: {audit_json}")
                except ImportError:
                    logger.warning("Could not save audit report - audit_pdf functions not available")
            else:
                logger.info("Audit report generation skipped by flag")

            if self.do_csv:
                logger.info("Step 8: Generating bilingual CSV")
                bilingual_csv = self.output_path.replace('.pdf', '_bilingual.csv')
                self._generate_bilingual_csv(extraction_result, translations, bilingual_csv)
            else:
                logger.info("Bilingual CSV generation skipped by flag")
            
            self.stats['end_time'] = datetime.now()
            duration = (self.stats['end_time'] - self.stats['start_time']).total_seconds()
            
            logger.info(f"PDF translation completed successfully in {duration:.2f} seconds")
            logger.info(f"  Total blocks processed: {self.stats['total_blocks']}")
            logger.info(f"  Translations applied: {self.stats['translated_blocks']}")
            logger.info(f"  Layout optimizations: {self.stats['layout_optimizations']}")
            logger.info(f"  Cache hits: {self.stats['cache_hits']}")
            logger.info(f"  API calls: {self.stats['api_calls']}")
            
            return True
            
        except Exception as e:
            logger.error(f"PDF translation failed: {e}")
            if self.verbose:
                import traceback
                traceback.print_exc()
            return False


def main():
    """Main CLI interface for PDF translation."""
    parser = argparse.ArgumentParser(
        description="Japanese-to-English PDF translation pipeline with layout preservation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Examples:
  python translate_pdf.py --in document.pdf --out translated.pdf
  python translate_pdf.py --in presentation.pdf --out en_presentation.pdf --model gpt-4o-mini
  python translate_pdf.py --in report.pdf --out en_report.pdf --pages 1-10 --glossary glossary.json
  python translate_pdf.py --in manual.pdf --out en_manual.pdf --cache-only --verbose

Production Models:
  Conservative: gpt-4o-2024-08-06 (highest quality)
  Balanced:     gpt-4o-2024-08-06 (recommended)
  Cost-lean:   gpt-4o-mini (good quality)

Cache & Glossary:
  Uses shared translation_cache.json with PPTX pipeline
  Supports shared glossary.json files
  Cache hit rate typically ~90% for repeated content

Environment:
  OPENAI_API_KEY must be set for translation
  OPENAI_MODEL overrides default model selection
        """
    )
    
    # Core arguments
    parser.add_argument('--in', required=True, dest='input_path',
                       help='Input PDF file path')
    parser.add_argument('--out', required=True, dest='output_path',
                       help='Output PDF file path')
    parser.add_argument('--model', default=os.getenv("OPENAI_MODEL", "gpt-4o-2024-08-06"),
                       help='OpenAI model for translation (default: gpt-4o-2024-08-06)')
    parser.add_argument('--temperature', type=float, default=0.6,
                       help='Translation temperature (0.0-1.0, default: 0.6)')
    parser.add_argument('--batch', type=int, default=10,
                       help='Batch size for API calls (default: 10)')
    
    # Page selection
    parser.add_argument('--pages', default=None,
                       help='Page range to translate (e.g., "1-10", "5", "3-")')
    
    # Cache and glossary
    parser.add_argument('--cache', default="translation_cache.json",
                       help='Translation cache file (default: translation_cache.json)')
    parser.add_argument('--glossary', default=None,
                       help='Optional glossary JSON file (format: {JA: EN})')
    parser.add_argument('--cache-only', action='store_true',
                       help='Use only cached translations, no API calls')
    parser.add_argument('--offline', action='store_true',
                       help='Run in offline mode (no API calls)')
    
    # Output options
    parser.add_argument('--no-csv', action='store_true',
                       help='Skip generating bilingual CSV file')
    parser.add_argument('--no-audit', action='store_true',
                       help='Skip generating audit report')
    
    # Utility options
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose logging')
    parser.add_argument('--fresh', action='store_true',
                       help='Backup existing output files with timestamps')
    parser.add_argument('--concurrency', type=int, default=1,
                       help='Number of concurrent API requests')
    
    args = parser.parse_args()
    
    # Validate arguments
    if not os.path.exists(args.input_path):
        print(f"Error: Input file not found: {args.input_path}")
        sys.exit(1)
    
    if not args.input_path.lower().endswith('.pdf'):
        print("Error: Input file must be a PDF")
        sys.exit(1)
    
    # Check for OpenAI API key if needed
    if not args.offline and not args.cache_only:
        if not os.getenv("OPENAI_API_KEY") and not os.getenv("GOOGLE_API_KEY") and not os.getenv("GEMINI_API_KEY"):
            print("Error: API KEY environment variable not set")
            print("Set OPENAI_API_KEY, GOOGLE_API_KEY, or GEMINI_API_KEY, or use --offline or --cache-only mode")
            sys.exit(1)
    
    # Check component availability
    if not PDF_COMPONENTS_AVAILABLE:
        print("Error: PDF translation components not available")
        print("Please ensure all PDF translation scripts are present in scripts/ directory")
        sys.exit(1)
    
    # Backup existing files if requested
    if args.fresh:
        files_to_backup = []
        if os.path.exists(args.output_path):
            files_to_backup.append(args.output_path)
        csv_candidate = args.output_path.replace('.pdf', '_bilingual.csv')
        if os.path.exists(csv_candidate):
            files_to_backup.append(csv_candidate)
        audit_candidate = args.output_path.replace('.pdf', '_audit.json')
        if os.path.exists(audit_candidate):
            files_to_backup.append(audit_candidate)
        if os.path.exists(args.cache):
            files_to_backup.append(args.cache)
        if os.path.exists("pdf_translation.log"):
            files_to_backup.append("pdf_translation.log")

        if files_to_backup:
            # Backup the translated PDF directly, then use shared helper for other outputs
            for file_path in files_to_backup:
                if file_path == args.output_path:
                    try:
                        backup_name = get_timestamped_filename(file_path)
                        shutil.copy2(file_path, backup_name)
                        print(f"Backed up {file_path} -> {backup_name}")
                    except Exception as exc:
                        print(f"Warning: failed to backup {file_path}: {exc}")

            backup_existing_files(
                args.cache,
                csv_candidate,
                audit_candidate,
                "pdf_translation.log",
            )
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    try:
        # Initialize and run orchestrator
        orchestrator = PDFTranslationOrchestrator(
            input_path=args.input_path,
            output_path=args.output_path,
            model=args.model,
            cache_file=args.cache,
            glossary_file=args.glossary,
            batch_size=args.batch,
            temperature=args.temperature,
            pages=args.pages,
            offline=args.offline,
            cache_only=args.cache_only,
            verbose=args.verbose,
            concurrency=args.concurrency,
            do_audit=not args.no_audit,
            do_csv=not args.no_csv,
        )
        
        success = orchestrator.translate_pdf()
        
        if success:
            print(f"\n✅ PDF translation completed successfully!")
            print(f"   Input:  {args.input_path}")
            print(f"   Output: {args.output_path}")
            if not args.no_csv:
                csv_path = args.output_path.replace('.pdf', '_bilingual.csv')
                print(f"   CSV:    {csv_path}")
            if not args.no_audit:
                audit_path = args.output_path.replace('.pdf', '_audit.json')
                print(f"   Audit:  {audit_path}")
            sys.exit(0)
        else:
            print(f"\n❌ PDF translation failed!")
            print("   Check the log file for details: pdf_translation.log")
            sys.exit(1)
            
    except Exception as e:
        print(f"Error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
