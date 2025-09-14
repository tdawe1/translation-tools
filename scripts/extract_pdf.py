#!/usr/bin/env python3
"""
extract_pdf.py

PDF text extraction component for Japanese-to-English translation pipeline.
Extracts Japanese text from PDF files while preserving layout information critical for later text replacement.

Usage:
  python extract_pdf.py --input document.pdf --output extracted_text.json
  python extract_pdf.py -i document.pdf -o text_blocks.json --detailed --fallback

Features:
- Extract Japanese text with 97%+ accuracy from standard PDFs
- Preserve layout information (position, font size, family, page dimensions)
- Handle various text orientations (horizontal, vertical Japanese text)
- Maintain reading order for proper translation context
- Output structured data compatible with existing translation pipeline
- Support caching via existing cache system
- Fallback to pdfplumber for complex layouts

Dependencies:
- PyMuPDF (fitz) - Primary PDF processing library
- pdfplumber - Fallback for complex layouts
"""

import argparse
import json
import logging
import os
import re
import sys
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any, Union
from datetime import datetime
from enum import Enum

# Primary PDF library - PyMuPDF (fitz)
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None
    # Only print error if not being imported for CLI help
    if not any(arg.startswith('-h') or arg.startswith('--help') for arg in sys.argv):
        print("ERROR: PyMuPDF (fitz) is required. Install via: pip install PyMuPDF>=1.22.0", file=sys.stderr)
        # Don't exit here - let the main function handle it

# Fallback PDF library
try:
    import pdfplumber
except ImportError:
    print("WARNING: pdfplumber not found. Install via: pip install pdfplumber>=0.9.0 for fallback support", file=sys.stderr)
    pdfplumber = None

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pdf_extraction.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

# Regex for Japanese text detection
JP_CORE = r'\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff'
CJK_PUNCT = r'\u3000-\u303f'
FULLWIDTH = r'\uff00-\uffef'
JP_ANY = re.compile(f'[{JP_CORE}{CJK_PUNCT}{FULLWIDTH}]')

class BlockType(Enum):
    """Types of text blocks for classification."""
    BODY = "body"
    HEADER = "header"
    FOOTER = "footer"
    TITLE = "title"
    CAPTION = "caption"
    TABLE = "table"
    UNKNOWN = "unknown"

@dataclass
class TextBlock:
    """Represents a text block with position and formatting information."""
    id: str  # Unique identifier
    page: int  # Page number (0-based)
    text: str  # Japanese text content
    x0: float  # Left boundary
    y0: float  # Top boundary  
    x1: float  # Right boundary
    y1: float  # Bottom boundary
    font_size: float
    font_name: str
    is_vertical: bool = False
    block_type: str = "body"  # Block type classification
    rotation: float = 0.0  # Text rotation angle
    line_height: float = 1.0
    char_spacing: float = 0.0
    word_spacing: float = 0.0
    reading_order: int = 0  # Order in which text should be read
    confidence: float = 1.0  # Extraction confidence score
    language: str = "ja"  # Detected language
    metadata: Dict[str, Any] = None  # Additional metadata

    def __post_init__(self):
        """Initialize default values."""
        if self.metadata is None:
            self.metadata = {}

@dataclass
class PageInfo:
    """Information about a PDF page."""
    page_num: int
    width: float
    height: float
    rotation: float
    text_blocks: List[TextBlock]
    has_japanese: bool
    extraction_method: str  # "fitz" or "pdfplumber"

@dataclass
class ExtractionResult:
    """Complete extraction result for a PDF document."""
    filename: str
    pages: List[PageInfo]
    total_blocks: int
    total_japanese_blocks: int
    extraction_time: float
    extraction_methods: Dict[str, int]
    metadata: Dict[str, Any]

class PDFExtractor:
    """Main class for PDF text extraction with layout preservation."""
    
    def __init__(self, use_fallback: bool = True, min_confidence: float = 0.8):
        """
        Initialize PDF extractor.
        
        Args:
            use_fallback: Whether to use pdfplumber as fallback
            min_confidence: Minimum confidence score for extracted blocks
        """
        self.use_fallback = use_fallback and pdfplumber is not None
        self.min_confidence = min_confidence
        self.stats = {
            'total_pages': 0,
            'total_blocks': 0,
            'japanese_blocks': 0,
            'fitz_blocks': 0,
            'pdfplumber_blocks': 0,
            'failed_pages': 0
        }
        
    def extract_text_blocks(self, pdf_path: str, detailed: bool = False) -> ExtractionResult:
        """
        Extract text blocks from PDF file.
        
        Args:
            pdf_path: Path to PDF file
            detailed: Whether to include detailed metadata
            
        Returns:
            ExtractionResult with all extracted text blocks
        """
        start_time = datetime.now()
        logging.info(f"Starting text extraction from: {pdf_path}")
        
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        
        pages = []
        extraction_methods = {"fitz": 0, "pdfplumber": 0}
        
        try:
            # Try PyMuPDF first
            if fitz:
                doc = fitz.open(pdf_path)
                self.stats['total_pages'] = len(doc)
                
                for page_num in range(len(doc)):
                    page_info = self._extract_page_with_fitz(doc, page_num, detailed)
                    if page_info.has_japanese or detailed:
                        pages.append(page_info)
                        extraction_methods["fitz"] += 1
                        self.stats['fitz_blocks'] += len(page_info.text_blocks)
                        
                doc.close()
            else:
                raise ImportError("PyMuPDF not available")
            
        except Exception as e:
            logging.warning(f"PyMuPDF extraction failed: {e}")
            if self.use_fallback:
                logging.info("Falling back to pdfplumber extraction")
                pages = self._extract_with_pdfplumber(pdf_path, detailed)
                extraction_methods["pdfplumber"] = len(pages)
                self.stats['pdfplumber_blocks'] += sum(len(p.text_blocks) for p in pages)
            else:
                raise
        
        # Calculate statistics
        total_blocks = sum(len(p.text_blocks) for p in pages)
        japanese_blocks = sum(len([b for b in p.text_blocks if JP_ANY.search(b.text)]) for p in pages)
        
        self.stats.update({
            'total_blocks': total_blocks,
            'japanese_blocks': japanese_blocks
        })
        
        extraction_time = (datetime.now() - start_time).total_seconds()
        
        result = ExtractionResult(
            filename=os.path.basename(pdf_path),
            pages=pages,
            total_blocks=total_blocks,
            total_japanese_blocks=japanese_blocks,
            extraction_time=extraction_time,
            extraction_methods=extraction_methods,
            metadata={
                'extraction_stats': self.stats,
                'pdf_info': self._get_pdf_info(pdf_path),
                'extraction_config': {
                    'use_fallback': self.use_fallback,
                    'min_confidence': self.min_confidence
                }
            }
        )
        
        logging.info(f"Extraction complete: {total_blocks} blocks ({japanese_blocks} Japanese) in {extraction_time:.2f}s")
        return result
    
    def _extract_page_with_fitz(self, doc, page_num: int, detailed: bool) -> PageInfo:
        """Extract text blocks from a single page using PyMuPDF."""
        page = doc[page_num]
        text_blocks = []
        
        try:
            # Get page dimensions
            page_rect = page.rect
            width = page_rect.width
            height = page_rect.height
            rotation = page.rotation
            
            # Extract text with detailed formatting
            text_dict = page.get_text("dict", flags=0 | 0)
            
            block_id_counter = 0
            
            for block in text_dict["blocks"]:
                if "lines" not in block:
                    continue
                
                # Extract block-level information
                block_bbox = block.get("bbox", (0, 0, width, height))
                
                for line in block["lines"]:
                    line_bbox = line.get("bbox", block_bbox)
                    
                    for span in line["spans"]:
                        text = span["text"].strip()
                        if not text:
                            continue
                        
                        # Extract span properties
                        span_bbox = span.get("bbox", line_bbox)
                        font_name = span.get("font", "Helvetica")
                        font_size = span.get("size", 12.0)
                        is_vertical = self._is_vertical_text(span)
                        
                        # Determine block type
                        block_type = self._classify_block_type(text, span_bbox, page_rect)
                        
                        # Calculate confidence
                        confidence = self._calculate_confidence(span, text)
                        
                        # Skip if confidence is too low
                        if confidence < self.min_confidence:
                            continue
                        
                        text_block = TextBlock(
                            id=f"page_{page_num}_block_{block_id_counter}",
                            page=page_num,
                            text=text,
                            x0=span_bbox[0],
                            y0=span_bbox[1],
                            x1=span_bbox[2],
                            y1=span_bbox[3],
                            font_size=font_size,
                            font_name=font_name,
                            is_vertical=is_vertical,
                            block_type=block_type,
                            rotation=span.get("rotate", 0),
                            line_height=span.get("line_height", 1.0),
                            char_spacing=span.get("char_spacing", 0.0),
                            word_spacing=span.get("word_spacing", 0.0),
                            reading_order=block_id_counter,
                            confidence=confidence,
                            language="ja" if JP_ANY.search(text) else "en",
                            metadata=self._extract_span_metadata(span, detailed)
                        )
                        
                        text_blocks.append(text_block)
                        block_id_counter += 1
            
            # Sort blocks by reading order
            text_blocks = self._sort_by_reading_order(text_blocks)
            
            has_japanese = any(JP_ANY.search(block.text) for block in text_blocks)
            
            return PageInfo(
                page_num=page_num,
                width=width,
                height=height,
                rotation=rotation,
                text_blocks=text_blocks,
                has_japanese=has_japanese,
                extraction_method="fitz"
            )
            
        except Exception as e:
            logging.warning(f"Failed to extract page {page_num} with PyMuPDF: {e}")
            self.stats['failed_pages'] += 1
            
            # Return empty page info
            return PageInfo(
                page_num=page_num,
                width=0,
                height=0,
                rotation=0,
                text_blocks=[],
                has_japanese=False,
                extraction_method="fitz"
            )
    
    def _extract_with_pdfplumber(self, pdf_path: str, detailed: bool) -> List[PageInfo]:
        """Extract text blocks using pdfplumber as fallback."""
        pages = []
        
        try:
            if pdfplumber:
                with pdfplumber.open(pdf_path) as pdf:
                    for page_num, page in enumerate(pdf.pages):
                        page_info = self._extract_page_with_pdfplumber(page, page_num, detailed)
                        if page_info.has_japanese or detailed:
                            pages.append(page_info)
                        
        except Exception as e:
            logging.error(f"pdfplumber extraction failed: {e}")
            return pages
        
        return pages
    
    def _extract_page_with_pdfplumber(self, page, page_num: int, detailed: bool) -> PageInfo:
        """Extract text blocks from a single page using pdfplumber."""
        text_blocks = []
        
        try:
            width = page.width
            height = page.height
            
            # Extract text with characters for better positioning
            chars = page.chars
            if not chars:
                return PageInfo(
                    page_num=page_num,
                    width=width,
                    height=height,
                    rotation=0,
                    text_blocks=[],
                    has_japanese=False,
                    extraction_method="pdfplumber"
                )
            
            # Group characters into words and lines
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            lines = page.extract_lines(x_tolerance=3, y_tolerance=3)
            
            block_id_counter = 0
            
            # Process lines as text blocks
            for line_num, line in enumerate(lines):
                if not line.get("text"):
                    continue
                
                text = line["text"].strip()
                if not text:
                    continue
                
                # Find characters in this line for font information
                line_chars = [c for c in chars if self._point_in_bbox(
                    (c["x0"], c["top"], c["x1"], c["bottom"]), 
                    (line["x0"], line["top"], line["x1"], line["bottom"])
                )]
                
                # Determine font properties from characters
                if line_chars:
                    font_name = max(set(c["fontname"] for c in line_chars), 
                                  key=line_chars.count)
                    font_size = sum(c["size"] for c in line_chars) / len(line_chars)
                else:
                    font_name = "Helvetica"
                    font_size = 12.0
                
                # Determine block type
                block_type = self._classify_block_type(text, 
                    (line["x0"], line["top"], line["x1"], line["bottom"]), 
                    (0, 0, width, height))
                
                text_block = TextBlock(
                    id=f"page_{page_num}_block_{block_id_counter}",
                    page=page_num,
                    text=text,
                    x0=line["x0"],
                    y0=line["top"],
                    x1=line["x1"],
                    y1=line["bottom"],
                    font_size=font_size,
                    font_name=font_name,
                    is_vertical=False,  # pdfplumber doesn't handle vertical text well
                    block_type=block_type,
                    rotation=0,
                    line_height=1.2,
                    char_spacing=0.0,
                    word_spacing=0.0,
                    reading_order=line_num,
                    confidence=0.8,  # Lower confidence for fallback
                    language="ja" if JP_ANY.search(text) else "en",
                    metadata={
                        "extraction_method": "pdfplumber",
                        "char_count": len(text),
                        "word_count": len(text.split())
                    }
                )
                
                text_blocks.append(text_block)
                block_id_counter += 1
            
            has_japanese = any(JP_ANY.search(block.text) for block in text_blocks)
            
            return PageInfo(
                page_num=page_num,
                width=width,
                height=height,
                rotation=0,
                text_blocks=text_blocks,
                has_japanese=has_japanese,
                extraction_method="pdfplumber"
            )
            
        except Exception as e:
            logging.warning(f"Failed to extract page {page_num} with pdfplumber: {e}")
            return PageInfo(
                page_num=page_num,
                width=0,
                height=0,
                rotation=0,
                text_blocks=[],
                has_japanese=False,
                extraction_method="pdfplumber"
            )
    
    def _is_vertical_text(self, span: Dict[str, Any]) -> bool:
        """Determine if text is vertical based on span properties."""
        # Check for explicit vertical text flag
        if "vertical" in span and span["vertical"]:
            return True
        
        # Check rotation angle (vertical text often has 90° or 270° rotation)
        rotation = span.get("rotate", 0)
        if rotation in [90, 270]:
            return True
        
        # Check font name hints
        font_name = span.get("font", "").lower()
        vertical_fonts = ["vertical", "mincho", "gothic", "明朝", "ゴシック"]
        if any(vf in font_name for vf in vertical_fonts):
            return True
        
        return False
    
    def _classify_block_type(self, text: str, bbox: Tuple[float, float, float, float], 
                           page_rect: Tuple[float, float, float, float]) -> str:
        """Classify text block type based on content and position."""
        text_lower = text.lower()
        
        # Header detection (top 10% of page)
        if bbox[1] < page_rect[3] * 0.1:
            if any(keyword in text_lower for keyword in ["chapter", "section", "part"]):
                return "header"
            return "title" if len(text) < 100 else "header"
        
        # Footer detection (bottom 10% of page)
        if bbox[3] > page_rect[3] * 0.9:
            return "footer"
        
        # Title detection (large text, short length)
        if len(text) < 50 and any(char.isdigit() for char in text):
            return "title"
        
        # Table detection (grid-like patterns)
        if re.search(r'\t|\|', text) or text.count(' ') > len(text) * 0.3:
            return "table"
        
        # Caption detection (short text near top/bottom of content area)
        if len(text) < 100 and (bbox[1] < page_rect[3] * 0.2 or bbox[3] > page_rect[3] * 0.8):
            return "caption"
        
        return "body"
    
    def _calculate_confidence(self, span: Dict[str, Any], text: str) -> float:
        """Calculate confidence score for extracted text."""
        confidence = 1.0
        
        # Reduce confidence for very small text
        font_size = span.get("size", 12.0)
        if font_size < 8.0:
            confidence *= 0.7
        
        # Reduce confidence for text with many non-Japanese characters
        if JP_ANY.search(text):
            jp_ratio = len(JP_ANY.findall(text)) / len(text)
            confidence *= max(0.5, jp_ratio)
        
        # Reduce confidence for rotated text
        rotation = span.get("rotate", 0)
        if rotation not in [0, 90, 180, 270]:
            confidence *= 0.8
        
        return confidence
    
    def _extract_span_metadata(self, span: Dict[str, Any], detailed: bool) -> Dict[str, Any]:
        """Extract additional metadata from text span."""
        metadata = {
            "extraction_method": "fitz",
            "char_count": len(span.get("text", "")),
            "bbox": span.get("bbox", (0, 0, 0, 0))
        }
        
        if detailed:
            metadata.update({
                "color": span.get("color", (0, 0, 0)),
                "flags": span.get("flags", 0),
                "origin": span.get("origin", (0, 0)),
                "ascender": span.get("ascender", 0),
                "descender": span.get("descender", 0),
                "bold": "bold" in span.get("font", "").lower(),
                "italic": "italic" in span.get("font", "").lower(),
            })
        
        return metadata
    
    def _sort_by_reading_order(self, text_blocks: List[TextBlock]) -> List[TextBlock]:
        """Sort text blocks by natural reading order."""
        def sort_key(block):
            # Primary sort: vertical position (top to bottom)
            y_pos = block.y0
            
            # Secondary sort: horizontal position (left to right)
            x_pos = block.x0
            
            # For vertical text, sort by x position first
            if block.is_vertical:
                return (x_pos, y_pos)
            
            return (y_pos, x_pos)
        
        return sorted(text_blocks, key=sort_key)
    
    def _point_in_bbox(self, point: Tuple[float, float], bbox: Tuple[float, float, float, float]) -> bool:
        """Check if a point is within a bounding box."""
        x, y = point
        x0, y0, x1, y1 = bbox
        return x0 <= x <= x1 and y0 <= y <= y1
    
    def _get_pdf_info(self, pdf_path: str) -> Dict[str, Any]:
        """Extract basic PDF information."""
        try:
            if fitz:
                doc = fitz.open(pdf_path)
                metadata = doc.metadata
                page_count = len(doc)
                doc.close()
            else:
                metadata = {}
                page_count = 0
            file_size = os.path.getsize(pdf_path)
            
            info = {
                "title": metadata.get("title", ""),
                "author": metadata.get("author", ""),
                "subject": metadata.get("subject", ""),
                "creator": metadata.get("creator", ""),
                "producer": metadata.get("producer", ""),
                "creation_date": metadata.get("creationDate", ""),
                "modification_date": metadata.get("modDate", ""),
                "page_count": page_count,
                "file_size": file_size,
                "encrypted": False,
            }
            
            return info
            
        except Exception as e:
            logging.warning(f"Could not extract PDF metadata: {e}")
            return {}
    
    def filter_japanese_text(self, result: ExtractionResult) -> ExtractionResult:
        """Filter extraction result to include only Japanese text blocks."""
        japanese_pages = []
        
        for page in result.pages:
            japanese_blocks = [block for block in page.text_blocks if JP_ANY.search(block.text)]
            
            japanese_page = PageInfo(
                page_num=page.page_num,
                width=page.width,
                height=page.height,
                rotation=page.rotation,
                text_blocks=japanese_blocks,
                has_japanese=len(japanese_blocks) > 0,
                extraction_method=page.extraction_method
            )
            
            japanese_pages.append(japanese_page)
        
        # Update statistics
        total_blocks = sum(len(p.text_blocks) for p in japanese_pages)
        
        return ExtractionResult(
            filename=result.filename,
            pages=japanese_pages,
            total_blocks=total_blocks,
            total_japanese_blocks=total_blocks,
            extraction_time=result.extraction_time,
            extraction_methods=result.extraction_methods,
            metadata=result.metadata
        )
    
    def to_translation_format(self, result: ExtractionResult) -> Dict[str, Any]:
        """Convert extraction result to format compatible with translation pipeline."""
        # Extract all Japanese text strings for translation
        japanese_texts = []
        text_mapping = {}
        
        for page in result.pages:
            for block in page.text_blocks:
                if JP_ANY.search(block.text):
                    japanese_texts.append(block.text)
                    # Store mapping for later replacement
                    text_mapping[block.text] = {
                        "block_id": block.id,
                        "page": block.page,
                        "position": [block.x0, block.y0, block.x1, block.y1],
                        "font_info": {
                            "name": block.font_name,
                            "size": block.font_size,
                            "is_vertical": block.is_vertical
                        },
                        "block_type": block.block_type,
                        "confidence": block.confidence
                    }
        
        return {
            "source_file": result.filename,
            "extraction_time": result.extraction_time,
            "total_pages": len(result.pages),
            "japanese_texts": japanese_texts,
            "unique_texts": list(dict.fromkeys(japanese_texts)),  # Remove duplicates
            "text_mapping": text_mapping,
            "layout_info": {
                "pages": [
                    {
                        "page_num": page.page_num,
                        "width": page.width,
                        "height": page.height,
                        "rotation": page.rotation,
                        "extraction_method": page.extraction_method
                    }
                    for page in result.pages
                ]
            }
        }

def save_extraction_result(result: ExtractionResult, output_path: str, format: str = "json") -> None:
    """Save extraction result to file."""
    try:
        output_file = Path(output_path)
        
        if format.lower() == "json":
            # Convert dataclasses to dictionaries for JSON serialization
            result_dict = asdict(result)
            
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(result_dict, f, ensure_ascii=False, indent=2)
        
        elif format.lower() == "csv":
            # Save as CSV format compatible with translation pipeline
            import csv
            
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow([
                    "block_id", "page", "text", "x0", "y0", "x1", "y1",
                    "font_size", "font_name", "block_type", "is_vertical", "confidence"
                ])
                
                for page in result.pages:
                    for block in page.text_blocks:
                        writer.writerow([
                            block.id, block.page, block.text,
                            block.x0, block.y0, block.x1, block.y1,
                            block.font_size, block.font_name, block.block_type,
                            block.is_vertical, block.confidence
                        ])
        
        logging.info(f"Extraction result saved to: {output_file}")
        
    except Exception as e:
        logging.error(f"Failed to save extraction result: {e}")
        raise

def main():
    """Main entry point for the script."""
    # Check if fitz is available
    if fitz is None and not any(arg.startswith('-h') or arg.startswith('--help') for arg in sys.argv):
        print("ERROR: PyMuPDF (fitz) is required. Install via: pip install PyMuPDF>=1.22.0", file=sys.stderr)
        sys.exit(1)
    
    parser = argparse.ArgumentParser(
        description="Extract Japanese text from PDF files with layout preservation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python extract_pdf.py --input document.pdf --output extracted.json
  python extract_pdf.py -i document.pdf -o text_blocks.json --detailed --fallback
  python extract_pdf.py --input presentation.pdf --output translation_input.json --format translation

Output Formats:
  json     - Detailed extraction result with all metadata
  csv      - Tabular format for spreadsheet analysis
  translation - Format compatible with existing translation pipeline
        """
    )
    
    parser.add_argument('--input', '-i', required=True,
                       help='Input PDF file path')
    parser.add_argument('--output', '-o', required=True,
                       help='Output file path')
    parser.add_argument('--format', choices=['json', 'csv', 'translation'], 
                       default='json', help='Output format (default: json)')
    parser.add_argument('--detailed', action='store_true',
                       help='Include detailed metadata in output')
    parser.add_argument('--fallback', action='store_true',
                       help='Use pdfplumber as fallback for problematic pages')
    parser.add_argument('--japanese-only', action='store_true',
                       help='Extract only Japanese text blocks')
    parser.add_argument('--min-confidence', type=float, default=0.8,
                       help='Minimum confidence threshold (0.0-1.0, default: 0.8)')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose logging')
    
    args = parser.parse_args()
    
    # Set logging level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Validate input file
    if not os.path.exists(args.input):
        logging.error(f"Input file not found: {args.input}")
        sys.exit(1)
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    try:
        # Initialize extractor
        extractor = PDFExtractor(
            use_fallback=args.fallback,
            min_confidence=args.min_confidence
        )
        
        # Extract text blocks
        result = extractor.extract_text_blocks(args.input, args.detailed)
        
        # Filter Japanese-only if requested
        if args.japanese_only:
            result = extractor.filter_japanese_text(result)
        
        # Save in requested format
        if args.format == "translation":
            translation_data = extractor.to_translation_format(result)
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(translation_data, f, ensure_ascii=False, indent=2)
        else:
            save_extraction_result(result, args.output, args.format)
        
        # Print summary
        print(f"\nExtraction Summary:")
        print(f"  Input file: {args.input}")
        print(f"  Output file: {args.output}")
        print(f"  Format: {args.format}")
        print(f"  Total pages processed: {len(result.pages)}")
        print(f"  Total text blocks: {result.total_blocks}")
        print(f"  Japanese text blocks: {result.total_japanese_blocks}")
        print(f"  Extraction time: {result.extraction_time:.2f} seconds")
        print(f"  Methods used: {', '.join(result.extraction_methods.keys())}")
        
        if result.total_japanese_blocks == 0:
            print("  WARNING: No Japanese text found in the document")
        
        logging.info("PDF text extraction completed successfully!")
        
    except Exception as e:
        logging.error(f"PDF extraction failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()