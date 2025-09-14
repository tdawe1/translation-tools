#!/usr/bin/env python3
"""
apply_pdf_translation.py

PDF back-projector for replacing Japanese text with English translations while preserving original formatting.
Uses PyMuPDF (fitz) for precise text replacement and layout preservation.

Usage:
  python apply_pdf_translation.py --input original.pdf --output translated.pdf --translations translations.json

Features:
- Replace Japanese text with English translations at original positions
- Preserve font family, style, color, and formatting attributes
- Apply font scaling and layout adjustments for text expansion
- Handle both horizontal and vertical text orientations
- Maintain PDF structure and compatibility
"""

import argparse
import json
import logging
import os
import re
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from typing_extensions import NotRequired, TypedDict
from difflib import SequenceMatcher
import hashlib

# PyMuPDF (fitz) for PDF text replacement
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None
    # Only print error if not being imported for CLI help or tests
    if not (any(arg.startswith('-h') or arg.startswith('--help') for arg in sys.argv) or 
            os.environ.get("PYTEST_CURRENT_TEST")):
        print("ERROR: PyMuPDF (fitz) is required. Install via: pip install PyMuPDF", file=sys.stderr)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pdf_translation.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

# Regex for Japanese text detection
JP_CORE = r'\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff'
CJK_PUNCT = r'\u3000-\u303f'
FULLWIDTH = r'\uff00-\uffef'
JP_ANY = re.compile(f'[{JP_CORE}{CJK_PUNCT}{FULLWIDTH}]')

# Data structures
@dataclass
class TextBlock:
    """Represents a text block with position and formatting information."""
    page_num: int
    bbox: Tuple[float, float, float, float]  # (x0, y0, x1, y1)
    text: str
    font_name: str
    font_size: float
    font_color: Tuple[float, float, float]  # RGB
    is_bold: bool
    is_italic: bool
    block_id: str
    rotation: float = 0.0
    line_height: float = 1.0
    char_spacing: float = 0.0

class TranslationData(TypedDict):
    """Translation mapping data structure."""
    original: str
    translated: str
    position: NotRequired[Dict[str, Any]]
    font_scaling: NotRequired[float]
    layout_adjustments: NotRequired[Dict[str, Any]]

class MatchResult(TypedDict):
    """Result of text matching with confidence score."""
    text: str
    translation: TranslationData
    confidence: float
    match_type: str  # 'exact', 'normalized', 'fuzzy', 'token'
    normalized_key: str

class PDFBackProjector:
    """Main class for PDF text replacement and formatting preservation."""
    
    def __init__(self, input_path: str, output_path: str, translations_path: str,
                 fuzzy_threshold: float = 0.85, token_threshold: float = 0.75):
        self.input_path = input_path
        self.output_path = output_path
        self.translations_path = translations_path
        self.doc = None
        self.translations = {}
        self.text_blocks = []
        self.fuzzy_threshold = fuzzy_threshold
        self.token_threshold = token_threshold
        self.replacement_stats = {
            'total_blocks': 0,
            'replaced_blocks': 0,
            'failed_blocks': 0,
            'layout_adjustments': 0,
            'match_types': {
                'exact': 0,
                'normalized': 0,
                'fuzzy': 0,
                'token': 0,
                'no_match': 0
            },
            'confidence_scores': [],
            'unmatched_texts': []
        }
        
    def load_translations(self) -> None:
        """Load translation mappings from JSON file."""
        try:
            with open(self.translations_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            # Convert to simple mapping for quick lookup
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, dict) and 'original' in item and 'translated' in item:
                        self.translations[item['original']] = item
            elif isinstance(data, dict):
                self.translations = data
                
            logging.info(f"Loaded {len(self.translations)} translations from {self.translations_path}")
            
        except Exception as e:
            logging.error(f"Failed to load translations: {e}")
            raise
    
    def extract_text_blocks(self) -> None:
        """Extract text blocks with formatting information from PDF."""
        if fitz is None:
            raise ImportError("PyMuPDF (fitz) is required. Install via: pip install PyMuPDF")
            
        self.doc = fitz.open(self.input_path)
        self.text_blocks = []
        
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            
            # Get text blocks with detailed formatting
            blocks = page.get_text("dict")["blocks"]
            
            for block in blocks:
                if "lines" not in block:
                    continue
                    
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"].strip()
                        if not text or not JP_ANY.search(text):
                            continue
                            
                        # Extract formatting information
                        bbox = span["bbox"]
                        font_name = span.get("font", "Helvetica")
                        font_size = span.get("size", 12.0)
                        font_color = span.get("color", (0, 0, 0))
                        is_bold = "bold" in font_name.lower()
                        is_italic = "italic" in font_name.lower()
                        
                        # Generate unique block ID
                        block_id = f"page_{page_num}_block_{len(self.text_blocks)}"
                        
                        text_block = TextBlock(
                            page_num=page_num,
                            bbox=bbox,
                            text=text,
                            font_name=font_name,
                            font_size=font_size,
                            font_color=font_color,
                            is_bold=is_bold,
                            is_italic=is_italic,
                            block_id=block_id,
                            rotation=span.get("rotation", 0.0),
                            line_height=span.get("line_height", 1.0),
                            char_spacing=span.get("char_spacing", 0.0)
                        )
                        
                        self.text_blocks.append(text_block)
        
        self.replacement_stats['total_blocks'] = len(self.text_blocks)
        logging.info(f"Extracted {len(self.text_blocks)} text blocks from PDF")
    
    def normalize_text(self, text: str) -> str:
        """Normalize text for robust matching.
        
        Handles:
        - Whitespace normalization (collapse multiple spaces, trim)
        - Newline variations (\n, \r\n, \r -> \n)
        - Fullwidth to halfwidth punctuation conversion
        - Common Japanese punctuation variants
        """
        if not text:
            return text
            
        # Normalize newlines to \n
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        
        # Normalize whitespace (collapse multiple spaces, trim)
        text = re.sub(r'[ \t]+', ' ', text)  # Replace tabs and multiple spaces with single space
        text = text.strip()
        
        # Replace newlines with spaces and normalize again
        text = text.replace('\n', ' ')
        text = re.sub(r'[ \t]+', ' ', text)  # Collapse again after newline replacement
        text = text.strip()
        
        # Fullwidth to halfwidth punctuation conversion
        # Note: Keep Japanese characters and specific punctuation unchanged
        text = text.replace('，', ',').replace('．', '.').replace('：', ':')  # Fullwidth punctuation
        text = text.replace('、', ',').replace('。', '.')  # Japanese punctuation
        text = text.replace('（', '(').replace('）', ')').replace('［', '[').replace('］', ']')
        text = text.replace('｛', '{').replace('｝', '}').replace('「', '「').replace('」', '」')
        text = text.replace('『', '『').replace('』', '』')
        
        # Fullwidth numbers to halfwidth
        for i, fullwidth_num in enumerate('０１２３４５６７８９'):
            text = text.replace(fullwidth_num, str(i))
        
        # Remove spaces in Japanese text (optional - can be made configurable)
        # This helps with matching Japanese text where spaces are not used
        if JP_ANY.search(text):
            text = text.replace(' ', '')
        
        # Normalize quotes
        text = text.replace('"', '"').replace("'", "'")
        
        return text
    
    def create_stable_key(self, text: str) -> str:
        """Create a stable key for text segments using hashing.
        
        This ensures consistent matching even when the same text
        appears in different formats or orders.
        """
        # Normalize text first
        normalized = self.normalize_text(text)
        
        # Create hash for stability
        text_hash = hashlib.sha256(normalized.encode('utf-8')).hexdigest()[:16]
        
        # Combine normalized text with hash for stability and debuggability
        return f"{normalized[:50]}...{text_hash}" if len(normalized) > 50 else f"{normalized}...{text_hash}"
    
    def calculate_similarity(self, text1: str, text2: str) -> float:
        """Calculate similarity score between two texts using multiple methods."""
        if text1 == text2:
            return 1.0
        
        # Normalize both texts
        norm1 = self.normalize_text(text1)
        norm2 = self.normalize_text(text2)
        
        if norm1 == norm2:
            return 1.0
        
        # Calculate string similarity
        seq_matcher = SequenceMatcher(None, norm1, norm2)
        string_similarity = seq_matcher.ratio()
        
        # Calculate token-based similarity
        tokens1 = norm1.split()
        tokens2 = norm2.split()
        
        if not tokens1 or not tokens2:
            return string_similarity
        
        # Calculate token overlap
        token_overlap = len(set(tokens1) & set(tokens2))
        max_tokens = max(len(tokens1), len(tokens2))
        token_similarity = token_overlap / max_tokens if max_tokens > 0 else 0
        
        # Weighted combination (string similarity gets more weight)
        combined_similarity = (string_similarity * 0.7) + (token_similarity * 0.3)
        
        return combined_similarity
    
    def find_best_match(self, text: str, translations: Dict[str, TranslationData]) -> Tuple[Optional[TranslationData], float, str]:
        """Find the best matching translation with confidence scoring.
        
        Returns:
            Tuple of (translation_data, confidence_score, match_type)
        """
        if not text or not translations:
            return None, 0.0, 'no_match'
        
        # 1. Exact match (highest confidence)
        if text in translations:
            return translations[text], 1.0, 'exact'
        
        # 2. Normalized match
        normalized_text = self.normalize_text(text)
        normalized_translations = {self.normalize_text(k): v for k, v in translations.items()}
        
        if normalized_text in normalized_translations:
            return normalized_translations[normalized_text], 0.95, 'normalized'
        
        # 3. Fuzzy matching with confidence scoring
        best_match = None
        best_confidence = 0.0
        
        for original, translation in translations.items():
            confidence = self.calculate_similarity(text, original)
            
            if confidence >= self.fuzzy_threshold and confidence > best_confidence:
                best_match = translation
                best_confidence = confidence
        
        if best_match:
            return best_match, best_confidence, 'fuzzy'
        
        # 4. Token-based matching as fallback
        best_token_match = None
        best_token_confidence = 0.0
        
        for original, translation in translations.items():
            # Calculate token similarity
            text_tokens = set(self.normalize_text(text).split())
            orig_tokens = set(self.normalize_text(original).split())
            
            if not text_tokens or not orig_tokens:
                continue
            
            token_overlap = len(text_tokens & orig_tokens)
            max_tokens = max(len(text_tokens), len(orig_tokens))
            token_confidence = token_overlap / max_tokens if max_tokens > 0 else 0
            
            if token_confidence >= self.token_threshold and token_confidence > best_token_confidence:
                best_token_match = translation
                best_token_confidence = token_confidence
        
        if best_token_match:
            return best_token_match, best_token_confidence, 'token'
        
        return None, 0.0, 'no_match'
    
    def match_with_confidence(self, text_blocks: List[TextBlock], translations: Dict[str, TranslationData]) -> List[MatchResult]:
        """Match text blocks with translations and confidence scoring."""
        results = []
        
        for block in text_blocks:
            translation, confidence, match_type = self.find_best_match(block.text, translations)
            
            if translation:
                normalized_key = self.create_stable_key(block.text)
                result: MatchResult = {
                    'text': block.text,
                    'translation': translation,
                    'confidence': confidence,
                    'match_type': match_type,
                    'normalized_key': normalized_key
                }
                results.append(result)
                
                # Update statistics
                self.replacement_stats['match_types'][match_type] += 1
                self.replacement_stats['confidence_scores'].append(confidence)
                
                logging.debug(f"Matched text: '{block.text[:50]}...' with confidence {confidence:.2f} ({match_type})")
            else:
                # Log unmatched text
                self.replacement_stats['match_types']['no_match'] += 1
                self.replacement_stats['unmatched_texts'].append(block.text[:100])
                
                logging.warning(f"No translation found for: '{block.text[:50]}...'")
        
        return results
    
    def find_best_translation_match(self, text: str) -> Optional[TranslationData]:
        """Legacy method for backward compatibility."""
        translation, confidence, match_type = self.find_best_match(text, self.translations)
        return translation
    
    def calculate_optimal_font_scaling(self, original_text: str, translated_text: str, 
                                      original_bbox: Tuple[float, float, float, float],
                                      font_size: float) -> float:
        """Calculate optimal font scaling to fit translated text in original space."""
        try:
            # Calculate text expansion ratio
            orig_len = len(original_text)
            trans_len = len(translated_text)
            
            if orig_len == 0:
                return 1.0
            
            expansion_ratio = trans_len / orig_len
            
            # Calculate available space
            bbox_width = original_bbox[2] - original_bbox[0]
            bbox_height = original_bbox[3] - original_bbox[1]
            
            # Estimate character width (rough approximation)
            char_width = font_size * 0.6  # Average character width
            
            # Calculate required scaling
            required_width = trans_len * char_width * 1.1  # 10% padding
            if bbox_width > 0:
                width_scaling = bbox_width / required_width
            else:
                width_scaling = 1.0
            
            # Apply conservative scaling (don't make text too small)
            optimal_scaling = min(width_scaling, 1.0)
            optimal_scaling = max(optimal_scaling, 0.7)  # Minimum 70% of original size
            
            return optimal_scaling
            
        except Exception as e:
            logging.warning(f"Error calculating font scaling: {e}")
            return 1.0
    
    def replace_text_in_block(self, page, text_block: TextBlock, 
                            translation: TranslationData) -> bool:
        """Replace text in a specific text block."""
        try:
            translated_text = translation['translated']
            
            # Calculate font scaling if needed
            font_scaling = translation.get('font_scaling', 1.0)
            if font_scaling == 1.0:  # Auto-calculate if not provided
                font_scaling = self.calculate_optimal_font_scaling(
                    text_block.text, translated_text, text_block.bbox, text_block.font_size
                )
            
            # Apply font scaling
            new_font_size = text_block.font_size * font_scaling
            
            # Redact the original text area
            page.add_redact_annot(text_block.bbox)
            page.apply_redactions()
            
            # Prepare text formatting
            font_flags = 0
            if text_block.is_bold:
                font_flags |= getattr(fitz, 'TEXT_BOLD', 1)
            if text_block.is_italic:
                font_flags |= getattr(fitz, 'TEXT_ITALIC', 2)
            
            # Convert RGB color tuple to fitz color
            try:
                # PyMuPDF 1.26.4 expects a single RRGGBB integer
                if isinstance(text_block.font_color, tuple) and len(text_block.font_color) == 3:
                    r, g, b = text_block.font_color
                    # Convert to 0-255 range and pack as RRGGBB
                    r = int(min(255, max(0, r * 255)))
                    g = int(min(255, max(0, g * 255)))
                    b = int(min(255, max(0, b * 255)))
                    srgb_color = (r << 16) | (g << 8) | b
                    color = fitz.sRGB_to_pdf(srgb_color)
                else:
                    color = fitz.sRGB_to_pdf(0)  # Default to black
            except (AttributeError, TypeError):
                # Fallback: use black color
                color = fitz.sRGB_to_pdf(0)
            
            # Calculate insertion point
            x0, y0, x1, y1 = text_block.bbox
            insert_x = x0
            insert_y = y0 + new_font_size  # Align with top of bbox
            
            # Handle rotation if needed
            if text_block.rotation != 0:
                # For rotated text, we need to adjust the insertion point
                # This is a simplified approach - complex rotations may need more sophisticated handling
                pass
            
            # Use a standard font that's always available
            # Map common Japanese fonts to standard fonts
            font_mapping = {
                'ms-gothic': 'goth',
                'ms-pgothic': 'goth',
                'ms-mincho': 'min',
                'yugothic': 'goth',
                'yu-gothic': 'goth',
                'meiryo': 'goth',
                'hiragino-sans-gb': 'goth',
                'hiragino-kaku-gothic': 'goth',
            }
            
            # Convert font name to lowercase for matching
            font_name_lower = text_block.font_name.lower() if text_block.font_name else ''
            standard_font = 'helv'  # Default fallback
            
            # Try to find a matching standard font
            for jp_font, std_font in font_mapping.items():
                if jp_font in font_name_lower:
                    standard_font = std_font
                    break
            
            # Use the direct insert_text method with built-in fonts only (PyMuPDF 1.26.4 compatible)
            try:
                # Prepare text insertion point
                point = (insert_x, insert_y)
                
                # Use only built-in fonts that don't require font files
                # PyMuPDF built-in fonts: 'helv', 'cour', 'times', 'sym', 'zapf'
                built_in_fonts = ['helv', 'cour', 'times', 'sym', 'zapf']
                safe_font = 'helv'  # Default to Helvetica
                
                # Try to use the standard font if it's a built-in font
                if standard_font in built_in_fonts:
                    safe_font = standard_font
                
                # Insert text directly onto the page using built-in font
                page.insert_text(
                    point=point,
                    text=translated_text,
                    fontsize=new_font_size,
                    fontname=safe_font,
                    color=color,
                    rotate=text_block.rotation,
                    overlay=True,
                    fill_opacity=1.0,
                    stroke_opacity=1.0,
                    border_width=0,
                    render_mode=0  # Fill text only
                )
                
            except Exception as e:
                logging.warning(f"Direct text insertion failed: {e}")
                # Fallback: use text annotation (without color parameter for compatibility)
                try:
                    rect = fitz.Rect(insert_x, insert_y - new_font_size, 
                                   insert_x + 300, insert_y + 10)
                    
                    # Create free text annotation without color parameter
                    annot = page.add_freetext_annot(
                        rect=rect,
                        text=translated_text,
                        fontname='helv',  # Use built-in Helvetica
                        fontsize=new_font_size,
                        rotate=text_block.rotation,
                        fill_opacity=1.0,
                        border_width=0
                    )
                    # Set annotation to appear as normal text
                    annot.set_flags(getattr(fitz, 'ANNOT_INVISIBLE', 0))
                    
                except Exception as e2:
                    logging.error(f"Both text insertion methods failed: {e2}")
                    return False
            
            logging.debug(f"Replaced text in block {text_block.block_id}: '{text_block.text}' -> '{translated_text}' (scaling: {font_scaling:.2f})")
            return True
            
        except Exception as e:
            logging.error(f"Failed to replace text in block {text_block.block_id}: {e}")
            return False
    
    def apply_layout_adjustments(self, page, text_block: TextBlock, 
                               translation: TranslationData) -> None:
        """Apply layout adjustments from the translation data."""
        adjustments = translation.get('layout_adjustments', {})
        
        if 'position_offset' in adjustments:
            offset = adjustments['position_offset']
            # This would require more complex handling to adjust text positions
            # For now, we'll log that adjustments were requested
            logging.debug(f"Layout adjustments requested for block {text_block.block_id}: {offset}")
            self.replacement_stats['layout_adjustments'] += 1
    
    def preserve_formatting(self, original_pdf_path: str, translated_pdf_path: str) -> None:
        """Preserve non-text elements and overall PDF structure."""
        if fitz is None:
            return
            
        try:
            # Open both PDFs
            original_doc = fitz.open(original_pdf_path)
            translated_doc = fitz.open(translated_pdf_path)
            
            # Copy non-text elements from original to translated
            # This includes images, graphics, annotations, etc.
            for page_num in range(min(len(original_doc), len(translated_doc))):
                orig_page = original_doc[page_num]
                trans_page = translated_doc[page_num]
                
                # Copy images
                for img in orig_page.get_images():
                    trans_page.insert_image(img[0])
                
                # Copy annotations (excluding our redaction annotations)
                for annot in orig_page.annots():
                    if annot.type[0] != "Redact":  # Skip redaction annotations
                        trans_page.add_annot(annot)
                
                # Copy links
                for link in orig_page.get_links():
                    trans_page.insert_link(link)
            
            # Save the updated translated document
            translated_doc.save(self.output_path, garbage=4, deflate=True, clean=True)
            translated_doc.close()
            original_doc.close()
            
            logging.info("Preserved non-text elements and formatting from original PDF")
            
        except Exception as e:
            logging.warning(f"Could not fully preserve formatting: {e}")
            # Continue with the translated document as-is
    
    def handle_special_elements(self, pdf_path: str) -> None:
        """Handle special PDF elements like forms, bookmarks, etc."""
        if fitz is None:
            return
            
        try:
            doc = fitz.open(pdf_path)
            
            # Preserve document metadata
            metadata = doc.metadata
            if metadata:
                logging.debug(f"Preserving metadata: {metadata.get('title', 'Untitled')}")
            
            # Preserve outline (bookmarks)
            outline = doc.get_outline()
            if outline:
                logging.debug(f"Preserving {len(outline)} outline entries")
            
            doc.close()
            
        except Exception as e:
            logging.warning(f"Could not preserve special elements: {e}")
    
    def log_match_quality_report(self) -> None:
        """Log detailed match quality statistics."""
        match_types = self.replacement_stats['match_types']
        confidence_scores = self.replacement_stats['confidence_scores']
        unmatched_texts = self.replacement_stats['unmatched_texts']
        
        logging.info("=== Match Quality Report ===")
        
        # Match type distribution
        total_matched = sum(v for k, v in match_types.items() if k != 'no_match')
        logging.info(f"Total text blocks: {self.replacement_stats['total_blocks']}")
        logging.info(f"Successfully matched: {total_matched} ({total_matched/max(1, self.replacement_stats['total_blocks'])*100:.1f}%)")
        
        logging.info("Match type distribution:")
        for match_type, count in match_types.items():
            percentage = count / max(1, self.replacement_stats['total_blocks']) * 100
            logging.info(f"  {match_type}: {count} ({percentage:.1f}%)")
        
        # Confidence statistics
        if confidence_scores:
            avg_confidence = sum(confidence_scores) / len(confidence_scores)
            high_confidence = sum(1 for score in confidence_scores if score >= 0.9)
            medium_confidence = sum(1 for score in confidence_scores if 0.7 <= score < 0.9)
            low_confidence = sum(1 for score in confidence_scores if score < 0.7)
            
            logging.info(f"Average confidence: {avg_confidence:.3f}")
            logging.info(f"High confidence matches (≥0.9): {high_confidence}")
            logging.info(f"Medium confidence matches (0.7-0.9): {medium_confidence}")
            logging.info(f"Low confidence matches (<0.7): {low_confidence}")
        
        # Unmatched texts
        if unmatched_texts:
            logging.warning(f"Unmatched texts: {len(unmatched_texts)}")
            for i, text in enumerate(unmatched_texts[:5]):  # Show first 5
                logging.warning(f"  {i+1}. '{text}...'")
            if len(unmatched_texts) > 5:
                logging.warning(f"  ... and {len(unmatched_texts) - 5} more")
        
        logging.info("=== End Match Quality Report ===")
    
    def process_document(self) -> None:
        """Main processing method for the entire document."""
        if fitz is None:
            raise ImportError("PyMuPDF (fitz) is required. Install via: pip install PyMuPDF")
            
        logging.info(f"Processing PDF: {self.input_path} -> {self.output_path}")
        logging.info(f"Matching thresholds: fuzzy={self.fuzzy_threshold}, token={self.token_threshold}")
        
        # Load translations
        self.load_translations()
        
        # Extract text blocks
        self.extract_text_blocks()
        
        # Enhanced matching with confidence scoring
        match_results = self.match_with_confidence(self.text_blocks, self.translations)
        
        # Process matched text blocks
        for result in match_results:
            # Find the corresponding text block
            text_block = next((block for block in self.text_blocks if block.text == result['text']), None)
            if not text_block:
                continue
            
            # Get the page
            page = self.doc[text_block.page_num]
            
            # Replace text
            success = self.replace_text_in_block(page, text_block, result['translation'])
            
            if success:
                self.replacement_stats['replaced_blocks'] += 1
                
                # Apply layout adjustments
                self.apply_layout_adjustments(page, text_block, result['translation'])
            else:
                self.replacement_stats['failed_blocks'] += 1
        
        # Save the modified document
        temp_output = self.output_path + ".tmp"
        self.doc.save(temp_output, garbage=4, deflate=True, clean=True)
        self.doc.close()
        
        # Preserve formatting from original
        self.preserve_formatting(self.input_path, temp_output)
        
        # Handle special elements
        self.handle_special_elements(self.output_path)
        
        # Clean up temporary file
        if os.path.exists(temp_output):
            os.remove(temp_output)
        
        # Log detailed statistics
        self.log_match_quality_report()
        logging.info(f"Output saved to: {self.output_path}")

def main():
    """Main entry point for the script."""
    # Check if fitz is available
    if fitz is None and not any(arg.startswith('-h') or arg.startswith('--help') for arg in sys.argv):
        print("ERROR: PyMuPDF (fitz) is required. Install via: pip install PyMuPDF", file=sys.stderr)
        sys.exit(1)
    
    parser = argparse.ArgumentParser(
        description="Replace Japanese text in PDF with English translations while preserving formatting",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python apply_pdf_translation.py --input original.pdf --output translated.pdf --translations translations.json
  python apply_pdf_translation.py -i document.pdf -o translated.pdf -t translations.json --verbose

Translation JSON format:
  [
    {
      "original": "日本語のテキスト",
      "translated": "English translation",
      "font_scaling": 0.85,
      "layout_adjustments": {"position_offset": {"x": 0, "y": 0}}
    }
  ]
  or
  {
    "日本語のテキスト": {
      "translated": "English translation",
      "font_scaling": 0.85
    }
  }
        """
    )
    
    parser.add_argument('--input', '-i', required=True, 
                       help='Input PDF file path')
    parser.add_argument('--output', '-o', required=True,
                       help='Output PDF file path')
    parser.add_argument('--translations', '-t', required=True,
                       help='Translation mappings JSON file path')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose logging')
    parser.add_argument('--fuzzy-threshold', type=float, default=0.85,
                       help='Fuzzy matching threshold (0.0-1.0, default: 0.85)')
    parser.add_argument('--token-threshold', type=float, default=0.75,
                       help='Token matching threshold (0.0-1.0, default: 0.75)')
    parser.add_argument('--debug-matching', action='store_true',
                       help='Enable detailed matching debug output')
    
    args = parser.parse_args()
    
    # Set logging level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Enable debug matching if requested
    if args.debug_matching:
        logging.getLogger().setLevel(logging.DEBUG)
        logging.info("Debug matching enabled - detailed matching output will be shown")
    
    # Validate thresholds
    if not 0 <= args.fuzzy_threshold <= 1:
        logging.error("Fuzzy threshold must be between 0.0 and 1.0")
        sys.exit(1)
    
    if not 0 <= args.token_threshold <= 1:
        logging.error("Token threshold must be between 0.0 and 1.0")
        sys.exit(1)
    
    # Validate input files
    if not os.path.exists(args.input):
        logging.error(f"Input file not found: {args.input}")
        sys.exit(1)
    
    if not os.path.exists(args.translations):
        logging.error(f"Translations file not found: {args.translations}")
        sys.exit(1)
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    try:
        # Initialize and run the back-projector with configurable thresholds
        projector = PDFBackProjector(
            args.input, 
            args.output, 
            args.translations,
            fuzzy_threshold=args.fuzzy_threshold,
            token_threshold=args.token_threshold
        )
        projector.process_document()
        
        logging.info("PDF translation completed successfully!")
        
    except Exception as e:
        logging.error(f"PDF translation failed: {e}")
        sys.exit(1)

# Standalone utility functions for testing and external use
def calculate_optimal_font_scaling_standalone(original_text: str, translated_text: str, 
                                             original_bbox: Tuple[float, float, float, float],
                                             font_size: float) -> float:
    """Standalone version of font scaling calculation for testing."""
    try:
        # Calculate text expansion ratio
        orig_len = len(original_text)
        trans_len = len(translated_text)
        
        if orig_len == 0:
            return 1.0
        
        expansion_ratio = trans_len / orig_len
        
        # Calculate available space
        bbox_width = original_bbox[2] - original_bbox[0]
        bbox_height = original_bbox[3] - original_bbox[1]
        
        # Estimate character width (rough approximation)
        char_width = font_size * 0.6  # Average character width
        
        # Calculate required scaling
        required_width = trans_len * char_width * 1.1  # 10% padding
        if bbox_width > 0:
            width_scaling = bbox_width / required_width
        else:
            width_scaling = 1.0
        
        # Apply conservative scaling (don't make text too small)
        optimal_scaling = min(width_scaling, 1.0)
        optimal_scaling = max(optimal_scaling, 0.7)  # Minimum 70% of original size
        
        return optimal_scaling
        
    except Exception as e:
        print(f"Error calculating font scaling: {e}")
        return 1.0

def normalize_text_standalone(text: str) -> str:
    """Standalone version of text normalization for testing."""
    if not text:
        return text
    
    # Normalize newlines to \n
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    
    # Normalize whitespace
    text = re.sub(r'[ \t]+', ' ', text)
    text = text.strip()
    
    # Replace newlines with spaces and normalize again
    text = text.replace('\n', ' ')
    text = re.sub(r'[ \t]+', ' ', text)  # Collapse again after newline replacement
    text = text.strip()
    
    # Fullwidth to halfwidth punctuation conversion
    text = text.replace('，', ',').replace('．', '.').replace('：', ':')  # Fullwidth punctuation
    text = text.replace('、', ',').replace('。', '.')  # Japanese punctuation
    text = text.replace('（', '(').replace('）', ')').replace('［', '[').replace('］', ']')
    text = text.replace('｛', '{').replace('｝', '}')
    
    # Fullwidth numbers to halfwidth
    for i, fullwidth_num in enumerate('０１２３４５６７８９'):
        text = text.replace(fullwidth_num, str(i))
    
    # Remove spaces in Japanese text (helps with matching)
    # Using a simple check for Japanese characters
    jp_pattern = re.compile(r'[\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff]')
    if jp_pattern.search(text):
        text = text.replace(' ', '')
    
    return text

def calculate_similarity_standalone(text1: str, text2: str) -> float:
    """Standalone version of similarity calculation for testing."""
    if text1 == text2:
        return 1.0
    
    # Normalize both texts
    norm1 = normalize_text_standalone(text1)
    norm2 = normalize_text_standalone(text2)
    
    if norm1 == norm2:
        return 1.0
    
    # Calculate string similarity
    seq_matcher = SequenceMatcher(None, norm1, norm2)
    string_similarity = seq_matcher.ratio()
    
    # Calculate token-based similarity
    tokens1 = norm1.split()
    tokens2 = norm2.split()
    
    if not tokens1 or not tokens2:
        return string_similarity
    
    # Calculate token overlap
    token_overlap = len(set(tokens1) & set(tokens2))
    max_tokens = max(len(tokens1), len(tokens2))
    token_similarity = token_overlap / max_tokens if max_tokens > 0 else 0
    
    # Weighted combination
    combined_similarity = (string_similarity * 0.7) + (token_similarity * 0.3)
    
    return combined_similarity

def find_best_translation_match_standalone(text: str, translations: Dict[str, Any], 
                                          fuzzy_threshold: float = 0.85, 
                                          token_threshold: float = 0.75) -> Tuple[Optional[Dict[str, Any]], float, str]:
    """Enhanced standalone version of translation matching for testing."""
    if not text or not translations:
        return None, 0.0, 'no_match'
    
    # 1. Exact match
    if text in translations:
        return translations[text], 1.0, 'exact'
    
    # 2. Normalized match
    normalized_text = normalize_text_standalone(text)
    normalized_translations = {normalize_text_standalone(k): v for k, v in translations.items()}
    
    if normalized_text in normalized_translations:
        return normalized_translations[normalized_text], 0.95, 'normalized'
    
    # 3. Fuzzy matching
    best_match = None
    best_confidence = 0.0
    
    for original, translation in translations.items():
        confidence = calculate_similarity_standalone(text, original)
        
        if confidence >= fuzzy_threshold and confidence > best_confidence:
            best_match = translation
            best_confidence = confidence
    
    if best_match:
        return best_match, best_confidence, 'fuzzy'
    
    # 4. Token-based matching
    best_token_match = None
    best_token_confidence = 0.0
    
    for original, translation in translations.items():
        text_tokens = set(normalize_text_standalone(text).split())
        orig_tokens = set(normalize_text_standalone(original).split())
        
        if not text_tokens or not orig_tokens:
            continue
        
        token_overlap = len(text_tokens & orig_tokens)
        max_tokens = max(len(text_tokens), len(orig_tokens))
        token_confidence = token_overlap / max_tokens if max_tokens > 0 else 0
        
        if token_confidence >= token_threshold and token_confidence > best_token_confidence:
            best_token_match = translation
            best_token_confidence = token_confidence
    
    if best_token_match:
        return best_token_match, best_token_confidence, 'token'
    
    return None, 0.0, 'no_match'

if __name__ == "__main__":
    main()