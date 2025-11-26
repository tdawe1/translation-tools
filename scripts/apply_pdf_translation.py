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
import shutil
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
    block_index: int
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
        self._font_cache: Dict[str, fitz.Font] = {}  # font_path -> Font object

        # Preferred font files (installed on the system)
        font_paths = {
            "sans_cjk_regular": "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
            "sans_cjk_bold": "/usr/share/fonts/noto-cjk/NotoSansCJK-Bold.ttc",
            "serif_cjk_regular": "/usr/share/fonts/noto-cjk/NotoSerifCJK-Regular.ttc",
            "serif_cjk_bold": "/usr/share/fonts/noto-cjk/NotoSerifCJK-Bold.ttc",
            "latin_sans_regular": "/usr/share/fonts/noto/NotoSans-Regular.ttf",
            "latin_sans_bold": "/usr/share/fonts/noto/NotoSans-Bold.ttf",
            # User-installed fonts frequently present in customer PDFs
            "helvetica_neue_regular": "/usr/local/share/fonts/h/Helvetica_Neue_Regular.otf",
            "helvetica_neue_bold": "/usr/local/share/fonts/h/Helvetica_Neue_Condensed_Bold.ttf",
            "hiragino_gothic": "/usr/local/share/fonts/h/Hiragino_Kaku_Gothic_Pro_W6.otf",
            "hiragino_maru": "/usr/local/share/fonts/h/Hiragino_Maru_Gothic_Pro_W4.otf",
            "hiragino_mincho": "/usr/local/share/fonts/h/Hiragino_Mincho_ProN_W3.otf",
            "rodin_regular": "/usr/local/share/fonts/f/FOT_Rodin_Pro_M.otf",
            "rodin_bold": "/usr/local/share/fonts/f/FOT_Rodin_Pro_DB.otf",
            "rodinntlg_regular": "/usr/local/share/fonts/f/FOT_RodinNTLG_Pro_M.otf",
            "rodinntlg_bold": "/usr/local/share/fonts/f/FOT_RodinNTLG_Pro_DB.otf",
        }
        # Only keep font paths that exist on this machine
        self.font_paths = {k: v for k, v in font_paths.items() if Path(v).exists()}

    def _get_font(self, font_path: str) -> Optional[fitz.Font]:
        """Register and return a Font object for PyMuPDF insertion."""
        if not font_path:
            return None
        if font_path in self._font_cache:
            return self._font_cache[font_path]
        try:
            font_obj = fitz.Font(fontfile=font_path)
            self._font_cache[font_path] = font_obj
            return font_obj
        except Exception:
            return None

    def _select_font_path(self, is_bold: bool, prefer_serif: bool = False) -> str:
        """Pick the best Noto font path based on weight and serif/sans preference."""
        if prefer_serif:
            return self.font_paths["serif_cjk_bold" if is_bold else "serif_cjk_regular"]
        return self.font_paths["sans_cjk_bold" if is_bold else "sans_cjk_regular"]

    def _font_candidates(self, text_block: TextBlock) -> List[str]:
        """Return a prioritized list of font file paths for this block."""
        candidates: List[str] = []
        name = (text_block.font_name or "").lower()

        def add(key: str):
            path = self.font_paths.get(key)
            if path and path not in candidates:
                candidates.append(path)

        # Prefer the original family if installed
        if "rodin" in name:
            add("rodin_bold" if text_block.is_bold else "rodin_regular")
            add("rodinntlg_bold" if text_block.is_bold else "rodinntlg_regular")
        if "hiragino" in name or "kaku" in name or "maru" in name:
            add("hiragino_gothic")
            add("hiragino_maru")
            add("hiragino_mincho")
        if "mincho" in name:
            add("hiragino_mincho")
        if "helvetica" in name or "galvji" in name:
            add("helvetica_neue_bold" if text_block.is_bold else "helvetica_neue_regular")

        # Stable CJK+Latin fallback
        prefer_serif = any(token in name for token in ["mincho", "serif"])
        add("serif_cjk_bold" if text_block.is_bold and prefer_serif else "serif_cjk_regular" if prefer_serif else "")
        add("sans_cjk_bold" if text_block.is_bold and not prefer_serif else "sans_cjk_regular" if not prefer_serif else "")
        add("latin_sans_bold" if text_block.is_bold else "latin_sans_regular")

        # Drop empty keys that might have been added
        candidates = [c for c in candidates if c]
        return candidates

    def _sample_average_color(self, page, rect: fitz.Rect) -> Tuple[float, float, float]:
        """Sample average color of a region to approximate background fill."""
        try:
            pix = page.get_pixmap(clip=rect, matrix=fitz.Matrix(0.3, 0.3), alpha=False)
            data = pix.samples
            if not data:
                return (1, 1, 1)
            import array
            arr = array.array("B", data)
            total_pixels = len(arr) // 3
            if total_pixels == 0:
                return (1, 1, 1)
            r = sum(arr[0::3]) / (255 * total_pixels)
            g = sum(arr[1::3]) / (255 * total_pixels)
            b = sum(arr[2::3]) / (255 * total_pixels)
            return (r, g, b)
        except Exception:
            return (1, 1, 1)

    def _region_variance(self, page, rect: fitz.Rect) -> float:
        """Estimate variance of background to decide redaction vs. cover."""
        try:
            pix = page.get_pixmap(clip=rect, matrix=fitz.Matrix(0.2, 0.2), alpha=False)
            data = pix.samples
            if not data:
                return 0.0
            import array
            arr = array.array("B", data)
            total_pixels = len(arr) // 3
            if total_pixels == 0:
                return 0.0
            rs = arr[0::3]; gs = arr[1::3]; bs = arr[2::3]
            import math
            def var(channel):
                n = len(channel)
                mean = sum(channel)/n
                return sum((c-mean)**2 for c in channel)/(n*255*255)
            return (var(rs)+var(gs)+var(bs))/3
        except Exception:
            return 0.0

    def _cover_block_background(self, page, rect: fitz.Rect, font_size: float) -> None:
        """Cover original text region with a sampled-background rectangle to erase JP text."""
        pad = max(0.3, font_size * 0.2)
        padded = rect + (-pad, -pad, pad, pad)
        color = self._sample_average_color(page, padded)
        try:
            page.draw_rect(padded, color=color, fill=color, width=0)
        except Exception:
            try:
                page.draw_rect(padded, color=(1, 1, 1), fill=(1, 1, 1), width=0)
            except Exception:
                pass
        
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

        # Strip ellipses and placeholder markers often present in extracted spans
        text = text.replace('...', '').replace('…', '')
        text = text.replace('＃', '#')  # Normalize fullwidth hash
        text = text.replace('XXXXX', '')
        text = text.replace('《', '').replace('》', '')
        
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
        
        for idx, block in enumerate(text_blocks):
            translation, confidence, match_type = self.find_best_match(block.text, translations)
            
            if translation:
                normalized_key = self.create_stable_key(block.text)
                result: MatchResult = {
                    'text': block.text,
                    'block_index': idx,
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
            orig_len = max(1, len(original_text))
            trans_len = max(1, len(translated_text))
            expansion_ratio = trans_len / orig_len

            bbox_width = max(1.0, original_bbox[2] - original_bbox[0])
            # Estimate char width; slightly optimistic to avoid shrinking when not needed
            char_width = font_size * 0.55
            required_width = trans_len * char_width * 1.02
            width_scaling = bbox_width / required_width

            if width_scaling >= 1.0:
                return 1.0  # plenty of space, keep original size

            # Only shrink when needed; clamp to avoid unreadably small text
            if expansion_ratio > 1.6:
                min_scale = 0.7
            elif expansion_ratio > 1.3:
                min_scale = 0.75
            else:
                min_scale = 0.8
            optimal_scaling = max(width_scaling, min_scale)
            return optimal_scaling
        except Exception as e:
            logging.warning(f"Error calculating font scaling: {e}")
            return 1.0
    
    def _is_latin_text(self, text: str) -> bool:
        """Check if text is primarily Latin/ASCII."""
        return all(ord(c) < 128 for c in text.replace('\n', '').replace(' ', ''))

    def replace_text_in_block(self, page, text_block: TextBlock, 
                            translation: TranslationData) -> bool:
        """Replace text in a specific text block."""
        try:
            # Capture original font name (may be empty)
            original_font = text_block.font_name or ""

            translated_text = translation['translated']
            
            # Calculate font scaling if needed
            font_scaling = translation.get('font_scaling', 1.0)
            if font_scaling == 1.0:  # Auto-calculate if not provided
                font_scaling = self.calculate_optimal_font_scaling(
                    text_block.text, translated_text, text_block.bbox, text_block.font_size
                )
            
            # Apply font scaling
            new_font_size = text_block.font_size * font_scaling

            # Convert color to 0-1 floats for fitz
            color = (0, 0, 0)
            try:
                if isinstance(text_block.font_color, tuple) and len(text_block.font_color) == 3:
                    r, g, b = text_block.font_color
                    if max(text_block.font_color) > 1:
                        color = (r / 255.0, g / 255.0, b / 255.0)
                    else:
                        color = (r, g, b)
                # If the color is very light and we drew a dark overlay, keep as-is.
                # If the color is very light and overlay was light, shift to dark for contrast.
                brightness = sum(color) / 3.0
                if brightness > 0.85:
                    color = (0, 0, 0)
            except Exception:
                color = (0, 0, 0)

            # Force a stable, full-coverage font: prefer original family if installed,
            # then fall back to Noto CJK/Latin.
            font_paths = self._font_candidates(text_block)
            
            # Force Helvetica for pure Latin text to avoid CJK font issues
            is_latin = self._is_latin_text(translated_text)
            
            def pick_builtin(base: str, bold: bool, italic: bool) -> str:
                if base == "times":
                    if bold and italic:
                        return "times-bolditalic"
                    if bold:
                        return "times-bold"
                    if italic:
                        return "times-italic"
                    return "times-roman"
                if base == "cour":
                    if bold and italic:
                        return "cour-boldoblique"
                    if bold:
                        return "cour-bold"
                    if italic:
                        return "cour-oblique"
                    return "cour"
                # helv/default
                if bold and italic:
                    return "helv-boldoblique"
                if bold:
                    return "helv-bold"
                if italic:
                    return "helv-oblique"
                return "helv"

            # Built-in fallback
            font_name_lower = original_font.lower()
            if any(token in font_name_lower for token in ["cour", "mono", "code"]):
                builtin_base = "cour"
            elif any(token in font_name_lower for token in ["mincho", "serif", "koz"]):
                builtin_base = "times"
            else:
                builtin_base = "helv"
            builtin_font = pick_builtin(builtin_base, text_block.is_bold, text_block.is_italic)

            # Insert text within original bounding box to preserve wrapping
            rect = fitz.Rect(text_block.bbox)
            
            # CRITICAL FIX: Draw a white background rectangle to cover original text
            # This ensures readability even if redaction failed or was incomplete
            try:
                # Draw a slightly larger white box to cover potential anti-aliasing artifacts
                # but respect the clip to avoid overwriting neighbors
                bg_rect = rect + (-1, -1, 1, 1) 
                page.draw_rect(bg_rect, fill=(1, 1, 1), color=None, overlay=True)
            except Exception as e:
                logging.warning(f"Failed to draw background cover: {e}")

            inserted = False
            last_err = None
            chosen_font_name = None
            # Try with font files first (embed Noto CJK / Latin if available)
            chosen_font_path = None
            
            # If text is Latin, prioritize built-in Helvetica/Times for reliability
            if is_latin:
                try:
                    page.insert_textbox(
                        rect,
                        translated_text,
                        fontname=builtin_font,
                        fontsize=new_font_size,
                        color=color,
                        rotate=text_block.rotation,
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
                    inserted = True
                    chosen_font_name = builtin_font
                except Exception as e:
                    last_err = e
            
            if not inserted:
                for path in font_paths:
                    if not path:
                        continue
                    try:
                        page.insert_textbox(
                            rect,
                            translated_text,
                            fontfile=path,
                            fontsize=new_font_size,
                            color=color,
                            rotate=text_block.rotation,
                            align=fitz.TEXT_ALIGN_LEFT,
                        )
                        inserted = True
                        chosen_font_path = path
                        break
                    except Exception as e:
                        last_err = e
                        continue

            # Fallback to built-in font if not already tried or failed
            if not inserted:
                try:
                    page.insert_textbox(
                        rect,
                        translated_text,
                        fontname=builtin_font,
                        fontsize=new_font_size,
                        color=color,
                        rotate=text_block.rotation,
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
                    inserted = True
                    chosen_font_name = builtin_font
                except Exception as e:
                    last_err = e

            if not inserted:
                logging.error(f"Text insertion failed for block {text_block.block_id}: {last_err}")
                return False

            logging.debug(
                f"Replaced text in block {text_block.block_id}: '{text_block.text}' -> '{translated_text}' "
                f"(scaling: {font_scaling:.2f}, font: {chosen_font_path or chosen_font_name})"
            )
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
    
    def preserve_formatting(self, original_pdf_path: str, translated_pdf_path: str) -> bool:
        """
        Preserve non-text elements and overall PDF structure.

        The previous implementation attempted to copy images/links directly and
        could throw xref errors on some PDFs. Here we take a safer path:
        - Keep the translated document as-is (images and vector content are already
          present because we edited the same document).
        - Simply save/rename the temp output to the final path with clean flags.
        """
        if fitz is None:
            return False

        try:
            doc = fitz.open(translated_pdf_path)
            doc.save(self.output_path, garbage=4, deflate=True, clean=True)
            doc.close()
            logging.info("Saved translated PDF with clean/deflate; skipped risky xref copies")
            return True
        except Exception as e:
            logging.warning(f"Could not finalize formatting save: {e}")
            return False
    
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
            outline = getattr(doc, "get_outline", None)
            if outline:
                try:
                    ol = outline()
                    if ol:
                        logging.debug(f"Preserving {len(ol)} outline entries")
                except Exception:
                    pass
            
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

        # First pass: add redactions for all matched blocks (remove JP text, keep backgrounds/images).
        redaction_map: dict[int, list[fitz.Rect]] = {}
        for result in match_results:
            block_idx = result.get('block_index')
            if block_idx is None or block_idx >= len(self.text_blocks):
                continue
            block = self.text_blocks[block_idx]
            rect = fitz.Rect(block.bbox)
            pad = max(0.5, block.font_size * 0.2)
            rect_padded = rect + (-pad, -pad, pad, pad)
            redaction_map.setdefault(block.page_num, []).append(rect_padded)

        for page_num, rects in redaction_map.items():
            page = self.doc[page_num]
            for rect in rects:
                try:
                    # Transparent fill: remove text only, keep background intact
                    page.add_redact_annot(rect, fill=None, text=None, cross_out=False)
                except Exception as e:
                    logging.debug(f"Redact annot failed on page {page_num + 1}: {e}")
            try:
                # Keep images/drawings; remove text.
                page.apply_redactions(images=getattr(fitz, "PDF_REDACT_IMAGE_NONE", 0))
            except Exception as e:
                logging.warning(f"Redaction apply failed on page {page_num + 1}: {e}")

        # Process matched text blocks
        for result in match_results:
            # Use the exact block index to avoid reusing the same block when text repeats
            block_idx = result.get('block_index')
            if block_idx is None or block_idx >= len(self.text_blocks):
                continue
            text_block = self.text_blocks[block_idx]
            
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
        
        # Preserve formatting from original (safe finalize)
        preserve_ok = self.preserve_formatting(self.input_path, temp_output)

        if not preserve_ok:
            try:
                shutil.copy(temp_output, self.output_path)
                logging.warning("Preserve formatting step failed; fell back to temp output")
            except Exception as e:
                logging.error(f"Failed to create output PDF fallback: {e}")
        
        # Handle special elements
        self.handle_special_elements(self.output_path)
        
        # Clean up temporary file
        if os.path.exists(temp_output):
            os.remove(temp_output)
        
        # Log detailed statistics
        self.log_match_quality_report()
        logging.info(
            "Replacement stats: replaced=%s failed=%s layout_adj=%s",
            self.replacement_stats.get('replaced_blocks'),
            self.replacement_stats.get('failed_blocks'),
            self.replacement_stats.get('layout_adjustments'),
        )
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

    # Strip ellipses and placeholder markers
    text = text.replace('...', '').replace('…', '')
    text = text.replace('＃', '#')
    text = text.replace('XXXXX', '')
    text = text.replace('《', '').replace('》', '')
    
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
