"""
Unit tests for apply_pdf_translation.py

Test suite for PDF back-projector functionality including:
- Text block extraction and parsing
- Translation matching and replacement
- Font scaling calculations
- Formatting preservation
- Error handling and edge cases

Test framework: pytest
- Uses fixtures and mocking for PDF operations
- Tests core logic without requiring actual PDF files
"""

import json
import os
import tempfile
import unittest
from unittest.mock import Mock, patch, MagicMock
from pathlib import Path
from typing import Dict, Any

import pytest

# Import the module under test
import sys
import os
sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

# Set environment variable to allow import without PyMuPDF
os.environ["PYTEST_CURRENT_TEST"] = "1"

from apply_pdf_translation import (
    PDFBackProjector, TextBlock, TranslationData, JP_ANY, 
    calculate_optimal_font_scaling_standalone, find_best_translation_match_standalone
)

# Mock PyMuPDF (fitz) for testing
class MockPage:
    def __init__(self, page_num: int):
        self.page_num = page_num
        self.redactions = []
        self.text_insertions = []
    
    def add_redact_annot(self, bbox):
        self.redactions.append(bbox)
    
    def apply_redactions(self):
        pass  # Mock implementation
    
    def insert_text(self, **kwargs):
        self.text_insertions.append(kwargs)
        return Mock()
    
    def get_text(self, format_type):
        # Mock text extraction data
        return {
            "blocks": [
                {
                    "lines": [
                        {
                            "spans": [
                                {
                                    "text": "こんにちは世界",
                                    "bbox": (100, 100, 200, 120),
                                    "font": "Arial",
                                    "size": 12.0,
                                    "color": (0, 0, 0),
                                    "rotation": 0.0,
                                    "line_height": 1.0,
                                    "char_spacing": 0.0
                                }
                            ]
                        }
                    ]
                }
            ]
        }

class MockDocument:
    def __init__(self, num_pages: int = 3):
        self.num_pages = num_pages
        self.pages = [MockPage(i) for i in range(num_pages)]
    
    def __len__(self):
        return self.num_pages
    
    def __getitem__(self, index):
        return self.pages[index]
    
    def save(self, path, **kwargs):
        pass  # Mock implementation
    
    def close(self):
        pass  # Mock implementation
    
    def get_outline(self):
        return []
    
    @property
    def metadata(self):
        return {"title": "Test Document", "author": "Test Author"}

# Test fixtures
@pytest.fixture
def sample_translations():
    """Sample translation data for testing."""
    return {
        "こんにちは世界": {
            "translated": "Hello World",
            "font_scaling": 0.9
        },
        "日本語のテキスト": {
            "translated": "Japanese text",
            "layout_adjustments": {"position_offset": {"x": 0, "y": 0}}
        },
        "これはテストです": {
            "translated": "This is a test"
        }
    }

@pytest.fixture
def sample_text_blocks():
    """Sample text blocks for testing."""
    return [
        TextBlock(
            page_num=0,
            bbox=(100, 100, 200, 120),
            text="こんにちは世界",
            font_name="Arial",
            font_size=12.0,
            font_color=(0, 0, 0),
            is_bold=False,
            is_italic=False,
            block_id="page_0_block_0"
        ),
        TextBlock(
            page_num=1,
            bbox=(150, 150, 300, 170),
            text="日本語のテキスト",
            font_name="Helvetica-Bold",
            font_size=14.0,
            font_color=(0.5, 0.5, 0.5),
            is_bold=True,
            is_italic=False,
            block_id="page_1_block_1"
        ),
        TextBlock(
            page_num=2,
            bbox=(200, 200, 350, 220),
            text="これはテストです",
            font_name="Times-Italic",
            font_size=16.0,
            font_color=(1, 0, 0),
            is_bold=False,
            is_italic=True,
            block_id="page_2_block_2"
        )
    ]

@pytest.fixture
def mock_fitz():
    """Mock fitz module."""
    mock_fitz = MagicMock()
    mock_fitz.TEXT_BOLD = 1
    mock_fitz.TEXT_ITALIC = 2
    mock_fitz.TEXT_ROTATE_ON_UPSIDE_DOWN = 0
    return mock_fitz

# Test classes
class TestTextBlock:
    """Test TextBlock dataclass functionality."""
    
    def test_text_block_creation(self):
        """Test TextBlock creation with all fields."""
        block = TextBlock(
            page_num=1,
            bbox=(0, 0, 100, 50),
            text="テスト",
            font_name="Arial",
            font_size=12.0,
            font_color=(0, 0, 0),
            is_bold=False,
            is_italic=False,
            block_id="test_block"
        )
        
        assert block.page_num == 1
        assert block.bbox == (0, 0, 100, 50)
        assert block.text == "テスト"
        assert block.font_name == "Arial"
        assert block.font_size == 12.0
        assert block.font_color == (0, 0, 0)
        assert block.is_bold is False
        assert block.is_italic is False
        assert block.block_id == "test_block"
    
    def test_text_block_defaults(self):
        """Test TextBlock with default values."""
        block = TextBlock(
            page_num=0,
            bbox=(0, 0, 50, 25),
            text="デフォルト",
            font_name="DefaultFont",
            font_size=10.0,
            font_color=(0.5, 0.5, 0.5),
            is_bold=True,
            is_italic=True,
            block_id="default_block"
        )
        
        # Check default values
        assert block.rotation == 0.0
        assert block.line_height == 1.0
        assert block.char_spacing == 0.0

class TestPDFBackProjector:
    """Test PDFBackProjector class functionality."""
    
    @patch('apply_pdf_translation.fitz')
    def test_initialization(self, mock_fitz, sample_translations):
        """Test PDFBackProjector initialization."""
        projector = PDFBackProjector("input.pdf", "output.pdf", "translations.json")
        
        assert projector.input_path == "input.pdf"
        assert projector.output_path == "output.pdf"
        assert projector.translations_path == "translations.json"
        assert projector.doc is None
        assert projector.translations == {}
        assert projector.text_blocks == []
        assert projector.replacement_stats['total_blocks'] == 0
    
    def test_load_translations_list_format(self, sample_translations):
        """Test loading translations in list format."""
        # Convert dict to list format
        translations_list = []
        for original, data in sample_translations.items():
            translation_data = {"original": original, **data}
            translations_list.append(translation_data)
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(translations_list, f, ensure_ascii=False)
            temp_path = f.name
        
        try:
            projector = PDFBackProjector("input.pdf", "output.pdf", temp_path)
            projector.load_translations()
            
            assert len(projector.translations) == 3
            assert "こんにちは世界" in projector.translations
            assert projector.translations["こんにちは世界"]["translated"] == "Hello World"
        finally:
            os.unlink(temp_path)
    
    def test_load_translations_dict_format(self, sample_translations):
        """Test loading translations in dictionary format."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(sample_translations, f, ensure_ascii=False)
            temp_path = f.name
        
        try:
            projector = PDFBackProjector("input.pdf", "output.pdf", temp_path)
            projector.load_translations()
            
            assert len(projector.translations) == 3
            assert "こんにちは世界" in projector.translations
            assert projector.translations["こんにちは世界"]["translated"] == "Hello World"
        finally:
            os.unlink(temp_path)
    
    @patch('apply_pdf_translation.fitz')
    def test_load_translations_file_error(self, mock_fitz):
        """Test handling of missing translations file."""
        projector = PDFBackProjector("input.pdf", "output.pdf", "nonexistent.json")
        
        with pytest.raises(FileNotFoundError):
            projector.load_translations()
    
    @patch('apply_pdf_translation.fitz')
    def test_extract_text_blocks(self, mock_fitz, sample_text_blocks):
        """Test text block extraction from PDF."""
        # Mock document with text data
        mock_doc = MockDocument(1)  # Single page document
        mock_fitz.open.return_value = mock_doc
        
        projector = PDFBackProjector("input.pdf", "output.pdf", "translations.json")
        projector.extract_text_blocks()
        
        assert len(projector.text_blocks) == 1
        assert projector.text_blocks[0].text == "こんにちは世界"
        assert projector.replacement_stats['total_blocks'] == 1
    
    def test_find_best_translation_match_exact(self, sample_translations):
        """Test exact translation matching."""
        projector = PDFBackProjector("input.pdf", "output.pdf", "translations.json")
        projector.translations = sample_translations
        
        translation = projector.find_best_translation_match("こんにちは世界")
        
        assert translation is not None
        assert translation["translated"] == "Hello World"
    
    def test_find_best_translation_match_partial(self, sample_translations):
        """Test partial translation matching."""
        projector = PDFBackProjector("input.pdf", "output.pdf", "translations.json")
        projector.translations = sample_translations
        
        # Test with extra whitespace
        translation = projector.find_best_translation_match("  こんにちは世界  ")
        
        assert translation is not None
        assert translation["translated"] == "Hello World"
    
    def test_find_best_translation_match_no_match(self, sample_translations):
        """Test case where no translation is found."""
        projector = PDFBackProjector("input.pdf", "output.pdf", "translations.json")
        projector.translations = sample_translations
        
        translation = projector.find_best_translation_match("未知のテキスト")
        
        assert translation is None
    
    def test_calculate_optimal_font_scaling(self):
        """Test font scaling calculations."""
        # Test normal expansion
        scaling = calculate_optimal_font_scaling_standalone(
            "短い", "longer text that expands significantly", 
            (0, 0, 100, 20), 12.0
        )
        assert scaling < 1.0
        
        # Test contraction
        scaling = calculate_optimal_font_scaling_standalone(
            "非常に長い日本語のテキストです", "short", 
            (0, 0, 200, 20), 12.0
        )
        assert scaling == 1.0  # No scaling needed for contraction
        
        # Test minimum scaling
        scaling = calculate_optimal_font_scaling_standalone(
            "短い", "extremely long text that would require very small font", 
            (0, 0, 50, 10), 12.0
        )
        assert scaling >= 0.7  # Minimum 70% scaling
    
    @patch('apply_pdf_translation.fitz')
    def test_replace_text_in_block_success(self, mock_fitz):
        """Test successful text replacement in a block."""
        mock_page = MockPage(0)
        
        text_block = TextBlock(
            page_num=0,
            bbox=(100, 100, 200, 120),
            text="こんにちは",
            font_name="Arial",
            font_size=12.0,
            font_color=(0, 0, 0),
            is_bold=False,
            is_italic=False,
            block_id="test_block"
        )
        
        translation = {"translated": "Hello", "font_scaling": 0.9}
        
        projector = PDFBackProjector("input.pdf", "output.pdf", "translations.json")
        
        success = projector.replace_text_in_block(mock_page, text_block, translation)
        
        assert success is True
        assert len(mock_page.redactions) == 1
        assert len(mock_page.text_insertions) == 1
    
    @patch('apply_pdf_translation.fitz')
    def test_replace_text_in_block_failure(self, mock_fitz):
        """Test text replacement failure."""
        mock_page = MockPage(0)
        mock_page.add_redact_annot = MagicMock(side_effect=Exception("Mock error"))
        
        text_block = TextBlock(
            page_num=0,
            bbox=(100, 100, 200, 120),
            text="こんにちは",
            font_name="Arial",
            font_size=12.0,
            font_color=(0, 0, 0),
            is_bold=False,
            is_italic=False,
            block_id="test_block"
        )
        
        translation = {"translated": "Hello"}
        
        projector = PDFBackProjector("input.pdf", "output.pdf", "translations.json")
        
        success = projector.replace_text_in_block(mock_page, text_block, translation)
        
        assert success is False
    
    def test_apply_layout_adjustments(self):
        """Test layout adjustments application."""
        mock_page = MockPage(0)
        
        text_block = TextBlock(
            page_num=0,
            bbox=(100, 100, 200, 120),
            text="こんにちは",
            font_name="Arial",
            font_size=12.0,
            font_color=(0, 0, 0),
            is_bold=False,
            is_italic=False,
            block_id="test_block"
        )
        
        translation = {
            "translated": "Hello",
            "layout_adjustments": {"position_offset": {"x": 10, "y": 5}}
        }
        
        projector = PDFBackProjector("input.pdf", "output.pdf", "translations.json")
        projector.apply_layout_adjustments(mock_page, text_block, translation)
        
        assert projector.replacement_stats['layout_adjustments'] == 1

class TestUtilityFunctions:
    """Test utility functions."""
    
    def test_japanese_text_detection(self):
        """Test Japanese text detection regex."""
        # Test Japanese text
        assert JP_ANY.search("こんにちは") is not None
        assert JP_ANY.search("日本語") is not None
        assert JP_ANY.search("漢字") is not None
        
        # Test English text
        assert JP_ANY.search("Hello") is None
        assert JP_ANY.search("English text") is None
        
        # Test mixed text
        assert JP_ANY.search("Hello こんにちは World") is not None
    
    def test_translation_data_structure(self):
        """Test TranslationData typed dict structure."""
        # Test minimal translation data
        translation: TranslationData = {
            "original": "こんにちは",
            "translated": "Hello"
        }
        assert translation["original"] == "こんにちは"
        assert translation["translated"] == "Hello"
        
        # Test complete translation data
        complete_translation: TranslationData = {
            "original": "こんにちは",
            "translated": "Hello",
            "font_scaling": 0.9,
            "layout_adjustments": {"position": {"x": 0, "y": 0}}
        }
        assert complete_translation["font_scaling"] == 0.9
        assert "layout_adjustments" in complete_translation

class TestIntegration:
    """Integration tests for the PDF back-projector."""
    
    @patch('apply_pdf_translation.fitz')
    def test_full_processing_workflow(self, mock_fitz, sample_translations, sample_text_blocks):
        """Test complete document processing workflow."""
        # Mock document
        mock_doc = MockDocument(3)
        mock_fitz.open.return_value = mock_doc
        
        # Create temporary translations file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(sample_translations, f, ensure_ascii=False)
            translations_path = f.name
        
        try:
            # Initialize projector
            projector = PDFBackProjector("input.pdf", "output.pdf", translations_path)
            
            # Mock text extraction
            projector.text_blocks = sample_text_blocks
            
            # Process document
            projector.process_document()
            
            # Check statistics
            assert projector.replacement_stats['total_blocks'] == 3
            assert projector.replacement_stats['replaced_blocks'] == 3
            assert projector.replacement_stats['failed_blocks'] == 0
            
        finally:
            os.unlink(translations_path)
    
    def test_error_handling(self):
        """Test error handling in various scenarios."""
        projector = PDFBackProjector("nonexistent.pdf", "output.pdf", "nonexistent.json")
        
        # Test with non-existent input file - should fail at file operations
        with pytest.raises((FileNotFoundError, ImportError)):
            projector.process_document()

# Main test execution
if __name__ == "__main__":
    pytest.main([__file__, "-v"])