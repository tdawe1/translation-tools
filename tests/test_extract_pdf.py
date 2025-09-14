#!/usr/bin/env python3
"""
Unit tests for PDF text extraction functionality.

Tests the extract_pdf.py script to ensure it correctly extracts Japanese text
with layout preservation and handles various edge cases.
"""

import json
import os
import tempfile
import unittest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Add the scripts directory to Python path
sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

from extract_pdf import (
    PDFExtractor, TextBlock, PageInfo, ExtractionResult, BlockType,
    save_extraction_result, JP_ANY
)

class TestTextBlock(unittest.TestCase):
    """Test TextBlock dataclass and functionality."""
    
    def test_text_block_creation(self):
        """Test basic TextBlock creation."""
        block = TextBlock(
            id="test_block",
            page=0,
            text="テストテキスト",
            x0=100.0,
            y0=200.0,
            x1=300.0,
            y1=250.0,
            font_size=12.0,
            font_name="Arial"
        )
        
        self.assertEqual(block.id, "test_block")
        self.assertEqual(block.page, 0)
        self.assertEqual(block.text, "テストテキスト")
        self.assertEqual(block.x0, 100.0)
        self.assertEqual(block.block_type, "body")
        self.assertFalse(block.is_vertical)
        self.assertEqual(block.confidence, 1.0)
    
    def test_text_block_with_defaults(self):
        """Test TextBlock with default values."""
        block = TextBlock(
            id="test_block",
            page=0,
            text="テスト",
            x0=0,
            y0=0,
            x1=100,
            y1=50,
            font_size=10,
            font_name="Helvetica"
        )
        
        # Check default values
        self.assertEqual(block.block_type, "body")
        self.assertFalse(block.is_vertical)
        self.assertEqual(block.rotation, 0.0)
        self.assertEqual(block.reading_order, 0)
        self.assertEqual(block.confidence, 1.0)
        self.assertEqual(block.language, "ja")
        self.assertEqual(block.metadata, {})

class TestPDFExtractor(unittest.TestCase):
    """Test PDFExtractor class functionality."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.extractor = PDFExtractor(use_fallback=True, min_confidence=0.8)
    
    def test_extractor_initialization(self):
        """Test PDFExtractor initialization."""
        self.assertTrue(self.extractor.use_fallback)
        self.assertEqual(self.extractor.min_confidence, 0.8)
        self.assertEqual(self.extractor.stats['total_pages'], 0)
    
    @patch('extract_pdf.fitz.open')
    def test_extract_nonexistent_file(self, mock_fitz_open):
        """Test extraction with non-existent file."""
        with self.assertRaises(FileNotFoundError):
            self.extractor.extract_text_blocks("/nonexistent/file.pdf")
    
    def test_is_vertical_text_detection(self):
        """Test vertical text detection logic."""
        # Test with explicit vertical flag
        span = {"vertical": True}
        self.assertTrue(self.extractor._is_vertical_text(span))
        
        # Test with rotation
        span = {"rotate": 90}
        self.assertTrue(self.extractor._is_vertical_text(span))
        
        span = {"rotate": 270}
        self.assertTrue(self.extractor._is_vertical_text(span))
        
        # Test with font name hints
        span = {"font": "mincho-vertical", "rotate": 0}
        self.assertTrue(self.extractor._is_vertical_text(span))
        
        # Test horizontal text
        span = {"font": "Arial", "rotate": 0}
        self.assertFalse(self.extractor._is_vertical_text(span))
    
    def test_block_type_classification(self):
        """Test block type classification."""
        page_rect = (0, 0, 600, 800)
        
        # Header detection (top of page)
        bbox = (0, 0, 600, 50)
        text = "Chapter 1"
        block_type = self.extractor._classify_block_type(text, bbox, page_rect)
        self.assertEqual(block_type, "header")
        
        # Footer detection (bottom of page)
        bbox = (0, 750, 600, 800)
        text = "Page 1"
        block_type = self.extractor._classify_block_type(text, bbox, page_rect)
        self.assertEqual(block_type, "footer")
        
        # Table detection
        bbox = (100, 200, 500, 400)
        text = "Col1\tCol2\tCol3"
        block_type = self.extractor._classify_block_type(text, bbox, page_rect)
        self.assertEqual(block_type, "table")
        
        # Body text (default)
        bbox = (100, 200, 500, 250)
        text = "This is regular body text content."
        block_type = self.extractor._classify_block_type(text, bbox, page_rect)
        self.assertEqual(block_type, "body")
    
    def test_confidence_calculation(self):
        """Test confidence score calculation."""
        # High confidence: Japanese text, good font size
        span = {"size": 12.0, "rotate": 0}
        text = "日本語のテスト"
        confidence = self.extractor._calculate_confidence(span, text)
        self.assertGreater(confidence, 0.8)
        
        # Low confidence: small text
        span = {"size": 6.0, "rotate": 0}
        text = "small text"
        confidence = self.extractor._calculate_confidence(span, text)
        self.assertLess(confidence, 0.8)
        
        # Medium confidence: rotated text
        span = {"size": 12.0, "rotate": 45}
        text = "日本語"
        confidence = self.extractor._calculate_confidence(span, text)
        self.assertLess(confidence, 1.0)
    
    def test_point_in_bbox(self):
        """Test point-in-bounding-box detection."""
        bbox = (100, 200, 300, 400)
        
        # Point inside
        self.assertTrue(self.extractor._point_in_bbox((150, 250), bbox))
        
        # Point outside
        self.assertFalse(self.extractor._point_in_bbox((50, 50), bbox))
        
        # Point on boundary
        self.assertTrue(self.extractor._point_in_bbox((100, 200), bbox))
    
    def test_reading_order_sorting(self):
        """Test text block sorting by reading order."""
        blocks = [
            TextBlock("block3", 0, "Text 3", 100, 200, 200, 250, 12, "Arial", reading_order=2),
            TextBlock("block1", 0, "Text 1", 100, 100, 200, 150, 12, "Arial", reading_order=0),
            TextBlock("block2", 0, "Text 2", 100, 150, 200, 200, 12, "Arial", reading_order=1),
        ]
        
        sorted_blocks = self.extractor._sort_by_reading_order(blocks)
        
        self.assertEqual(sorted_blocks[0].id, "block1")
        self.assertEqual(sorted_blocks[1].id, "block2")
        self.assertEqual(sorted_blocks[2].id, "block3")
    
    def test_filter_japanese_text(self):
        """Test filtering to Japanese text only."""
        # Create mock extraction result
        page1 = PageInfo(
            page_num=0,
            width=600,
            height=800,
            rotation=0,
            text_blocks=[
                TextBlock("jp_block", 0, "日本語", 0, 0, 100, 50, 12, "Arial"),
                TextBlock("en_block", 0, "English", 0, 50, 100, 100, 12, "Arial"),
            ],
            has_japanese=True,
            extraction_method="fitz"
        )
        
        result = ExtractionResult(
            filename="test.pdf",
            pages=[page1],
            total_blocks=2,
            total_japanese_blocks=1,
            extraction_time=1.0,
            extraction_methods={"fitz": 1},
            metadata={}
        )
        
        filtered_result = self.extractor.filter_japanese_text(result)
        
        self.assertEqual(filtered_result.total_blocks, 1)
        self.assertEqual(filtered_result.total_japanese_blocks, 1)
        self.assertEqual(len(filtered_result.pages[0].text_blocks), 1)
        self.assertEqual(filtered_result.pages[0].text_blocks[0].id, "jp_block")
    
    def test_to_translation_format(self):
        """Test conversion to translation pipeline format."""
        # Create mock extraction result
        page = PageInfo(
            page_num=0,
            width=600,
            height=800,
            rotation=0,
            text_blocks=[
                TextBlock("jp_block", 0, "日本語テキスト", 100, 200, 300, 250, 12, "Arial"),
                TextBlock("en_block", 0, "English text", 100, 250, 300, 300, 12, "Arial"),
            ],
            has_japanese=True,
            extraction_method="fitz"
        )
        
        result = ExtractionResult(
            filename="test.pdf",
            pages=[page],
            total_blocks=2,
            total_japanese_blocks=1,
            extraction_time=1.0,
            extraction_methods={"fitz": 1},
            metadata={}
        )
        
        translation_data = self.extractor.to_translation_format(result)
        
        self.assertIn("japanese_texts", translation_data)
        self.assertIn("text_mapping", translation_data)
        self.assertIn("layout_info", translation_data)
        
        # Check that only Japanese text is included
        self.assertEqual(len(translation_data["japanese_texts"]), 1)
        self.assertEqual(translation_data["japanese_texts"][0], "日本語テキスト")
        
        # Check text mapping
        self.assertIn("日本語テキスト", translation_data["text_mapping"])

class TestIntegration(unittest.TestCase):
    """Integration tests with mock PDF data."""
    
    @patch('extract_pdf.fitz.open')
    def test_extract_with_mock_pdf(self, mock_fitz_open):
        """Test extraction with mocked PDF data."""
        # Mock PDF document
        mock_doc = Mock()
        mock_doc.__len__ = Mock(return_value=1)
        
        # Mock page
        mock_page = Mock()
        mock_page.rect = Mock()
        mock_page.rect.width = 600
        mock_page.rect.height = 800
        mock_page.rotation = 0
        
        # Mock text extraction result
        mock_text_dict = {
            "blocks": [
                {
                    "bbox": [0, 0, 600, 800],
                    "lines": [
                        {
                            "bbox": [100, 200, 500, 250],
                            "spans": [
                                {
                                    "text": "日本語のテスト",
                                    "bbox": [100, 200, 500, 250],
                                    "font": "Arial",
                                    "size": 12.0,
                                    "rotate": 0
                                }
                            ]
                        }
                    ]
                }
            ]
        }
        
        mock_page.get_text = Mock(return_value=mock_text_dict)
        mock_doc.__getitem__ = Mock(return_value=mock_page)
        
        mock_fitz_open.return_value = mock_doc
        
        # Test extraction
        extractor = PDFExtractor()
        result = extractor.extract_text_blocks("test.pdf")
        
        # Verify results
        self.assertEqual(len(result.pages), 1)
        self.assertEqual(result.total_blocks, 1)
        self.assertEqual(result.total_japanese_blocks, 1)
        self.assertTrue(result.pages[0].has_japanese)

class TestUtilityFunctions(unittest.TestCase):
    """Test utility functions."""
    
    def test_japanese_text_detection(self):
        """Test Japanese text regex pattern."""
        # Japanese text
        self.assertTrue(JP_ANY.search("日本語のテキスト"))
        self.assertTrue(JP_ANY.search("漢字"))
        self.assertTrue(JP_ANY.search("ひらがな"))
        self.assertTrue(JP_ANY.search("カタカナ"))
        
        # English text
        self.assertFalse(JP_ANY.search("English text"))
        self.assertFalse(JP_ANY.search("123 numbers"))
        
        # Mixed text
        self.assertTrue(JP_ANY.search("日本語とEnglish"))
    
    @patch('builtins.open', create=True)
    @patch('json.dump')
    def test_save_extraction_result_json(self, mock_json_dump, mock_open):
        """Test saving extraction result as JSON."""
        # Create mock result
        result = ExtractionResult(
            filename="test.pdf",
            pages=[],
            total_blocks=0,
            total_japanese_blocks=0,
            extraction_time=1.0,
            extraction_methods={},
            metadata={}
        )
        
        save_extraction_result(result, "output.json", "json")
        
        mock_open.assert_called_once()
        mock_json_dump.assert_called_once()
    
    @patch('builtins.open', create=True)
    @patch('csv.writer')
    def test_save_extraction_result_csv(self, mock_csv_writer, mock_open):
        """Test saving extraction result as CSV."""
        # Create mock result with data
        page = PageInfo(
            page_num=0,
            width=600,
            height=800,
            rotation=0,
            text_blocks=[
                TextBlock("test_block", 0, "テスト", 100, 200, 200, 250, 12, "Arial")
            ],
            has_japanese=True,
            extraction_method="fitz"
        )
        
        result = ExtractionResult(
            filename="test.pdf",
            pages=[page],
            total_blocks=1,
            total_japanese_blocks=1,
            extraction_time=1.0,
            extraction_methods={"fitz": 1},
            metadata={}
        )
        
        save_extraction_result(result, "output.csv", "csv")
        
        mock_open.assert_called_once()
        mock_csv_writer.assert_called_once()

class TestErrorHandling(unittest.TestCase):
    """Test error handling and edge cases."""
    
    def test_empty_pdf_handling(self):
        """Test handling of empty PDF files."""
        # This would test the extractor's behavior with empty PDFs
        # In a real scenario, you might create a temporary empty PDF file
        pass
    
    def test_corrupted_pdf_handling(self):
        """Test handling of corrupted PDF files."""
        # This would test error handling when PDF cannot be parsed
        pass
    
    def test_unicode_text_handling(self):
        """Test handling of Unicode and special characters."""
        # Test with various Unicode characters
        test_texts = [
            "日本語",  # Japanese
            "中文",   # Chinese
            "한국어",  # Korean
            "English with emoji 📎",
            "Mixed 日本語とEnglish",
            "Text with\nnewlines",
            "Text with\ttabs"
        ]
        
        for text in test_texts:
            # Create text block and verify it handles Unicode properly
            block = TextBlock(
                id="test",
                page=0,
                text=text,
                x0=0, y0=0, x1=100, y1=50,
                font_size=12,
                font_name="Arial"
            )
            self.assertEqual(block.text, text)

if __name__ == '__main__':
    # Run tests
    unittest.main(verbosity=2)