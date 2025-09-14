#!/usr/bin/env python3
"""
Unit tests for PDF translation orchestration logic.

Tests the main PDFTranslationOrchestrator class and its core functionality.
"""

import unittest
import json
import os
import tempfile
import shutil
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime

# Import the orchestrator
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Mock the components before importing
sys.modules['extract_pdf'] = Mock()
sys.modules['pdf_layout_engine'] = Mock()
sys.modules['apply_pdf_translation'] = Mock()
sys.modules['audit_pdf'] = Mock()
sys.modules['translate_pptx_inplace'] = Mock()

from scripts.translate_pdf import PDFTranslationOrchestrator


class MockExtractionResult:
    """Mock extraction result for testing."""
    def __init__(self, total_blocks=10, total_japanese_blocks=8):
        self.filename = "test.pdf"
        self.pages = []
        self.total_blocks = total_blocks
        self.total_japanese_blocks = total_japanese_blocks
        self.extraction_time = 1.0
        self.extraction_methods = ["mock"]
        self.metadata = {}


class MockPage:
    """Mock page for testing."""
    def __init__(self, page_num=0, blocks=None):
        self.page_num = page_num
        self.text_blocks = blocks or []


class MockTextBlock:
    """Mock text block for testing."""
    def __init__(self, text="テスト", block_type="body", font_size=12.0):
        self.id = f"block_{block_type}_0"
        self.text = text
        self.block_type = block_type
        self.font_size = font_size
        self.font_name = "Arial"
        self.line_height = 14.0
        self.char_spacing = 0.0
        self.x0 = 100
        self.y0 = 100
        self.x1 = 200
        self.y1 = 150


class TestPDFTranslationOrchestrator(unittest.TestCase):
    """Test cases for PDFTranslationOrchestrator."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.input_pdf = os.path.join(self.temp_dir, "test.pdf")
        self.output_pdf = os.path.join(self.temp_dir, "test_out.pdf")
        self.cache_file = os.path.join(self.temp_dir, "cache.json")
        self.glossary_file = os.path.join(self.temp_dir, "glossary.json")
        
        # Create dummy input file
        with open(self.input_pdf, 'w') as f:
            f.write("dummy pdf content")
        
        # Initialize orchestrator
        self.orchestrator = PDFTranslationOrchestrator(
            input_path=self.input_pdf,
            output_path=self.output_pdf,
            cache_file=self.cache_file,
            glossary_file=None,
            offline=True,  # Use offline mode for testing
            cache_only=False
        )
    
    def tearDown(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.temp_dir)
    
    def test_init_basic(self):
        """Test basic initialization."""
        self.assertEqual(self.orchestrator.input_path, self.input_pdf)
        self.assertEqual(self.orchestrator.output_path, self.output_pdf)
        self.assertEqual(self.orchestrator.cache_file, self.cache_file)
        self.assertIsInstance(self.orchestrator.cache, dict)
        self.assertIsInstance(self.orchestrator.glossary, dict)
        self.assertEqual(len(self.orchestrator.cache), 0)
        self.assertEqual(len(self.orchestrator.glossary), 0)
    
    def test_load_cache_empty(self):
        """Test loading cache when file doesn't exist."""
        cache = self.orchestrator._load_cache()
        self.assertEqual(cache, {})
    
    def test_load_cache_existing(self):
        """Test loading existing cache file."""
        # Create cache file
        cache_data = {"テスト": "Test", "こんにちは": "Hello"}
        with open(self.cache_file, 'w', encoding='utf-8') as f:
            json.dump(cache_data, f)
        
        # Load cache
        cache = self.orchestrator._load_cache()
        self.assertEqual(cache, cache_data)
    
    def test_load_glossary_empty(self):
        """Test loading glossary when file doesn't exist."""
        glossary = self.orchestrator._load_glossary()
        self.assertEqual(glossary, {})
    
    def test_load_glossary_dict_format(self):
        """Test loading glossary in dict format."""
        glossary_data = {"テスト": "Test", "こんにちは": "Hello"}
        with open(self.glossary_file, 'w', encoding='utf-8') as f:
            json.dump(glossary_data, f)
        
        # Set glossary file path
        self.orchestrator.glossary_file = self.glossary_file
        glossary = self.orchestrator._load_glossary()
        self.assertEqual(glossary, glossary_data)
    
    def test_load_glossary_list_format(self):
        """Test loading glossary in list format."""
        glossary_data = [
            {"original": "テスト", "translated": "Test"},
            {"original": "こんにちは", "translated": "Hello"}
        ]
        with open(self.glossary_file, 'w', encoding='utf-8') as f:
            json.dump(glossary_data, f)
        
        # Set glossary file path
        self.orchestrator.glossary_file = self.glossary_file
        glossary = self.orchestrator._load_glossary()
        expected = {"テスト": "Test", "こんにちは": "Hello"}
        self.assertEqual(glossary, expected)
    
    def test_save_cache(self):
        """Test saving cache to file."""
        self.orchestrator.cache = {"テスト": "Test", "こんにちは": "Hello"}
        self.orchestrator._save_cache()
        
        # Verify cache file was created
        self.assertTrue(os.path.exists(self.cache_file))
        
        # Verify cache content
        with open(self.cache_file, 'r', encoding='utf-8') as f:
            saved_cache = json.load(f)
        self.assertEqual(saved_cache, self.orchestrator.cache)
    
    def test_parse_page_range_none(self):
        """Test parsing None page range."""
        result = self.orchestrator._parse_page_range(None)
        self.assertEqual(result, (0, -1))
    
    def test_parse_page_range_single(self):
        """Test parsing single page range."""
        result = self.orchestrator._parse_page_range("5")
        self.assertEqual(result, (5, 5))
    
    def test_parse_page_range_range(self):
        """Test parsing page range with dash."""
        result = self.orchestrator._parse_page_range("1-10")
        self.assertEqual(result, (1, 10))
    
    def test_contains_japanese(self):
        """Test Japanese character detection."""
        # Test Japanese text
        self.assertTrue(self.orchestrator._contains_japanese("テスト"))
        self.assertTrue(self.orchestrator._contains_japanese("こんにちは"))
        self.assertTrue(self.orchestrator._contains_japanese("漢字"))
        
        # Test English text
        self.assertFalse(self.orchestrator._contains_japanese("Hello"))
        self.assertFalse(self.orchestrator._contains_japanese("Test"))
        
        # Test mixed text
        self.assertTrue(self.orchestrator._contains_japanese("Hello テスト"))
    
    def test_extract_unique_japanese_text(self):
        """Test extraction of unique Japanese text."""
        # Create mock extraction result
        blocks = [
            MockTextBlock("テスト"),  # Japanese
            MockTextBlock("Hello"),  # English
            MockTextBlock("テスト"),  # Duplicate Japanese
            MockTextBlock("こんにちは"),  # Japanese
            MockTextBlock("こんにちは")  # Duplicate Japanese
        ]
        pages = [MockPage(page_num=0, blocks=blocks)]
        
        extraction_result = MockExtractionResult()
        extraction_result.pages = pages
        
        # Extract unique Japanese text
        unique_texts = self.orchestrator._extract_unique_japanese_text(extraction_result)
        
        # Should have only unique Japanese texts, no duplicates
        self.assertEqual(len(unique_texts), 2)
        self.assertIn("テスト", unique_texts)
        self.assertIn("こんにちは", unique_texts)
        self.assertNotIn("Hello", unique_texts)
    
    def test_translate_with_cache_all_cached(self):
        """Test translation when all texts are cached."""
        self.orchestrator.cache = {"テスト": "Test", "こんにちは": "Hello"}
        text_list = ["テスト", "こんにちは"]
        
        translations = self.orchestrator._translate_with_cache(text_list)
        
        expected = {"テスト": "Test", "こんにちは": "Hello"}
        self.assertEqual(translations, expected)
        self.assertEqual(self.orchestrator.stats['cache_hits'], 2)
        self.assertEqual(self.orchestrator.stats['api_calls'], 0)
    
    def test_translate_with_cache_mixed(self):
        """Test translation with mixed cached and uncached texts."""
        self.orchestrator.cache = {"テスト": "Test"}
        text_list = ["テスト", "こんにちは"]  # One cached, one not
        
        translations = self.orchestrator._translate_with_cache(text_list)
        
        # In offline mode, uncached text should be returned as-is
        expected = {"テスト": "Test", "こんにちは": "こんにちは"}
        self.assertEqual(translations, expected)
        self.assertEqual(self.orchestrator.stats['cache_hits'], 1)
        self.assertEqual(self.orchestrator.stats['api_calls'], 0)
    
    def test_translate_cache_only_missing(self):
        """Test cache-only mode with missing translations."""
        self.orchestrator.cache_only = True
        text_list = ["テスト"]  # Not in cache
        
        with self.assertRaises(ValueError) as context:
            self.orchestrator._translate_with_cache(text_list)
        
        self.assertIn("Cache-only mode", str(context.exception))
    
    @patch('scripts.translate_pdf.PDF_COMPONENTS_AVAILABLE', True)
    @patch('scripts.translate_pdf.PPTX_SYSTEM_AVAILABLE', True)
    @patch('scripts.translate_pdf.OPENAI_AVAILABLE', True)
    def test_translate_with_api_mock(self):
        """Test translation with API call (mocked)."""
        # Mock the OpenAI client and batch_translate
        with patch('scripts.translate_pdf.OpenAI') as mock_openai, \
             patch('scripts.translate_pdf.batch_translate') as mock_batch_translate:
            
            mock_openai.return_value = Mock(api_key="test_key")
            mock_batch_translate.return_value = ["Test"]
            
            self.orchestrator.cache = {}
            text_list = ["テスト"]
            
            translations = self.orchestrator._translate_with_cache(text_list)
            
            self.assertEqual(translations, {"テスト": "Test"})
            self.assertEqual(self.orchestrator.stats['cache_hits'], 0)
            self.assertEqual(self.orchestrator.stats['api_calls'], 1)
            
            # Verify OpenAI client was created
            mock_openai.assert_called_once()
            
            # Verify batch_translate was called
            mock_batch_translate.assert_called_once()
    
    def test_filter_pages_by_range_none(self):
        """Test page filtering when no range specified."""
        extraction_result = MockExtractionResult()
        extraction_result.pages = [
            MockPage(page_num=0),
            MockPage(page_num=1),
            MockPage(page_num=2)
        ]
        
        result = self.orchestrator._filter_pages_by_range(extraction_result)
        self.assertEqual(len(result.pages), 3)
    
    def test_filter_pages_by_range_specific(self):
        """Test page filtering with specific range."""
        extraction_result = MockExtractionResult()
        extraction_result.pages = [
            MockPage(page_num=0),  # Page 1 (1-based)
            MockPage(page_num=1),  # Page 2 (1-based)
            MockPage(page_num=2)   # Page 3 (1-based)
        ]
        
        self.orchestrator.pages = "2-3"
        result = self.orchestrator._filter_pages_by_range(extraction_result)
        
        # Should have pages 2 and 3 (1-based) which are pages 1 and 2 (0-based)
        self.assertEqual(len(result.pages), 2)
        self.assertEqual(result.pages[0].page_num, 1)
        self.assertEqual(result.pages[1].page_num, 2)
    
    def test_generate_bilingual_csv(self):
        """Test bilingual CSV generation."""
        # Create mock extraction result
        blocks = [
            MockTextBlock("テスト", "title", 14.0),
            MockTextBlock("こんにちは", "body", 12.0)
        ]
        pages = [MockPage(page_num=0, blocks=blocks)]
        
        extraction_result = MockExtractionResult()
        extraction_result.pages = pages
        
        translations = {"テスト": "Test", "こんにちは": "Hello"}
        csv_path = os.path.join(self.temp_dir, "bilingual.csv")
        
        # Generate CSV
        self.orchestrator._generate_bilingual_csv(extraction_result, translations, csv_path)
        
        # Verify CSV file was created
        self.assertTrue(os.path.exists(csv_path))
        
        # Verify CSV content
        with open(csv_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn("テスト", content)
        self.assertIn("Test", content)
        self.assertIn("こんにちは", content)
        self.assertIn("Hello", content)
        self.assertIn("title", content)
        self.assertIn("body", content)
        self.assertIn("14.0", content)
        self.assertIn("12.0", content)
    
    def test_copy_pdf_for_translation(self):
        """Test PDF copying for translation."""
        # Create source PDF content
        with open(self.input_pdf, 'w') as f:
            f.write("PDF content")
        
        # Copy PDF
        self.orchestrator._copy_pdf_for_translation()
        
        # Verify output file was created
        self.assertTrue(os.path.exists(self.output_pdf))
        
        # Verify content is the same
        with open(self.input_pdf, 'r') as f1, open(self.output_pdf, 'r') as f2:
            self.assertEqual(f1.read(), f2.read())


class TestPDFTranslationOrchestratorIntegration(unittest.TestCase):
    """Integration tests for PDFTranslationOrchestrator."""
    
    def setUp(self):
        """Set up integration test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.input_pdf = os.path.join(self.temp_dir, "test.pdf")
        self.output_pdf = os.path.join(self.temp_dir, "test_out.pdf")
        self.cache_file = os.path.join(self.temp_dir, "cache.json")
        
        # Create dummy input file
        with open(self.input_pdf, 'w') as f:
            f.write("dummy pdf content")
    
    def tearDown(self):
        """Clean up integration test fixtures."""
        shutil.rmtree(self.temp_dir)
    
    @patch('scripts.translate_pdf.PDF_COMPONENTS_AVAILABLE', True)
    def test_complete_orchestration_offline(self):
        """Test complete orchestration in offline mode."""
        # Mock the components
        with patch('scripts.translate_pdf.PDFExtractor') as mock_extractor, \
             patch('scripts.translate_pdf.PDFLayoutEngine') as mock_layout, \
             patch('scripts.translate_pdf.PDFBackProjector') as mock_backprojector, \
             patch('scripts.translate_pdf.PDFAuditor') as mock_auditor:
            
            # Set up mock extractor
            mock_extractor_instance = Mock()
            mock_extractor.return_value = mock_extractor_instance
            
            # Create mock extraction result
            blocks = [MockTextBlock("テスト", "body", 12.0)]
            pages = [MockPage(page_num=0, blocks=blocks)]
            extraction_result = MockExtractionResult(total_blocks=1, total_japanese_blocks=1)
            extraction_result.pages = pages
            mock_extractor_instance.extract_text_blocks.return_value = extraction_result
            
            # Set up mock layout engine
            mock_layout_instance = Mock()
            mock_layout.return_value = mock_layout_instance
            mock_layout_instance.optimize_font_sizes.return_value = []
            
            # Set up mock back-projector
            mock_backprojector_instance = Mock()
            mock_backprojector.return_value = mock_backprojector_instance
            
            # Set up mock auditor
            mock_auditor_instance = Mock()
            mock_auditor.return_value = mock_auditor_instance
            mock_audit_report = Mock()
            mock_auditor_instance.generate_audit_report.return_value = mock_audit_report
            
            # Initialize orchestrator
            orchestrator = PDFTranslationOrchestrator(
                input_path=self.input_pdf,
                output_path=self.output_pdf,
                cache_file=self.cache_file,
                offline=True,
                cache_only=False
            )
            
            # Run complete orchestration
            success = orchestrator.translate_pdf()
            
            # Verify orchestration succeeded
            self.assertTrue(success)
            
            # Verify extraction was called
            mock_extractor_instance.extract_text_blocks.assert_called_once()
            
            # Verify layout optimization was called
            mock_layout_instance.optimize_font_sizes.assert_called_once()
            
            # Verify back-projection was called
            mock_backprojector_instance.process_document.assert_called_once()
            
            # Verify audit was called
            mock_auditor_instance.generate_audit_report.assert_called_once()
            
            # Verify statistics were updated
            self.assertEqual(orchestrator.stats['total_blocks'], 1)
            self.assertEqual(orchestrator.stats['translated_blocks'], 1)


if __name__ == '__main__':
    unittest.main()