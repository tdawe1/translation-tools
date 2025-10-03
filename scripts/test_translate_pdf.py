#!/usr/bin/env python3
"""
test_translate_pdf.py

Unit tests for the PDF translation orchestrator.
Tests the main workflow, CLI interface, and integration with existing components.
"""

import unittest
import json
import os
import sys
import tempfile
import shutil
from unittest.mock import Mock, patch, MagicMock
from pathlib import Path

# Add the scripts directory to Python path to import modules
sys.path.insert(0, '/home/thomas/translation-tools/translations-pptx-pipeline/scripts')

# Mock PyMuPDF (fitz) if not available
try:
    import fitz
except ImportError:
    sys.modules['fitz'] = Mock()

# Mock other dependencies that might not be available
try:
    import pdfplumber
except ImportError:
    sys.modules['pdfplumber'] = Mock()

try:
    import pypdf
except ImportError:
    sys.modules['pypdf'] = Mock()

try:
    from pdfminer.high_level import extract_text_to_fp
    from pdfminer.layout import LAParams
except ImportError:
    sys.modules['pdfminer'] = Mock()
    sys.modules['pdfminer.high_level'] = Mock()
    sys.modules['pdfminer.layout'] = Mock()

# Import the translator under test
try:
    from translate_pdf import PDFTranslator, parse_page_range
except ImportError as e:
    print(f"Warning: Could not import translate_pdf: {e}")
    # Create a minimal mock for testing
    PDFTranslator = Mock
    parse_page_range = Mock()


class TestPDFTranslator(unittest.TestCase):
    """Test cases for PDFTranslator class."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.input_pdf = os.path.join(self.temp_dir, "test.pdf")
        self.output_pdf = os.path.join(self.temp_dir, "test_en.pdf")
        self.cache_file = os.path.join(self.temp_dir, "test_cache.json")
        self.glossary_file = os.path.join(self.temp_dir, "test_glossary.json")
        
        # Create test files
        self._create_test_files()
        
        # Mock PDF content
        self.mock_extraction_result = self._create_mock_extraction_result()
        
    def tearDown(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.temp_dir)
    
    def _create_test_files(self):
        """Create test input files."""
        # Create a minimal PDF file (just empty file for testing)
        Path(self.input_pdf).touch()
        
        # Create test cache
        cache_data = {
            "テスト": "Test",
            "日本語": "Japanese"
        }
        with open(self.cache_file, 'w', encoding='utf-8') as f:
            json.dump(cache_data, f)
        
        # Create test glossary
        glossary_data = {
            "専門用語": "technical term",
            "会社名": "company name"
        }
        with open(self.glossary_file, 'w', encoding='utf-8') as f:
            json.dump(glossary_data, f)
    
    def _create_mock_extraction_result(self):
        """Create mock extraction result for testing."""
        from translate_pdf import ExtractionResult, PageInfo
        from extract_pdf import TextBlock
        
        # Create mock text blocks
        text_blocks = [
            TextBlock(
                id="page_0_block_0",
                page=0,
                text="これはテストです",
                x0=100, y0=100, x1=200, y1=120,
                font_size=12.0,
                font_name="Arial",
                block_type="body"
            ),
            TextBlock(
                id="page_0_block_1",
                page=0,
                text="日本語の文章",
                x0=100, y0=150, x1=250, y1=170,
                font_size=14.0,
                font_name="Helvetica",
                block_type="title"
            )
        ]
        
        # Create page info
        page_info = PageInfo(
            page_num=0,
            width=595, height=842,
            rotation=0,
            text_blocks=text_blocks,
            has_japanese=True,
            extraction_method="fitz"
        )
        
        # Create extraction result
        return ExtractionResult(
            filename="test.pdf",
            pages=[page_info],
            total_blocks=2,
            total_japanese_blocks=2,
            extraction_time=1.0,
            extraction_methods={"fitz": 1},
            metadata={}
        )
    
    def test_parse_page_range(self):
        """Test page range parsing."""
        # Test single page
        self.assertEqual(parse_page_range("5"), (5, 5))
        
        # Test page range
        self.assertEqual(parse_page_range("1-5"), (1, 5))
        
        # Test None input
        self.assertIsNone(parse_page_range(None))
        self.assertIsNone(parse_page_range(""))
        
        # Test invalid input
        with self.assertRaises(ValueError):
            parse_page_range("invalid")
    
    def test_translator_initialization(self):
        """Test PDFTranslator initialization."""
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf,
            cache_path=self.cache_file,
            glossary_path=self.glossary_file,
            model="gpt-4o-mini",
            batch_size=5,
            offline_mode=True,
            cache_only=False
        )
        
        self.assertEqual(translator.input_path, self.input_pdf)
        self.assertEqual(translator.output_path, self.output_pdf)
        self.assertEqual(translator.cache_path, self.cache_file)
        self.assertEqual(translator.glossary_path, self.glossary_file)
        self.assertEqual(translator.model, "gpt-4o-mini")
        self.assertEqual(translator.batch_size, 5)
        self.assertTrue(translator.offline_mode)
        self.assertFalse(translator.cache_only)
    
    def test_load_resources(self):
        """Test loading cache and glossary files."""
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf,
            cache_path=self.cache_file,
            glossary_path=self.glossary_file
        )
        
        translator.load_resources()
        
        # Check cache was loaded
        self.assertIn("テスト", translator.cache)
        self.assertEqual(translator.cache["テスト"], "Test")
        
        # Check glossary was loaded
        self.assertIn("専門用語", translator.glossary)
        self.assertEqual(translator.glossary["専門用語"], "technical term")
    
    def test_load_resources_missing_files(self):
        """Test handling of missing resource files."""
        # Remove cache and glossary files
        os.remove(self.cache_file)
        os.remove(self.glossary_file)
        
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf,
            cache_path=self.cache_file,
            glossary_path=self.glossary_file
        )
        
        translator.load_resources()
        
        # Should initialize empty dictionaries
        self.assertEqual(translator.cache, {})
        self.assertEqual(translator.glossary, {})
    
    @patch('translate_pdf.PDFExtractor')
    def test_extract_text(self, mock_extractor_class):
        """Test text extraction from PDF."""
        # Mock the extractor
        mock_extractor = Mock()
        mock_extractor.extract_text_blocks.return_value = self.mock_extraction_result
        mock_extractor_class.return_value = mock_extractor
        
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf
        )
        
        result = translator.extract_text()
        
        # Verify extractor was called correctly
        mock_extractor.extract_text_blocks.assert_called_once_with(self.input_pdf, detailed=True)
        
        # Verify result
        self.assertEqual(result.total_blocks, 2)
        self.assertEqual(result.total_japanese_blocks, 2)
        self.assertEqual(len(result.pages), 1)
    
    @patch('translate_pdf.PDFExtractor')
    def test_extract_text_with_page_range(self, mock_extractor_class):
        """Test text extraction with page range filtering."""
        # Mock extractor to return multiple pages
        mock_result = self._create_mock_extraction_result()
        
        # Add a second page
        from extract_pdf import TextBlock, PageInfo
        
        page2_text = TextBlock(
            id="page_1_block_0",
            page=1,
            text="2ページ目",
            x0=100, y0=100, x1=200, y1=120,
            font_size=12.0,
            font_name="Arial",
            block_type="body"
        )
        
        page2_info = PageInfo(
            page_num=1,
            width=595, height=842,
            rotation=0,
            text_blocks=[page2_text],
            has_japanese=True,
            extraction_method="fitz"
        )
        
        mock_result.pages.append(page2_info)
        mock_result.total_blocks = 3
        mock_result.total_japanese_blocks = 3
        
        mock_extractor = Mock()
        mock_extractor.extract_text_blocks.return_value = mock_result
        mock_extractor_class.return_value = mock_extractor
        
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf
        )
        
        # Test filtering to first page only
        result = translator.extract_text(page_range=(1, 1))
        
        # Verify only first page is included
        self.assertEqual(len(result.pages), 1)
        self.assertEqual(result.pages[0].page_num, 0)  # 0-based
    
    def test_prepare_for_translation(self):
        """Test preparation of texts for translation."""
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf
        )
        
        japanese_texts = translator.prepare_for_translation(self.mock_extraction_result)
        
        # Should extract both Japanese texts
        self.assertEqual(len(japanese_texts), 2)
        self.assertIn("これはテストです", japanese_texts)
        self.assertIn("日本語の文章", japanese_texts)
        
        # Should remove duplicates
        # Add a duplicate text to the mock result
        self.mock_extraction_result.pages[0].text_blocks.append(
            self.mock_extraction_result.pages[0].text_blocks[0]
        )
        
        japanese_texts = translator.prepare_for_translation(self.mock_extraction_result)
        self.assertEqual(len(japanese_texts), 2)  # Should still be 2 (no duplicates)
    
    def test_translate_texts_with_cache(self):
        """Test text translation with cache hits."""
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf,
            cache_path=self.cache_file
        )
        
        translator.load_resources()
        
        texts = ["テスト", "日本語", "新しい文章"]
        translations = translator.translate_texts(texts)
        
        # Check cached translations
        self.assertEqual(translations["テスト"], "Test")
        self.assertEqual(translations["日本語"], "Japanese")
        
        # Check statistics
        self.assertEqual(translator.translation_stats['cache_hits'], 2)
        self.assertEqual(translator.translation_stats['translated_blocks'], 3)
    
    def test_translate_texts_offline_mode(self):
        """Test text translation in offline mode."""
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf,
            offline_mode=True
        )
        
        texts = ["テスト", "日本語"]
        translations = translator.translate_texts(texts)
        
        # Should generate mock translations
        self.assertEqual(translations["テスト"], "[OFFLINE] テスト")
        self.assertEqual(translations["日本語"], "[OFFLINE] 日本語")
        
        # Should update cache
        self.assertIn("テスト", translator.cache)
        self.assertIn("日本語", translator.cache)
    
    def test_translate_texts_cache_only(self):
        """Test cache-only mode."""
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf,
            cache_path=self.cache_file,
            cache_only=True
        )
        
        translator.load_resources()
        
        # This should work (all texts in cache)
        texts = ["テスト", "日本語"]
        translations = translator.translate_texts(texts)
        self.assertEqual(len(translations), 2)
        
        # This should fail (text not in cache)
        texts = ["キャッシュにない文章"]
        with self.assertRaises(ValueError):
            translator.translate_texts(texts)
    
    @patch('translate_pdf.PDFLayoutEngine')
    def test_optimize_layout(self, mock_engine_class):
        """Test layout optimization."""
        # Mock layout engine
        mock_engine = Mock()
        mock_engine.optimize_font_sizes.return_value = []
        mock_engine.handle_overflow.return_value = []
        mock_engine_class.return_value = mock_engine
        
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf
        )
        
        translations = {"これはテストです": "This is a test"}
        
        optimized_blocks = translator.optimize_layout(self.mock_extraction_result, translations)
        
        # Verify layout engine was called
        mock_engine.optimize_font_sizes.assert_called_once()
        mock_engine.handle_overflow.assert_called_once()
    
    @patch('translate_pdf.PDFBackProjector')
    def test_apply_translations(self, mock_backprojector_class):
        """Test application of translations to PDF."""
        # Mock back-projector
        mock_backprojector = Mock()
        mock_backprojector_class.return_value = mock_backprojector
        
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf
        )
        
        optimized_blocks = []  # Mock optimized blocks
        translations = {"これはテストです": "This is a test"}
        
        translator.apply_translations(self.mock_extraction_result, optimized_blocks, translations)
        
        # Verify back-projector was instantiated and called
        mock_backprojector_class.assert_called_once()
        mock_backprojector.process_document.assert_called_once()
    
    @patch('translate_pdf.PDFAuditor')
    def test_generate_outputs(self, mock_auditor_class):
        """Test generation of output files."""
        # Mock auditor
        mock_auditor = Mock()
        mock_auditor.generate_audit_report.return_value = Mock()
        mock_auditor_class.return_value = mock_auditor
        
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf
        )
        
        translations = {"これはテストです": "This is a test"}
        
        bilingual_csv, audit_json = translator.generate_outputs(self.mock_extraction_result, translations)
        
        # Verify output file paths
        expected_csv = os.path.splitext(self.output_pdf)[0] + "_bilingual.csv"
        expected_json = os.path.splitext(self.output_pdf)[0] + "_audit.json"
        
        self.assertEqual(bilingual_csv, expected_csv)
        self.assertEqual(audit_json, expected_json)
        
        # Verify files were created
        self.assertTrue(os.path.exists(bilingual_csv))
        self.assertTrue(os.path.exists(audit_json))
        
        # Verify auditor was called
        mock_auditor.generate_audit_report.assert_called_once_with(self.output_pdf, self.input_pdf)
    
    def test_classify_content_type(self):
        """Test content type classification."""
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf
        )
        
        # Test different block types
        from pdf_layout_engine import ContentType
        
        self.assertEqual(
            translator._classify_content_type("text", "title"),
            ContentType.TITLE
        )
        self.assertEqual(
            translator._classify_content_type("text", "header"),
            ContentType.HEADER
        )
        self.assertEqual(
            translator._classify_content_type("text", "unknown"),
            ContentType.BODY
        )
    
    def test_get_layout_constraint(self):
        """Test layout constraint determination."""
        translator = PDFTranslator(
            input_path=self.input_pdf,
            output_path=self.output_pdf
        )
        
        # Create mock text blocks
        from extract_pdf import TextBlock
        
        table_block = TextBlock("id", 0, "text", 0, 0, 0, 0, 12, "Arial", "table")
        header_block = TextBlock("id", 0, "text", 0, 0, 0, 0, 12, "Arial", "header")
        body_block = TextBlock("id", 0, "text", 0, 0, 0, 0, 12, "Arial", "body")
        
        self.assertEqual(translator._get_layout_constraint(table_block), 'fixed')
        self.assertEqual(translator._get_layout_constraint(header_block), 'flexible')
        self.assertEqual(translator._get_layout_constraint(body_block), 'flexible')


class TestCLIInterface(unittest.TestCase):
    """Test CLI interface integration."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.test_pdf = os.path.join(self.temp_dir, "test.pdf")
        Path(self.test_pdf).touch()
    
    def tearDown(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.temp_dir)
    
    @patch('translate_pdf.PDFTranslator')
    def test_cli_basic_usage(self, mock_translator_class):
        """Test basic CLI usage."""
        # Mock translator
        mock_translator = Mock()
        mock_translator_class.return_value = mock_translator
        
        # Mock sys.argv
        test_args = [
            'translate_pdf.py',
            '--in', self.test_pdf,
            '--out', os.path.join(self.temp_dir, 'output.pdf')
        ]
        
        with patch('sys.argv', test_args):
            from translate_pdf import main
            
            try:
                main()
            except SystemExit:
                pass  # CLI may exit with code 0
        
        # Verify translator was instantiated with correct parameters
        mock_translator_class.assert_called_once()
        args, kwargs = mock_translator_class.call_args
        
        self.assertEqual(kwargs['input_path'], self.test_pdf)
        self.assertTrue(kwargs['input_path'].endswith('test.pdf'))
    
    @patch('translate_pdf.PDFTranslator')
    def test_cli_with_all_options(self, mock_translator_class):
        """Test CLI with all options."""
        mock_translator = Mock()
        mock_translator_class.return_value = mock_translator
        
        cache_file = os.path.join(self.temp_dir, 'cache.json')
        glossary_file = os.path.join(self.temp_dir, 'glossary.json')
        output_file = os.path.join(self.temp_dir, 'output.pdf')
        
        test_args = [
            'translate_pdf.py',
            '--in', self.test_pdf,
            '--out', output_file,
            '--cache', cache_file,
            '--glossary', glossary_file,
            '--model', 'gpt-4o-mini',
            '--batch', '15',
            '--pages', '1-10',
            '--offline',
            '--no-backup',
            '--verbose'
        ]
        
        with patch('sys.argv', test_args):
            from translate_pdf import main
            
            try:
                main()
            except SystemExit:
                pass
        
        # Verify all parameters were passed correctly
        mock_translator_class.assert_called_once()
        args, kwargs = mock_translator_class.call_args
        
        self.assertEqual(kwargs['model'], 'gpt-4o-mini')
        self.assertEqual(kwargs['batch_size'], 15)
        self.assertTrue(kwargs['offline_mode'])
        self.assertEqual(kwargs['cache_path'], cache_file)
        self.assertEqual(kwargs['glossary_path'], glossary_file)


class TestIntegration(unittest.TestCase):
    """Integration tests with mocked external dependencies."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.test_pdf = os.path.join(self.temp_dir, "test.pdf")
        Path(self.test_pdf).touch()
    
    def tearDown(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.temp_dir)
    
    @patch('translate_pdf.PDFExtractor')
    @patch('translate_pdf.PDFBackProjector')
    @patch('translate_pdf.PDFAuditor')
    def test_full_workflow_mocked(self, mock_auditor, mock_backprojector, mock_extractor):
        """Test full workflow with mocked dependencies."""
        from translate_pdf import ExtractionResult, PageInfo
        from extract_pdf import TextBlock
        
        # Mock extraction result
        text_block = TextBlock(
            id="test_block",
            page=0,
            text="テスト",
            x0=100, y0=100, x1=200, y1=120,
            font_size=12.0,
            font_name="Arial",
            block_type="body"
        )
        
        page_info = PageInfo(
            page_num=0,
            width=595, height=842,
            rotation=0,
            text_blocks=[text_block],
            has_japanese=True,
            extraction_method="fitz"
        )
        
        extraction_result = ExtractionResult(
            filename="test.pdf",
            pages=[page_info],
            total_blocks=1,
            total_japanese_blocks=1,
            extraction_time=1.0,
            extraction_methods={"fitz": 1},
            metadata={}
        )
        
        mock_extractor_instance = Mock()
        mock_extractor_instance.extract_text_blocks.return_value = extraction_result
        mock_extractor.return_value = mock_extractor_instance
        
        # Mock back-projector
        mock_backprojector_instance = Mock()
        mock_backprojector.return_value = mock_backprojector_instance
        
        # Mock auditor
        mock_auditor_instance = Mock()
        mock_auditor_instance.generate_audit_report.return_value = Mock()
        mock_auditor.return_value = mock_auditor_instance
        
        # Create translator and run workflow
        translator = PDFTranslator(
            input_path=self.test_pdf,
            output_path=os.path.join(self.temp_dir, "output.pdf"),
            offline_mode=True
        )
        
        # Should complete without errors
        translator.translate()
        
        # Verify all components were called
        mock_extractor_instance.extract_text_blocks.assert_called_once()
        mock_backprojector_instance.process_document.assert_called_once()
        mock_auditor_instance.generate_audit_report.assert_called_once()


if __name__ == '__main__':
    # Configure logging for tests
    import logging
    logging.basicConfig(level=logging.WARNING)
    
    # Run tests
    unittest.main(verbosity=2)