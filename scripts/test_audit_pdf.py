#!/usr/bin/env python3
"""
Unit tests for PDF audit functionality.
Tests the PDFAuditor class and its methods.
"""
import unittest
import tempfile
import os
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
import json
import csv
import io

# Import the PDFAuditor class (adjust path if needed)
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.audit_pdf import PDFAuditor, LayoutCheckResult, QualityAssessment, AuditReport


class TestPDFAuditor(unittest.TestCase):
    """Test cases for PDFAuditor class."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.auditor = PDFAuditor()
        
    def test_japanese_pattern_detection(self):
        """Test Japanese character pattern detection."""
        # Test Hiragana
        self.assertTrue(self.auditor.jp_core_pattern.search('こんにちは'))
        
        # Test Katakana  
        self.assertTrue(self.auditor.jp_core_pattern.search('コンニチハ'))
        
        # Test Kanji
        self.assertTrue(self.auditor.jp_core_pattern.search('日本語'))
        
        # Test English (should not match)
        self.assertFalse(self.auditor.jp_core_pattern.search('Hello'))
        
        # Test mixed content
        text = "Hello こんにちは World 日本語"
        matches = self.auditor.jp_core_pattern.findall(text)
        self.assertEqual(len(matches), 8)  # こ + ん + に + ち + は + 日 + 本 + 語
    
    def test_count_residual_jp_mock(self):
        """Test residual Japanese counting with mock PDF."""
        # Mock the text extraction
        with patch.object(self.auditor, '_extract_text') as mock_extract:
            mock_extract.return_value = "Hello こんにちは World 日本語"
            
            result = self.auditor.count_residual_jp("dummy.pdf")
            self.assertEqual(result, 8)
    
    def test_layout_integrity_mock(self):
        """Test layout integrity checking with mock PDFs."""
        with patch.object(self.auditor, '_get_page_count') as mock_page_count, \
             patch.object(self.auditor, '_extract_text') as mock_extract:
            
            # Mock successful case
            mock_page_count.return_value = 10
            mock_extract.return_value = "Sample text content"
            
            result = self.auditor.check_layout_integrity("original.pdf", "translated.pdf")
            
            self.assertIsInstance(result, LayoutCheckResult)
            self.assertEqual(result.score, 1.0)
            self.assertTrue(result.page_count_match)
            self.assertTrue(result.similar_structure)
            self.assertEqual(len(result.issues), 0)
    
    def test_layout_integrity_page_mismatch(self):
        """Test layout integrity with page count mismatch."""
        with patch.object(self.auditor, '_get_page_count') as mock_page_count, \
             patch.object(self.auditor, '_extract_text') as mock_extract:
            
            def mock_page_count_side_effect(path):
                return 5 if "original" in path else 7
            
            mock_page_count.side_effect = mock_page_count_side_effect
            mock_extract.return_value = "Sample text content"
            
            result = self.auditor.check_layout_integrity("original.pdf", "translated.pdf")
            
            self.assertIsInstance(result, LayoutCheckResult)
            self.assertLess(result.score, 1.0)
            self.assertFalse(result.page_count_match)
            self.assertIn("Page count mismatch", result.issues[0])
    
    def test_quality_assessment_no_japanese(self):
        """Test quality assessment with no residual Japanese."""
        with patch.object(self.auditor, '_extract_text') as mock_extract:
            mock_extract.return_value = "This is a properly translated document with no Japanese characters."
            
            result = self.auditor.assess_translation_quality("dummy.pdf")
            
            self.assertIsInstance(result, QualityAssessment)
            self.assertEqual(result.residual_japanese_count, 0)
            self.assertEqual(result.residual_japanese_percentage, 0.0)
            self.assertGreater(result.overall_quality_score, 0.8)
    
    def test_quality_assessment_with_japanese(self):
        """Test quality assessment with residual Japanese."""
        with patch.object(self.auditor, '_extract_text') as mock_extract:
            mock_extract.return_value = "This document has こんにちは Japanese characters remaining."
            
            result = self.auditor.assess_translation_quality("dummy.pdf")
            
            self.assertIsInstance(result, QualityAssessment)
            self.assertGreater(result.residual_japanese_count, 0)
            self.assertGreater(result.residual_japanese_percentage, 0.0)
            self.assertIn("Remove residual Japanese characters", result.recommendations[0])
    
    def test_quality_assessment_short_document(self):
        """Test quality assessment with very short document."""
        with patch.object(self.auditor, '_extract_text') as mock_extract:
            mock_extract.return_value = "Short"
            
            result = self.auditor.assess_translation_quality("dummy.pdf")
            
            self.assertIsInstance(result, QualityAssessment)
            self.assertLess(result.text_completeness_score, 1.0)
            self.assertIn("incomplete", result.recommendations[0].lower())
    
    def test_compare_with_original_mock(self):
        """Test comparison with original PDF."""
        with patch.object(self.auditor, '_extract_text') as mock_extract:
            # Mock Japanese original and English translation
            def mock_extract_side_effect(path):
                if "original" in path:
                    return "こんにちは世界"
                else:
                    return "Hello World"
            
            mock_extract.side_effect = mock_extract_side_effect
            
            result = self.auditor.compare_with_original("original.pdf", "translated.pdf")
            
            self.assertIn('original_stats', result)
            self.assertIn('translated_stats', result)
            self.assertIn('expansion_ratio', result)
            self.assertIn('residual_japanese_removed', result)
            
            # Should show significant Japanese removal
            self.assertGreater(result['residual_japanese_removed'], 50)
    
    def test_analyze_character_types(self):
        """Test character type analysis."""
        text = "Hello 123 こんにちは！"
        result = self.auditor._analyze_character_types(text)
        
        self.assertIsInstance(result, dict)
        self.assertIn('japanese', result)
        self.assertIn('english', result)
        self.assertIn('digits', result)
        self.assertIn('punctuation', result)
        self.assertIn('whitespace', result)
        
        self.assertEqual(result['english'], 5)  # H,e,l,l,o
        self.assertEqual(result['digits'], 3)    # 1,2,3
        self.assertEqual(result['japanese'], 5)  # こ,ん,に,ち,は
    
    def test_analyze_text_structure(self):
        """Test text structure analysis."""
        text = "First paragraph.\n\nSecond paragraph with\nmultiple lines."
        result = self.auditor._analyze_text_structure(text)
        
        self.assertIsInstance(result, dict)
        self.assertIn('line_count', result)
        self.assertIn('paragraph_count', result)
        self.assertIn('avg_line_length', result)
        self.assertIn('avg_paragraph_length', result)
        
        self.assertEqual(result['paragraph_count'], 2)
        self.assertGreater(result['line_count'], 1)
    
    def test_assess_formatting_consistency(self):
        """Test formatting consistency assessment."""
        # Good formatting (avoid hyphenated words)
        good_text = "This is well formatted text.\nWith proper line breaks.\nAnd consistent spacing."
        score = self.auditor._assess_formatting_consistency(good_text)
        self.assertEqual(score, 1.0)
        
        # Bad formatting - multiple spaces
        bad_text_spaces = "Bad   formatting with multiple spaces"
        score_spaces = self.auditor._assess_formatting_consistency(bad_text_spaces)
        self.assertLess(score_spaces, 1.0)
        
        # Bad formatting - mixed line endings
        bad_text_endings = "Formatting\r\nWith mixed line endings"
        score_endings = self.auditor._assess_formatting_consistency(bad_text_endings)
        self.assertLess(score_endings, 1.0)
    
    @patch('scripts.audit_pdf.pypdf.PdfReader')
    @patch('builtins.open')
    def test_get_page_count(self, mock_open, mock_pdf_reader):
        """Test page count extraction."""
        # Mock PDF reader with 5 pages
        mock_reader = Mock()
        mock_reader.pages = [Mock() for _ in range(5)]
        mock_pdf_reader.return_value = mock_reader
        
        result = self.auditor._get_page_count("dummy.pdf")
        self.assertEqual(result, 5)
    
    @patch('scripts.audit_pdf.pypdf.PdfReader')
    @patch('builtins.open')
    def test_analyze_pages(self, mock_open, mock_pdf_reader):
        """Test page-level analysis."""
        # Mock pages with different content
        mock_pages = []
        for i in range(3):
            mock_page = Mock()
            mock_page.extract_text.return_value = f"Page {i+1} content with こんにちは"
            mock_pages.append(mock_page)
        
        mock_reader = Mock()
        mock_reader.pages = mock_pages
        mock_pdf_reader.return_value = mock_reader
        
        result = self.auditor._analyze_pages("dummy.pdf")
        
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 3)
        
        # Check first page
        first_page = result[0]
        self.assertEqual(first_page['page_number'], 1)
        self.assertGreater(first_page['word_count'], 0)
        self.assertEqual(first_page['japanese_chars'], 5)  # こ,ん,に,ち,は


class TestReportGeneration(unittest.TestCase):
    """Test cases for report generation."""
    
    def setUp(self):
        """Set up test fixtures."""
        # Create a sample audit report
        self.sample_report = AuditReport(
            file_path="/test/translated.pdf",
            original_file_path="/test/original.pdf",
            timestamp="2024-01-01T12:00:00",
            total_pages=10,
            extracted_text_length=5000,
            layout_check=LayoutCheckResult(
                score=0.9,
                issues=["Minor layout issue"],
                page_count_match=True,
                similar_structure=True
            ),
            quality_assessment=QualityAssessment(
                residual_japanese_count=5,
                residual_japanese_percentage=0.1,
                text_completeness_score=0.95,
                formatting_consistency_score=0.9,
                overall_quality_score=0.85,
                recommendations=["Review formatting"]
            ),
            page_details=[
                {"page_number": 1, "word_count": 100, "japanese_chars": 2, "text_length": 500}
            ]
        )
    
    def test_save_report_csv(self):
        """Test CSV report generation."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
            csv_path = f.name
        
        try:
            from scripts.audit_pdf import save_report_csv
            save_report_csv(self.sample_report, csv_path)
            
            # Verify file was created
            self.assertTrue(os.path.exists(csv_path))
            
            # Read and verify content
            with open(csv_path, 'r', encoding='utf-8') as f:
                content = f.read()
                self.assertIn('file_path', content)
                self.assertIn('residual_japanese_count', content)
                self.assertIn('overall_quality_score', content)
                
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
    
    def test_save_report_json(self):
        """Test JSON report generation."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json_path = f.name
        
        try:
            from scripts.audit_pdf import save_report_json
            save_report_json(self.sample_report, json_path)
            
            # Verify file was created
            self.assertTrue(os.path.exists(json_path))
            
            # Read and verify content
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.assertEqual(data['file_path'], "/test/translated.pdf")
                self.assertEqual(data['total_pages'], 10)
                self.assertEqual(data['quality_assessment']['residual_japanese_count'], 5)
                
        finally:
            if os.path.exists(json_path):
                os.unlink(json_path)


class TestCLIIntegration(unittest.TestCase):
    """Test cases for CLI integration."""
    
    def setUp(self):
        """Set up test fixtures."""
        # Create temporary PDF-like files for testing
        self.temp_dir = tempfile.mkdtemp()
        self.translated_path = os.path.join(self.temp_dir, "translated.pdf")
        self.original_path = os.path.join(self.temp_dir, "original.pdf")
        
        # Create dummy files (not real PDFs, but exist for file checking)
        with open(self.translated_path, 'w') as f:
            f.write("dummy content")
        with open(self.original_path, 'w') as f:
            f.write("dummy content")
    
    def tearDown(self):
        """Clean up test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir)
    
    @patch('scripts.audit_pdf.PDFAuditor.generate_audit_report')
    def test_cli_main_function_exists(self, mock_generate):
        """Test that main function exists and can be imported."""
        from scripts.audit_pdf import main
        
        # Mock the audit report generation
        mock_generate.return_value = Mock(
            file_path=self.translated_path,
            original_file_path=self.original_path,
            timestamp="2024-01-01T12:00:00",
            total_pages=1,
            extracted_text_length=100,
            layout_check=None,
            quality_assessment=Mock(
                residual_japanese_count=0,
                residual_japanese_percentage=0.0,
                text_completeness_score=1.0,
                formatting_consistency_score=1.0,
                overall_quality_score=0.9,
                recommendations=[]
            ),
            page_details=[]
        )
        
        # Mock sys.argv for CLI testing
        test_args = ["audit_pdf.py", self.translated_path, self.original_path]
        
        with patch('sys.argv', test_args), \
             patch('scripts.audit_pdf.save_report_csv') as mock_save:
            
            # This should not raise an exception
            try:
                main()
            except SystemExit as e:
                # Expecting SystemExit from sys.exit() calls
                self.assertIn(e.code, [0, 1])
    
    def test_cli_file_not_found(self):
        """Test CLI behavior with non-existent files."""
        from scripts.audit_pdf import main
        
        # Test with non-existent file
        test_args = ["audit_pdf.py", "/nonexistent/file.pdf"]
        
        with patch('sys.argv', test_args), \
             patch('sys.exit') as mock_exit:
            main()
            mock_exit.assert_called_with(1)


if __name__ == '__main__':
    # Run tests with verbose output
    unittest.main(verbosity=2)