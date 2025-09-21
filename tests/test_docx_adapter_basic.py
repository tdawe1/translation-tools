"""
Basic unit tests for DOCX adapter security and functionality.
Tests core security checks and basic document processing.
"""
import pytest
import tempfile
import os
from unittest.mock import Mock, patch, mock_open

# Add scripts to path
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

from scripts.docx_adapter import DocxAdapter


class TestDocxAdapterBasic:
    """Basic security and functionality tests for DocxAdapter."""

    def test_file_size_check(self):
        """Test file size security check."""
        adapter = DocxAdapter()
        # Test within limits
        assert adapter._validate_file_size(1024) == True

        # Test oversized file
        with pytest.raises(ValueError, match="File too large"):
            adapter._validate_file_size(adapter.MAX_FILE_SIZE + 1)

    def test_xml_size_check(self):
        """Test XML size security check."""
        adapter = DocxAdapter()
        # Test within limits
        assert adapter._validate_xml_size(1024) == True

        # Test oversized XML
        with pytest.raises(ValueError, match="XML document too large"):
            adapter._validate_xml_size(adapter.MAX_XML_SIZE + 1)

    def test_japanese_detection(self):
        """Test Japanese text detection."""
        adapter = DocxAdapter()

        # Test Japanese text
        japanese_text = "これは日本語のテキストです"
        assert adapter._contains_japanese(japanese_text) == True

        # Test English text
        english_text = "This is English text"
        assert adapter._contains_japanese(english_text) == False

        # Test mixed text
        mixed_text = "This is mixed 日本語 text"
        assert adapter._contains_japanese(mixed_text) == True

    def test_extract_segments_file_size_validation(self):
        """Test that extract_segments validates file size."""
        adapter = DocxAdapter()

        # Test with oversized file
        with patch('scripts.docx_adapter.os.path.exists', return_value=True), \
             patch('scripts.docx_adapter.os.path.getsize', return_value=adapter.MAX_FILE_SIZE + 1):

            with pytest.raises(ValueError, match="File too large"):
                adapter.extract_segments("dummy.docx")

    def test_security_constants(self):
        """Test that security constants are properly defined."""
        adapter = DocxAdapter()

        # Check constants exist and are reasonable
        assert adapter.MAX_FILE_SIZE > 0
        assert adapter.MAX_XML_SIZE > 0
        assert adapter.MAX_FILE_SIZE > adapter.MAX_XML_SIZE  # File should be larger than XML