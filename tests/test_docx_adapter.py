#!/usr/bin/env python3
"""Unit tests for docx_adapter.py"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import io
import zipfile
from pathlib import Path
from typing import List
from unittest.mock import Mock, patch

import pytest

from scripts.docx_adapter import DocxAdapter
from backend.document_adapter import Segment, SegmentType


class TestDocxAdapter:
    def setup_method(self):
        self.adapter = DocxAdapter()

    def test_supported_formats(self):
        assert self.adapter.supported_formats() == ["docx"]

    def test_extract_segments_nonexistent_file(self):
        segments = self.adapter.extract_segments("nonexistent.docx")
        assert segments == []

    def test_extract_segments_no_document_xml(self):
        # Create a mock zip without word/document.xml
        mock_zip = Mock()
        mock_zip.namelist.return_value = ["other/file.txt"]

        with patch("zipfile.ZipFile", return_value=mock_zip):
            segments = self.adapter.extract_segments("test.docx")
            assert segments == []

    def test_extract_segments_simple_docx(self):
        # Create mock DOCX content
        document_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello</w:t>
      </w:r>
      <w:r>
        <w:t> world</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>日本語</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>'''

        # Create a mock zip file
        mock_zip = Mock()
        mock_zip.__enter__ = Mock(return_value=mock_zip)
        mock_zip.__exit__ = Mock(return_value=None)
        mock_zip.namelist.return_value = ["word/document.xml"]
        mock_zip.read.return_value = document_xml.encode('utf-8')

        with patch("zipfile.ZipFile", return_value=mock_zip), \
             patch("pathlib.Path.exists", return_value=True):
            segments = self.adapter.extract_segments("test.docx")

            assert len(segments) == 3
            # First segment
            assert segments[0].id == "word/document.xml:0:0"
            assert segments[0].text == "Hello"
            assert segments[0].segment_type == SegmentType.PARAGRAPH
            assert segments[0].has_japanese == False
            assert segments[0].word_count == 1

            # Second segment
            assert segments[1].id == "word/document.xml:0:1"
            assert segments[1].text == " world"
            assert segments[1].word_count == 1

            # Third segment (Japanese)
            assert segments[2].id == "word/document.xml:1:0"
            assert segments[2].text == "日本語"
            assert segments[2].has_japanese == True

    def test_collect_metadata(self):
        # Mock extract_segments
        mock_segments = [
            Segment(
                id="test:0:0",
                text="Hello world",
                segment_type=SegmentType.PARAGRAPH,
                file_path="word/document.xml",
                paragraph_index=0,
                run_index=0,
                word_count=2
            ),
            Segment(
                id="test:1:0",
                text="日本語",
                segment_type=SegmentType.PARAGRAPH,
                file_path="word/document.xml",
                paragraph_index=1,
                run_index=0,
                word_count=1,
                has_japanese=True
            )
        ]

        with patch.object(self.adapter, 'extract_segments', return_value=mock_segments):
            metadata = self.adapter.collect_metadata("test.docx")

            assert metadata.file_path == "test.docx"
            assert metadata.format == "docx"
            assert metadata.word_count == 3
            assert metadata.character_count == 14
            assert metadata.segment_count == 2
            assert metadata.has_headers_footers == False
            assert metadata.has_footnotes == False
            assert metadata.has_comments == False
            assert metadata.has_fields == False

    def test_apply_translations(self):
        # Create mock DOCX content
        document_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>'''

        translations = [
            Segment(
                id="word/document.xml:0:0",
                text="こんにちは",
                segment_type=SegmentType.PARAGRAPH,
                file_path="word/document.xml",
                paragraph_index=0,
                run_index=0,
                has_japanese=True,
                word_count=1
            )
        ]

        # Mock input zip
        mock_zip_in = Mock()
        mock_zip_in.__enter__ = Mock(return_value=mock_zip_in)
        mock_zip_in.__exit__ = Mock(return_value=None)
        mock_zip_in.namelist.return_value = ["word/document.xml"]
        mock_zip_in.read.return_value = document_xml.encode('utf-8')

        # Mock output zip
        mock_zip_out = Mock()
        mock_zip_out.__enter__ = Mock(return_value=mock_zip_out)
        mock_zip_out.__exit__ = Mock(return_value=None)

        with patch("zipfile.ZipFile") as mock_zip_class, \
             patch("pathlib.Path") as mock_path, \
             patch("os.makedirs"):

            # Configure mocks - return different mocks for input and output
            def zip_file_side_effect(path, mode):
                if mode == 'r':
                    return mock_zip_in
                else:
                    return mock_zip_out
            mock_zip_class.side_effect = zip_file_side_effect
            mock_path_instance = Mock()
            mock_path_instance.with_name.return_value = mock_path_instance
            mock_path_instance.parent = Mock()
            mock_path.return_value = mock_path_instance

            result = self.adapter.apply_translations(
                "input.docx",
                translations,
                "output.docx"
            )

            assert result.segments_translated == 1
            assert result.total_segments == 1
            assert result.words_translated == 1
            assert result.total_words == 1

    def test_has_japanese_helper(self):
        from scripts.docx_adapter import _has_japanese

        # Test various cases
        assert _has_japanese("こんにちは") == True
        assert _has_japanese("カタカナ") == True
        assert _has_japanese("漢字") == True
        assert _has_japanese("Hello") == False
        assert _has_japanese("") == False
        assert _has_japanese("Hello 日本語 World") == True


if __name__ == "__main__":
    pytest.main([__file__, "-v"])