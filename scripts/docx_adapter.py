#!/usr/bin/env python3
"""
DOCX adapter for document processing.
"""

import re
import shutil
import uuid
from pathlib import Path
from typing import List
import zipfile
try:
    from defusedxml.ElementTree import fromstring as safe_fromstring, tostring as safe_tostring
except ImportError:
    from xml.etree.ElementTree import fromstring as safe_fromstring, tostring as safe_tostring

# Add parent directory to path for imports
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from backend.document_adapter import BaseDocumentAdapter, Segment, SegmentType


class DocxAdapter(BaseDocumentAdapter):
    """Adapter for processing DOCX documents."""

    MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
    MAX_XML_SIZE = 10 * 1024 * 1024   # 10MB

    def __init__(self):
        super().__init__()

    def _validate_file_size(self, file_size: int) -> bool:
        """Validate file size is within limits."""
        return file_size <= self.MAX_FILE_SIZE

    def extract_segments(self, file_path: str) -> List[Segment]:
        """Extract text segments from DOCX file."""
        segments = []
        docx_path = Path(file_path)

        with zipfile.ZipFile(docx_path, 'r') as zf:
            # Read document.xml
            with zf.open('word/document.xml') as f:
                xml_content = f.read()

            if len(xml_content) > self.MAX_XML_SIZE:
                raise ValueError("XML content too large")

            # Parse XML
            root = safe_fromstring(xml_content)

            # Extract text from <w:t> elements
            namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            segment_id = 0

            for text_elem in root.findall('.//w:t', namespaces):
                if text_elem.text:
                    segment_id += 1
                    has_japanese = self._contains_japanese(text_elem.text)
                    word_count = self.count_words(text_elem.text)

                    segment = Segment(
                        id=str(segment_id),
                        text=text_elem.text,
                        segment_type=SegmentType.PARAGRAPH,
                        metadata={'position': f'w:t[{segment_id}]'},
                        has_japanese=has_japanese,
                        word_count=word_count
                    )
                    segments.append(segment)

        return segments

    def apply_translations(self, input_path: str, segments: List[Segment], output_path: str) -> None:
        """Apply translations to document."""
        # For now, just copy the file
        shutil.copy2(input_path, output_path)

    def get_metadata(self, file_path: str):
        """Get document metadata."""
        from backend.document_adapter import DocumentMetadata
        path = Path(file_path)
        stat = path.stat()

        return DocumentMetadata(
            filename=path.name,
            file_size=stat.st_size,
            created_at=None,
            modified_at=None
        )

    def _contains_japanese(self, text: str) -> bool:
        """Check if text contains Japanese characters."""
        # Check for Hiragana, Katakana, and Kanji
        japanese_pattern = re.compile(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF]')
        return bool(japanese_pattern.search(text))

    def count_words(self, text: str) -> int:
        """Count words in text, handling Japanese properly."""
        # For Japanese, count characters as words
        if self._contains_japanese(text):
            # Count Japanese characters and words separately
            japanese_chars = len(re.findall(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF]', text))
            english_words = len(re.findall(r'\b[a-zA-Z]+\b', text))
            return japanese_chars + english_words
        else:
            # For English, count words normally
            return len(text.split())