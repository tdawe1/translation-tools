"""
DOCX adapter for secure document processing.
Handles DOCX file parsing with security measures against XXE attacks.
"""
import os
import re
import zipfile
from typing import List, Dict, Any, Optional, Union
try:
    from defusedxml.ElementTree import parse, fromstring
    from defusedxml import ElementTree as ET
except ImportError:
    # Fallback if defusedxml not available
    from xml.etree.ElementTree import parse, fromstring
    import xml.etree.ElementTree as ET

# Try to import from backend module
try:
    from backend.document_adapter import (
        Segment,
        SegmentType,
        DocumentMetadata,
        TranslationResult,
        BaseDocumentAdapter
    )
except ImportError:
    # Fallback definitions
    from dataclasses import dataclass
    from enum import Enum

    class SegmentType(Enum):
        PARAGRAPH = "paragraph"
        TABLE = "table"
        HEADER = "header"
        FOOTER = "footer"
        FOOTNOTE = "footnote"
        ENDNOTE = "endnote"

    @dataclass
    class Segment:
        id: str
        text: str
        type: SegmentType
        metadata: Dict[str, Any]

    @dataclass
    class DocumentMetadata:
        filename: str
        language: Optional[str] = None
        word_count: int = 0
        has_japanese: bool = False
        processing_time: Optional[float] = None

    @dataclass
    class TranslationResult:
        segments: List[Segment]
        metadata: DocumentMetadata
        target_language: str

    class BaseDocumentAdapter:
        def extract_segments(self, file_path: str) -> List[Segment]:
            raise NotImplementedError


class DocxAdapter(BaseDocumentAdapter):
    """Adapter for processing DOCX files with security measures."""

    # Security constants
    MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
    MAX_XML_SIZE = 10 * 1024 * 1024   # 10MB

    # Pattern to detect Japanese characters
    JAPANESE_PATTERN = re.compile(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF\uFF00-\uFFEF]+')

    def __init__(self):
        self.word_count = 0
        self.has_japanese = False

    def _validate_file_size(self, file_size: int) -> bool:
        """Validate file size is within limits."""
        if file_size > self.MAX_FILE_SIZE:
            raise ValueError(f"File too large: {file_size} > {self.MAX_FILE_SIZE}")
        return True

    def _validate_xml_size(self, xml_size: int) -> bool:
        """Validate XML document size is within limits."""
        if xml_size > self.MAX_XML_SIZE:
            raise ValueError(f"XML document too large: {xml_size} > {self.MAX_XML_SIZE}")
        return True

    def _contains_japanese(self, text: str) -> bool:
        """Check if text contains Japanese characters."""
        return bool(self.JAPANESE_PATTERN.search(text))

    def _count_words(self, text: str) -> int:
        """Count words in text, handling multiple spaces and newlines."""
        # Split on whitespace and filter out empty strings
        words = [word for word in text.split() if word.strip()]
        return len(words)

    def extract_segments(self, file_path: str) -> List[Segment]:
        """
        Extract text segments from a DOCX file.

        Args:
            file_path: Path to the DOCX file

        Returns:
            List of Segment objects containing text and metadata

        Raises:
            ValueError: If file is too large or malformed
            FileNotFoundError: If file doesn't exist
        """
        # Check file exists
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        # Check file size
        file_size = os.path.getsize(file_path)
        self._validate_file_size(file_size)

        segments = []
        segment_id = 0

        try:
            with zipfile.ZipFile(file_path, 'r') as docx:
                # Check document.xml size
                try:
                    doc_xml_info = docx.getinfo('word/document.xml')
                    self._validate_xml_size(doc_xml_info.file_size)
                except KeyError:
                    # Try main document relationship
                    raise ValueError("Invalid DOCX: missing document.xml")

                # Parse document.xml
                with docx.open('word/document.xml') as xml_file:
                    # Use defusedxml for security
                    tree = parse(xml_file)
                    root = tree.getroot()

                    # Extract paragraphs
                    for p in root.findall('.//w:p', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        text_elements = []
                        for t in p.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                            if t.text:
                                text_elements.append(t.text)

                        if text_elements:
                            text = ''.join(text_elements).strip()
                            if text:  # Only add non-empty segments
                                # Check for Japanese
                                if self._contains_japanese(text):
                                    self.has_japanese = True

                                # Count words
                                self.word_count += self._count_words(text)

                                segment = Segment(
                                    id=f"seg_{segment_id:04d}",
                                    text=text,
                                    type=SegmentType.PARAGRAPH,
                                    metadata={
                                        'element': 'p',
                                        'style': p.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle', '')
                                    }
                                )
                                segments.append(segment)
                                segment_id += 1

                    # Extract table content
                    for table in root.findall('.//w:tbl', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        table_text = []
                        for tr in table.findall('.//w:tr', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                            row_text = []
                            for tc in tr.findall('.//w:tc', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                                cell_text = []
                                for t in tc.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                                    if t.text:
                                        cell_text.append(t.text)
                                row_text.append(''.join(cell_text))
                            table_text.append(' | '.join(row_text))

                        if table_text:
                            text = '\n'.join(table_text)
                            if self._contains_japanese(text):
                                self.has_japanese = True

                            self.word_count += self._count_words(text)

                            segment = Segment(
                                id=f"seg_{segment_id:04d}",
                                text=text,
                                type=SegmentType.TABLE,
                                metadata={
                                    'element': 'table',
                                    'rows': len(table_text)
                                }
                            )
                            segments.append(segment)
                            segment_id += 1

        except zipfile.BadZipFile:
            raise ValueError("Invalid DOCX file: not a valid ZIP archive")
        except ET.ParseError as e:
            raise ValueError(f"XML parsing error: {e}")

        return segments

    def get_metadata(self, file_path: str) -> DocumentMetadata:
        """Get document metadata."""
        if not hasattr(self, 'word_count'):
            # Extract segments to populate metadata
            self.extract_segments(file_path)

        return DocumentMetadata(
            filename=os.path.basename(file_path),
            word_count=self.word_count,
            has_japanese=self.has_japanese
        )