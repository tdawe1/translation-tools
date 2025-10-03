#!/usr/bin/env python3
"""
DOCX Adapter for Japanese-to-English translation pipeline.

Extracts text segments with run-level metadata from DOCX documents,
translates them using the existing batch translation system,
and applies translations back while preserving formatting.

Usage:
  python docx_adapter.py --extract input.docx
  python docx_adapter.py --translate segments.json
  python docx_adapter.py --apply input.docx translations.json output.docx
"""

import argparse
import asyncio
import json
import logging
import os
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
from xml.etree import ElementTree as ET

# Add parent directory for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import shared protocol definitions
from backend.document_adapter import (
    BaseDocumentAdapter,
    Segment,
    SegmentType,
    TranslationResult,
)
from backend.document_adapter import DocumentMetadata as SharedDocumentMetadata

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# XML namespaces
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _create_secure_xml_parser():
    """Create a secure XML parser that prevents XXE attacks."""
    try:
        # Try to create parser with entity resolution disabled
        parser = ET.XMLParser(resolve_entities=False)
        return parser
    except TypeError:
        # Older Python versions - use defusedxml for security if available
        try:
            import defusedxml.ElementTree as safe_ET
            # Return a wrapper that works like ET.XMLParser
            class SafeParserWrapper:
                def __init__(self):
                    self._safe_ET = safe_ET

                def feed(self, data):
                    pass  # Not needed for fromstring

                def close(self):
                    pass
            return SafeParserWrapper()
        except ImportError:
            logger.warning("defusedxml not installed - falling back to insecure XML parsing")
            # Last resort - use basic parser but at least document the risk
            return ET.XMLParser()


@dataclass
class DocxSpecificMetadata:
    """DOCX-specific metadata for internal use."""
    title: Optional[str] = None
    author: Optional[str] = None
    created: Optional[str] = None
    modified: Optional[str] = None
    language: Optional[str] = None
    has_headers: bool = False
    has_footers: bool = False
    has_footnotes: bool = False
    has_endnotes: bool = False
    paragraph_count: int = 0
    table_count: int = 0


class DocxAdapter(BaseDocumentAdapter):
    """Adapter for extracting and applying translations to DOCX documents."""

    # Security limits
    MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB max file size
    MAX_XML_SIZE = 50 * 1024 * 1024    # 50MB max for individual XML files

    def __init__(self, docx_path: Union[str, Path] = None):
        super().__init__()
        self.docx_path = str(docx_path) if docx_path else None
        self.segments: List[Segment] = []
        self.docx_metadata = DocxSpecificMetadata()
        self._run_index_map: Dict[str, int] = {}  # Track run indices per paragraph

    def extract_segments(self, file_path: Union[str, Path]) -> List[Segment]:
        """Extract text segments from the DOCX document."""
        if file_path is None:
            raise ValueError("File path cannot be None")

        docx_path = str(file_path) if file_path != self.docx_path else self.docx_path

        # Security: Validate file path to prevent directory traversal
        # Allow absolute paths but check for traversal attempts
        if not os.path.exists(docx_path):
            raise ValueError(f"File not found: {docx_path}")

        # Convert to absolute path and normalize
        abs_path = os.path.abspath(docx_path)
        if not abs_path.startswith(os.getcwd()) and not abs_path.startswith('/tmp/') and not abs_path.startswith('/var/tmp/'):
            # Only allow files in current directory or temp directories
            raise ValueError(f"Invalid file location: {docx_path}")

        # Check for path traversal patterns
        if '../' in docx_path or '..\\' in docx_path:
            raise ValueError(f"Path traversal not allowed: {docx_path}")

        # Validate file size
        path = Path(docx_path)
        if path.exists() and path.stat().st_size > self.MAX_FILE_SIZE:
            raise ValueError(f"File too large: {path.stat().st_size} bytes exceeds limit of {self.MAX_FILE_SIZE} bytes")

        logger.info(f"Extracting segments from {docx_path}")

        segments = []

        try:
            with zipfile.ZipFile(docx_path, 'r') as docx:
                # Security: Validate ZIP file contents
                for name in docx.namelist():
                    if name.startswith('/') or '..' in name:
                        raise ValueError(f"Invalid ZIP entry: {name}")

                # Extract document metadata
                self._extract_document_metadata(docx)

                # Extract main document content
                if 'word/document.xml' in docx.namelist():
                    document_xml = docx.read('word/document.xml')
                    segments.extend(self._extract_from_document(document_xml, 'word/document.xml'))

                # Extract headers/footers if they exist
                header_files = [f for f in docx.namelist() if f.startswith('word/header')]
                footer_files = [f for f in docx.namelist() if f.startswith('word/footer')]

                for header_file in header_files:
                    header_xml = docx.read(header_file)
                    segments.extend(self._extract_from_document(header_xml, header_file, SegmentType.HEADER))

                for footer_file in footer_files:
                    footer_xml = docx.read(footer_file)
                    segments.extend(self._extract_from_document(footer_xml, footer_file, SegmentType.FOOTER))

                # Extract footnotes and endnotes
                if 'word/footnotes.xml' in docx.namelist():
                    footnotes_xml = docx.read('word/footnotes.xml')
                    segments.extend(self._extract_from_document(footnotes_xml, 'word/footnotes.xml', SegmentType.FOOTNOTE))

                if 'word/endnotes.xml' in docx.namelist():
                    endnotes_xml = docx.read('word/endnotes.xml')
                    segments.extend(self._extract_from_document(endnotes_xml, 'word/endnotes.xml', SegmentType.ENDNOTE))

        except (zipfile.BadZipFile, FileNotFoundError, PermissionError) as e:
            logger.error(f"Failed to process DOCX file {docx_path}: {e}")
            return []

        logger.info(f"Extracted {len(segments)} segments")
        self.segments = segments
        return segments

    def collect_metadata(self, file_path: Union[str, Path]) -> SharedDocumentMetadata:
        """Collect metadata about the DOCX file."""
        path = Path(file_path)
        # Validate file path inline since _validate_file_path doesn't exist
        if not path.exists():
            raise ValueError(f"File not found: {path}")

        # Always extract metadata when called directly
        # This ensures collect_metadata works even without extract_segments being called first
        if self.docx_path != str(path):
            # Different file path, reset and re-extract
            self.docx_metadata = DocxSpecificMetadata()
            self.docx_path = str(path)

        with zipfile.ZipFile(path, 'r') as docx:
            # Always extract metadata to ensure fresh data
            self._extract_document_metadata(docx)

            # Extract paragraph and table counts from document.xml
            # Always re-extract to ensure fresh data when called directly
            if 'word/document.xml' in docx.namelist():
                try:
                    doc_xml_content = docx.read('word/document.xml')
                    # Use secure XML parsing
                    parser = _create_secure_xml_parser()
                    if hasattr(parser, '_safe_ET'):
                        # Using defusedxml
                        doc_xml = parser._safe_ET.fromstring(doc_xml_content)
                    else:
                        # Using standard ET
                        doc_xml = ET.fromstring(doc_xml_content, parser=parser)
                    w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

                    # Count paragraphs (only those with text)
                    paragraph_count = 0
                    for p in doc_xml.iter(f'{w_ns}p'):
                        if p.find(f'.//{w_ns}t') is not None:
                            paragraph_count += 1

                    if self.docx_metadata.paragraph_count == 0:
                        self.docx_metadata.paragraph_count = paragraph_count

                    # Count tables
                    if self.docx_metadata.table_count == 0:
                        self.docx_metadata.table_count = len(list(doc_xml.iter(f'{w_ns}tbl')))
                except Exception:
                    # If parsing fails, counts remain 0
                    pass

        # Convert to shared DocumentMetadata format
        return SharedDocumentMetadata(
            file_path=str(path),
            format="docx",
            word_count=self.docx_metadata.paragraph_count * 15,  # Rough estimate
            character_count=None,
            segment_count=self.docx_metadata.paragraph_count,
            has_headers_footers=self.docx_metadata.has_headers or self.docx_metadata.has_footers,
            has_footnotes=self.docx_metadata.has_footnotes,
            has_comments=False,  # TODO: Implement comment extraction
            has_fields=False,  # TODO: Implement field detection
            tables=[],  # TODO: Extract table information
            languages=[self.docx_metadata.language] if self.docx_metadata.language else [],
            custom_properties={
                'title': self.docx_metadata.title,
                'author': self.docx_metadata.author,
                'created': self.docx_metadata.created,
                'modified': self.docx_metadata.modified,
                'paragraph_count': self.docx_metadata.paragraph_count,
                'table_count': self.docx_metadata.table_count
            }
        )

    def supported_formats(self) -> List[str]:
        """Return supported file formats."""
        return ["docx"]

    def _extract_document_metadata(self, docx: zipfile.ZipFile):
        """Extract document-level metadata."""
        # Core properties
        if 'docProps/core.xml' in docx.namelist():
            core_xml_content = docx.read('docProps/core.xml')
            # Use secure XML parsing
            parser = _create_secure_xml_parser()
            if hasattr(parser, '_safe_ET'):
                # Using defusedxml
                core_xml = parser._safe_ET.fromstring(core_xml_content)
            else:
                # Using standard ET
                core_xml = ET.fromstring(core_xml_content, parser=parser)

            # ElementTree automatically expands namespaces
            title = core_xml.find('{http://purl.org/dc/elements/1.1/}title')
            if title is not None:
                self.docx_metadata.title = title.text if title is not None else ''

            creator = core_xml.find('{http://purl.org/dc/elements/1.1/}creator')
            if creator is not None:
                self.docx_metadata.author = creator.text if creator is not None else ''

            created = core_xml.find('{http://purl.org/dc/terms/}created')
            if created is not None:
                self.docx_metadata.created = created.text if created is not None else ''

            modified = core_xml.find('{http://purl.org/dc/terms/}modified')
            if modified is not None:
                self.docx_metadata.modified = modified.text if modified is not None else ''

        # Check for special sections
        self.docx_metadata.has_headers = any(f.startswith('word/header') for f in docx.namelist())
        self.docx_metadata.has_footers = any(f.startswith('word/footer') for f in docx.namelist())
        self.docx_metadata.has_footnotes = 'word/footnotes.xml' in docx.namelist()
        self.docx_metadata.has_endnotes = 'word/endnotes.xml' in docx.namelist()

        # Document settings
        if 'word/settings.xml' in docx.namelist():
            settings_content = docx.read('word/settings.xml')
            # Use secure XML parsing
            parser = _create_secure_xml_parser()
            if hasattr(parser, '_safe_ET'):
                # Using defusedxml
                settings_xml = parser._safe_ET.fromstring(settings_content)
            else:
                # Using standard ET
                settings_xml = ET.fromstring(settings_content, parser=parser)

            # Build namespace map
            ns_map = {}
            for key, value in settings_xml.attrib.items():
                if key.startswith('xmlns:'):
                    if value == 'http://schemas.openxmlformats.org/wordprocessingml/2006/main':
                        ns_map['w'] = f'{{{value}}}'

            w_ns = ns_map.get('w', W_NS)
            default_lang = settings_xml.find(f'.//{w_ns}defaultLanguage')
            self.docx_metadata.language = default_lang.get(f'{w_ns}val') if default_lang is not None else ''

        # Set defaults to empty strings if still None
        self.docx_metadata.title = self.docx_metadata.title or ''
        self.docx_metadata.author = self.docx_metadata.author or ''
        self.docx_metadata.created = self.docx_metadata.created or ''
        self.docx_metadata.modified = self.docx_metadata.modified or ''
        self.docx_metadata.language = self.docx_metadata.language or ''

    def _extract_from_document(self, xml_content: bytes, file_path: str = 'word/document.xml', segment_type: SegmentType = SegmentType.PARAGRAPH) -> List[Segment]:
        """Extract segments from a document XML file."""
        try:
            # Security: Always use a secure parser that prevents XXE attacks
            parser = _create_secure_xml_parser()
            try:
                if hasattr(parser, '_safe_ET'):
                    # Using defusedxml
                    root = parser._safe_ET.fromstring(xml_content)
                else:
                    # Using standard ET
                    root = ET.fromstring(xml_content, parser=parser)
            except Exception as e:
                # Handle defusedxml security exceptions
                if "EntitiesForbidden" in str(e) or "ExternalReferenceForbidden" in str(e):
                    raise ValueError(f"Security risk detected in XML: {e}")
                raise
        except (ET.ParseError, ValueError) as e:
            logger.error(f"Failed to parse XML in {file_path}: {e}")
            return []
        segments = []

        # Extract namespace mapping from the root element
        # Handle both prefixed and unprefixed namespaces
        w_ns = W_NS  # Default namespace
        for key, value in root.attrib.items():
            if key == 'xmlns:w':
                w_ns = f"{{{value}}}"
            elif key.startswith('xmlns:'):
                # Store other namespace mappings if needed
                pass

        # Track paragraph indices
        paragraph_idx = 0

        # Process all paragraphs
        for p in root.iter(f'{w_ns}p'):
            # Skip empty paragraphs
            if p.find(f'.//{w_ns}t') is None:
                paragraph_idx += 1
                continue

            # Get paragraph style and list properties
            p_style = None
            list_properties = None
            p_pr = p.find(f'{w_ns}pPr')
            if p_pr is not None:
                p_style_elem = p_pr.find(f'{w_ns}pStyle')
                if p_style_elem is not None:
                    p_style = p_style_elem.get(f'{w_ns}val')

                # Extract list numbering properties
                numPr = p_pr.find(f'{w_ns}numPr')
                if numPr is not None:
                    ilvl_elem = numPr.find(f'{w_ns}ilvl')
                    numId_elem = numPr.find(f'{w_ns}numId')
                    list_properties = {
                        'ilvl': ilvl_elem.get(f'{w_ns}val') if ilvl_elem is not None else None,
                        'numId': numId_elem.get(f'{w_ns}val') if numId_elem is not None else None
                    }

            # Check if paragraph is in a table
            table_context = None
            if file_path == 'word/document.xml':
                parent_table = self._find_parent_table(p, root, w_ns)
                if parent_table is not None:
                    table_context = {
                        'type': 'table',
                        'nesting_depth': parent_table['nesting_depth'],
                        'structure': parent_table['structure']
                    }

            # Extract runs from paragraph
            run_idx = 0
            for r in p.iter(f'{w_ns}r'):
                # Get run properties and store full rPr XML for preservation
                metadata = self._extract_run_metadata(r, w_ns)
                r_pr_elem = r.find(f'{w_ns}rPr')
                original_r_pr = ET.tostring(r_pr_elem, encoding='unicode') if r_pr_elem is not None else None
                metadata['original_rPr'] = original_r_pr  # Store raw XML for exact copy

                # Get paragraph properties for pPr preservation
                p_pr_elem = p.find(f'{w_ns}pPr')
                original_p_pr = ET.tostring(p_pr_elem, encoding='unicode') if p_pr_elem is not None else None

                # Get text elements
                text_elements = []
                for t in r.iter(f'{w_ns}t'):
                    if t.text:
                        text_elements.append(t.text)

                # Join text elements (handling cases where text is split across multiple <w:t> elements)
                if text_elements:
                    text = ''.join(text_elements)

                    # Always create segment ID for consistent indexing
                    segment_id = f"{file_path.replace('/', '_')}_{paragraph_idx}_{run_idx}"

                    # Only create segment if there's Japanese text
                    has_japanese, _ = self._contains_japanese(text)
                    if has_japanese:
                        segment = Segment(
                            id=segment_id,
                            text=text.strip(),
                            segment_type=segment_type,
                            file_path=file_path,
                            paragraph_index=paragraph_idx,
                            run_index=run_idx,
                            metadata=metadata,
                            context={
                                'p_style': p_style,
                                'table_context': table_context,
                                'original_pPr': original_p_pr,
                                'list_properties': list_properties
                            }
                        )
                        segments.append(segment)

                    run_idx += 1  # Always increment run index

            paragraph_idx += 1

        # Update metadata counts
        if file_path == 'word/document.xml':
            self.docx_metadata.paragraph_count = paragraph_idx
            self.docx_metadata.table_count = len(list(root.iter(f'{w_ns}tbl')))

        return segments

    def _analyze_table_structure(self, table_element, w_ns: str) -> Dict[str, Any]:
        """Analyze table structure for merged cells and nested tables."""
        structure = {
            'grid_span': {},  # Maps (row_idx, col_idx) -> span
            'v_merge': {},     # Maps (row_idx, col_idx) -> merge info
            'nested_tables': {},  # Maps (row_idx, col_idx) -> nested table
            'tbl_grid': None,
            'tc_pr_map': {}    # Maps (row_idx, col_idx) -> cell properties
        }

        # Extract table grid
        tbl_grid = table_element.find(f'{w_ns}tblGrid')
        if tbl_grid is not None:
            structure['tbl_grid'] = tbl_grid
            grid_cols = tbl_grid.findall(f'{w_ns}gridCol')
            for col_idx, grid_col in enumerate(grid_cols):
                structure['grid_span'][(0, col_idx)] = 1  # Default span

        # Analyze each row
        rows = table_element.findall(f'{w_ns}tr')
        for row_idx, row in enumerate(rows):
            cells = row.findall(f'{w_ns}tc')
            col_idx = 0

            for cell in cells:
                # Check for gridSpan (horizontal merge)
                grid_span = cell.find(f'{w_ns}tcPr/{w_ns}gridSpan')
                if grid_span is not None:
                    span_val = int(grid_span.get(f'{w_ns}val', 1))
                    structure['grid_span'][(row_idx, col_idx)] = span_val
                    col_idx += span_val
                else:
                    structure['grid_span'][(row_idx, col_idx)] = 1
                    col_idx += 1

                # Check for vMerge (vertical merge)
                v_merge = cell.find(f'{w_ns}tcPr/{w_ns}vMerge')
                if v_merge is not None:
                    merge_type = v_merge.get(f'{w_ns}val', 'continue')
                    structure['v_merge'][(row_idx, col_idx)] = merge_type

                # Store cell properties
                tc_pr = cell.find(f'{w_ns}tcPr')
                if tc_pr is not None:
                    structure['tc_pr_map'][(row_idx, col_idx)] = tc_pr

                # Check for nested tables
                nested_tbl = cell.find(f'{w_ns}tbl')
                if nested_tbl is not None:
                    structure['nested_tables'][(row_idx, col_idx)] = nested_tbl

        return structure

    def _find_parent_table(self, element, root, w_ns=None):
        """Find if an element is inside a table and return table structure info."""
        # ElementTree doesn't have parent navigation, so we'll use a different approach
        parent_map = {}

        # Build the parent map recursively
        def build_parent_map(parent):
            for child in parent:
                parent_map[child] = parent
                build_parent_map(child)

        build_parent_map(root)

        # Find the table and track nesting depth
        current = element
        table_stack = []
        nesting_depth = 0

        while current in parent_map and current != root:
            parent = parent_map[current]
            if parent.tag.startswith(f'{w_ns or W_NS}tbl'):
                table_stack.append(parent)
                nesting_depth += 1
            current = parent

        if not table_stack:
            return None

        # Return the innermost table (deepest nesting)
        table = table_stack[0]

        # Analyze table structure
        structure = self._analyze_table_structure(table, w_ns or W_NS)
        structure['nesting_depth'] = nesting_depth

        return {
            'table': table,
            'structure': structure,
            'nesting_depth': nesting_depth
        }

    def _extract_run_metadata(self, run_element, w_ns: str) -> Dict[str, Any]:
        """Extract formatting metadata from a run element."""
        metadata: Dict[str, Any] = {}

        r_pr = run_element.find(f'{w_ns}rPr')
        if r_pr is None:
            return metadata

        # Bold
        b = r_pr.find(f'{w_ns}b')
        metadata['bold'] = b is not None

        # Italic
        i = r_pr.find(f'{w_ns}i')
        metadata['italic'] = i is not None

        # Underline
        u = r_pr.find(f'{w_ns}u')
        metadata['underline'] = u is not None

        # Color
        color = r_pr.find(f'{w_ns}color')
        if color is not None:
            metadata['color'] = color.get(f'{w_ns}val')

        # Size
        sz = r_pr.find(f'{w_ns}sz')
        if sz is not None:
            size_val = sz.get(f'{w_ns}val')
            if size_val:
                metadata['size'] = float(size_val) / 2  # Convert from half-points to points

        # Font
        r_fonts = r_pr.find(f'{w_ns}rFonts')
        if r_fonts is not None:
            metadata['font'] = r_fonts.get(f'{w_ns}ascii') or r_fonts.get(f'{w_ns}hAnsi')

        # Language
        lang = r_pr.find(f'{w_ns}lang')
        if lang is not None:
            metadata['language'] = lang.get(f'{w_ns}val')

        return metadata

    def _contains_japanese(self, text: str) -> tuple[bool, float]:
        """Check if text contains Japanese characters and estimate expansion ratio."""
        # Unicode ranges for Japanese characters
        jp_ranges = [
            (0x3040, 0x309F),  # Hiragana
            (0x30A0, 0x30FF),  # Katakana
            (0x31F0, 0x31FF),  # Katakana phonetic extensions
            (0x3400, 0x4DBF),  # CJK Unified Ideographs Extension A
            (0x4E00, 0x9FFF),  # CJK Unified Ideographs
            (0xFF00, 0xFFEF),  # Fullwidth forms
        ]

        jp_count = 0
        total_chars = len(text)
        for char in text:
            code = ord(char)
            for start, end in jp_ranges:
                if start <= code <= end:
                    jp_count += 1
                    break

        has_japanese = jp_count > 0
        if total_chars == 0:
            return has_japanese, 1.0

        # Estimate expansion ratio (JP ~0.7x EN length; 1.5x for full JP)
        jp_ratio = jp_count / total_chars
        expansion_ratio = 1.0 + (0.5 * jp_ratio)  # 1.5x for 100% JP, 1.0 for 0%
        return has_japanese, expansion_ratio

    def apply_translations(self,
                         file_path: Union[str, Path],
                         translations: List[Segment],
                         output_path: Optional[Union[str, Path]] = None) -> TranslationResult:
        """Apply translations to the DOCX document."""
        input_path = Path(file_path)
        if output_path is None:
            output_path = input_path.with_suffix(f'_translated{input_path.suffix}')
        else:
            output_path = Path(output_path)

        logger.info(f"Applying translations to {input_path} -> {output_path}")

        # Create a copy of the input file
        import shutil
        shutil.copy2(input_path, output_path)

        # Create translation lookup by segment ID
        translation_map = {seg.id: seg.text for seg in translations}

        # Process each file that contains segments
        processed_files = set()
        segments_translated = 0
        words_translated = 0

        # Create a temporary file for the output
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_file:
            temp_path = temp_file.name

        # First, collect all file data with size limits
        file_data = {}
        with zipfile.ZipFile(output_path, 'r') as docx_in:
            for name in docx_in.namelist():
                info = docx_in.getinfo(name)

                # Skip files that are too large
                if info.file_size > self.MAX_FILE_SIZE:
                    logger.warning(f"Skipping file larger than MAX_FILE_SIZE: {name} ({info.file_size} bytes)")
                    continue

                # For XML files, check against XML size limit
                if name.startswith('word/') and name.endswith('.xml'):
                    if info.file_size > self.MAX_XML_SIZE:
                        logger.warning(f"Skipping large XML file: {name} ({info.file_size} bytes)")
                        file_data[name] = docx_in.read(name)
                        continue

                file_data[name] = docx_in.read(name)

        # Now process the files to temporary output
        with zipfile.ZipFile(temp_path, 'w') as docx_out:
            for name, data in file_data.items():
                if name.startswith('word/') and name.endswith('.xml'):
                    try:
                        # Use secure XML parsing
                        parser = _create_secure_xml_parser()

                        # Parse based on parser type
                        if hasattr(parser, '_safe_ET'):
                            # Using defusedxml
                            root = parser._safe_ET.fromstring(data)
                        else:
                            # Using standard ET
                            root = ET.fromstring(data, parser=parser)

                        modified, seg_count, word_count = self._apply_translations_to_xml_v2(
                            root, translation_map, name, self.segments
                        )

                        if modified:
                            xml_str = ET.tostring(root, encoding='utf-8', method='xml')
                            docx_out.writestr(name, xml_str)
                            processed_files.add(name)
                            segments_translated += seg_count
                            words_translated += word_count
                        else:
                            docx_out.writestr(name, data)
                    except ET.ParseError as e:
                        logger.warning(f"Failed to parse {name}: {e}")
                        docx_out.writestr(name, data)
                else:
                    docx_out.writestr(name, data)

        logger.info(f"Applied translations to {len(processed_files)} files")

        # Replace the original file with the temporary file
        import shutil
        shutil.move(temp_path, output_path)

        # Calculate actual word count from segments
        actual_word_count = sum(len(seg.text.split()) for seg in self.segments if seg.text)

        # Convert DocxSpecificMetadata to SharedDocumentMetadata
        shared_metadata = SharedDocumentMetadata(
            file_path=str(output_path),
            format="docx",
            word_count=actual_word_count or self.docx_metadata.paragraph_count * 15,  # Use actual or fallback
            character_count=None,
            segment_count=len(translations),
            has_headers_footers=self.docx_metadata.has_headers or self.docx_metadata.has_footers,
            has_footnotes=self.docx_metadata.has_footnotes,
            has_comments=False,
            has_fields=False,
            tables=[],
            languages=[self.docx_metadata.language] if self.docx_metadata.language else [],
            custom_properties={
                'title': self.docx_metadata.title,
                'author': self.docx_metadata.author,
                'paragraph_count': self.docx_metadata.paragraph_count,
                'table_count': self.docx_metadata.table_count
            }
        )

        return TranslationResult(
            output_path=str(output_path),
            segments_translated=segments_translated,
            total_segments=len(translations),
            words_translated=words_translated,
            total_words=sum(seg.word_count for seg in translations),
            cache_hits=0,
            processing_time=0.0,
            warnings=[],
            artifacts={'metadata': shared_metadata}
        )

    def _apply_translations_to_xml(self, root: ET.Element, translation_map: Dict[str, str], file_path: str) -> bool:
        """Apply translations to XML content."""
        modified = False
        paragraph_idx = 0

        # Extract namespace from root element
        w_ns = W_NS  # Default namespace
        for key, value in root.attrib.items():
            if key == 'xmlns:w':
                w_ns = f"{{{value}}}"
                break

        for p in root.iter(f'{w_ns}p'):
            run_idx = 0
            for r in p.iter(f'{w_ns}r'):
                # Generate segment ID
                segment_id = f"{file_path.replace('/', '_')}_{paragraph_idx}_{run_idx}"

                if segment_id in translation_map:
                    # Find all text elements in this run
                    text_elements = list(r.iter(f'{w_ns}t'))
                    if text_elements:
                        # Apply translation to first text element
                        translation = translation_map[segment_id]
                        text_elements[0].text = translation

                        # Clear other text elements in this run
                        for t in text_elements[1:]:
                            t.text = ""

                        modified = True

                run_idx += 1

            if any(r.find(f'{w_ns}t') is not None for r in p.iter(f'{w_ns}r')):
                paragraph_idx += 1

        return modified

    def _apply_translations_to_xml_v2(self, root: ET.Element, translation_map: Dict[str, str],
                                     file_path: str, segments: List[Segment] = None) -> tuple[bool, int, int]:
        """Apply translations to XML content, preserving styles exactly, returning stats."""
        modified = False
        segments_translated = 0
        words_translated = 0

        # Use provided segments or fall back to self.segments
        source_segments = segments if segments is not None else self.segments

        # Create segment lookup map for O(1) access instead of O(n) search
        segment_map = {seg.id: seg for seg in source_segments}

        w_ns = W_NS  # Default namespace
        for key, value in root.attrib.items():
            if key == 'xmlns:w':
                w_ns = f"{{{value}}}"
                break

        paragraph_idx = 0

        for p in root.iter(f'{w_ns}p'):
            # Skip empty paragraphs (matches extraction pattern)
            if p.find(f'.//{w_ns}t') is None:
                paragraph_idx += 1
                continue

            run_idx = 0
            paragraph_modified = False

            for r in p.iter(f'{w_ns}r'):
                # Generate segment ID (matches extraction pattern exactly)
                segment_id = f"{file_path.replace('/', '_')}_{paragraph_idx}_{run_idx}"

                if segment_id in translation_map:
                    # Get the original segment metadata (O(1) lookup)
                    original_segment = segment_map.get(segment_id)

                    # Check if this run is in a nested table - if so, skip translation
                    if original_segment and original_segment.context:
                        table_context = original_segment.context.get('table_context')
                        if table_context and table_context.get('nesting_depth', 0) > 0:
                            # Check if this run contains table structure elements
                            # Only translate leaf text elements (not table structure)
                            has_table_elements = any(
                                elem.tag.startswith(f'{w_ns}tbl') or
                                elem.tag.startswith(f'{w_ns}tr') or
                                elem.tag.startswith(f'{w_ns}tc')
                                for elem in r.iter()
                            )

                            if has_table_elements:
                                # Skip table structure elements
                                run_idx += 1
                                continue

                    if original_segment and 'original_rPr' in original_segment.metadata:
                        # Preserve exact rPr structure
                        original_r_pr_xml = original_segment.metadata['original_rPr']
                        if original_r_pr_xml:
                            # Remove existing rPr
                            existing_r_pr = r.find(f'{w_ns}rPr')
                            if existing_r_pr is not None:
                                r.remove(existing_r_pr)
                            # Parse and insert original rPr
                            try:
                                parser = _create_secure_xml_parser()
                                if hasattr(parser, '_safe_ET'):
                                    # Using defusedxml
                                    new_r_pr = parser._safe_ET.fromstring(original_r_pr_xml)
                                else:
                                    # Using standard ET
                                    new_r_pr = ET.fromstring(original_r_pr_xml, parser=parser)

                                # Validate the parsed XML
                                if new_r_pr.tag is None:
                                    raise ValueError("Invalid XML structure in rPr")

                                r.insert(0, new_r_pr)
                            except Exception as e:
                                logger.warning(f"Failed to restore rPr for {segment_id}: {e}")

                    # Find all text elements in this run
                    text_elements = list(r.iter(f'{w_ns}t'))
                    if text_elements:
                        # Apply translation to first text element
                        translation = translation_map[segment_id]
                        text_elements[0].text = translation

                        # Clear other text elements in this run
                        for t in text_elements[1:]:
                            t.text = ""

                        modified = True
                        paragraph_modified = True
                        segments_translated += 1
                        words_translated += len(translation.split())

                run_idx += 1  # Always increment run index (matches extraction)

            # Preserve paragraph properties if available
            if paragraph_modified:
                # Check if we have original paragraph data
                # Use exact match instead of startswith to avoid ID conflicts
                original_segment = None
                # Look for any segment from this paragraph that has pPr data
                for run_idx in range(100):  # Reasonable limit for runs per paragraph
                    test_id = f"{file_path.replace('/', '_')}_{paragraph_idx}_{run_idx}"
                    candidate = segment_map.get(test_id)
                    if candidate and candidate.context and 'original_pPr' in candidate.context:
                        original_segment = candidate
                        break

                if original_segment and original_segment.context.get('original_pPr'):
                    # Restore original pPr
                    original_p_pr_xml = original_segment.context['original_pPr']
                    if original_p_pr_xml:
                        existing_p_pr = p.find(f'{w_ns}pPr')
                        if existing_p_pr is not None:
                            p.remove(existing_p_pr)
                        try:
                            parser = _create_secure_xml_parser()
                            if hasattr(parser, '_safe_ET'):
                                # Using defusedxml
                                new_p_pr = parser._safe_ET.fromstring(original_p_pr_xml)
                            else:
                                # Using standard ET
                                new_p_pr = ET.fromstring(original_p_pr_xml, parser=parser)

                            # Validate the parsed XML
                            if new_p_pr.tag is None:
                                raise ValueError("Invalid XML structure in pPr")

                            p.insert(0, new_p_pr)
                        except Exception as e:
                            logger.warning(f"Failed to restore pPr for paragraph {paragraph_idx}: {e}")

                # Preserve list numbering properties
                if original_segment and original_segment.context and 'list_properties' in original_segment.context:
                    list_props = original_segment.context['list_properties']
                    if list_props and (list_props.get('numId') or list_props.get('ilvl')):
                        # Ensure pPr exists
                        p_pr = p.find(f'{w_ns}pPr')
                        if p_pr is None:
                            p_pr = ET.SubElement(p, f'{w_ns}pPr')

                        # Ensure numPr exists
                        num_pr = p_pr.find(f'{w_ns}numPr')
                        if num_pr is None:
                            num_pr = ET.SubElement(p_pr, f'{w_ns}numPr')

                        # Set list level
                        if list_props.get('ilvl'):
                            ilvl = num_pr.find(f'{w_ns}ilvl')
                            if ilvl is None:
                                ilvl = ET.SubElement(num_pr, f'{w_ns}ilvl')
                            ilvl.set(f'{w_ns}val', str(list_props['ilvl']))

                        # Set list number reference
                        if list_props.get('numId'):
                            num_id = num_pr.find(f'{w_ns}numId')
                            if num_id is None:
                                num_id = ET.SubElement(num_pr, f'{w_ns}numId')
                            num_id.set(f'{w_ns}val', str(list_props['numId']))

            paragraph_idx += 1  # Always increment paragraph index (matches extraction)

        return modified, segments_translated, words_translated

    def generate_bilingual_json(self, translations: List[Dict[str, Any]], output_path: str):
        """Generate a bilingual JSON file for QA."""
        bilingual_data = []

          # Create segment lookup map for O(1) access
        segment_lookup = {s.id: s for s in self.segments}

        for trans in translations:
            segment = segment_lookup.get(trans['id'])
            if segment:
                bilingual_data.append({
                    'id': segment.id,
                    'file_path': segment.file_path,
                    'paragraph_index': segment.paragraph_index,
                    'run_index': segment.run_index,
                    'original': segment.text,
                    'translated': trans['translation'],
                    'context': segment.context,
                    'metadata': segment.metadata
                })

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(bilingual_data, f, ensure_ascii=False, indent=2)

        logger.info(f"Generated bilingual JSON: {output_path}")


def segment_to_dict(segment):
    """Convert Segment to JSON-serializable dict."""
    from dataclasses import asdict
    result = asdict(segment)
    # Convert SegmentType enum to string
    if 'segment_type' in result and hasattr(result['segment_type'], 'value'):
        result['segment_type'] = result['segment_type'].value
    return result


def main():
    parser = argparse.ArgumentParser(description='DOCX Translation Adapter')
    parser.add_argument('--extract', help='Extract segments from DOCX file')
    parser.add_argument('--translate', help='Translate segments JSON file')
    parser.add_argument('--apply', nargs=3, metavar=('INPUT_DOCX', 'TRANSLATIONS_JSON', 'OUTPUT_DOCX'),
                       help='Apply translations to DOCX file')
    parser.add_argument('--bilingual', help='Generate bilingual JSON from translations')

    args = parser.parse_args()

    if args.extract:
        adapter = DocxAdapter(args.extract)
        segments = adapter.extract_segments(args.extract)
        metadata = adapter.collect_metadata(args.extract)

        # Save segments
        segments_file = args.extract.replace('.docx', '_segments.json')
        with open(segments_file, 'w', encoding='utf-8') as f:
            json.dump({
                'metadata': asdict(metadata),
                'segments': [segment_to_dict(s) for s in segments]
            }, f, ensure_ascii=False, indent=2)

        print(f"Extracted {len(segments)} segments to {segments_file}")

    elif args.translate:
        # Load segments from JSON file
        with open(args.translate, 'r', encoding='utf-8') as f:
            data = json.load(f)

        segments = data.get('segments', [])
        if not segments:
            print("No segments found in JSON file")
            return

        # Extract Japanese text for translation
        import os
        if not os.getenv("OPENAI_API_KEY"):
            print("ERROR: OPENAI_API_KEY environment variable is required")
            return

        # Import translation function
        sys.path.insert(0, str(Path(__file__).parent))
        import openai
        from translate_pptx_inplace import translate_batch

        # Prepare texts for translation
        texts_to_translate = []
        segment_map = []  # Map to track which segment each text belongs to

        for seg_data in segments:
            if seg_data.get('text') and any('\u3040' <= c <= '\u309f' or '\u30a0' <= c <= '\u30ff' or '\u4e00' <= c <= '\u9fff' for c in seg_data['text']):
                texts_to_translate.append(seg_data['text'])
                segment_map.append(seg_data)

        if not texts_to_translate:
            print("No Japanese text found to translate")
            return

        print(f"Translating {len(texts_to_translate)} segments...")

        # Create a simple args object for translate_batch
        class Args:
            max_output_tokens = None
            max_retries = 3
            on_batch_fail = "split"
            json_debug_dir = None
            concurrency = 1

        # Translate using batch translation
        client = openai.OpenAI()
        translated_texts = asyncio.run(translate_batch(texts_to_translate, args=Args(), model="gpt-4o-2024-08-06", client=client))

        # Create translations file with proper format
        translations = []
        for i, seg_data in enumerate(segment_map):
            translations.append({
                'id': seg_data['id'],
                'translation': translated_texts[i]
            })

        # Save translations
        translations_file = args.translate.replace('_segments.json', '_translations.json')
        with open(translations_file, 'w', encoding='utf-8') as f:
            json.dump({
                'translations': translations
            }, f, ensure_ascii=False, indent=2)

        print(f"Translated {len(translations)} segments to {translations_file}")
        print("Use --apply to apply translations to the document")

    elif args.apply:
        input_docx, translations_json, output_docx = args.apply

        # Load translations
        with open(translations_json, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Convert JSON data back to Segment objects
        translations = []
        for t_data in data.get('translations', []):
            # Handle both dictionary format and full segment format
            if isinstance(t_data, dict):
                if 'id' in t_data and 'translation' in t_data:
                    # Simple format: {id: ..., translation: ...}
                    translations.append(Segment(
                        id=t_data['id'],
                        text=t_data['translation'],
                        segment_type=SegmentType.PARAGRAPH
                    ))
                else:
                    # Full segment format - convert dict keys to Segment fields
                    translations.append(Segment(
                        id=t_data.get('id', ''),
                        text=t_data.get('text', ''),
                        segment_type=SegmentType(t_data.get('segment_type', 'paragraph')),
                        file_path=t_data.get('file_path'),
                        paragraph_index=t_data.get('paragraph_index'),
                        run_index=t_data.get('run_index'),
                        metadata=t_data.get('metadata', {}),
                        context=t_data.get('context')
                    ))

        adapter = DocxAdapter(input_docx)
        result = adapter.apply_translations(input_docx, translations, output_docx)

        # Generate bilingual file if requested
        if args.bilingual:
            adapter.generate_bilingual_json(data.get('translations', []), args.bilingual)

        print(f"Applied translations to {result}")

    else:
        parser.print_help()


# Register this adapter with the global registry (import at module level)
try:
    from backend.document_adapter import adapter_registry
    # Register the adapter class (registry will instantiate as needed)
    adapter_registry.register('docx', DocxAdapter)
except ImportError:
    pass  # Graceful fallback if backend not in path

if __name__ == '__main__':
    main()
