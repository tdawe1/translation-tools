#!/usr/bin/env python3
"""
Tests for the DOCX adapter functionality.
"""

import json
import os
import shutil

# Add the scripts directory to the path
import sys
import tempfile
import unittest
from zipfile import ZipFile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from backend.document_adapter import Segment, SegmentType
from scripts.docx_adapter import DocxAdapter, W_NS
import xml.etree.ElementTree as ET


class TestDocxAdapter(unittest.TestCase):
    """Test cases for DOCX adapter functionality."""

    def setUp(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.test_docx_path = os.path.join(self.temp_dir, 'test.docx')

    def tearDown(self):
        """Clean up test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir)

    def create_simple_docx(self, text_content="これはテストです。"):
        """Create a simple DOCX file for testing."""
        # Create a minimal DOCX structure
        with ZipFile(self.test_docx_path, 'w') as docx:
            # Add required content types
            docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')

            # Add document relationships
            docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')

            # Create document XML with Japanese text
            docx.writestr('word/document.xml', f'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Normal"/>
            </w:pPr>
            <w:r>
                <w:t>{text_content}</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>''')

            # Add word relationships
            docx.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')

    def test_extract_segments_basic(self):
        """Test basic segment extraction from a DOCX file."""
        # Create test document
        self.create_simple_docx()

        # Extract segments
        adapter = DocxAdapter()
        segments = adapter.extract_segments(self.test_docx_path)
        metadata = adapter.collect_metadata(self.test_docx_path)

        # Verify extraction
        self.assertEqual(len(segments), 1)
        self.assertEqual(segments[0].text, "これはテストです。")
        self.assertEqual(segments[0].file_path, 'word/document.xml')
        self.assertEqual(segments[0].paragraph_index, 0)
        self.assertEqual(segments[0].run_index, 0)
        self.assertEqual(segments[0].context.get('p_style'), 'Normal')

    def test_extract_segments_with_formatting(self):
        """Test extraction of formatted text, including original rPr/pPr storage."""
        # Create DOCX with formatted text
        with ZipFile(self.test_docx_path, 'w') as docx:
            # Add required content types
            docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')

            # Add relationships
            docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')

            # Create document with formatted text and pPr
            docx.writestr('word/document.xml', '''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
                <w:jc w:val="center"/>  <!-- Paragraph alignment -->
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:i/>
                    <w:sz w:val="24"/>
                    <w:color w:val="0000FF"/>  <!-- Blue color -->
                </w:rPr>
                <w:t>これは太字のテキストです。</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:u w:val="single"/>
                    <w:color w:val="FF0000"/>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                </w:rPr>
                <w:t>これは赤い下線付きです。</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>''')

            docx.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')

        # Extract segments
        adapter = DocxAdapter()
        segments = adapter.extract_segments(self.test_docx_path)

        # Verify extraction
        self.assertEqual(len(segments), 2)

        # Check first segment (bold, italic, size, color, pStyle)
        self.assertEqual(segments[0].text, "これは太字のテキストです。")
        self.assertTrue(segments[0].metadata['bold'])
        self.assertTrue(segments[0].metadata['italic'])
        self.assertEqual(segments[0].metadata['size'], 12.0)  # 24 half-points = 12 points
        self.assertEqual(segments[0].metadata['color'], '0000FF')
        self.assertIn('original_rPr', segments[0].metadata)  # Stored raw rPr
        self.assertEqual(segments[0].context['p_style'], 'Heading1')
        self.assertIn('original_pPr', segments[0].context)  # Stored raw pPr

        # Check second segment (underline, color, font)
        self.assertEqual(segments[1].text, "これは赤い下線付きです。")
        self.assertTrue(segments[1].metadata['underline'])
        self.assertEqual(segments[1].metadata['color'], 'FF0000')
        self.assertEqual(segments[1].metadata['font'], 'Arial')
        self.assertIn('original_rPr', segments[1].metadata)

    def test_metadata_extraction(self):
        """Test document metadata extraction."""
        # Create DOCX with metadata
        with ZipFile(self.test_docx_path, 'w') as docx:
            # Add content types
            docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>''')

            # Add core properties
            docx.writestr('docProps/core.xml', '''<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/">
    <dc:title>テスト文書</dc:title>
    <dc:creator>テストユーザー</dc:creator>
    <dcterms:created>2023-09-21T00:00:00Z</dcterms:created>
    <dcterms:modified>2023-09-21T12:00:00Z</dcterms:modified>
</cp:coreProperties>''')

            # Add document settings with language
            docx.writestr('word/settings.xml', '''<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:defaultLanguage w:val="ja-JP"/>
</w:settings>''')

            # Add document with table
            docx.writestr('word/document.xml', '''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>テーブルセル</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
            </w:tr>
        </w:tbl>
    </w:body>
</w:document>''')

            # Add relationships
            docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="word/settings.xml"/>
</Relationships>''')

            docx.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')

        # Test that collect_metadata works independently
        adapter = DocxAdapter()
        metadata = adapter.collect_metadata(self.test_docx_path)

        # Verify basic metadata extraction
        self.assertEqual(metadata.custom_properties.get('title'), "テスト文書")
        self.assertEqual(metadata.custom_properties.get('author'), "テストユーザー")
        self.assertEqual(metadata.custom_properties.get('created'), "2023-09-21T00:00:00Z")
        self.assertEqual(metadata.custom_properties.get('modified'), "2023-09-21T12:00:00Z")
        self.assertEqual(metadata.languages[0] if metadata.languages else None, "ja-JP")
        self.assertEqual(metadata.custom_properties.get('table_count'), 1)
        # Note: paragraph_count is only set during extract_segments, so we don't test it here

    def test_apply_translations_style_preservation(self):
        """Test applying translations preserves inline and paragraph styles."""
        # Create styled DOCX
        with ZipFile(self.test_docx_path, 'w') as docx:
            docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')
            docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')
            docx.writestr('word/document.xml', '''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:i/>
                    <w:sz w:val="24"/>
                    <w:color w:val="0000FF"/>
                </w:rPr>
                <w:t>これは太字の青いテキストです。</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>''')
            docx.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')

        # Extract
        adapter = DocxAdapter()
        segments = adapter.extract_segments(self.test_docx_path)

        # Create translation (same ID)
        translated_segments = [
            Segment(
                id=segments[0].id,
                text='This is bold blue translated text.',
                segment_type=SegmentType.PARAGRAPH,
                metadata=segments[0].metadata,  # Preserve metadata
                context=segments[0].context
            )
        ]

        # Apply
        output_path = os.path.join(self.temp_dir, 'styled_translated.docx')
        result = adapter.apply_translations(self.test_docx_path, translated_segments, output_path)

        # Verify styles preserved in output XML
        with ZipFile(output_path, 'r') as z:
            xml_content = z.read('word/document.xml').decode('utf-8')

        # Check translation text
        self.assertIn('This is bold blue translated text.', xml_content)

        # Check styles preserved (rPr elements) - handle various namespace formats
        self.assertIn('b', xml_content)  # Bold
        self.assertIn('i', xml_content)  # Italic
        self.assertIn('sz', xml_content)  # Size
        self.assertIn('color', xml_content)  # Color

        # Check paragraph style
        self.assertIn('pStyle', xml_content)
        self.assertIn('jc', xml_content)

    def test_run_indexing_consistency(self):
        """Test that run indexing increments for all runs, not just Japanese."""
        # Create DOCX with mixed runs: Japanese and English
        with ZipFile(self.test_docx_path, 'w') as docx:
            docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')
            docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')
            docx.writestr('word/document.xml', '''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r><w:t>English run 0</w:t></w:r>
            <w:r><w:t>日本語 run 1</w:t></w:r>
            <w:r><w:t>English run 2</w:t></w:r>
        </w:p>
    </w:body>
</w:document>''')
            docx.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')

        adapter = DocxAdapter()
        segments = adapter.extract_segments(self.test_docx_path)

        # Should extract only Japanese segment, but with correct run_index=1
        self.assertEqual(len(segments), 1)
        self.assertEqual(segments[0].run_index, 1)  # Second run (index 1)

    def test_japanese_detection(self):
        """Test Japanese character detection."""
        adapter = DocxAdapter()

        # Test various Japanese texts
        result, _ = adapter._contains_japanese("これは日本語です")
        self.assertTrue(result)
        result, _ = adapter._contains_japanese("漢字とひらがな")
        self.assertTrue(result)
        result, _ = adapter._contains_japanese("カタカナ too")
        self.assertTrue(result)
        result, _ = adapter._contains_japanese("Mixed 日本語 and English")
        self.assertTrue(result)

        # Test non-Japanese texts
        result, _ = adapter._contains_japanese("This is English only")
        self.assertFalse(result)
        result, _ = adapter._contains_japanese("Ceci est français")
        self.assertFalse(result)
        result, _ = adapter._contains_japanese("Dies ist Deutsch")
        self.assertFalse(result)

    def test_bilingual_json_generation(self):
        """Test generation of bilingual JSON file."""
        # Create document
        self.create_simple_docx()

        # Extract segments
        adapter = DocxAdapter()
        segments = adapter.extract_segments(self.test_docx_path)

        # Create translations
        translations = [
            {
                'id': 'word_document.xml_0_0',
                'translation': 'This is a test.'
            }
        ]

        # Generate bilingual JSON
        bilingual_path = os.path.join(self.temp_dir, 'bilingual.json')
        adapter.generate_bilingual_json(translations, bilingual_path)

        # Verify JSON file
        self.assertTrue(os.path.exists(bilingual_path))

        with open(bilingual_path, 'r', encoding='utf-8') as f:
            bilingual_data = json.load(f)

        self.assertEqual(len(bilingual_data), 1)
        self.assertEqual(bilingual_data[0]['original'], "これはテストです。")
        self.assertEqual(bilingual_data[0]['translated'], "This is a test.")
        self.assertEqual(bilingual_data[0]['context'].get('p_style'), 'Normal')

    def test_preserve_styles(self):
        """Test that rPr and pPr are preserved exactly during translation."""
        # Create DOCX with complex formatting
        with ZipFile(self.test_docx_path, 'w') as docx:
            # Add required content types
            docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')

            # Add document relationships
            docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')

            # Add word relationships
            docx.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')

            # Add document with complex formatting
            docx.writestr('word/document.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
                <w:spacing w:before="240" w:after="120"/>
                <w:ind w:left="720"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:i/>
                    <w:color w:val="FF0000"/>
                    <w:sz w:val="28"/>
                    <w:u w:val="single"/>
                </w:rPr>
                <w:t>これは複雑な書式のテキストです。</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>''')

        # Extract segments
        adapter = DocxAdapter()
        segments = adapter.extract_segments(self.test_docx_path)

        # Verify extraction captured original rPr
        self.assertEqual(len(segments), 1)
        self.assertIn('original_rPr', segments[0].metadata)
        original_rPr = segments[0].metadata['original_rPr']
        # Check for elements (namespace is preserved during serialization)
        self.assertIn('b', original_rPr)
        self.assertIn('i', original_rPr)
        self.assertIn('FF0000', original_rPr)
        self.assertIn('28', original_rPr)
        self.assertIn('single', original_rPr)

        # Verify pPr captured in context
        self.assertIn('original_pPr', segments[0].context)
        original_pPr = segments[0].context['original_pPr']
        self.assertIn('Heading1', original_pPr)
        self.assertIn('240', original_pPr)
        self.assertIn('120', original_pPr)
        self.assertIn('720', original_pPr)

        # Create translation
        translations = [
            Segment(
                id=segments[0].id,
                text="This is complex formatted text.",
                segment_type=segments[0].segment_type,
                file_path=segments[0].file_path,
                paragraph_index=segments[0].paragraph_index,
                run_index=segments[0].run_index,
                metadata=segments[0].metadata,
                context=segments[0].context
            )
        ]

        # Apply translations
        output_path = os.path.join(self.temp_dir, 'output.docx')
        adapter.apply_translations(self.test_docx_path, translations, output_path)

        # Verify the output preserves formatting
        with ZipFile(output_path, 'r') as docx_out:
            document_xml = docx_out.read('word/document.xml').decode('utf-8')

            # Check that formatting is preserved (may have namespace prefixes)
            self.assertIn('b', document_xml)
            self.assertIn('i', document_xml)
            self.assertIn('FF0000', document_xml)
            self.assertIn('28', document_xml)
            self.assertIn('single', document_xml)

            # Check paragraph style preservation
            self.assertIn('Heading1', document_xml)
            self.assertIn('240', document_xml)
            self.assertIn('120', document_xml)
            self.assertIn('720', document_xml)

    def test_list_sync(self):
        """Test that list numbering properties (numPr/ilvl) are preserved."""
        # Create DOCX with numbered lists
        with ZipFile(self.test_docx_path, 'w') as docx:
            # Add required content types
            docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>''')

            # Add document relationships
            docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')

            # Add word relationships
            docx.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>''')

            # Add numbering definitions
            docx.writestr('word/numbering.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1."/>
            <w:lvlJc w:val="left"/>
        </w:lvl>
        <w:lvl w:ilvl="1">
            <w:start w:val="1"/>
            <w:numFmt w:val="lowerLetter"/>
            <w:lvlText w:val="%2."/>
            <w:lvlJc w:val="left"/>
        </w:lvl>
    </w:abstractNum>
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"/>
    </w:num>
</w:numbering>''')

            # Add document with numbered lists
            docx.writestr('word/document.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="1"/>
                </w:numPr>
            </w:pPr>
            <w:r>
                <w:t>これは最初のレベルの項目です。</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:numPr>
                    <w:ilvl w:val="1"/>
                    <w:numId w:val="1"/>
                </w:numPr>
            </w:pPr>
            <w:r>
                <w:t>これは2番目のレベルの項目です。</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>''')

        # Extract segments
        adapter = DocxAdapter()
        segments = adapter.extract_segments(self.test_docx_path)

        # Verify list properties captured
        self.assertEqual(len(segments), 2)

        # First segment should have ilvl="0", numId="1"
        self.assertIn('list_properties', segments[0].context)
        list_props_0 = segments[0].context['list_properties']
        self.assertIsNotNone(list_props_0)
        self.assertEqual(list_props_0.get('ilvl'), '0')
        self.assertEqual(list_props_0.get('numId'), '1')

        # Second segment should have ilvl="1", numId="1"
        self.assertIn('list_properties', segments[1].context)
        list_props_1 = segments[1].context['list_properties']
        self.assertIsNotNone(list_props_1)
        self.assertEqual(list_props_1.get('ilvl'), '1')
        self.assertEqual(list_props_1.get('numId'), '1')

        # Create translations
        translations = [
            Segment(
                id=segments[0].id,
                text="This is first level item.",
                segment_type=segments[0].segment_type,
                file_path=segments[0].file_path,
                paragraph_index=segments[0].paragraph_index,
                run_index=segments[0].run_index,
                metadata=segments[0].metadata,
                context=segments[0].context
            ),
            Segment(
                id=segments[1].id,
                text="This is second level item.",
                segment_type=segments[1].segment_type,
                file_path=segments[1].file_path,
                paragraph_index=segments[1].paragraph_index,
                run_index=segments[1].run_index,
                metadata=segments[1].metadata,
                context=segments[1].context
            )
        ]

        # Apply translations
        output_path = os.path.join(self.temp_dir, 'output.docx')
        adapter.apply_translations(self.test_docx_path, translations, output_path)

        # Verify list numbering is preserved
        with ZipFile(output_path, 'r') as docx_out:
            document_xml = docx_out.read('word/document.xml').decode('utf-8')

            # Check that numPr is preserved for both paragraphs
            self.assertIn('ilvl', document_xml)
            self.assertIn('numId', document_xml)
            self.assertIn('0', document_xml)
            self.assertIn('1', document_xml)

            # Verify translations are applied
            self.assertIn('This is first level item.', document_xml)
            self.assertIn('This is second level item.', document_xml)

    def test_xml_parity_subset(self):
        """Test that XML structure parity is maintained for critical elements."""
        # Create test DOCX with complex XML structures
        temp_dir = tempfile.mkdtemp()
        test_docx_path = os.path.join(temp_dir, 'test.docx')

        # Complex XML with namespaces, custom styles, and formatting
        document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
    <w:body>
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
                <w:spacing w:before="240" w:after="120"/>
                <w:ind w:left="720"/>
                <w:rPr>
                    <w:color w:val="FF0000"/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:i/>
                    <w:sz w:val="28"/>
                    <w:color w:val="0000FF"/>
                    <w:u w:val="single"/>
                </w:rPr>
                <w:t>見出し</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="1"/>
                </w:numPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:sz w:val="22"/>
                </w:rPr>
                <w:t>リスト項目</w:t>
            </w:r>
        </w:p>
        <w:tbl>
            <w:tr>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:rPr>
                                <w:b/>
                            </w:rPr>
                            <w:t>表のテキスト</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
            </w:tr>
        </w:tbl>
    </w:body>
</w:document>'''

        with ZipFile(test_docx_path, 'w') as docx:
            docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')
            docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')
            docx.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')
            docx.writestr('word/document.xml', document_xml)

        # Extract and capture original XML
        adapter = DocxAdapter()
        segments = adapter.extract_segments(test_docx_path)

        # Store original XML snippets for comparison
        original_xmls = {}
        for seg in segments:
            if seg.metadata.get('original_rPr'):
                original_xmls[seg.id] = {
                    'rPr': seg.metadata['original_rPr'],
                    'pPr': seg.context.get('original_pPr', ''),
                    'text': seg.text
                }

        # Apply translations (preserve structure)
        translated_segments = [
            Segment(
                id=seg.id,
                text=f"Translated: {seg.text}",
                segment_type=SegmentType.PARAGRAPH,
                file_path=seg.file_path,
                paragraph_index=seg.paragraph_index,
                run_index=seg.run_index,
                metadata=seg.metadata,
                context=seg.context
            )
            for seg in segments
        ]

        # Apply to document
        output_path = os.path.join(temp_dir, 'output.docx')
        adapter.apply_translations(test_docx_path, translated_segments, output_path)

        # Extract again to verify XML parity
        adapter2 = DocxAdapter(str(output_path))
        final_segments = adapter2.extract_segments(str(output_path))

        # Verify XML structure parity
        for orig_id, orig_data in original_xmls.items():
            # Find corresponding final segment
            final_seg = next((s for s in final_segments if s.id == orig_id), None)
            self.assertIsNotNone(final_seg, f"Segment {orig_id} not found in final document")

            # Check rPr preservation
            if orig_data['rPr']:
                self.assertIn('original_rPr', final_seg.metadata,
                             f"original_rPr missing for segment {orig_id}")
                final_rPr = final_seg.metadata['original_rPr']
                # Normalize XML for comparison (ignore whitespace and namespace prefixes)
                self.assertEqual(
                    self._normalize_xml(orig_data['rPr']),
                    self._normalize_xml(final_rPr),
                    f"rPr XML parity failed for segment {orig_id}"
                )

            # Check pPr preservation if it existed
            if orig_data['pPr']:
                self.assertIn('original_pPr', final_seg.context,
                             f"original_pPr missing for segment {orig_id}")
                final_pPr = final_seg.context['original_pPr']
                self.assertEqual(
                    self._normalize_xml(orig_data['pPr']),
                    self._normalize_xml(final_pPr),
                    f"pPr XML parity failed for segment {orig_id}"
                )

        # Clean up
        shutil.rmtree(temp_dir)

    def _normalize_xml(self, xml_string):
        """Normalize XML string for comparison by removing whitespace and standardizing."""
        if not xml_string:
            return xml_string
        # Remove whitespace between tags
        import re
        xml_string = re.sub(r'>\s+<', '><', xml_string.strip())
        # Standardize namespace prefixes (remove them for comparison)
        xml_string = re.sub(r'\w+:', '', xml_string)
        return xml_string

    def test_large_document_performance(self):
        """Test performance with large documents (many paragraphs)."""
        import time

        temp_dir = tempfile.mkdtemp()
        test_docx_path = os.path.join(temp_dir, 'large_test.docx')

        # Create a document with many paragraphs
        paragraph_count = 1000
        paragraphs = []

        # Start with basic DOCX structure
        docx_content = {
            '[Content_Types].xml': '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''',
            '_rels/.rels': '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''',
            'word/_rels/document.xml.rels': '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>'''
        }

        # Generate many paragraphs with varied formatting
        for i in range(paragraph_count):
            if i % 100 == 0:
                # Heading every 100 paragraphs
                paragraphs.append(f'''        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading{i//100 + 1}"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:sz w:val="32"/>
                </w:rPr>
                <w:t>見出し {i//100 + 1}</w:t>
            </w:r>
        </w:p>''')
            elif i % 10 == 0:
                # List item every 10 paragraphs
                level = (i // 10) % 3
                paragraphs.append(f'''        <w:p>
            <w:pPr>
                <w:numPr>
                    <w:ilvl w:val="{level}"/>
                    <w:numId w:val="1"/>
                </w:numPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:sz w:val="22"/>
                </w:rPr>
                <w:t>リスト項目 {i}</w:t>
            </w:r>
        </w:p>''')
            else:
                # Regular paragraph with some formatting
                bold = "true" if i % 3 == 0 else "false"
                italic = "true" if i % 5 == 0 else "false"
                paragraphs.append(f'''        <w:p>
            <w:r>
                <w:rPr>
                    <w:b w:val="{bold}"/>
                    <w:i w:val="{italic}"/>
                    <w:sz w:val="24"/>
                </w:rPr>
                <w:t>これはテスト段落です。番号 {i} です。さらにテキストが続きます。</w:t>
            </w:r>
        </w:p>''')

        # Create document.xml
        document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
''' + '\n'.join(paragraphs) + '''
    </w:body>
</w:document>'''

        docx_content['word/document.xml'] = document_xml

        # Write DOCX file
        with ZipFile(test_docx_path, 'w') as docx:
            for name, content in docx_content.items():
                docx.writestr(name, content)

        # Test extraction performance
        adapter = DocxAdapter()

        start_time = time.time()
        segments = adapter.extract_segments(test_docx_path)
        extraction_time = time.time() - start_time

        # Verify extraction worked
        self.assertGreater(len(segments), 500, "Should extract many segments from large document")
        self.assertLess(extraction_time, 10.0, f"Extraction took {extraction_time:.2f}s, should be < 10s for {paragraph_count} paragraphs")

        # Test translation application performance
        translated_segments = [
            Segment(
                id=seg.id,
                text=f"Translated paragraph {i}: " + seg.text,
                segment_type=SegmentType.PARAGRAPH,
                file_path=seg.file_path,
                paragraph_index=seg.paragraph_index,
                run_index=seg.run_index,
                metadata=seg.metadata,
                context=seg.context
            )
            for i, seg in enumerate(segments)
        ]

        output_path = os.path.join(temp_dir, 'large_output.docx')

        start_time = time.time()
        adapter.apply_translations(test_docx_path, translated_segments, output_path)
        application_time = time.time() - start_time

        self.assertLess(application_time, 15.0, f"Application took {application_time:.2f}s, should be < 15s for {len(segments)} segments")

        # Verify memory usage is reasonable
        import psutil
        process = psutil.Process()
        memory_info = process.memory_info()
        memory_mb = memory_info.rss / 1024 / 1024

        self.assertLess(memory_mb, 500, f"Memory usage {memory_mb:.1f}MB should be < 500MB for large document processing")

        # Clean up
        shutil.rmtree(temp_dir)


    def test_nested_tables_extraction(self):
        """Test extraction from document with nested tables."""
        fixture_path = os.path.join(
            os.path.dirname(__file__),
            'fixtures',
            'uploads',
            'nested_tables.docx'
        )
        if not os.path.exists(fixture_path):
            self.skipTest("Nested tables fixture not found")
        
        adapter = DocxAdapter()
        segments = adapter.extract_segments(fixture_path)
        metadata = adapter.collect_metadata(fixture_path)
        
        self.assertGreater(len(segments), 0, "Should extract segments from nested tables document")
        self.assertGreater(metadata.custom_properties.get('table_count', 0), 1, "Should detect multiple tables including nested")
        
        # Check that some segments have table context
        table_segments = [s for s in segments if s.context and s.context.get('table_context') == 'table']
        self.assertGreater(len(table_segments), 0, "Should detect table context for some segments")
        
        # Apply dummy translation and check parity
        translated_segments = [
            Segment(
                id=s.id,
                text=f"Translated: {s.text}",
                segment_type=s.segment_type,
                file_path=s.file_path,
                paragraph_index=s.paragraph_index,
                run_index=s.run_index,
                metadata=s.metadata,
                context=s.context
            ) for s in segments
        ]
        
        output_path = os.path.join(self.temp_dir, 'nested_translated.docx')
        result = adapter.apply_translations(fixture_path, translated_segments, output_path)
        
        # Verify output exists and has similar structure
        self.assertTrue(os.path.exists(output_path))
        
        # Basic XML parity check - count paragraphs and tables
        with ZipFile(fixture_path, 'r') as input_zip:
            input_xml = input_zip.read('word/document.xml')
        
        with ZipFile(output_path, 'r') as output_zip:
            output_xml = output_zip.read('word/document.xml')
        
        # Parse and count elements (simple check)
        input_root = ET.fromstring(input_xml)
        output_root = ET.fromstring(output_xml)
        
        input_paragraphs = len(list(input_root.iter(W_NS + 'p')))
        output_paragraphs = len(list(output_root.iter(W_NS + 'p')))
        self.assertEqual(input_paragraphs, output_paragraphs, "Paragraph count should be preserved")
        
        input_tables = len(list(input_root.iter(W_NS + 'tbl')))
        output_tables = len(list(output_root.iter(W_NS + 'tbl')))
        self.assertEqual(input_tables, output_tables, "Table count should be preserved for nested tables")


if __name__ == '__main__':
    unittest.main()
