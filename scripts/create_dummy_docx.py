"""
Create a minimal valid DOCX file for testing purposes.
This script generates a basic DOCX structure with Japanese text.
"""
import os
import sys
import zipfile
import tempfile
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

def create_minimal_docx(output_path: str):
    """Create a minimal valid DOCX file with Japanese text."""

    # Create a temporary directory to build the DOCX structure
    with tempfile.TemporaryDirectory() as temp_dir:
        docx_dir = Path(temp_dir) / "docx"
        docx_dir.mkdir()

        # Create the required directory structure
        word_dir = docx_dir / "word"
        word_dir.mkdir()
        _rels_dir = docx_dir / "_rels"
        _rels_dir.mkdir()
        word_rels_dir = word_dir / "_rels"
        word_rels_dir.mkdir()
        docProps_dir = docx_dir / "docProps"
        docProps_dir.mkdir()

        # Create [Content_Types].xml
        content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""

        with open(docx_dir / "[Content_Types].xml", "w", encoding="utf-8") as f:
            f.write(content_types)

        # Create _rels/.rels
        rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""

        with open(_rels_dir / ".rels", "w", encoding="utf-8") as f:
            f.write(rels)

        # Create word/_rels/document.xml.rels
        doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

        with open(word_rels_dir / "document.xml.rels", "w", encoding="utf-8") as f:
            f.write(doc_rels)

        # Create word/document.xml with Japanese text
        document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Normal"/>
            </w:pPr>
            <w:r>
                <w:t>これは日本語のテスト文書です。</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>This is a test document with 日本語 text.</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>DOCXアダプターのセキュリティテスト用です。</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>"""

        with open(word_dir / "document.xml", "w", encoding="utf-8") as f:
            f.write(document)

        # Create word/styles.xml (minimal)
        styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:docDefaults>
        <w:rPrDefault>
            <w:rPr>
                <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
        </w:rPrDefault>
    </w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:name w:val="Normal"/>
        <w:pPr>
            <w:widowControl w:val="0"/>
            <w:jc w:val="both"/>
        </w:pPr>
        <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
            <w:sz w:val="24"/>
            <w:szCs w:val="24"/>
        </w:rPr>
    </w:style>
</w:styles>"""

        with open(word_dir / "styles.xml", "w", encoding="utf-8") as f:
            f.write(styles)

        # Create docProps/core.xml
        core_props = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title>Test Document</dc:title>
    <dc:creator>Test System</dc:creator>
    <cp:lastModifiedBy>Test System</cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""

        with open(docx_dir / "docProps/core.xml", "w", encoding="utf-8") as f:
            f.write(core_props)

        # Create docProps/app.xml
        app_props = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Template>Normal.dotm</Template>
    <TotalTime>0</TotalTime>
    <Pages>1</Pages>
    <Words>10</Words>
    <Characters>60</Characters>
    <Application>Test System</Application>
    <DocSecurity>0</DocSecurity>
    <Lines>1</Lines>
    <Paragraphs>1</Paragraphs>
    <ScaleCrop>false</ScaleCrop>
    <LinksUpToDate>false</LinksUpToDate>
    <CharactersWithSpaces>70</CharactersWithSpaces>
    <SharedDoc>false</SharedDoc>
    <HyperlinksChanged>false</HyperlinksChanged>
    <AppVersion>16.0000</AppVersion>
</Properties>"""

        with open(docx_dir / "docProps/app.xml", "w", encoding="utf-8") as f:
            f.write(app_props)

        # Create the DOCX file by zipping the contents
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            for root, dirs, files in os.walk(docx_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, docx_dir)
                    docx.write(file_path, arcname)

        print(f"Created minimal DOCX file: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python create_dummy_docx.py <output_path>")
        sys.exit(1)

    output_path = sys.argv[1]
    create_minimal_docx(output_path)