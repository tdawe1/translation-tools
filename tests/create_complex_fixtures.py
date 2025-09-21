#!/usr/bin/env python3
"""
Create minimal DOCX fixtures for complex features.
Creates small fixtures (<5KB) with specific XML elements.
"""

import zipfile
import tempfile
import os
from pathlib import Path

def create_minimal_docx_with_xml(filename, document_xml, relations_xml=None, content_types=None):
    """Create a minimal DOCX with custom XML content."""
    # Create temporary directory structure
    with tempfile.TemporaryDirectory() as temp_dir:
        # Create required directories
        word_dir = Path(temp_dir) / "word"
        word_dir.mkdir()
        _rels_dir = word_dir / "_rels"
        _rels_dir.mkdir()
        root_rels_dir = Path(temp_dir) / "_rels"
        root_rels_dir.mkdir()

        # Create [Content_Types].xml
        content_types_xml = content_types or '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

        with open(Path(temp_dir) / "[Content_Types].xml", "w") as f:
            f.write(content_types_xml)

        # Create document.xml
        with open(word_dir / "document.xml", "w") as f:
            f.write(document_xml)

        # Create _rels/.rels
        rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

        with open(Path(temp_dir) / "_rels" / ".rels", "w") as f:
            f.write(rels_xml)

        # Create word/_rels/document.xml.rels
        if relations_xml:
            with open(word_dir / "_rels" / "document.xml.rels", "w") as f:
                f.write(relations_xml)

        # Create the DOCX by zipping
        output_path = Path("tests/fixtures/uploads/complex") / filename
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(output_path, "w") as zf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(temp_dir)
                    zf.write(file_path, arcname)

        print(f"Created: {output_path} ({output_path.stat().st_size} bytes)")
        return output_path

def create_tracked_changes_docx():
    """Create minimal DOCX with tracked changes."""
    document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a document with </w:t>
      </w:r>
      <w:ins w:id="1" w:author="Test Author" w:date="2025-09-21T00:00:00Z">
        <w:r>
          <w:t>inserted text</w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t> and some normal text.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>'''

    create_minimal_docx_with_xml("tracked_changes.docx", document_xml)

def create_comments_docx():
    """Create minimal DOCX with comments."""
    document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:commentRangeStart w:id="0"/>
      <w:r>
        <w:t>This text has a comment</w:t>
      </w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r>
        <w:commentReference w:id="0"/>
      </w:r>
    </w:p>
  </w:body>
</w:document>'''

    # We need comments.xml and the relationship
    relations_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>'''

    # Create DOCX with comments.xml
    with tempfile.TemporaryDirectory() as temp_dir:
        word_dir = Path(temp_dir) / "word"
        word_dir.mkdir()
        _rels_dir = word_dir / "_rels"
        _rels_dir.mkdir()
        root_rels_dir = Path(temp_dir) / "_rels"
        root_rels_dir.mkdir()

        # [Content_Types].xml
        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>'''

        with open(Path(temp_dir) / "[Content_Types].xml", "w") as f:
            f.write(content_types)

        # document.xml
        with open(word_dir / "document.xml", "w") as f:
            f.write(document_xml)

        # comments.xml
        comments_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Commenter" w:date="2025-09-21T00:00:00Z" w:initials="C">
    <w:p>
      <w:r>
        <w:t>This is a comment.</w:t>
      </w:r>
    </w:p>
  </w:comment>
</w:comments>'''

        with open(word_dir / "comments.xml", "w") as f:
            f.write(comments_xml)

        # _rels/.rels
        rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

        with open(Path(temp_dir) / "_rels" / ".rels", "w") as f:
            f.write(rels_xml)

        # word/_rels/document.xml.rels
        with open(word_dir / "_rels" / "document.xml.rels", "w") as f:
            f.write(relations_xml)

        # Create the DOCX
        output_path = Path("tests/fixtures/uploads/complex") / "comments.docx"
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(output_path, "w") as zf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(temp_dir)
                    zf.write(file_path, arcname)

        print(f"Created: {output_path} ({output_path.stat().st_size} bytes)")

def create_hyperlinks_docx():
    """Create minimal DOCX with hyperlinks."""
    document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Visit </w:t>
      </w:r>
      <w:hyperlink r:id="rId1">
        <w:r>
          <w:t>this link</w:t>
        </w:r>
      </w:hyperlink>
      <w:r>
        <w:t> for more info.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>'''

    relations_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>'''

    create_minimal_docx_with_xml("hyperlinks.docx", document_xml, relations_xml)

def create_fields_docx():
    """Create minimal DOCX with fields."""
    document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Current date: </w:t>
      </w:r>
      <w:fldSimple w:instr="DATE">
        <w:r>
          <w:t>21-Sep-25</w:t>
        </w:r>
      </w:fldSimple>
    </w:p>
    <w:p>
      <w:fldSimple w:instr="HYPERLINK "https://example.com"">
        <w:r>
          <w:t>Link text</w:t>
        </w:r>
      </w:fldSimple>
    </w:p>
  </w:body>
</w:document>'''

    create_minimal_docx_with_xml("fields.docx", document_xml)

if __name__ == "__main__":
    print("Creating minimal complex DOCX fixtures...")
    create_tracked_changes_docx()
    create_comments_docx()
    create_hyperlinks_docx()
    create_fields_docx()
    print("Done!")