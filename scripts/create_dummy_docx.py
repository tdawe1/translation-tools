#!/usr/bin/env python3
"""
Create a minimal valid DOCX file for testing.
"""

import sys
import zipfile
from pathlib import Path

def create_dummy_docx(output_path):
    """Create a minimal valid DOCX file."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Create a minimal DOCX structure
    with zipfile.ZipFile(output_path, 'w') as zf:
        # Document XML with Japanese text
        zf.writestr('word/document.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>これはテストです。</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>翻訳が必要な文章です。</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>''')

        # Content types
        zf.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')

        # Relationships
        zf.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')

    print(f"Created dummy DOCX: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python create_dummy_docx.py <output_path>")
        sys.exit(1)

    create_dummy_docx(sys.argv[1])