# Set 1 Fixtures

## Introduction

The Set 1 fixtures test advanced DOCX features in the translation pipeline. They include specific documents for nested tables, merged cells, and complex scenarios in the `uploads/complex/` subdirectory.

These fixtures are used in unit and integration tests to verify parsing, extraction, and preservation during translation.

## Fixture Details

### nested_tables.docx

**Location:** tests/fixtures/set1/nested_tables.docx

**Purpose:** Verify handling of nested tables.

**Structure:**
- Outer table: 3 rows x 2 columns
- Nested table in cell (row 1, col 0): 2 rows x 3 columns
- Text in cells: Headers and data.

**Archive Info (unzip -l):**
- 17 files
- Total size: approximately 828 KB
- Key files: [Content_Types].xml, word/document.xml (3750 bytes), word/styles.xml (349458 bytes)

**Expected XML Paths and Snippets:**
```xml
<!-- word/document.xml -->
<w:tbl>
  <w:tblPr>
    <w:tblGrid>
      <w:gridCol w:w="4320"/>
      <w:gridCol w:w="4320"/>
    </w:tblGrid>
  </w:tblPr>
  <!-- ... -->
</w:tbl>
```
- Nested table: Inside `<w:tc>`: another `<w:tbl>` with 3 columns
- Example text: `<w:t>Nested Header 1</w:t>`

**Test Assertions:**
- `unzip -l` returns 17 files.
- `python-docx`: `from docx import Document; doc = Document(path); assert len(doc.tables) == 1; nested = doc.tables[0].rows[1].cells[0].tables[0]; assert nested.rows[0].cells[0].text == 'Nested Header 1'`
- Loads without exceptions.
- Adapter extracts all text segments correctly, preserving hierarchy.

### merged_cells.docx

**Location:** tests/fixtures/set1/merged_cells.docx

**Purpose:** Verify handling of merged cells in tables.

**Structure:**
- Table: 3 rows x 3 columns
- Merged cells: Row 2, columns 0 and 1 merged (gridSpan=2)
- Text: &quot;Merged Item 1&quot; in merged cell.

**Archive Info:**
- 17 files
- Total size: approximately 827 KB
- Key files: word/document.xml (2957 bytes)

**Expected XML Paths and Snippets:**
```xml
<!-- word/document.xml -->
<w:tbl>
  <w:tblGrid>
    <w:gridCol w:w="2880"/>
    <w:gridCol w:w="2880"/>
    <w:gridCol w:w="2880"/>
  </w:tblGrid>
  <!-- Merged cell example -->
  <w:tc>
    <w:tcPr>
      <w:tcW w:type="dxa" w:w="5760"/>
      <w:gridSpan w:val="2"/>
    </w:tcPr>
    <w:p>
      <w:r>
        <w:t>Merged Item 1</w:t>
      </w:r>
    </w:p>
  </w:tc>
</w:tbl>
```

**Test Assertions:**
- `unzip -l` returns 17 files.
- `python-docx`: `doc = Document(path); table = doc.tables[0]; assert table.rows[1].cells[0].text == 'Merged Item 1'; assert len(table.rows[1].cells) == 2`
- No load errors.
- Extraction treats merged as single unit.

## Complex Fixtures

The `tests/fixtures/uploads/complex/` directory contains fixtures for more intricate features.

### comments.docx

**Location:** tests/fixtures/uploads/complex/comments.docx

**Purpose:** Test preservation of comments during translation.

**Structure:**
- One paragraph with a comment range around &quot;Text with comment.&quot;
- Comment ID 0.

**Archive Info:**
- 5 files
- Total size: approximately 2 KB
- Key files: word/document.xml (551 bytes), word/comments.xml (356 bytes)

**Expected XML Paths and Snippets:**
- word/document.xml: `<w:commentRangeStart w:id=&quot;0&quot;/>&lt;w:r>&lt;w:t>Text with comment.&lt;/w:t>&lt;/w:r>&lt;w:commentReference w:id=&quot;0&quot;/>&lt;w:commentRangeEnd w:id=&quot;0&quot;/>`
- word/comments.xml: Contains `<w:comment w:id=&quot;0&quot;>` with comment text.

**Test Assertions:**
- `unzip -l` returns 5 files.
- `python-docx`: Loads without error.
- Adapter ignores comment text for translation, preserves structure.
- Post-translation, comments remain intact.

### hyperlinks.docx

**Location:** tests/fixtures/uploads/complex/hyperlinks.docx

**Purpose:** Test hyperlink preservation.

**Structure:**
- Paragraph with hyperlink &quot;Click here&quot; to external URL.

**Archive Info:**
- 4 files
- Total size: approximately 1.5 KB
- Key files: word/document.xml (561 bytes), word/_rels/document.xml.rels (284 bytes)

**Expected XML Paths and Snippets:**
- word/document.xml: `<w:hyperlink r:id=&quot;rId1&quot;>&lt;w:r>&lt;w:rPr>&lt;w:rStyle w:val=&quot;Hyperlink&quot;/>&lt;/w:rPr>&lt;w:t>Click here&lt;/w:t>&lt;/w:r>&lt;/w:hyperlink>`
- word/_rels/document.xml.rels: `<Relationship Id=&quot;rId1&quot; Type=&quot;http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink&quot; Target=&quot;https://example.com&quot;/>&quot;`

**Test Assertions:**
- `unzip -l` returns 4 files.
- `python-docx`: `doc = Document(path); assert 'https://example.com' in doc.paragraphs[0].runs[0].hyperlink.target if supported`
- Links unchanged after translation.

### fields.docx

**Location:** tests/fixtures/uploads/complex/fields.docx

**Purpose:** Test field code preservation.

**Structure:**
- Paragraph with fldSimple for hyperlink field.

**Archive Info:**
- 4 files
- Total size: approximately 1.4 KB
- Key files: word/document.xml (548 bytes)

**Expected XML Paths and Snippets:**
- word/document.xml: `<w:fldSimple w:instr=&quot;HYPERLINK \&quot;https://example.com\&quot;&quot;>&lt;w:r>&lt;w:t>Link text&lt;/w:t>&lt;/w:r>&lt;/w:fldSimple>`

**Test Assertions:**
- `unzip -l` returns 4 files.
- `python-docx`: Loads, fields preserved.
- Fields not translated; codes intact.

### tracked_changes.docx

**Location:** tests/fixtures/uploads/complex/tracked_changes.docx

**Purpose:** Test preservation of tracked changes.

**Structure:**
- Paragraph with inserted text (w:ins).

**Archive Info:**
- 4 files
- Total size: approximately 1.4 KB
- Key files: word/document.xml (502 bytes)

**Expected XML Paths and Snippets:**
- word/document.xml: `<w:ins w:id=&quot;1&quot; w:author=&quot;Test Author&quot; w:date=&quot;2025-09-21T00:00:00Z&quot;>&lt;w:t> This is inserted text. &lt;/w:t>&lt;/w:ins>`

**Test Assertions:**
- `unzip -l` returns 4 files.
- `python-docx`: Loads without error.
- Revisions preserved, text in insertions translated if policy allows.

## General Notes

- All fixtures are minimal to isolate features.
- Use evidence from S1-T1 (unzip counts/sizes), S1-T2 (XML paths/snippets), S1-T3 (python-docx load tests).
- Verify in CI: pytest tests/test_docx_adapter.py -k &quot;set1&quot;