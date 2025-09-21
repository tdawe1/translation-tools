# Manual DOCX Fixture Creation Guide

This document describes how to create DOCX fixtures manually using Microsoft Word for testing the translation pipeline.

## Basic Fixtures

### simple_japanese.docx
1. Open Microsoft Word
2. Type: "これは日本語のサンプルテキストです。"
3. Save as `simple_japanese.docx` in `tests/fixtures/`

### mixed_language.docx
1. Open Microsoft Word
2. Type: "This is English text mixed with 日本語 text."
3. Save as `mixed_language.docx` in `tests/fixtures/`

### complex_japanese.docx
1. Open Microsoft Word
2. Create a title: "複雑な日本語ドキュメント"
3. Add paragraphs with mixed content:
   - "段落1：これは複雑な日本語のテキストを含んでいます。"
   - "Paragraph 2: This contains English and 日本語."
4. Insert a table (2x2):
   - Cell 1: "名前"
   - Cell 2: "Name"
   - Cell 3: "値"
   - Cell 4: "Value"
5. Save as `complex_japanese.docx` in `tests/fixtures/`

### empty.docx
1. Open Microsoft Word
2. Save the empty document as `empty.docx` in `tests/fixtures/`

### cli_sample.docx
1. Open Microsoft Word
2. Type: "CLI sample document for testing."
3. Save as `cli_sample.docx` in `tests/fixtures/`

## Set 1 Fixtures (Advanced Features)

### nested_tables.docx
1. Open Microsoft Word
2. Insert a table (3 rows, 2 columns)
3. In the first cell of the second row:
   - Click inside the cell
   - Go to Insert > Table
   - Insert a 2x3 table
4. Fill in the nested table:
   - Header row: "Nested Header 1", "Nested Header 2", "Nested Header 3"
   - Data row: "Data 1", "Data 2", "Data 3"
5. Fill in the outer table with sample text
6. Save as `set1/nested_tables.docx`

### merged_cells.docx
1. Open Microsoft Word
2. Insert a table (3 rows, 3 columns)
3. Select cells in the second row, first and second columns
4. Right-click > Merge Cells
5. Type "Merged Item 1" in the merged cell
6. Fill other cells with sample text
7. Save as `set1/merged_cells.docx`

## Complex Fixtures (uploads/complex/)

### comments.docx
1. Open Microsoft Word
2. Type: "This is text with comment."
3. Select "with comment"
4. Go to Review > New Comment
5. Add comment: "This is a test comment"
6. Save as `uploads/complex/comments.docx`

### hyperlinks.docx
1. Open Microsoft Word
2. Type: "Click here"
3. Select the text
4. Right-click > Link
5. Enter URL: "https://example.com"
6. Save as `uploads/complex/hyperlinks.docx`

### fields.docx
1. Open Microsoft Word
2. Press Ctrl+F9 to insert field braces
3. Type between braces: `HYPERLINK "https://example.com"`
4. Press F9 to update field
5. Type "Link text" after the field
6. Save as `uploads/complex/fields.docx`

### tracked_changes.docx
1. Open Microsoft Word
2. Turn on Track Changes (Review tab)
3. Type: "This is normal text."
4. Type: " This is inserted text." (with tracking on)
5. Save as `uploads/complex/tracked_changes.docx`

## Verification

After creating fixtures, verify them using:

```bash
# Check file structure
unzip -l tests/fixtures/<filename>.docx

# Check XML content (optional)
python -c "
import zipfile
with zipfile.ZipFile('tests/fixtures/<filename>.docx') as z:
    with z.open('word/document.xml') as f:
        print(f.read()[:500])  # First 500 chars
"

# Test with python-docx (if available)
python -c "
from docx import Document
doc = Document('tests/fixtures/<filename>.docx')
print(f'Paragraphs: {len(doc.paragraphs)}')
print(f'Tables: {len(doc.tables)}')
"
```

## Tips

1. Keep fixtures minimal and focused on specific features
2. Use consistent naming conventions
3. Document any special characteristics in the fixtures
4. Test fixtures with the docx_adapter to ensure they work correctly
5. Commit fixtures to version control for reproducible tests