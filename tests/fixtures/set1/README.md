# Set 1 DOCX Fixtures for Advanced Features

This directory contains targeted DOCX fixtures for testing specific advanced features in the DOCX adapter and translation pipeline. Each file focuses on one feature to allow isolated testing.

## Files and Descriptions

### nested_tables.docx
- **Structure**: Contains an outer table (3x2) with a nested table (2x3) inside the cell at row 1, column 0.
- **Expected Assertions**:
  - The adapter should extract text from both outer and inner tables correctly.
  - Nested structure preserved in segments (e.g., table cells as separate segments).
  - No loss of hierarchy during parsing.
  - Translation should apply to all text in tables without breaking layout.

### merged_cells.docx
- **Structure**: A 3x3 table where the first two cells in row 1 are merged, containing "Merged Item 1".
- **Expected Assertions**:
  - Merged cells treated as single segment for text extraction.
  - Text in merged cell translated as one unit.
  - Layout preserved post-translation (no splitting of merged cells).
  - Adapter handles merge spans correctly.

### hyperlinks.docx
- **Structure**: Paragraphs with hyperlinks to "<https://example.com>" and "#anchor", styled in blue and underlined.
- **Expected Assertions**:
  - Hyperlinks preserved in output DOCX (URLs unchanged).
  - Surrounding text translated, but link text remains as is or translated if needed (depending on policy).
  - No breakage of link functionality.
  - Adapter extracts link text separately if required.

### comments.docx
- **Structure**: A paragraph with placeholder text indicating a comment location. (Note: Full comments simulated; real comments may need manual creation in Word.)
- **Expected Assertions**:
  - Comments ignored during translation (not translated, preserved as is).
  - Comment text and metadata (author, date) unchanged.
  - Adapter skips comment content in text extraction.
  - Output DOCX retains all comments.

### tracked_changes.docx
- **Structure**: Paragraphs with placeholder text for insertions and deletions. (Note: Simulated; real tracked changes require Word.)
- **Expected Assertions**:
  - Tracked changes preserved (including revision types: insert, delete).
  - Only accepted/unresolved changes' text translated if policy allows; typically preserve revisions.
  - Adapter detects and handles revision markup without altering it.
  - No translation of revision metadata.

### fields.docx
- **Structure**: Paragraphs with placeholder text for fields like {DATE} and {PAGE}. (Note: Simulated; real fields need XML.)
- **Expected Assertions**:
  - Fields preserved unchanged (e.g., {DATE} not evaluated or translated).
  - Surrounding text translated.
  - Adapter recognizes field codes and skips them in extraction.
  - Output maintains field functionality.

## General Notes
- These fixtures are small (under 1 page each) to focus on specific behaviors.
- Use in unit/integration tests to verify docx_adapter.py parsing and translation_orchestrator.py handling.
- For full features (comments, tracked changes, fields), consider generating in Microsoft Word and zipping as fixtures if python-docx limitations persist.
- Smoke checks: Load each via Document(fixture_path) and assert no errors; verify text extraction via adapter.

## Testing Commands
- Run pytest tests/test_docx_adapter.py::test_load_fixture -k "nested_tables" (adapt as needed).
- Verify XML: unzip -l fixture.docx | grep document.xml; check for expected elements (e.g., <w:tbl> for tables, <w:hyperlink> for links).