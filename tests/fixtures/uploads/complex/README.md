# Complex DOCX Fixtures

This directory contains minimal DOCX fixtures (<5KB each) demonstrating specific OOXML elements for testing translation preservation.

## Files

### tracked_changes.docx (1.8KB)
Contains tracked insertions using `<w:ins>` elements.

**XML Structure:**
```xml
<w:ins w:id="1" w:author="Test Author" w:date="2025-09-21T00:00:00Z">
  <w:r>
    <w:t>inserted text</w:t>
  </w:r>
</w:ins>
```

**Verification:**
```bash
unzip -p tracked_changes.docx word/document.xml | grep '<w:ins'
```

### comments.docx (2.6KB)
Contains document comments with `<w:comment>` elements and comment ranges.

**XML Structure:**
```xml
<!-- In word/document.xml -->
<w:commentRangeStart w:id="0"/>
<w:r>
  <w:t>This text has a comment</w:t>
</w:r>
<w:commentRangeEnd w:id="0"/>
<w:r>
  <w:commentReference w:id="0"/>
</w:r>

<!-- In word/comments.xml -->
<w:comment w:id="0" w:author="Commenter" w:date="2025-09-21T00:00:00Z" w:initials="C">
  <w:p>
    <w:r>
      <w:t>This is a comment.</w:t>
    </w:r>
  </w:p>
</w:comment>
```

**Verification:**
```bash
unzip -p comments.docx word/comments.xml | grep '<w:comment'
```

### hyperlinks.docx (2.0KB)
Contains external hyperlinks using `<w:hyperlink>` elements.

**XML Structure:**
```xml
<w:hyperlink r:id="rId1">
  <w:r>
    <w:t>this link</w:t>
  </w:r>
</w:hyperlink>

<!-- In word/_rels/document.xml.rels -->
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
              Target="https://example.com" TargetMode="External"/>
```

**Verification:**
```bash
unzip -p hyperlinks.docx word/document.xml | grep '<w:hyperlink'
```

### fields.docx (1.9KB)
Contains field codes using `<w:fldSimple>` elements.

**XML Structure:**
```xml
<w:fldSimple w:instr="DATE">
  <w:r>
    <w:t>21-Sep-25</w:t>
  </w:r>
</w:fldSimple>

<w:fldSimple w:instr="HYPERLINK "https://example.com"">
  <w:r>
    <w:t>Link text</w:t>
  </w:r>
</w:fldSimple>
```

**Verification:**
```bash
unzip -p fields.docx word/document.xml | grep '<w:fldSimple'
```

## Usage

These fixtures are designed for testing the DOCX translation pipeline's handling of complex OOXML features:

1. **Load each fixture** using python-docx or directly parse the XML
2. **Extract and translate text** while preserving the OOXML structure
3. **Verify the output** maintains all original elements with translated text content

## Test Integration

Example test assertions:
```python
def test_tracked_changes_preserved():
    # After translation, w:ins elements should remain with translated content
    doc = Document(translated_path)
    # Check insertion elements are preserved

def test_hyperlinks_preserved():
    # After translation, hyperlinks should maintain target URLs
    doc = Document(translated_path)
    # Check hyperlink relationships are unchanged
```

## Size Optimization

These fixtures are minimal (<5KB) to:
- Keep test suites fast
- Focus on specific features without noise
- Enable precise assertions about XML structure