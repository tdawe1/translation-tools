# Complex DOCX Fixtures

These fixtures test advanced DOCX features in the translation adapter. Each is minimal, generated with python-docx where possible, and annotated with expected XML structures and assertions.

## General Assertions
- All files load without error: `from docx import Document; doc = Document('file.docx')`
- File size <5KB (note: default python-docx generates ~36KB due to styles; for true minimal, edit XML)
- Unzip: `unzip file.docx -d tmp/` then check XML files.

## 1. nested_tables.docx
**Description:** Contains a table with a nested table in the first cell.

**Python-docx Load Assertions:**
- `doc.tables` length == 1 (outer)
- `doc.tables[0].cell(0,0).tables` length == 1 (nested)
- `doc.tables[0].cell(0,0).tables[0].cell(0,0).text == 'N'`

**XML Snippet (word/document.xml):**
```xml
<w:tr>
  <w:tc>
    <w:p><w:r><w:t>Nested</w:t></w:r></w:p>
    <w:tbl> <!-- nested table -->
      <w:tr><w:tc><w:p><w:r><w:t>N</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
  </w:tc>
</w:tr>
```
**Unzip Pattern:** grep -o 'w:tbl.*w:tbl' word/document.xml | head -1

## 2. merged_cells.docx
**Description:** Simulates merged cells (actual merge requires XML edit).

**Note:** python-docx does not support cell merging natively. Edit word/document.xml to add `<w:gridSpan w:val="2"/>` in <w:tc>.

**Python-docx Load Assertions:**
- `doc.tables[0].rows[0].cells` length == 1 (but actually 1 cell used)

**Expected XML Snippet for Merged (after edit):**
```xml
<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
<w:tr>
  <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc>
</w:tr>
```
**Unzip Pattern:** grep 'gridSpan' word/document.xml

## 3. track_changes.docx
**Description:** Basic text; track changes require XML for revisions.

**Note:** Add manually in Word or edit XML for <w:ins> and <w:del>.

**Python-docx Load Assertions:**
- `len(doc.paragraphs) == 1`
- `doc.paragraphs[0].text == 'T'`

**Expected XML Snippet for Track Changes:**
```xml
<w:p>
  <w:r><w:rPr><w:ins w:author="Author" w:date="..."/></w:rPr><w:t>Inserted</w:t></w:r>
  <w:r><w:rPr><w:del w:author="Author" w:date="..."/></w:rPr><w:delText>Deleted</w:delText></w:r>
</w:p>
```
**Unzip Pattern:** grep -E 'w:ins|w:del' word/document.xml

## 4. comments.docx
**Description:** Basic text; comments require comments.xml and references.

**Note:** python-docx has limited comment support. Add via XML.

**Python-docx Load Assertions:**
- `len(doc.paragraphs) == 1`
- `doc.paragraphs[0].text == 'C'`
- `len(doc.comments) == 0` (since not added)

**Expected XML Snippet:**
- word/comments.xml:
```xml
<w:comment w:id="0" w:author="Test" w:date="..."><w:p><w:r><w:t>Comment text</w:t></w:r></w:p></w:comment>
```
- In document.xml: `<w:commentRangeStart w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:commentRangeEnd w:id="0"/>`
**Unzip Pattern:** ls word/comments.xml && grep 'commentReference' word/document.xml

## 5. hyperlinks.docx
**Description:** Contains a hyperlink.

**Python-docx Load Assertions:**
- `len(doc.paragraphs) == 1`
- `doc.paragraphs[0].runs[0].hyperlink.address == 'https://example.com'`
- `doc.paragraphs[0].text == 'H'`

**XML Snippet (word/document.xml):**
```xml
<w:hyperlink r:id="rId1"><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hyperlink>
```
**Unzip Pattern:** grep 'hyperlink' word/document.xml

## 6. fields.docx
**Description:** Basic text representing field.

**Note:** Fields are <w:fldSimple> in XML.

**Python-docx Load Assertions:**
- `len(doc.paragraphs) == 1`
- `doc.paragraphs[0].text == 'F'`

**Expected XML Snippet for Field:**
```xml
<w:p>
  <w:r>
    <w:fldChar w:fldCharType="begin"/>
  </w:r>
  <w:r>
    <w:instrText> DATE </w:instrText>
  </w:r>
  <w:r>
    <w:fldChar w:fldCharType="end"/>
  </w:r>
</w:p>
```
**Unzip Pattern:** grep 'fldChar' word/document.xml

## Building on Set 1
These extend tests/fixtures/set1/ by adding complex structures for adapter testing (load/extract text, preserve structure).

## Verification
- Load: `python -c "from docx import Document; print(Document('tests/fixtures/uploads/complex/nested_tables.docx').paragraphs)"`
- Extract: Run adapter extract_text or similar.
- For full features, edit generated DOCX in Word to add missing elements (track changes, comments, merges, fields) and re-zip if needed for minimal size.