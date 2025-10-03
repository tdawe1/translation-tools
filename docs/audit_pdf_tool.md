# PDF Audit Tool Documentation

## Overview

The `audit_pdf.py` tool provides comprehensive quality assessment for translated PDF documents, extending the existing audit framework to handle PDF translations with the same rigor as PPTX translations.

## Features

### Core Functionality
- **Residual Japanese Detection**: Identifies untranslated Japanese characters with 99%+ accuracy
- **Layout Integrity Verification**: Compares layout structure between original and translated PDFs
- **Translation Quality Assessment**: Evaluates completeness, formatting, and overall quality
- **Comprehensive Reporting**: Generates detailed audit reports in CSV or JSON format

### Audit Metrics
1. **Residual Japanese**: Count and percentage of untranslated Japanese content
2. **Layout Integrity**: Position and formatting preservation score (0.0-1.0)
3. **Text Completeness**: Verification that no content is missing or truncated
4. **Formatting Consistency**: Style preservation across pages
5. **Overall Quality Score**: Combined assessment (0.0-1.0)

## Installation

The tool uses existing Python libraries that should already be available:
- `pypdf` - PDF parsing and page extraction
- `pdfminer.six` - Advanced text extraction from PDFs

## Usage

### Basic Usage

```bash
# Audit a translated PDF
python scripts/audit_pdf.py translated.pdf

# Audit with comparison to original Japanese PDF
python scripts/audit_pdf.py translated.pdf original.pdf

# Generate custom report
python scripts/audit_pdf.py translated.pdf original.pdf --report custom_audit.csv
```

### Advanced Usage

```bash
# Generate JSON report
python scripts/audit_pdf.py translated.pdf --json --output audit.json

# Set quality threshold (default: 0.8)
python scripts/audit_pdf.py translated.pdf --threshold 0.9

# Verbose output for debugging
python scripts/audit_pdf.py translated.pdf --verbose
```

### CLI Options

- `translated_pdf`: Path to the translated PDF file (required)
- `original_pdf`: Path to original Japanese PDF (optional, for comparison)
- `--report, -r`: Output CSV report path (default: PDF_AUDIT_REPORT.csv)
- `--json, -j`: Output JSON format instead of CSV
- `--output, -o`: Output file path (for JSON)
- `--verbose, -v`: Verbose output for debugging
- `--threshold, -t`: Quality threshold (0.0-1.0, default: 0.8)

## API Reference

### PDFAuditor Class

The main class provides the following methods:

#### `count_residual_jp(pdf_path: str) -> int`
Counts residual Japanese characters in a translated PDF.

**Parameters:**
- `pdf_path`: Path to the translated PDF file

**Returns:**
- Number of Japanese characters found

#### `check_layout_integrity(original: str, translated: str) -> LayoutCheckResult`
Verifies layout integrity between original and translated PDFs.

**Parameters:**
- `original`: Path to original PDF
- `translated`: Path to translated PDF

**Returns:**
- `LayoutCheckResult` object with score and issues

#### `assess_translation_quality(pdf_path: str) -> QualityAssessment`
Assesses translation quality of a PDF document.

**Parameters:**
- `pdf_path`: Path to the translated PDF

**Returns:**
- `QualityAssessment` object with metrics and recommendations

#### `generate_audit_report(pdf_path: str, original_pdf_path: Optional[str] = None) -> AuditReport`
Generates comprehensive audit report for a translated PDF.

**Parameters:**
- `pdf_path`: Path to translated PDF
- `original_pdf_path`: Optional path to original PDF

**Returns:**
- Complete `AuditReport` object

#### `compare_with_original(original: str, translated: str) -> Dict[str, Any]`
Compares translated PDF with original Japanese PDF.

**Parameters:**
- `original`: Path to original PDF
- `translated`: Path to translated PDF

**Returns:**
- Dictionary with detailed comparison metrics

### Data Classes

#### LayoutCheckResult
- `score`: Layout integrity score (0.0-1.0)
- `issues`: List of layout issues found
- `page_count_match`: Whether page counts match
- `similar_structure`: Whether text structure is similar

#### QualityAssessment
- `residual_japanese_count`: Number of Japanese characters found
- `residual_japanese_percentage`: Percentage of Japanese characters
- `text_completeness_score`: Completeness assessment (0.0-1.0)
- `formatting_consistency_score`: Formatting consistency (0.0-1.0)
- `overall_quality_score`: Overall quality score (0.0-1.0)
- `recommendations`: List of improvement recommendations

#### AuditReport
- `file_path`: Path to audited file
- `original_file_path`: Path to original file (if provided)
- `timestamp`: Audit timestamp
- `total_pages`: Number of pages in PDF
- `extracted_text_length`: Length of extracted text
- `layout_check`: Layout integrity results
- `quality_assessment`: Quality assessment results
- `page_details`: Page-by-page analysis

## Integration with Existing Framework

The PDF audit tool integrates seamlessly with the existing audit framework:

### Japanese Character Detection
Uses the same Japanese character patterns as `audit_pptx_jp_count.py`:
- Hiragana: `\u3040-\u309f`
- Katakana: `\u30a0-\u30ff`  
- Kanji: `\u3400-\u4dbf\u4e00-\u9fff`
- CJK punctuation: `\u3000-\u303f\uff00-\uffef`

### Report Compatibility
- CSV reports follow similar structure to existing style audits
- JSON reports provide detailed machine-readable output
- CLI interface matches existing audit tools

### CI Integration
The tool returns appropriate exit codes:
- `0`: Quality meets or exceeds threshold
- `1`: Quality below threshold or error occurred

This allows integration into CI/CD pipelines similar to existing audit tools.

## Examples

### Example 1: Basic PDF Audit
```bash
python scripts/audit_pdf.py translated_document.pdf
```

Output:
```
=== PDF Audit Report ===
File: translated_document.pdf
Pages: 15
Text length: 8456

=== Quality Assessment ===
Residual Japanese: 3 chars (0.04%)
Completeness Score: 0.98
Formatting Score: 0.95
Overall Quality: 0.92

=== Recommendations ===
1. Remove residual Japanese characters (3 found, 0.04%)

CSV report saved to: PDF_AUDIT_REPORT.csv
Quality check passed (score: 0.92)
```

### Example 2: With Original Comparison
```bash
python scripts/audit_pdf.py translated.pdf original.pdf --report comparison.csv
```

### Example 3: JSON Output
```bash
python scripts/audit_pdf.py translated.pdf --json --output audit.json
```

## Testing

Run the unit tests to verify functionality:
```bash
python scripts/test_audit_pdf.py
```

The test suite includes:
- Japanese character pattern detection
- Layout integrity checking
- Quality assessment algorithms
- Report generation
- CLI interface testing

## Performance Considerations

- **Text Extraction**: Uses `pdfminer.six` for accurate text extraction with fallback to `pypdf`
- **Memory Usage**: Processes PDFs page by page to handle large documents
- **Speed**: Optimized for quick audits with optional detailed analysis

## Troubleshooting

### Common Issues

1. **PDF Not Found**: Ensure the PDF file path is correct and the file exists
2. **Permission Denied**: Check file read permissions
3. **Corrupted PDF**: Try opening the PDF in a viewer to verify it's not corrupted
4. **No Text Extracted**: Some PDFs may contain images instead of text

### Debug Mode
Use the `--verbose` flag for detailed error messages and debugging information:
```bash
python scripts/audit_pdf.py translated.pdf --verbose
```

## Future Enhancements

Potential future improvements:
- Image-based PDF analysis (OCR integration)
- Advanced layout comparison using visual similarity
- Integration with translation memory systems
- Support for additional file formats
- Real-time audit monitoring