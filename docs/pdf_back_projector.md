# PDF Back-Projector Documentation

## Overview

The `apply_pdf_translation.py` script is a sophisticated PDF back-projector designed to replace Japanese text with English translations while preserving the original formatting, layout, and visual appearance of the document.

## Features

### Core Functionality
- **Precise Text Replacement**: Replace Japanese text at exact original positions
- **Formatting Preservation**: Maintain font family, size, color, bold, italic attributes
- **Font Scaling**: Automatically calculate optimal font scaling to accommodate text expansion
- **Layout Adjustments**: Apply position and spacing adjustments for optimal readability
- **Multi-page Support**: Process documents with multiple pages consistently
- **Non-text Element Preservation**: Preserve images, graphics, annotations, and hyperlinks

### Advanced Features
- **Rotation Handling**: Support for rotated text elements
- **Color Preservation**: Maintain original text colors and highlighting
- **Metadata Preservation**: Keep document metadata, bookmarks, and outlines
- **Fallback Strategies**: Graceful handling of translation mismatches
- **Comprehensive Logging**: Detailed logging for debugging and quality assurance

## Installation

### Prerequisites
```bash
# Install PyMuPDF for PDF processing
pip install PyMuPDF

# Optional: Install for enhanced color handling
pip install pillow
```

### Dependencies
- **PyMuPDF (fitz)**: Core PDF processing library
- **Standard Library**: json, logging, argparse, pathlib, re, typing

## Usage

### Basic Usage
```bash
python scripts/apply_pdf_translation.py \
  --input original.pdf \
  --output translated.pdf \
  --translations translations.json
```

### Advanced Usage
```bash
python scripts/apply_pdf_translation.py \
  --input document.pdf \
  --output translated_document.pdf \
  --translations translations.json \
  --verbose
```

### Command Line Options

| Option | Short | Description | Required |
|--------|-------|-------------|----------|
| `--input` | `-i` | Input PDF file path | Yes |
| `--output` | `-o` | Output PDF file path | Yes |
| `--translations` | `-t` | Translation mappings JSON file | Yes |
| `--verbose` | `-v` | Enable verbose logging | No |

## Translation Data Format

### JSON Structure
The translations file should contain mappings from Japanese text to English translations with optional formatting adjustments.

#### List Format (Recommended)
```json
[
  {
    "original": "こんにちは世界",
    "translated": "Hello World",
    "font_scaling": 0.9,
    "layout_adjustments": {
      "position_offset": {
        "x": 0,
        "y": 0
      }
    }
  },
  {
    "original": "日本語のテキスト",
    "translated": "Japanese text",
    "font_scaling": 1.0
  }
]
```

#### Dictionary Format
```json
{
  "こんにちは世界": {
    "translated": "Hello World",
    "font_scaling": 0.9
  },
  "日本語のテキスト": {
    "translated": "Japanese text",
    "font_scaling": 1.0
  }
}
```

### Translation Fields

| Field | Type | Description | Required |
|-------|------|-------------|----------|
| `original` | string | Original Japanese text | Yes (list format) |
| `translated` | string | English translation | Yes |
| `font_scaling` | float | Font size scaling factor (0.7-1.0) | No |
| `layout_adjustments` | object | Position and layout adjustments | No |

### Layout Adjustments
```json
{
  "layout_adjustments": {
    "position_offset": {
      "x": 5,    // Horizontal offset in points
      "y": 2     // Vertical offset in points
    },
    "line_spacing": 1.2,  // Line spacing multiplier
    "char_spacing": 0.1   // Character spacing adjustment
  }
}
```

## Text Processing Details

### Japanese Text Detection
The script uses a comprehensive regex pattern to detect Japanese characters:
- **Hiragana**: \u3040-\u309f
- **Katakana**: \u30a0-\u30ff
- **Kanji**: \u3400-\u4dbf, \u4e00-\u9fff
- **CJK Punctuation**: \u3000-\u303f
- **Fullwidth Characters**: \uff00-\uffef

### Font Scaling Algorithm
The font scaling algorithm considers:
1. **Text Expansion Ratio**: Ratio of English to Japanese text length
2. **Available Space**: Original text bounding box dimensions
3. **Minimum Font Size**: Prevents text from becoming unreadable (70% minimum)
4. **Character Width**: Estimated based on font size and typeface

### Text Replacement Process
1. **Redaction**: Original text area is redacted (cleared)
2. **Scaling**: Font size is adjusted based on expansion ratio
3. **Insertion**: Translated text is inserted at original position
4. **Formatting**: Original font attributes are applied
5. **Adjustments**: Layout adjustments are applied if specified

## Error Handling

### Common Issues and Solutions

#### Translation Not Found
```
WARNING: No translation found for: '未知のテキスト'
```
**Solution**: Add the missing text to your translations file.

#### Font Scaling Issues
```
WARNING: Error calculating font scaling: division by zero
```
**Solution**: Ensure original text has measurable length.

#### PDF Processing Errors
```
ERROR: Failed to replace text in block page_0_block_0: [error details]
```
**Solution**: Check PDF file integrity and permissions.

### Log Files
- `pdf_translation.log`: Detailed processing log
- Console output: Real-time progress and error reporting

## Integration with Translation Pipeline

### Input Sources
The back-projector accepts translations from:
- **Manual Translation Files**: JSON files created by translators
- **Translation APIs**: Output from machine translation services
- **Bilingual Databases**: Export from translation management systems

### Output Quality Assurance
- **Accuracy Verification**: Text replacement at correct positions
- **Formatting Validation**: Original formatting preserved
- **Layout Integrity**: No unintended layout changes
- **Readability Assessment**: Font scaling produces readable text

## Performance Considerations

### Processing Speed
- **Small Documents** (< 10 pages): 2-5 seconds
- **Medium Documents** (10-50 pages): 5-15 seconds
- **Large Documents** (> 50 pages): 15-60 seconds

### Memory Usage
- **Typical Usage**: 50-200 MB RAM
- **Large Documents**: Up to 1 GB RAM for complex layouts

### Optimization Tips
1. **Batch Processing**: Process multiple documents sequentially
2. **Memory Management**: Close documents when not in use
3. **Cache Management**: Clear temporary files regularly
4. **Parallel Processing**: Use multiple processes for batch jobs

## Testing

### Running Tests
```bash
# Run all tests
python -m pytest tests/test_apply_pdf_translation.py -v

# Run specific test categories
python -m pytest tests/test_apply_pdf_translation.py::TestPDFBackProjector -v

# Run with coverage
python -m pytest tests/test_apply_pdf_translation.py --cov=scripts.apply_pdf_translation
```

### Test Coverage
- **Unit Tests**: Individual component functionality
- **Integration Tests**: End-to-end document processing
- **Error Handling**: Exception scenarios and edge cases
- **Performance Tests**: Large document processing

## Troubleshooting

### Common Issues

#### PDF Not Opening
```python
# Error: PyMuPDF not found
# Solution: Install PyMuPDF
pip install PyMuPDF
```

#### Translation Mismatches
```python
# Issue: Text not matching exactly
# Solution: Normalize whitespace and punctuation
text = text.strip().replace('\n', ' ').replace('  ', ' ')
```

#### Layout Problems
```python
# Issue: Text overflow after replacement
# Solution: Adjust font_scaling in translations file
"font_scaling": 0.8  # Reduce for better fit
```

### Debug Mode
Enable verbose logging for detailed debugging:
```bash
python scripts/apply_pdf_translation.py --input test.pdf --output out.pdf --translations trans.json --verbose
```

## API Reference

### PDFBackProjector Class

#### Methods
- `__init__(input_path, output_path, translations_path)`
- `load_translations()`
- `extract_text_blocks()`
- `process_document()`
- `find_best_translation_match(text)`
- `replace_text_in_block(page, text_block, translation)`
- `calculate_optimal_font_scaling(...)`
- `apply_layout_adjustments(page, text_block, translation)`

#### Attributes
- `input_path`: Source PDF file path
- `output_path`: Target PDF file path
- `translations_path`: Translation mappings file path
- `text_blocks`: Extracted text blocks
- `replacement_stats`: Processing statistics

### TextBlock Dataclass
- `page_num`: Page number (0-indexed)
- `bbox`: Bounding box coordinates (x0, y0, x1, y1)
- `text`: Original text content
- `font_name`: Font family name
- `font_size`: Font size in points
- `font_color`: RGB color tuple
- `is_bold`: Bold formatting flag
- `is_italic`: Italic formatting flag
- `block_id`: Unique block identifier

## Best Practices

### Translation Quality
1. **Context Awareness**: Provide translations that fit the document context
2. **Length Considerations**: Account for text expansion in translations
3. **Terminology Consistency**: Use consistent terminology throughout the document
4. **Cultural Adaptation**: Adapt translations for the target audience

### Document Preparation
1. **PDF Optimization**: Ensure PDFs are optimized for text extraction
2. **Font Embedding**: Use embedded fonts for consistent appearance
3. **Structure Preservation**: Maintain logical document structure
4. **Quality Control**: Review output documents for accuracy

### Performance Optimization
1. **Batch Processing**: Process multiple documents together
2. **Memory Management**: Monitor memory usage for large documents
3. **Error Recovery**: Implement retry logic for transient errors
4. **Logging**: Maintain detailed logs for debugging and audit trails

## Limitations

### Current Limitations
- **Complex Layouts**: May have issues with highly complex layouts
- **Image-based Text**: Cannot process text embedded in images
- **Dynamic Content**: Cannot process dynamic PDF form elements
- **Encryption**: Does not support password-protected PDFs

### Future Enhancements
- **OCR Integration**: Add support for image-based text extraction
- **Advanced Layout**: Better handling of complex multi-column layouts
- **Interactive Elements**: Support for form fields and interactive content
- **Batch Processing**: Enhanced batch processing capabilities

## Support

### Documentation Updates
This documentation is maintained alongside the source code. For the latest version, refer to the repository.

### Bug Reports
Report bugs and issues through the project's issue tracker with:
- Description of the issue
- Sample files (if possible)
- Error messages and logs
- System information

### Feature Requests
Submit feature requests through the project's issue tracker with:
- Detailed description of the requested feature
- Use case examples
- Implementation suggestions (if any)