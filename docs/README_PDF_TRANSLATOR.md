# PDF Translation Orchestrator

A complete end-to-end PDF translation system that translates Japanese PDFs to English while preserving layout and formatting.

## Overview

The `translate_pdf.py` script provides a production-ready PDF translation workflow that:

1. **Extracts** Japanese text with layout information using PyMuPDF
2. **Translates** text using the existing PPTX translation pipeline
3. **Optimizes** layout for English text expansion
4. **Back-projects** translations to PDF with formatting preservation
5. **Audits** quality and generates comprehensive reports

## Features

### Core Translation Pipeline
- **Text Extraction**: High-precision Japanese text extraction with layout preservation
- **Translation API Integration**: Uses existing PPTX batch translation system
- **Layout Optimization**: Intelligent font scaling and spacing adjustments
- **PDF Reconstruction**: Precise text replacement while maintaining formatting
- **Quality Assurance**: Automated audit reports and bilingual CSV generation

### CLI Interface (PPTX-Compatible)
```bash
# Basic translation
python translate_pdf.py --in document.pdf --out translated.pdf

# Page range selection
python translate_pdf.py -i input.pdf -o output.pdf --pages 1-5

# Offline testing mode
python translate_pdf.py --in doc.pdf --out doc_en.pdf --offline

# Cache-only mode
python translate_pdf.py --in doc.pdf --out doc_en.pdf --cache-only --glossary custom.json
```

### Advanced Features
- **Unified Cache System**: Shares translation cache with PPTX pipeline
- **Page Range Support**: Translate specific pages or ranges
- **Offline Mode**: Mock translations for testing
- **Glossary Integration**: Custom terminology management
- **Batch Optimization**: Configurable batch sizes for API efficiency
- **Layout Preservation**: Maintains fonts, colors, positioning, and formatting

## Installation Requirements

### Core Dependencies
```bash
pip install PyMuPDF>=1.22.0     # PDF processing
pip install pdfplumber>=0.9.0    # Fallback extraction
pip install pypdf>=3.0.0         # PDF audit tools
```

### Translation Pipeline (Optional)
For full API integration:
```bash
pip install openai>=1.0.0        # OpenAI API client
```

## Usage Examples

### Basic Translation
```bash
python scripts/translate_pdf.py \
  --in examples/sample.pdf \
  --out examples/sample_en.pdf
```

### Advanced Usage
```bash
python scripts/translate_pdf.py \
  --in examples/document.pdf \
  --out examples/document_en.pdf \
  --model gpt-4o-2024-08-06 \
  --batch 12 \
  --pages 1-10 \
  --glossary custom_glossary.json \
  --cache shared_cache.json
```

### Offline Testing
```bash
python scripts/translate_pdf.py \
  --in examples/test.pdf \
  --out examples/test_en.pdf \
  --offline \
  --verbose
```

## Output Files

| File | Description |
|------|-------------|
| `output_en.pdf` | Translated PDF with preserved formatting |
| `output_en_bilingual.csv` | Side-by-side Japanese-English mapping |
| `output_en_audit.json` | Quality assessment and metrics |
| `translation_cache.json` | Updated translation cache |

## Integration Components

### PDF Components Used
- **`extract_pdf.py`**: Text extraction with layout preservation
- **`pdf_layout_engine.py`**: Layout optimization for text expansion
- **`apply_pdf_translation.py`**: PDF text replacement and formatting
- **`audit_pdf.py`**: Quality assessment and audit reporting

### PPTX Pipeline Integration
- **`batch_translate()`**: Shared translation function
- **Cache System**: Unified translation cache
- **Style Guides**: Consistent terminology and formatting
- **Glossary Support**: Custom terminology management

## Configuration

### Environment Variables
```bash
# API Configuration
export OPENAI_API_KEY="your-api-key-here"
export OPENAI_MODEL="gpt-4o-2024-08-06"

# Style Configuration (optional)
export STYLE_PRESET="gengo"
export STYLE_GUIDE_FILE="path/to/style.json"
```

### Translation Models
- **Conservative**: `gpt-4o-2024-08-06` (highest quality)
- **Balanced**: `gpt-4o-2024-08-06` (recommended)
- **Cost-effective**: `gpt-4o-mini` (good quality, lower cost)

## CLI Reference

```
usage: translate_pdf.py [-h] --in INP --out OUTP [--cache CACHE]
                       [--glossary GLOSSARY] [--model MODEL] [--batch BATCH]
                       [--pages PAGES] [--offline] [--cache-only]
                       [--no-backup] [--verbose]

PDF Japanese-to-English translator with layout preservation

options:
  -h, --help           show this help message and exit
  --in, -i INP         Input PDF file path
  --out, -o OUTP       Output PDF file path
  --cache CACHE        Translation cache file (default: translation_cache.json)
  --glossary GLOSSARY  Optional glossary JSON file
  --model MODEL        Translation model (default: gpt-4o-2024-08-06)
  --batch BATCH        Batch size for API calls (default: 10)
  --pages PAGES        Page range to translate (e.g., "1-5" or "3")
  --offline            Run in offline mode with mock translations
  --cache-only         Use only cached translations (no API calls)
  --no-backup          Skip backing up existing output files
  --verbose, -v        Enable verbose logging
```

## Quality Assurance

### Audit Metrics
- **Residual Japanese Detection**: Identifies untranslated Japanese text
- **Layout Integrity**: Verifies formatting and structure preservation
- **Text Completeness**: Ensures no content was lost in translation
- **Expansion Analysis**: Monitors English text growth ratios

### Output Quality Checks
- Automatic detection of translation failures
- Layout overflow prevention
- Font scaling optimization
- Formatting consistency verification

## Troubleshooting

### Common Issues

**Import Errors**
```
ERROR: PyMuPDF (fitz) is required
```
Solution: `pip install PyMuPDF>=1.22.0`

**Cache Issues**
- Verify cache file permissions
- Use `--fresh` flag to create new cache
- Check JSON format validity

**API Failures**
- Verify `OPENAI_API_KEY` is set
- Check model availability
- Reduce batch size for rate limiting

### Performance Optimization

**Large Documents**
- Use `--pages` to process in chunks
- Optimize batch size based on document complexity
- Enable verbose logging for progress monitoring

**Memory Management**
- Process large PDFs page by page
- Clear temporary files after processing
- Monitor system resources during translation

## Development

### Running Tests
```bash
# Basic functionality tests
python scripts/test_translate_pdf_simple.py

# Unit tests (requires full dependencies)
python -m pytest scripts/test_translate_pdf.py -v
```

### Component Architecture
```
translate_pdf.py (Main Orchestrator)
├── extract_pdf.py (Text Extraction)
├── pdf_layout_engine.py (Layout Optimization)
├── apply_pdf_translation.py (PDF Reconstruction)
├── audit_pdf.py (Quality Assessment)
└── translate_pptx_inplace.py (Translation Pipeline)
```

## Contributing

1. **Component Updates**: Maintain compatibility with existing interfaces
2. **CLI Changes**: Preserve PPTX translator compatibility
3. **Cache System**: Use unified cache for consistency
4. **Error Handling**: Follow existing error handling patterns
5. **Testing**: Include unit tests for new functionality

## License

This project follows the same license as the parent translation pipeline.