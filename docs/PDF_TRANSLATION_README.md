# PDF Translation Pipeline

This document describes the complete PDF translation pipeline for Japanese-to-English translation with layout preservation.

## Overview

The PDF translation pipeline provides a complete end-to-end solution for translating Japanese PDF documents to English while preserving layout, formatting, and visual elements. The system integrates seamlessly with the existing PPTX translation pipeline, sharing cache and glossary systems.

## Architecture

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   PDF Extractor │ → │  Translation    │ → │  Layout Engine  │
│                 │    │    Engine       │    │                 │
│ - Text blocks   │    │ - Cache system  │    │ - Font scaling  │
│ - Layout info   │    │ - Glossary      │    │ - Overflow      │
│ - Page metadata │    │ - API calls     │    │   handling      │
└─────────────────┘    └─────────────────┘    └─────────────────┘
         │                                               │
         └───────────────────────┬───────────────────────┘
                                 │
                ┌─────────────────┐    ┌─────────────────┐
                │  Back-Projector │ → │     Auditor      │
                │                 │    │                 │
                │ - Apply text    │    │ - Quality       │
                │   replacements │    │   assessment    │
                │ - Preserve      │    │ - Residual JP   │
                │   formatting    │    │ - Layout check  │
                └─────────────────┘    └─────────────────┘
```

## Key Components

### 1. PDF Extractor (`scripts/extract_pdf.py`)
- Extracts Japanese text with precise layout information
- Identifies different content types (titles, headers, body text, tables)
- Preserves font metrics and positioning data
- Supports multiple extraction methods for robustness

### 2. Translation Engine (`scripts/translate_pdf.py`)
- Main orchestrator that coordinates all components
- Reuses PPTX pipeline's translation cache and glossary
- Supports batch translation for efficiency
- Handles offline and cache-only modes

### 3. Layout Engine (`scripts/pdf_layout_engine.py`)
- Optimizes font sizes to prevent text overflow
- Handles text expansion from Japanese to English
- Applies content-specific layout constraints
- Preserves visual hierarchy and readability

### 4. Back-Projector (`scripts/apply_pdf_translation.py`)
- Applies translations to the original PDF
- Preserves formatting and layout
- Uses precise text replacement algorithms
- Handles special cases like tables and headers

### 5. Auditor (`scripts/audit_pdf.py`)
- Generates comprehensive quality reports
- Detects residual Japanese characters
- Assesses layout integrity
- Provides actionable recommendations

## Usage

### Command Line Interface

```bash
# Basic translation
python scripts/translate_pdf.py --in document.pdf --out translated.pdf

# With specific model
python scripts/translate_pdf.py --in doc.pdf --out en_doc.pdf --model gpt-4o-mini

# Page range selection
python scripts/translate_pdf.py --in manual.pdf --out en_manual.pdf --pages 1-10

# Using glossary
python scripts/translate_pdf.py --in doc.pdf --out en_doc.pdf --glossary custom_glossary.json

# Cache-only mode
python scripts/translate_pdf.py --in doc.pdf --out en_doc.pdf --cache-only

# Offline mode
python scripts/translate_pdf.py --in doc.pdf --out en_doc.pdf --offline
```

### Makefile Commands

```bash
# Translate PDF using make
make translate-pdf INPUT=document.pdf OUTPUT=translated.pdf

# With options
make translate-pdf INPUT=doc.pdf OUTPUT=en_doc.pdf MODEL=gpt-4o-mini PAGES=1-5

# Test PDF translation pipeline
make test-pdf

# Run unit tests
make test
```

## Configuration

### Environment Variables

```bash
# Required for translation
export OPENAI_API_KEY="your-api-key-here"

# Optional model selection
export OPENAI_MODEL="gpt-4o-2024-08-06"

# Optional reasoning effort for GPT-5
export OPENAI_REASONING_EFFORT="high"
```

### Model Selection

| Model | Quality | Cost | Use Case |
|-------|---------|------|----------|
| `gpt-4o-2024-08-06` | High | Medium | Production, critical documents |
| `gpt-4o-mini` | Good | Low | Drafts, internal documents |
| `gpt-5` | Highest | High | Complex content, specialized terms |

### Glossary Support

Create a `glossary.json` file to ensure consistent terminology:

```json
{
  "マジセミ": "Majisemi",
  "ウェビナー": "webinar",
  "株式会社": "Corporation",
  "AI": "AI",
  "API": "API"
}
```

Or use list format:

```json
[
  {"original": "マジセミ", "translated": "Majisemi"},
  {"original": "ウェビナー", "translated": "webinar"}
]
```

## Output Files

The translation pipeline generates multiple output files:

### 1. Translated PDF
- Main output: `[name]_en.pdf`
- Complete translated document with preserved layout

### 2. Bilingual CSV
- Format: `[name]_bilingual.csv`
- Side-by-side Japanese and English text mapping
- Includes block-level metadata (position, font size, content type)

### 3. Audit Report
- Format: `[name]_audit.json`
- Comprehensive quality assessment
- Residual Japanese character detection
- Layout integrity verification

### 4. Updated Cache
- File: `translation_cache.json`
- Shared with PPTX pipeline
- Cumulative translation database

### 5. Log File
- File: `pdf_translation.log`
- Detailed execution log
- Error tracking and debugging information

## Performance Optimization

### Cache System
- Shared with PPTX pipeline for ~90% hit rate
- Automatic deduplication of identical text
- Persistent storage across translation sessions

### Batch Processing
- Intelligent batching based on model limits
- Automatic retry with reduced batch size on failures
- Progress tracking and ETA calculation

### Layout Optimization
- Smart font scaling to prevent overflow
- Content type-specific constraints
- Hierarchical optimization (titles > body > footers)

## Advanced Features

### Page Range Selection
```bash
# Single page
--pages 5

# Page range
--pages 1-10

# From page to end
--pages 3-
```

### Content Type Handling
- **Titles**: Preserve visual hierarchy, minimal scaling
- **Headers/Footers**: Flexible layout, can wrap
- **Body Text**: Optimal readability, moderate scaling
- **Tables**: Fixed layout, precise positioning
- **Captions**: Flexible, can expand downward

### Error Handling
- Graceful degradation on component failures
- Fallback translation modes
- Comprehensive error reporting
- Automatic retry mechanisms

## Integration with PPTX Pipeline

### Shared Resources
- **Translation Cache**: `translation_cache.json`
- **Glossary**: `glossary.json`
- **API Configuration**: Environment variables
- **Style Guides**: Consistent terminology

### Consistent Translation
- Same models and prompts
- Shared glossary terms
- Consistent formatting rules
- Unified quality standards

## Testing

### Unit Tests
```bash
# Run all tests
make test

# Run PDF-specific tests
make test-pdf

# Run specific test file
python -m pytest tests/test_translate_pdf.py -v
```

### Integration Tests
```bash
# Test complete pipeline (requires sample PDF)
python scripts/translate_pdf.py --in test/sample.pdf --out test/out.pdf --offline

# Validate output quality
python scripts/audit_pdf.py test/out.pdf
```

## Troubleshooting

### Common Issues

1. **Import Errors**
   ```
   ERROR: PDF translation components not found
   ```
   - Ensure all PDF scripts are in `scripts/` directory
   - Check Python path configuration

2. **API Key Issues**
   ```
   ERROR: OPENAI_API_KEY environment variable not set
   ```
   - Set environment variable: `export OPENAI_API_KEY="your-key"`
   - Use `--offline` mode for testing

3. **Component Failures**
   ```
   PDF translation failed: component not available
   ```
   - Install required dependencies
   - Check component availability logs

4. **Layout Issues**
   ```
   Layout optimization failed
   ```
   - Review page-by-page in audit report
   - Check for complex tables or graphics

### Debug Mode
```bash
# Enable verbose logging
python scripts/translate_pdf.py --in doc.pdf --out en_doc.pdf --verbose

# Check log file
tail -f pdf_translation.log
```

## Development

### Adding New Features

1. **Component Integration**
   - Implement new component in `scripts/`
   - Update imports in `translate_pdf.py`
   - Add integration tests

2. **CLI Options**
   - Add argument to parser in `main()`
   - Update help text and examples
   - Add validation logic

3. **Testing**
   - Add unit tests for new functionality
   - Update integration tests
   - Document breaking changes

### Code Style
- Follow existing patterns in PPTX pipeline
- Use consistent logging and error handling
- Maintain type hints and documentation
- Add comprehensive test coverage

## License

This project follows the same license as the main translation tools project.