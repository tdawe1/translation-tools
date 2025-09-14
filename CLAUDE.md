# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **production-ready Japanese-to-English document translation pipeline** that supports both **PowerPoint presentations and PDF documents**. The system specializes in preserving layout, formatting, and visual elements while translating content. It processes documents by extracting Japanese text, batching content for optimal AI model usage, translating using OpenAI's GPT models, and replacing text while maintaining original formatting.

## Development Commands

### Core Translation
```bash
# Basic translation (recommended)
python scripts/translate_pptx_inplace.py --in input.pptx --out output_en.pptx --model gpt-4o-2024-08-06

# Cost-optimized option
python scripts/translate_pptx_inplace.py --in input.pptx --out output_en.pptx --model gpt-4o-mini

# Offline translation using cache
python scripts/translate_pptx_inplace.py --offline --in input.pptx --out output_en.pptx

# PDF Translation
python scripts/translate_pdf.py --in input.pdf --out output_en.pdf --model gpt-4o-2024-08-06

# PDF with page range
python scripts/translate_pdf.py --in input.pdf --out output_en.pdf --pages 1-10

# PDF offline translation
python scripts/translate_pdf.py --offline --in input.pdf --out output_en.pdf
```

### Development Tools
```bash
# Cost estimation
make estimate
# or
python tools/estimate_cost.py input.pptx --pricing pricing.example.json --producer openai:gpt-5

# PDF cost estimation
make estimate-pdf PDF_INPUT=input.pdf
# or
python tools/estimate_cost_pdf.py input.pdf --model openai:gpt-5 --pages 1-20

# Tone analysis
make derive-tone
# or  
python tools/derive_deck_tone.py input.pptx

# Style checking
python scripts/audit_style.py output_en.pptx

# Cache management
python scripts/scrub_cache.py

# Translation-only audit (excludes images/charts)
python scripts/audit_translated_only.py output_en.pptx

# PDF text extraction (standalone)
python scripts/extract_pdf.py --input document.pdf --output extracted.json --format translation

# PDF layout auditing
python scripts/audit_pdf.py translated_document.pdf
```

### Testing
```bash
# Run all tests
python -m pytest tests/

# Run specific test
python -m pytest tests/test_estimate_cost.py

# Test with verbose output
python -m pytest tests/ -v

# PDF-specific tests
python -m pytest tests/test_translate_pdf.py -v

# PDF integration tests
make test-pdf
```

### Build and Deployment
```bash
# Clean up artifacts
make clean
# or
./scripts/cleanup.sh aggressive

# Clean PDF artifacts only
make clean-pdf

# Clean all artifacts
make clean-all

# GitHub Actions workflow (manual trigger)
# Uses .github/workflows/translate-pptx.yml
```

## Architecture Overview

### Core Components

**Main Translation Engines**:
- `scripts/translate_pptx_inplace.py` - Direct XML manipulation of PPTX files (no PowerPoint libraries required)
- `scripts/translate_pdf.py` - PDF document translation with layout preservation
- Smart batch processing with token-aware sizing for both formats
- Translation caching (JSON-based) for cost control (~90% hit rate)
- Word-aware text replacement maintaining formatting

**Style System**:
- `scripts/style_checker.py` - Style consistency validation
- `scripts/style_mechanics_normalize.py` - Mechanical style fixes
- `scripts/style_autofix_from_report.py` - Automated style corrections
- `scripts/audit_style.py` - Quality analysis and reporting

**Audit Tools**:
- `scripts/audit_pptx_jp_count.py` - Japanese character counting for PPTX
- `scripts/audit_pdf.py` - PDF document quality auditing
- `scripts/audit_translated_only.py` - Focused audit on translated text only
- `scripts/scrub_cache.py` - Cache cleanup and optimization

**Layout Management**:
- `scripts/scale_pptx_fonts.py` - Font scaling for overflow prevention
- `scripts/expansion_policy.py` - Text expansion handling
- `scripts/extract_pdf.py` - PDF text extraction with layout preservation
- Auto-fit modes: `norm`, `shape`, `none`
- Page range processing for PDFs

### Key Technologies

- **Language**: Python 3.12+
- **AI Models**: OpenAI GPT-4o, GPT-4o-mini, GPT-5
- **Dependencies**: `openai`, `google-api-python-client`, `google-auth`, `tiktoken`
- **File Processing**: 
  - PPTX: Direct ZIP/XML parsing 
  - PDF: PyMuPDF (fitz) with pdfplumber fallback
- **Testing**: pytest with comprehensive mocking
- **CI/CD**: GitHub Actions with Google Drive integration

### Data Flow

1. **Extraction**: 
   - PPTX: Parse XML to extract Japanese text from slides and notes
   - PDF: Extract text blocks with layout information using PyMuPDF/pdfplumber
2. **Batching**: Group content optimally based on model token limits
3. **Translation**: Send batches to OpenAI API with strict JSON response format
4. **Caching**: Store translations locally to avoid re-translation costs
5. **Replacement**: Word-aware text replacement preserving formatting
6. **Style**: Multi-stage style processing and normalization
7. **Audit**: Quality checks and residual Japanese character detection

## Configuration

### Environment Variables (`.env`)
```bash
# Required
OPENAI_API_KEY=sk-your-openai-key

# Optional AI model settings
OPENAI_MODEL=gpt-5
OPENAI_TEMPERATURE=0.6

# Feature flags
ENABLE_STYLE_CHECKING=1
ENABLE_EXPANSION_POLICY=1
ENABLE_FORMATTING_PROFILE=1

# Google Drive integration
GOOGLE_OAUTH_CLIENT_ID=
GOOGLE_OAUTH_CLIENT_SECRET=
GOOGLE_OAUTH_REFRESH_TOKEN=
# or service account
GDRIVE_SA_JSON=
```

### Glossary (`glossary.json`)
Terminology consistency mapping for key terms:
```json
{
  "マジセミ": "Majisemi",
  "ウェビナー": "webinar",
  "株式会社": "Corporation"
}
```

### Model Pricing (`pricing.example.json`)
Configuration for cost estimation across different AI providers and models.

## Production Features

### Smart Batch Processing
- **Auto-optimization**: Calculates optimal batch sizes based on model limits
- **Dynamic adjustment**: Reduces batch size on high retry rates
- **Token-aware**: Respects model token limits (8k-16k target range)

### Error Resilience
- **Progressive backoff**: 1s, 2s, 3s delays on retries
- **Graceful degradation**: Falls back to smaller batches on failures
- **Cache recovery**: Preserves work through interruptions

### Layout Preservation
- **Word-aware replacement**: Inserts `<a:br/>` correctly, never cuts words mid-run
- **Auto-fit modes**: Shrink-to-fit, shape-aware, or none
- **Font scaling**: Prevents text overflow with configurable minimums

### Quality Assurance
- **Multi-stage processing**: Pre-translation normalization → translation → post-processing
- **Bilingual output**: CSV mapping for quality review
- **Audit trails**: Comprehensive statistics and residual detection

## Testing Strategy

### Test Structure
- `tests/test_estimate_cost.py` - Cost estimation and token counting
- `tests/test_style_checker.py` - Style validation and correction logic

### Testing Approach
- **pytest framework** with fixtures and mocking
- **External dependencies mocked** via sys.modules for reliable testing
- **Comprehensive coverage** of core algorithms and edge cases
- **Fake PPTX generation** for testing file parsing logic

## Development Workflow

### Local Development
1. Set up `.env` file with API keys
2. Use `make estimate` to preview costs
3. Run translation with small batch sizes initially
4. Audit results with `scripts/audit_style.py`
5. Apply style fixes as needed

### CI/CD Pipeline
- **GitHub Actions** workflow for automated translation
- **Google Drive integration** for input/output files
- **Artifact preservation** of all generated files
- **Configurable models** and parameters via workflow inputs

### Performance Optimization
- **Cache-first approach**: Minimizes API costs through local caching
- **Batch optimization**: Reduces API call overhead
- **Model selection**: Balance between quality (gpt-4o) and cost (gpt-4o-mini)

## Common Patterns

### Error Handling
- Structured logging with timestamps
- Graceful degradation on failures
- Comprehensive error messages with recovery suggestions

### Configuration Management
- Environment variables for secrets and settings
- JSON files for static configuration (glossary, pricing)
- Feature flags for optional functionality

### Output Organization
- Timestamped backups of existing files
- Consistent output file naming conventions
- Comprehensive audit trails and metrics