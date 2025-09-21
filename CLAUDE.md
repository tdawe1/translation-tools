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

### Makefile Commands (Recommended)
```bash
# Translation
make translate-pptx INPUT=input.pptx OUTPUT=output.pptx MODEL=gpt-4o-mini
make translate-pdf INPUT=document.pdf OUTPUT=document_en.pdf MODEL=gpt-4o-mini

# Cost estimation
make estimate                                    # PPTX estimation
make estimate-pdf PDF_INPUT=document.pdf         # PDF estimation
make estimate-all                               # Both formats

# Testing
make test                                       # Run all tests
make test-pdf                                   # PDF-specific tests
make test-all                                   # All test suites

# Cleanup
make clean                                      # Remove artifacts
make clean-pdf                                  # PDF artifacts only
make clean-all                                  # All artifacts

# Development tools
make derive-tone                                # Analyze document tone
```

### Development Tools
```bash
# Cost estimation (detailed)
python tools/estimate_cost.py input.pptx --pricing pricing.example.json --producer openai:gpt-5
python tools/estimate_cost_pdf.py input.pdf --model openai:gpt-5 --pages 1-20

# Style checking and audit
python scripts/audit_style.py output_en.pptx                    # Full style audit
python scripts/audit_translated_only.py output_en.pptx          # Text-only audit
python scripts/style_autofix_from_report.py report.json        # Auto-fix issues

# Cache management
python scripts/scrub_cache.py                                  # Clean cache
python scripts/apply_cache_only.py --in input.pptx --out output.pptx  # Offline apply

# PDF tools
python scripts/extract_pdf.py --input document.pdf --output extracted.json --format translation
python scripts/audit_pdf.py translated_document.pdf
```

### Frontend Development (Next.js)
```bash
cd frontend
npm run dev        # Development server
npm run build      # Production build
npm run start      # Production server
npm run lint       # ESLint check
```

### Backend Development (FastAPI)
```bash
cd backend
./start_backend.sh           # Start FastAPI server
# Server runs on http://localhost:8000
```

### Testing
```bash
# Python tests
python -m pytest tests/ -v                           # All tests with verbose output
python -m pytest tests/test_estimate_cost.py         # Specific test file
python -m pytest tests/test_translate_pdf.py -v     # PDF tests

# Integration tests
python test_integration.py                            # Full pipeline test
./run_tests.py                                       # Custom test runner
```

## Architecture Overview

### Monorepo Structure
```
.
├── scripts/               # Core translation scripts
│   ├── translate_pptx_inplace.py    # Main PPTX engine
│   ├── translate_pdf.py            # PDF translation engine
│   ├── style_checker.py            # Style validation
│   └── audit_*.py                  # Audit tools
├── tools/                 # Development utilities
├── tests/                 # Test suite
├── frontend/              # Next.js web interface
│   ├── src/              # React components
│   └── package.json      # Frontend dependencies
├── backend/               # FastAPI API server
│   ├── main.py           # FastAPI app
│   └── requirements.txt  # Backend dependencies
├── inputs/               # Source documents
├── outputs/              # Translated results
└── data/                 # Configuration files
```

### Core Translation Engines

**PPTX Processing**:
- Direct ZIP/XML manipulation without PowerPoint dependencies
- Word-aware text replacement preserving formatting
- Smart batch processing with token optimization
- Cache-first approach (~90% hit rate)

**PDF Processing**:
- PyMuPDF (fitz) primary extraction with pdfplumber fallback
- Layout preservation with block-level positioning
- Page range processing support
- Same batch translation system as PPTX

**Style System**:
- Multi-stage processing: normalize → translate → post-process
- Mechanical fixes (ASCII/full-width, punctuation, spacing)
- Title case preservation for headings
- Configurable style presets (gengo, minimal)

### Key Technologies

- **Backend**: Python 3.12+, FastAPI, OpenAI API
- **Frontend**: Next.js 15, React 19, TypeScript, Tailwind CSS
- **AI Models**: GPT-4o, GPT-4o-mini, GPT-5 with fallback chains
- **PDF Processing**: PyMuPDF, pdfplumber, pdfminer.six
- **Testing**: pytest with comprehensive mocking
- **Real-time**: WebSocket servers for progress tracking

### Data Flow

1. **Document Analysis**: Extract Japanese text while preserving layout context
2. **Batch Optimization**: Calculate optimal batch sizes based on model limits
3. **Translation**: Send to OpenAI with strict JSON response requirements
4. **Caching**: Store translations to minimize API costs
5. **Layout Application**: Replace text while maintaining formatting constraints
6. **Style Processing**: Apply mechanical fixes and consistency rules
7. **Audit & QA**: Generate reports and identify residual issues

## Configuration

### Environment Variables
```bash
# Required
OPENAI_API_KEY=sk-your-openai-key

# AI Model Settings
OPENAI_MODEL=gpt-5
OPENAI_TEMPERATURE=0.6
OPENAI_USE_RESPONSES=1

# Feature Flags
ENABLE_STYLE_CHECKING=1
ENABLE_EXPANSION_POLICY=1
ENABLE_FORMATTING_PROFILE=1

# Google Drive Integration
GOOGLE_OAUTH_CLIENT_ID=
GOOGLE_OAUTH_CLIENT_SECRET=
GOOGLE_OAUTH_REFRESH_TOKEN=
GDRIVE_SA_JSON=
```

### Configuration Files
- `glossary.json` - Terminology mappings
- `pricing.example.json` - Model cost configurations
- `prompt.example.md` - Custom translation prompts
- `.env` - Environment variables (see `.env.example`)

## Production Features

### Smart Processing
- **Auto-batch sizing**: Dynamically adjusts based on content complexity
- **Progressive backoff**: 1s, 2s, 3s delays on API failures
- **Graceful degradation**: Falls back to smaller batches on errors
- **Layout preservation**: Font scaling, auto-fit modes, margin control

### Quality Assurance
- **Multi-audit system**: Style, residual JP, and translation-only audits
- **Bilingual output**: CSV mapping for human review
- **Cache optimization**: Scrubbing and merging utilities
- **Real-time progress**: WebSocket tracking with ETA estimates

### Error Resilience
- **Retry logic**: Exponential backoff with jitter
- **Cache recovery**: Preserves work through interruptions
- **Structured logging**: Comprehensive error tracking
- **Fallback models**: Automatic model switching on failures

## Testing Strategy

### Test Categories
- **Unit tests**: Core algorithms and utilities
- **Integration tests**: Full pipeline with mocked APIs
- **PDF tests**: Document extraction and layout preservation
- **Style tests**: Validation and correction logic
- **Performance tests**: Batch optimization and caching

### Testing Tools
- **pytest**: Primary test framework
- **Mocking**: sys.modules injection for external dependencies
- **Fixtures**: Test data generation for PPTX/PDF
- **Integration**: Full end-to-end test scripts

## Development Workflow

### Setup
1. Copy `.env.example` to `.env` and configure API keys
2. Install dependencies: `pip install -r requirements.txt`
3. For frontend: `cd frontend && npm install`
4. For backend: `cd backend && pip install -r requirements.txt`

### Daily Development
1. Use `make estimate` to preview translation costs
2. Start with small batches for testing
3. Run audits after translation: `scripts/audit_style.py`
4. Apply style fixes: `scripts/style_autofix_from_report.py`
5. Use `--offline` mode for cached runs during development

### Git Workflow
- Feature branches: `feature/translation-pdf`
- Commit messages: Conventional commits format
- PR reviews: Required with automated checks
- CI/CD: GitHub Actions with artifact preservation

## Common Patterns

### Error Handling
```python
# Pattern: Graceful degradation with fallbacks
try:
    result = translate_batch(items)
except OpenAIError:
    # Retry with smaller batch
    result = translate_batch(items[:len(items)//2])
```

### Configuration Loading
```python
# Pattern: Environment variables with defaults
model = os.getenv('OPENAI_MODEL', 'gpt-4o-2024-08-06')
enable_style = os.getenv('ENABLE_STYLE_CHECKING', '1') == '1'
```

### Cache Management
```python
# Pattern: Normalized keys for cache lookup
cache_key = normalize_japanese(text)
if cache_key in translation_cache:
    return translation_cache[cache_key]
```

## Performance Considerations

### Batch Optimization
- Target token range: 8k-16k per batch
- Auto-adjust based on retry rates
- Split batches on JSON parsing failures
- Minimum batch size: 6 items

### Cost Management
- Cache hit rate target: >90%
- Model selection: gpt-4o-mini for cost savings
- Concurrent processing: 4-8 workers
- Token tracking and budget limits

### Scaling
- PDF processing: Memory-efficient block extraction
- Large files: Page range processing
- Concurrent jobs: Redis-based queue in backend
- WebSockets: Non-blocking progress updates