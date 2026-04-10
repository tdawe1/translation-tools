# Translation PPTX Pipeline

Legacy lightweight translation utilities for PPTX and DOCX.

This repository is still usable for simple document-translation workflows, but new development has moved to [`translation-app`](https://github.com/tdawe1/translation-app), which covers a broader scope.

## Scope

Utilities for translating Japanese PPTX and DOCX files to English while preserving document structure.

## Primary Entrypoints

- `scripts/translate_pptx_inplace.py`: main PPTX translation pipeline
- `scripts/apply_cache_only.py`: apply an existing PPTX translation cache
- `scripts/manual_docx_translation.py`: prepare/apply manual DOCX translations
- `scripts/translate_docx.py`: DOCX CLI via backend orchestrator
- `backend/app/main.py`: minimal FastAPI DOCX API
- `tools/estimate_cost.py`: cost estimation utility

## Status

- PPTX path is the main implemented workflow.
- Manual DOCX path is usable.
- Backend DOCX API is minimal and primarily covered by tests.
- `scripts/translate_pptx_inplace.py --help` is currently broken due to an unescaped `%` in an argparse help string.
- PPTX `--offline` mode emits mock translations, not real offline translation.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install openai requests python-docx fastapi uvicorn pydantic-settings defusedxml pytest jsonschema
cp .env.example .env
```

Set `OPENAI_API_KEY` for online translation runs.

## Usage

### PPTX Translation

```bash
python scripts/translate_pptx_inplace.py \
  --in input.pptx \
  --out output_en.pptx \
  --glossary glossary.json
```

Outputs:

- `translation_cache.json`
- `bilingual.csv`
- `audit.json`
- `translation.log`

### PPTX Cache Apply

```bash
python scripts/apply_cache_only.py \
  --in input.pptx \
  --out output_en.pptx \
  --cache translation_cache.json
```

### DOCX Manual Workflow

Prepare template:

```bash
python scripts/manual_docx_translation.py prepare \
  --input input.docx \
  --template translations/template.json
```

Apply completed translations:

```bash
python scripts/manual_docx_translation.py apply \
  --input input.docx \
  --translations translations/filled.json \
  --output output_en.docx
```

### DOCX API

```bash
uvicorn backend.app.main:app --reload
```

Endpoints:

- `GET /healthz`
- `GET /readyz`
- `POST /api/translate`
- `GET /api/jobs/{job_id}`
- `GET /api/jobs/{job_id}/download`

## Tests

```bash
python -m pytest tests -q
```

Focused suites:

```bash
python -m pytest tests/test_integration.py -q
python -m pytest tests/test_docx_adapter.py -q
python -m pytest tests/test_smoke_translate_docx.py -q
```

## References

- [Architecture](docs/ARCHITECTURE.md)
- [Repo Structure](docs/REPO_STRUCTURE.md)
- [Restructure Plan](docs/RESTRUCTURE_PLAN.md)
- [Style Guide](STYLE_GUIDE.md)
