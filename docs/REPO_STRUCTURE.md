**Repository Structure**

- `scripts/` — CLI tools and helpers
  - `translate_pptx_inplace.py` — main PPTX translator (cache, style, formatting)
  - `manual_docx_translation.py` — DOCX prepare/apply (local/manual)
  - `docx_adapter.py` — DOCX adapter used by backend/tests
  - `apply_cache_only.py` — apply {JP:EN} cache to PPTX without API
  - `audit_pptx_jp_count.py` — residual JP counter for PPTX
  - `style_*`, `pptx_format.py` — style mechanics and formatting helpers
- `backend/` — orchestration + API prototype
- `inputs/` — source decks (not committed)
- `outputs/` — translated outputs + caches (not committed)
- `data/` — domain data (glossaries, configs)
- `tests/` — pytest suites + fixtures
- `schemas/` — schema definitions
- `tools/` — automation assets
- `frontend/` — optional UI prototypes
- `archived_runs/` — run artifacts
- `glossary.json` — project glossary
- `STYLE_GUIDE.md` — Gengo‑aligned style guide (or use `$STYLE_GUIDE_FILE`)

