**Proposed Restructure (Review Before Changes)**

Goals
- Clarify responsibilities by format (pptx vs docx)
- Separate runtime artifacts from source
- Reduce top‑level clutter; add docs for discoverability

Moves (non‑breaking paths kept by updating imports)
- `scripts/` → split by format:
  - `scripts/pptx/translate_inplace.py` (from translate_pptx_inplace.py)
  - `scripts/pptx/apply_cache_only.py`
  - `scripts/pptx/style_*`, `scripts/pptx/pptx_format.py`
  - `scripts/docx/manual_translation.py` (from manual_docx_translation.py)
  - `scripts/docx/adapter.py` (from docx_adapter.py)
- Group audits under `scripts/audit/`
  - `audit_pptx_jp_count.py` → `scripts/audit/pptx_jp_count.py`
- Add `logs/` for `translation.log` and similar (git‑ignored)
- Add `caches/` for cache snapshots (git‑ignored); keep `outputs/` for final deliverables
- Keep top‑level CLIs as shims if needed (import new modules and forward args)

Docs/Config
- Keep `STYLE_GUIDE.md` and `glossary.json` at root
- Add `docs/USAGE.md` with copy‑paste recipes
- Add `docs/CONTRIBUTING.md` (pytest, ruff, commit conventions)

Open Questions
- Keep `backend/` in this repo or split into a separate service repository?
- Preferred title casing policy by template vs audit suggestions

Next Steps
1) Approve the above mapping
2) I’ll create shims for renamed scripts to avoid breaking existing commands
3) Update imports + README/docs links
4) Run quick smoke: offline PPTX and manual DOCX

