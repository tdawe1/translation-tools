# Repository Guidelines

## Project Structure & Module Organization
- `scripts/translate_pptx_inplace.py` drives PPTX translation with cache/style helpers; shared adapters live in `utils/`, `schemas/`, and `backend/app/`.
- Place source decks in `inputs/`, results in `outputs/`, and domain data under `data/`; tests and fixtures stay in `tests/`, run artifacts in `archived_runs/`.
- Optional UI prototypes sit in `frontend/`, while automation assets and schema definitions live in `tools/` and `schemas/`.

## Build, Test, and Development Commands
- Translate a deck (online/API): `python scripts/translate_pptx_inplace.py --in inputs/demo.pptx --out outputs/demo_en.pptx --model gpt-4o-2024-08-06`.
- Translate a deck (offline/no API; Codex‑friendly): `python scripts/translate_pptx_inplace.py --in inputs/demo.pptx --out outputs/demo_en.pptx --offline`.
- Run regression checks: `python -m pytest tests -q`; focus suites with `-k translate_docx_cli_end_to_end`.
- Clean artifacts after experiments: `make clean MODE=light`; full DOCX pipeline smoke: `make docx-ci`.

Notes
- Offline runs do not require a real API key. You may set `OPENAI_API_KEY=dummy` for tooling that expects it; it is not used when `--offline` is provided.
- Manual DOCX flow (fully local):
  - Prepare: `python scripts/manual_docx_translation.py prepare --input inputs/source.docx --template translations/source_template.json`
  - Apply:   `python scripts/manual_docx_translation.py apply --input inputs/source.docx --translations translations/source_translations.json --output outputs/source_en.docx`

## Coding Style & Naming Conventions
- Use 4-space indentation, snake_case modules/functions, PascalCase classes, and keep pipeline steps small, returning slide/block keyed dicts.
- Prefer f-strings, explicit `Path` usage, and `logging.getLogger(__name__)` for loggers.
- Run `ruff check .` before committing; keep docstrings imperative and avoid introducing global state.

## Testing Guidelines
- Pytest is the standard; fixtures live in `tests/fixtures/` and `tests/data/`.
- Export `OPENAI_API_KEY` (real for online tests; `dummy` is fine for offline smoke). Assert on bilingual CSV and cache artifacts when feasible.
- New adapters should add at least one regression case and guard external API hits with marks or environment flags.

## Commit & Pull Request Guidelines
- Follow Conventional Commits (`feat:`, `fix:`, `test:`) as in history, and keep each commit scoped.
- Exclude generated PPTX, logs, and caches from commits; reference stored outputs by path instead.
- PRs must list commands run, residual JP counts or audit diffs, and call out any configuration changes required.

## Security & Configuration Tips
- Keep secrets in `.env` (see `.env.example`) and document required variables alongside new modules.
- Sync updates to `glossary.json` and `pricing.example.json` so sample configurations stay current.

## Codex CLI vs API
- Default to offline/manual workflows when developing locally or pairing via Codex CLI. Use `--offline` for PPTX and the manual DOCX prepare/apply flow for DOCX.
- Only enable online/API translation when intentional. For online runs, export a real `OPENAI_API_KEY` and pass `--model` (e.g., `gpt-4o-2024-08-06`).
- The offline path still applies cache/style/formatting; it just skips network calls.
