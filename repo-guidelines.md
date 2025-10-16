# Repository Guidelines

## Project Structure & Module Organization
- `scripts/translate_pptx_inplace.py` drives PPTX translation with cache/style helpers; shared adapters live in `utils/`, `schemas/`, and `backend/app/`.
- Place source decks in `inputs/`, results in `outputs/`, and domain data under `data/`; tests and fixtures stay in `tests/`, run artifacts in `archived_runs/`.
- Optional UI prototypes sit in `frontend/`, while automation assets and schema definitions live in `tools/` and `schemas/`.

## Build, Test, and Development Commands
- Translate a deck: `python scripts/translate_pptx_inplace.py --in inputs/demo.pptx --out outputs/demo_en.pptx --model gpt-4o-2024-08-06`.
- Run regression checks: `python -m pytest tests -q`; focus suites with `-k translate_docx_cli_end_to_end`.
- Clean artifacts after experiments: `make clean MODE=light`; full DOCX pipeline smoke: `make docx-ci`.

## Coding Style & Naming Conventions
- Use 4-space indentation, snake_case modules/functions, PascalCase classes, and keep pipeline steps small, returning slide/block keyed dicts.
- Prefer f-strings, explicit `Path` usage, and `logging.getLogger(__name__)` for loggers.
- Run `ruff check .` before committing; keep docstrings imperative and avoid introducing global state.

## Testing Guidelines
- Pytest is the standard; fixtures live in `tests/fixtures/` and `tests/data/`.
- Export `OPENAI_API_KEY` (real or `dummy`) before smoke tests; assert on bilingual CSV and cache artifacts when feasible.
- New adapters should add at least one regression case and guard external API hits with marks or environment flags.

## Commit & Pull Request Guidelines
- Follow Conventional Commits (`feat:`, `fix:`, `test:`) as in history, and keep each commit scoped.
- Exclude generated PPTX, logs, and caches from commits; reference stored outputs by path instead.
- PRs must list commands run, residual JP counts or audit diffs, and call out any configuration changes required.

## Security & Configuration Tips
- Keep secrets in `.env` (see `.env.example`) and document required variables alongside new modules.
- Sync updates to `glossary.json` and `pricing.example.json` so sample configurations stay current.
