**Architecture**

- Pipeline Overview
  - Extract segments (PPTX/DOCX)
  - Batch + translate (online API or offline/manual)
  - Style normalization + optional audit
  - Apply translations back preserving runs/layout
  - Format-tightening (autofit, margins, line spacing)
  - Emit artifacts (cache, bilingual CSV, audit)

- Key Components
  - `scripts/translate_pptx_inplace.py`: End‑to‑end PPTX pipeline with cache, style, and layout controls. Supports `--offline`.
  - `scripts/manual_docx_translation.py`: Two‑step DOCX flow (prepare template → apply filled translations) for fully local/manual work.
  - `scripts/docx_adapter.py`: Rich DOCX adapter (segment extraction, application) used by backend and tests.
  - `backend/translation_orchestrator.py`: Orchestration layer (service-facing wrapper) for adapters.
  - `scripts/style_*` and `scripts/pptx_format.py`: Deterministic style mechanics and formatting helpers.
  - `glossary.json`: Project glossary applied in prompts and manual reviews.
  - `STYLE_GUIDE.md` or `$STYLE_GUIDE_FILE`: Gengo‑aligned style guidance loaded into prompts.

- Artifacts
  - `translation_cache.json`: {JP: EN} mapping used for idempotent re‑runs and offline apply.
  - `bilingual.csv`: Slide/paragraph indexed JP↔EN pairs for QA.
  - `audit.json`: Residual JP counts (before/after) with per‑slide breakdown.
  - `translation.log`: Run log.

- Online vs Offline
  - Online: Uses OpenAI client (Responses/Chat) with JSON array contract; retries + batch splitting.
  - Offline: `--offline` mode returns mock translations; or use manual templates + `apply_cache_only.py`/DOCX apply.

- Safety & Layout
  - Masking of URLs/numbers/codes; unmask after.
  - PPTX XML: `<a:normAutofit>`, body insets, line spacing normalization.
  - Optional "spill to notes" policy when expansion exceeds thresholds (online mode).

