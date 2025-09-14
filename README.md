# 🚀 Document Translation Pipeline (JA→EN)

A production-ready translation system for converting Japanese PowerPoint presentations **and PDF documents** to English while preserving layout, formatting, and visual elements.

<details>
<summary><strong>📊 Project Status Summary (Today)</strong></summary>

## What we fixed/changed

* **Stopped word-splitting & newline bugs:** Replaced `set_para_text` with a word-aware version that inserts `<a:br/>` correctly and never cuts words mid-run.
* **Hardened JSON handling:** Replaced fragile bracket-counting with `JSONDecoder.raw_decode`, multi-strategy extraction, batch splitting on failure, and clampable auto-batch (`--max-array-items`) so chatty outputs don't kill runs.
* **Killed identity/non-translated cache entries:** Added `scrub_cache.py` and improved JP-count audit to ignore punctuation (e.g., `・`), reducing false positives.
* **Concurrency & batching tuned:** Demonstrated stable settings (`--concurrency 4–8`, small auto-batches) and removed accidental full re-runs (`--fresh`).
* **Style: mechanics only (not voice):**

  * Added `style_mechanics_normalize.py` and a stronger `style_autofix_from_report.py` to fix ASCII/full-width, dashes, %/¥ spacing, units, stray punctuation, ellipses, bullet punctuation—without altering tone.
  * Added a summarizer to see which rules the checker complains about most.
* **Scope-correct residuals:** New `audit_translated_only.py` counts JP **only** in translated EN text; ignores SmartArt/charts/images by design.
* **Overflow solved at the XML layer:** Introduced slide-safe layout knobs:

  * `<a:normAutofit>` (shrink to fit),
  * tighter `<a:bodyPr>` insets (optional),
  * normalized `<a:lnSpc>` (line spacing).
  * Exposed flags: `--autofit-mode {norm,shape,none}`, `--font-scale-min`, `--line-spacing-pct`, `--tight-margins`.
* **ET warning future-proofed:** Replaced `r.find(...) or ET.SubElement(...)` with explicit `if t is None: …`.

## Translation quality improvements

* **Cache refinement pass:** Normalized punctuation (NFKC), ranges (`–`), yen/percent formatting, time ranges, pluralization, and consistent webinar terminology (attendee, registrant, operations, etc.). Curated overrides for key headlines and Majisemi terms.
* **Title consistency:** Restored **Title Case** on headings only (acronym-aware, hyphen-aware), leaving bullets/body untouched.

## What we ran / artifacts produced

* **Runs:** Full online translations (4o/4o-mini) with batch split & retries → cache filled; offline apply to preserve layout.
* **Artifacts:**

  * `translation_cache.refined.json` — mechanics/terminology upgrades.
  * `translation_cache.retitled.json` — refined + Title Case for headings.
  * `cache_diff.csv` — JP | old EN | new EN.
  * Updated PPTX outputs (e.g., `outputs/styled_offline.pptx`).

## Current state

* **Formatting:** Fixed via autofit + margins + line spacing; only a **short manual pass** needed where English still pushes boundaries.
* **Style checker:** Mechanical issues substantially reduced; remaining flags are mostly preference/tone or out-of-scope artifacts if the old audit is used.
* **Cache:** Ready to use.

### Apply the cache offline

```bash
cp translation_cache.retitled.json translation_cache.json
python3 scripts/translate_pptx_inplace.py --offline \
  --in inputs/68b42f175c652_f711fcda865b11f0b6cecace4a312dcf.pptx \
  --out outputs/final_retitled.pptx
```

## Optional next steps (small but high-ROI)

* **Bake title-case at write-time:** Only for Title/CenteredTitle placeholders.
* **Use translated-only audit in CI:** Drop legacy residual counters that include non-text artifacts.
* **Set sensible defaults:** `--autofit-mode norm --font-scale-min 90000 --line-spacing-pct 100000 --tight-margins`.
* **Clamp batching:** Respect `--max-array-items` min=6 to reduce JSON hiccups on short blurbs.

</details>

<details>
<summary><strong>🎯 Next Steps & Roadmap (Zero-Touch, GPT-5, More Formats)</strong></summary>

Here's a tight, no-nonsense plan to make this zero-touch, faster, and broader.

## A) Zero-touch "Drive-in / Drive-out" pipeline

**Goal:** User drops a file in Drive → system detects → translates → uploads finished pack to Drive (final PPTX/Doc + bilingual CSV + audit) → optional Slack/email ping.

1. **Folder contract (no UI needed)**

* `Drive:/TranslationInbox/` (incoming, read-only to users)
* `Drive:/TranslationOut/` (deliverables)
* `Drive:/TranslationArchive/` (originals + logs)

2. **Detection**

* **Simplest & robust:** GitHub Actions (cron every 2–5 min) + Drive Changes API using a stored `startPageToken`.
* Keep a `jobs/STATE.json` in the repo (or Redis) with processed file IDs to avoid duplicates.

3. **Job manifest**

* On new file: create `job_<fileId>.json` with:

  ```json
  {"fileId":"...", "name":"...", "mime":"application/vnd.openxmlformats-officedocument.presentationml.presentation",
   "created":"...", "status":"QUEUED", "model":"gpt-4o-mini", "style":"gengo"}
  ```
* Status transitions: QUEUED → EXTRACTING → TRANSLATING → QA → DELIVERED (or FAILED, with reason).

4. **Processing runner (idempotent)**

* **Extract** → **Batch translate** (cache-first, slide/block-level) → **Autofit & layout pass** → **Style/autofix** → **Translated-only audit** → **Package**.
* Drive upload results to `TranslationOut/` with a suffix: `originalName.en-US.[timestamp].pptx` plus CSV/JSON.
* Always move source to `TranslationArchive/` and attach a `job.log`.

5. **Notifications**

* Optional: Email via Gmail API or Slack webhook with links to Drive outputs + a small summary (residual=0, changed=N, cost estimate).

6. **Observability**

* Write compact metrics per job: `tokens_in/out`, `cache_hit_rate`, `api_errors`, total duration.
* (Optional) OpenAI **Webhooks**: add a `/webhook` endpoint (tiny Flask/Cloud Run) to receive translate batch updates; mirror to the job manifest.

**Definition of Done**

* Drop a PPTX in `TranslationInbox/` → see final artifacts in `TranslationOut/` with residual JP=0 (translated-only audit) and a Drive comment or Slack ping.

---

## B) Make GPT-5 work (without breaking anything)

**Problem you saw:** `response_format` caused failures; client/model mismatch.

1. **Adapter layer**

* Add `llm_adapter.py` with a single `translate_batch(items, model, sys, temp)` that:

  * Uses `/chat/completions` for 4o/4o-mini/4.1; **no `response_format`**.
  * If `model.startswith("gpt-5")`, choose the correct endpoint & params (no unsupported args).
  * Always wrap prompts with a strict **JSON-array only** instruction and validate with `JSONDecoder.raw_decode` (you already added).
  * Feature flag: `--primary-model gpt-5` with **fallback chain** (`5 → 4.1 → 4o → 4o-mini`) on capability/HTTP errors.

2. **Compatibility switch**

* Centralize all OpenAI kwargs in one place; forbid stray params.
* Add `--dry-run` to print composed payloads for a single batch to verify.

3. **Resilience**

* Per-batch retries with backoff; if still chatty → auto-split (`--on-batch-fail split` already there).
* **Cost guard:** `--max-output-tokens` and a per-job token budget; abort gracefully if exceeded.

4. **Tests**

* Golden tests: 10 representative JP lines → verify strict JSON array parse and stable output across models.

**DoD**

* `--primary-model gpt-5` runs end-to-end; if not available, auto-fallback without failing the job.

---

## C) Expand to other document types (shared engine)

**Unify on a "Document Abstraction Layer" (DAL)**

1. **Common interfaces**

```python
class Extractor:
    def extract(self, path) -> list[Block]  # Block = {id, kind, meta, jp_text}
class BackProjector:
    def apply(self, path_in, path_out, translations: dict[id, en_text]) -> None
```

2. **Handlers (start with easiest)**

* **DOCX**: `python-docx`. Extract paragraphs, headings, tables (cell text). Back-project by run order; preserve styles.
* **Markdown / TXT**: trivial; line/block based.
* **XLSX**: `openpyxl`. Translate **values only** (skip formulas). Keep data types; don't touch numbers/dates.
* **SRT/VTT**: segment by cue; preserve timestamps.
* **PDF (text-only first)**: `pdfminer.six` for text; back-project as **bilingual PDF** or export to DOCX then reassemble (full layout-faithful PDF is a separate project—defer).
* **Google Docs/Slides**: fetch via Drive export (DOCX/PPTX) and reuse the above; native API mapping optional later.

3. **Re-use your core**

* Same **batch translator**, cache, glossary, style autofix, and translated-only audit.
* Same **autofit** concept where applicable:

  * DOCX: allow "Automatically adjust right indent when grid is defined"; tighten spacing; avoid font size drops below a floor.
  * XLSX: enable wrap, column autosize (optional).

**DoD**

* Drop a DOCX/XLSX/TXT → get translated file + bilingual CSV in `TranslationOut/`.

---

## D) Friction killers (small but high-impact)

* **One command for everything:** `tt submit <local_file>` → uploads to `TranslationInbox/` and pings the runner.
* **Cache across projects:** move cache to a shared KV (SQLite/Redis) with normalized JP keys (`NFKC + whitespace fold`) and optional fuzzy (rapidfuzz) for ≥0.96 similarity.
* **Defaults baked in:** `--autofit-mode norm --font-scale-min 90000 --line-spacing-pct 100000 --style-preset gengo`.
* **Strict gates:** Use **translated-only** audit in CI; fail if residual>0; warn on style mechanics only.
* **Cost estimator:** quick preflight on extracted blocks: estimated tokens × model price → attach to job manifest and Slack ping.

---

## E) Concrete next tasks (merge-friendly)

1. **Jobs + Drive poller (GH Action + small script)**

   * `scripts/drive_poller.py` (changes.list → manifest → enqueue).
   * Action workflow `on: schedule`: runs every few minutes; uses SA creds; posts status.

2. **LLM adapter + fallback**

   * `llm_adapter.py` with endpoint/param matrix; add `--primary-model` and fallback list.

3. **DAL + DOCX handler**

   * `extract_docx.py` / `apply_docx.py`; wire into `translate_any.py` driver (detect by MIME/extension).

4. **Notifier**

   * Simple Slack webhook or Gmail email with Drive links and cost/timing.

5. **Cache sharing**

   * `cache_store.py` with SQLite file `cache.db` (table: jp\_norm TEXT PK, en TEXT, ts INT, src TEXT).

6. **Defaults + flags cleanup**

   * Config file `.translationrc` (YAML/JSON) for model, style, autofit defaults; CLI reads it so you don't have to pass flags.

If you want, I can turn this into three small PRs: (1) Drive poller + job runner, (2) LLM adapter + GPT-5 fallback, (3) DAL with DOCX handler.

</details>

## ✨ Features

### 🎯 **Production-Ready Translation**
- **Multi-format support**: PowerPoint presentations AND PDF documents
- **Smart batch sizing**: Auto-optimizes API requests per model
- **Comprehensive logging**: Real-time progress with ETA estimates  
- **Robust error handling**: Auto-retry with intelligent backoff
- **Layout preservation**: Maintains original formatting and design

### 🧠 **AI-Powered Quality**
- **Style consistency**: Unified tone and terminology across slides
- **Content-aware processing**: Handles titles, bullets, tables differently
- **Expansion management**: Prevents text overflow with smart compression
- **Glossary integration**: Ensures consistent translation of key terms

### 📊 **Advanced Features**
- **Translation caching**: Avoids re-translating identical content
- **Bilingual output**: CSV mapping for quality assurance
- **Performance metrics**: Detailed audit reports and statistics
- **Page range support**: Translate specific sections of PDFs
- **Fallback extraction**: PyMuPDF + pdfplumber for complex layouts
- **Webhook integration**: Real-time progress tracking (optional)

## 🚀 Quick Start

### Prerequisites
```bash
export OPENAI_API_KEY=your_key_here
```

### Basic Usage
```bash
# Production presets (recommended)
python scripts/translate_pptx_inplace.py \
  --in input.pptx \
  --out output_en.pptx \
  --model gpt-4o-2024-08-06

# Cost-optimized option
python scripts/translate_pptx_inplace.py \
  --in input.pptx \
  --out output_en.pptx \
  --model gpt-4o-mini
```

## 📄 PDF Translation

The pipeline now supports Japanese-to-English translation for PDF documents with the same quality and layout preservation as the PPTX system.

### Quick Start - PDF Translation

```bash
# Basic PDF translation
python scripts/translate_pdf.py --in document.pdf --out document_en.pdf

# Cost estimation for PDF
python tools/estimate_cost_pdf.py document.pdf

# Using Makefile (recommended)
make translate-pdf INPUT=document.pdf OUTPUT=document_en.pdf
make estimate-pdf PDF_INPUT=document.pdf
```

### PDF-Specific Features

- **Layout Preservation**: Maintains original formatting, fonts, and positioning
- **Page Range Support**: Translate specific pages or ranges (e.g., `--pages 1-10`)
- **Smart Text Extraction**: Uses PyMuPDF with pdfplumber fallback for complex layouts
- **Japanese Text Detection**: Automatically identifies and extracts Japanese content
- **Batch Processing**: Same intelligent batching as PPTX for optimal API usage
- **Cache Integration**: Reuses existing translation cache for cost savings

### PDF Command Line Options

```bash
python scripts/translate_pdf.py [OPTIONS]

Required:
  --in INPUT.pdf       Input PDF file
  --out OUTPUT.pdf     Output translated PDF file

Optional:
  --model MODEL        AI model (default: gpt-4o-2024-08-06)
  --pages PAGES        Page range (e.g., 1-10, 5)
  --glossary FILE      Custom glossary file
  --cache FILE         Cache file path
  --cache-only         Use only cached translations
  --offline            Run in offline mode
  --verbose            Enable detailed logging
```

### PDF Cost Estimation

```bash
# Basic cost estimation
python tools/estimate_cost_pdf.py document.pdf

# With specific model and pages
python tools/estimate_cost_pdf.py document.pdf --model openai:gpt-5 --pages 1-20

# Using Makefile
make estimate-pdf PDF_INPUT=document.pdf MODEL=openai:gpt-5 PAGES=1-20
```

## 🎛️ Production Presets

| Preset | Model | Batch Size | Use Case |
|--------|-------|------------|----------|
| **Conservative** | `gpt-4o-2024-08-06` | 8-12 (auto) | Maximum reliability |
| **Balanced** | `gpt-4o-2024-08-06` | 10-14 (auto) | **Recommended** |
| **Cost-lean** | `gpt-4o-mini` | 12-16 (auto) | Good quality, lower cost |

*Batch sizes are automatically calculated based on content complexity and token limits.*

## 📋 Command Line Options

```bash
python scripts/translate_pptx_inplace.py [OPTIONS]

Required:
  --in INPUT.pptx          Input PowerPoint file
  --out OUTPUT.pptx        Output translated file

Optional:
  --model MODEL           AI model (default: auto-optimized)
  --batch N               Batch size (default: auto-calculated)
  --cache FILE            Translation cache (default: translation_cache.json)
  --glossary FILE         Terminology glossary (default: glossary.json)
  --slides RANGE          Process specific slides (e.g., "1-10")
  --style-preset PRESET   Style guide preset (gengo, minimal)
```

## 📁 Project Structure

```
├── scripts/
│   ├── translate_pptx_inplace.py  # Main translation engine
│   ├── style_checker.py           # Style consistency system
│   ├── eta.py                     # Progress estimation
│   ├── webhook_server.py          # Real-time progress tracking
│   └── audit_style.py            # Quality analysis
├── tools/
│   ├── derive_deck_tone.py       # Tone analysis
│   └── estimate_cost.py          # Cost estimation
├── inputs/                       # Source presentations
├── outputs/                      # Translated results
└── data/                        # Glossaries and configs
```

## 🔧 Advanced Configuration

### Custom Glossary
Create `glossary.json` for consistent terminology:
```json
{
  "株式会社": "Corporation",
  "取締役": "Director",
  "戦略": "Strategy"
}
```

### Style Consistency
Configure tone and style preferences:
```json
{
  "formality": "business_formal",
  "technical_terms": "preserve_english",
  "bullet_style": "concise_fragments"
}
```

### Webhook Progress Tracking
Run the webhook server for real-time updates:
```bash
# Terminal 1: Start webhook server
uvicorn scripts.webhook_server:app --port 8000

# Terminal 2: Run translation
python scripts/translate_pptx_inplace.py --in input.pptx --out output.pptx
```

## 📊 Output Files

Each translation run generates:

| File | Description |
|------|-------------|
| `output_en.pptx` | Translated presentation |
| `bilingual.csv` | Side-by-side translation mapping |
| `audit.json` | Translation statistics and metrics |
| `translation_cache.json` | Cached translations for efficiency |
| `translation.log` | Detailed execution log |

## 🛠️ System Architecture

### Smart Batch Processing
- **Token-aware sizing**: Calculates optimal batch sizes based on model limits
- **Dynamic adjustment**: Reduces batch size automatically on high retry rates
- **Content analysis**: Adjusts for complex content (tables, technical text)

### Style Consistency Engine
- **Multi-stage processing**: Pre-translation normalization → Translation → Post-processing
- **Authority corrections**: Deterministic style fixes based on diagnostics
- **Tone preservation**: Maintains consistent voice across the document

### Error Resilience
- **Progressive backoff**: 1s, 2s, 3s delays on retries
- **Graceful degradation**: Falls back to smaller batches on failures
- **Cache recovery**: Preserves work through interruptions

## 📈 Performance Optimization

### Batch Size Guidelines
- **gpt-4o models**: 8-14 items (10k token target)
- **gpt-4o-mini**: 12-18 items (8k token target)
- **Complex content**: Use lower end of ranges
- **Simple text**: Can use higher batch sizes

### Cost Management
- **Cache efficiency**: ~90% cache hit rate on re-runs
- **Model selection**: gpt-4o-mini offers 10x cost savings
- **Batch optimization**: Reduces API call overhead

## 🚨 Troubleshooting

### Common Issues

**High retry rates (>5%)**
- System automatically reduces batch size
- Check API key limits and quotas
- Consider using gpt-4o-mini for better stability

**Text overflow in slides**
- Enable PowerPoint's "Shrink text on overflow"
- Use style presets for more concise translations
- Adjust font sizes manually if needed

**Cache corruption**
- Delete `translation_cache.json` to reset
- Use `--cache new_cache.json` for fresh cache

### Debug Mode
```bash
# Enable verbose logging
export PYTHONPATH=scripts
python -u scripts/translate_pptx_inplace.py --in input.pptx --out output.pptx 2>&1 | tee debug.log
```

## 🔮 Future Enhancements

- **OCR integration**: Translate text in images
- **Multi-language support**: Beyond JA→EN
- **Real-time collaboration**: Shared translation sessions  
- **Template management**: Reusable style configurations
- **Quality scoring**: Automatic translation assessment

## 📄 License

MIT License - see LICENSE file for details.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests and documentation
5. Submit a pull request

---

*Built with ❤️ for efficient, high-quality presentation translation.*