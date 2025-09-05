# STYLE_GUIDE (Gengo-aligned, JA→EN)

**Goal:** natural, faithful English that mirrors the source tone. Enforce mechanics per Gengo’s guidelines; do not impose a house voice.

## Tone & Register
- Infer and mirror tone from the Japanese (formality, directness, technicality).
- If ambiguous, default to **neutral–professional**.
- Do **not** add hype or soften/strengthen claims. Avoid word-by-word literalism; translate the **whole block** naturally.

## Punctuation & Mechanics
- **Quotation marks:** use **double** quotes for speech; **single** quotes only inside double. Place periods and commas **inside** closing quotes.
- **Lists/commas:** use commas to disambiguate lists (“a, b, and c”). Serial (Oxford) comma is preferred to avoid ambiguity.
- **Dates:** Month Day, Year → e.g., **August 1st, 2013**.
- **Numbers:** use **thousands separators** (e.g., 3,000). Preserve numerals/percents exactly.
- **Ranges:** convert Japanese tilde `〜` ranges to **en dash**: `5–10%`, `April–June`.
- **JP punctuation → EN:**  、 → “,”  /  。 → “.”  /  「」 → quotes  /  ／ → “/”  /  ・ → “•” (use sparingly).
- Use ASCII punctuation in the final text; collapse duplicate spaces.

## Slide Structure
- **Dates:** Month Day, Year (e.g., **August 1, 2013**).
- **Numbers:** use **thousands separators** for numbers (e.g., 3,000). Preserve numerals/percents exactly.

### Structure

- **Titles:** concise headline (≤ 10–12 words). Prefer Title Case for slide/section titles; if a client template mandates otherwise, mirror the template.
- **Bullets:** fragments, not full sentences. Capitalize the first word. No full stops unless a bullet contains multiple sentences.
- **Parallelism:** ensure sibling bullets share the same grammatical form (e.g., all start with a verb, or all are noun phrases).

## Terminology & Numbers
- Enforce the **glossary** exactly (case-sensitive). Do not translate product/brand names; keep source casing.
- Keep **numbers/URLs/model codes** exactly (no normalization beyond thousands separators and punctuation spacing).

## Overflow Policy (mechanical; quality first)
1) **Condense** by ~15% while preserving meaning/figures/URLs.
2) If still long, add a short stub `(detail → Notes)` and put overflow in **Notes**.
3) Enable **shrink-to-fit**; clamp minimum sizes (≈ **18 pt title / 11 pt body**). Tighten spacing/indents instead of large font drops.

## Line Breaks & Formatting
- Preserve intended line breaks (`\n` → `<a:br/>`).
- Normalize spacing: collapse runs of whitespace; fix stray spaces before punctuation.

---

### Reviewer diagnostics (for JSON checks; no rewrites)
Return flags only (matches scripts/style_checker.py):
```json
{
  "style": {
    "title_case_violations": [],
    "bullet_terminal_punctuation": [],
    "parallelism_issues": [],
    "glossary_violations": [],
    "banned_phrases": [],
    "punctuation_errors": []
  },
  "tone_flags": {
    "added_hype": [],
    "softened_claims": [],
    "over_formalized": false,
    "over_casual": false,
    "deviation_from_deck_profile": []
  }
}
```
Authority applies only deterministic mechanics fixes; tone flags are surfaced for human review.
