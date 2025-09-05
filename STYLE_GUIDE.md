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
- **Titles:** concise headline (≤ 10–12 words). Use Title Case if your template requires; otherwise mirror source capitalization.
- **Bullets:** use **fragments**, not full sentences. Capitalize first word. **No period at end** unless multiple sentences are truly needed.
- **Parallelism:** sibling bullets keep the same grammatical form.

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
Return flags only:
```json
{
  "mechanics": {
    "quotes_rule": true,
    "periods_commas_inside_quotes": true,
    "serial_comma_missed": [3, 5],
    "date_style_violations": [12],
    "thousands_separator_missed": [7],
    "range_dash_needed": [9]
  },
  "structure": {
    "bullet_terminal_punct": [2],
    "parallelism_mismatch": [{ "slide": 14, "bullets": [1,2,3] }]
  },
  "tone": {
    "over_formalized": false,
    "over_casual": false,
    "added_hype_terms": []
  }
}
```
Authority applies only deterministic mechanics fixes; tone flags are surfaced for human review.
