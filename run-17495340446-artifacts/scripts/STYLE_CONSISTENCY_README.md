# Style Consistency System

Comprehensive formatting and tone consistency for PPTX translations. Bakes **one source of truth** into the pipeline with deterministic rules and model-based checking.

## System Overview

The style consistency system applies a 6-stage workflow:

1. **Producer** → Tagged English with style-aware prompts
2. **Style Normalizer** → Deterministic punctuation, case, and format fixes  
3. **Style Checker** → Model-based diagnostics (JSON output)
4. **Authority Fixes** → Deterministic corrections based on diagnostics
5. **PPTX Formatting** → Consistent layout profile applied during write-back
6. **Style Audit** → CI-ready consistency checks with failure conditions

## Components

### 1. Style Normalization (`style_normalize.py`)

Deterministic, fast rules for consistent formatting:

- **Punctuation**: Japanese punctuation (、。・「」) → English (, . • "")  
- **Typography**: Ranges use en dash (5–10%), straight quotes, proper spacing
- **Case**: Title Case for titles, bullet fragments without terminal periods
- **Tone**: Banned phrase replacement (utilize→use, cutting-edge→advanced)
- **Content Types**: Auto-detection (title/bullet/table) with appropriate rules

**Key Functions:**
- `normalize_block(text, content_type)` - Main normalization entry point
- `title_case(text)` - Marketing-appropriate Title Case
- `bullet_fragment(text)` - Remove terminal punctuation from fragments

### 2. PPTX Formatting Profile (`pptx_format.py`)

Single formatting profile for deck-wide consistency:

- **Margins**: Tight 2pt left/right, 1pt top/bottom for maximum text space
- **Autofit**: Shrink-to-fit with scaling limits (85% min font, 15% max line reduction)
- **Line Spacing**: 110% (vs default 120%) for better density
- **Fonts**: Consistent font family (Inter default), minimum sizes (18pt titles, 11pt body)  
- **Bullets**: Optimized indentation (0.3" base, 0.2" hanging)

**Integration Points:**
- Applied during PPTX write-back after content translation
- Works directly with XML for maximum control
- Respects existing autofit while adding consistency

### 3. Style Checker (`style_checker.py`)

Model-based diagnostics with structured JSON output:

- **Title Case Violations**: Detects non-conforming title capitalization
- **Bullet Punctuation**: Flags inappropriate terminal punctuation in fragments  
- **Glossary Consistency**: Checks for terminology drift across slides
- **Banned Phrases**: Identifies tone inconsistencies (utilize, world-class, etc.)
- **Parallelism**: Detects mixed verb forms in bullet groups

**Authority Fixes:**
- Deterministic corrections based on JSON diagnostics
- Token-level replacements without sentence rewriting
- Preserves formatting tags and placeholders

### 4. Style Audit (`audit_style.py`)

CI-ready consistency checks with failure conditions:

- **Capitalization Audit**: Title Case compliance across deck
- **Punctuation Audit**: Terminal punctuation in bullets
- **Terminology Audit**: Consistent glossary usage  
- **Length Violations**: Title/bullet length limits (12 words / 200 chars)
- **Japanese Residual**: Remaining untranslated characters
- **Tag Balance**: Formatting tag integrity

**CI Integration:**
```bash
# Fail CI on critical issues
python audit_style.py bilingual.csv --glossary glossary.json --fail-on-critical

# Exit codes: 0 = pass, 1 = fail
```

## Environment Configuration

Control system behavior with environment variables:

```bash
# Core translation model (affects style checking availability)
export OPENAI_MODEL="gpt-5"                    # Enable advanced features
export OPENAI_TEMPERATURE="0.6"                # Producer temperature  

# Style system controls
export ENABLE_EXPANSION_POLICY="1"             # 3-stage expansion policy
export ENABLE_STYLE_CHECKING="1"               # Model-based style checks
export ENABLE_FORMATTING_PROFILE="1"           # Consistent PPTX formatting
export ENABLE_STYLE_AUDIT="1"                  # Post-process audit

# Style guide customization
export STYLE_PRESET="marketing"                # Built-in style presets
export STYLE_GUIDE_FILE="custom_style.txt"     # Custom style guide path
```

## Integration Workflow

### Translation Pipeline Integration

1. **Enhanced Prompts**: Style guide automatically embedded in system prompts
2. **Post-Translation Processing**: 
   - Deterministic normalization applied to all translations
   - Model-based style checking (GPT-5 only)
   - Authority fixes applied based on diagnostics
3. **PPTX Write-back**:
   - Consistent formatting profile applied
   - Layout tightening for space-constrained slides
   - Autofit safety nets preserved

### Usage Examples

```bash
# Full translation with style consistency
python translate_pptx_inplace.py \
  --in input.pptx --out output_en.pptx \
  --model gpt-5 --glossary glossary.json

# Style audit only (for CI)
python audit_style.py bilingual.csv \
  --glossary glossary.json \
  --max-issues 5 \
  --fail-on-critical

# Test style normalization
python -c "
from style_normalize import normalize_block
print(normalize_block('This is a title example', 'title'))
print(normalize_block('Bullet point ending.', 'bullet'))
"
```

## Style Guide Reference

**Voice & Tone:**
- US English, marketing slide voice
- Benefits > features; plain verbs; no hype
- Avoid: "utilize", "cutting-edge", "world-class", "leverage"

**Structure:**
- **Titles**: Title Case, ≤ 10-12 words
- **Bullets**: Fragments, not sentences. Capitalize first word. No terminal periods
- **Parallelism**: Sibling bullets share grammatical form

**Typography:**
- ASCII punctuation. Convert JP punctuation to EN equivalents
- Ranges: en dash (5–10%), never hyphen  
- Quotes: straight `" "` 
- Numbers: numerals everywhere; preserve original values

**Terminology:**
- Enforce glossary exactly (導入 → implementation, not introduction)
- Keep product/brand names in source casing

## Performance Notes

- **Deterministic Operations**: Style normalization and PPTX formatting add ~5-10ms per slide
- **Model-based Checking**: Adds 1-2 API calls per batch (only for GPT-5)  
- **Memory Usage**: Minimal additional overhead (~10MB for style modules)
- **Fallback Behavior**: System gracefully degrades when style modules unavailable

## Troubleshooting

### Common Issues

1. **Style modules not found**: Install dependencies or set `ENABLE_*` vars to "0"
2. **Model-based checking disabled**: Requires GPT-5 model and Responses API
3. **Formatting not applied**: Check `ENABLE_FORMATTING_PROFILE="1"`
4. **CI failures**: Review audit thresholds in `should_fail_ci()` function

### Debug Output

```bash
# Enable verbose style checking
STYLE_DEBUG=1 python translate_pptx_inplace.py [args]

# Check style module availability  
python -c "
try:
    from style_normalize import *
    from style_checker import *
    from pptx_format import *
    print('Style modules: OK')
except ImportError as e:
    print(f'Style modules: ERROR - {e}')
"
```

## File Structure

```
scripts/
├── translate_pptx_inplace.py    # Main translation with integrated style system
├── style_normalize.py           # Deterministic style normalization  
├── style_checker.py             # Model-based style checking
├── pptx_format.py              # PPTX formatting profile
├── audit_style.py              # Style consistency audit
└── STYLE_CONSISTENCY_README.md  # This documentation
```

## Future Enhancements

- **Content-aware Detection**: Improved title/bullet/table detection using slide layout
- **Brand Customization**: Style profiles for different brand guidelines
- **Multilingual Support**: Extend to other language pairs beyond JP→EN  
- **Advanced Parallelism**: Semantic analysis for better parallelism detection
- **Performance Optimization**: Batch style checking for large decks