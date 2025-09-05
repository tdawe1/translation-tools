#!/usr/bin/env python3
"""
Deterministic style normalizer for consistent PPTX translation formatting.
Applies predictable rules without model calls for speed and reliability.
"""
import re

# Small words that stay lowercase in Title Case (except first/last position)
SMALL_WORDS = {"a", "an", "and", "as", "at", "but", "by", "for", "in", "nor", "of", "on", "or", "per", "the", "to", "via", "vs"}

# Japanese punctuation to English mapping
JP_PUNCT_MAP = str.maketrans({
    "、": ", ",    # JP comma -> EN comma + space
    "。": ". ",    # JP period -> EN period + space  
    "「": '"',     # JP open quote -> straight quote
    "」": '"',     # JP close quote -> straight quote
    "／": "/",     # JP slash -> EN slash
    "・": "•",     # JP bullet -> EN bullet
    "　": " "      # JP full-width space -> EN space
})

# Banned phrases for tone consistency
BANNED_PHRASES = {
    "utilize": "use",
    "utilise": "use", 
    "cutting-edge": "advanced",
    "world-class": "high-quality",
    "leverage": "use",
    "synergy": "collaboration",
    "robust": "reliable",
    "cutting edge": "advanced"
}

def collapse_whitespace(s: str) -> str:
    """Normalize all whitespace to single spaces and trim."""
    return re.sub(r"\s+", " ", s).strip()

def normalize_punctuation(s: str) -> str:
    """Convert Japanese punctuation and fix common spacing issues."""
    # Apply Japanese punctuation mapping
    s = s.translate(JP_PUNCT_MAP)
    
    # Fix spacing around punctuation
    s = s.replace(" ,", ",").replace(" .", ".").replace(" %", "%")
    s = s.replace(" )", ")").replace("( ", "(")
    
    # Fix en dash spacing (ranges like 5–10%)
    s = re.sub(r"\s*–\s*", "–", s)  # Remove spaces around existing en dashes
    s = re.sub(r"(\d)\s*-\s*(\d)", r"\1–\2", s)  # Convert hyphen ranges to en dash
    
    # Fix quote spacing
    s = re.sub(r'"\s+', '"', s)  # Remove space after opening quote
    s = re.sub(r'\s+"', '"', s)  # Remove space before closing quote
    
    return collapse_whitespace(s)

def title_case(s: str) -> str:
    """Apply Title Case following standard rules for marketing materials."""
    words = s.split()
    if not words:
        return s
    
    result = []
    for i, word in enumerate(words):
        # Check if word contains formatting tags
        if '[' in word and ']' in word:
            result.append(word)  # Skip formatting tags
            continue
            
        lower_word = word.lower()
        is_first_or_last = (i == 0 or i == len(words) - 1)
        
        if is_first_or_last or lower_word not in SMALL_WORDS:
            # Capitalize first letter, preserve rest
            if word:
                result.append(word[0].upper() + word[1:])
            else:
                result.append(word)
        else:
            result.append(lower_word)
    
    return " ".join(result)

def bullet_fragment(s: str) -> str:
    """Convert to bullet fragment: remove terminal punctuation, keep capitalization."""
    s = s.strip()
    
    # Remove terminal punctuation (period, semicolon, colon) but preserve if multiple sentences
    sentence_count = len(re.findall(r'[.!?]+', s))
    if sentence_count <= 1:
        s = re.sub(r"[.;:]\s*$", "", s)
    
    return s

def replace_banned_phrases(s: str) -> str:
    """Replace banned phrases with preferred alternatives."""
    result = s
    for banned, replacement in BANNED_PHRASES.items():
        # Case-insensitive replacement preserving original case pattern
        pattern = re.compile(re.escape(banned), re.IGNORECASE)
        
        def replace_match(match):
            original = match.group(0)
            if original.isupper():
                return replacement.upper()
            elif original.istitle():
                return replacement.capitalize()
            else:
                return replacement
        
        result = pattern.sub(replace_match, result)
    
    return result

def detect_content_type(tagged_text: str, context_hints: dict = None) -> str:
    """Detect content type from text and context hints."""
    if context_hints:
        if context_hints.get("is_title", False):
            return "title"
        if context_hints.get("is_bullet", False):
            return "bullet"
    
    # Heuristic detection
    text = re.sub(r'\[/?[^\]]+\]', '', tagged_text)  # Remove tags for analysis
    
    # Short text likely title
    if len(text.split()) <= 12 and not text.strip().endswith(('.', ':', ';')):
        return "title"
    
    # Bullet indicators
    if any(indicator in tagged_text.lower() for indicator in ['[li-', '•', '◦', '▪']):
        return "bullet"
    
    # Default to bullet for body content
    return "bullet"

def normalize_block(tagged_en: str, content_type: str = None, context_hints: dict = None) -> str:
    """
    Main normalization function - applies deterministic style rules.
    
    Args:
        tagged_en: English text with formatting tags preserved
        content_type: "title", "bullet", or None for auto-detection  
        context_hints: Additional context like {"is_title": bool, "is_bullet": bool}
    
    Returns:
        Normalized text with consistent style and formatting
    """
    if not tagged_en.strip():
        return tagged_en
    
    # Detect content type if not provided
    if content_type is None:
        content_type = detect_content_type(tagged_en, context_hints)
    
    # Start with original text
    result = tagged_en
    
    # Always normalize punctuation and banned phrases
    result = normalize_punctuation(result)
    result = replace_banned_phrases(result)
    
    # Apply type-specific formatting
    if content_type == "title":
        # Apply Title Case while preserving formatting tags
        parts = re.split(r'(\[/?[^\]]+\])', result)
        normalized_parts = []
        
        for part in parts:
            if part.startswith('[') and part.endswith(']'):
                # Keep formatting tags intact
                normalized_parts.append(part)
            else:
                # Apply title case to text content
                normalized_parts.append(title_case(part))
        
        result = ''.join(normalized_parts)
    
    elif content_type == "bullet":
        result = bullet_fragment(result)
    
    # Final whitespace cleanup
    result = collapse_whitespace(result)
    
    return result

def get_style_guide() -> str:
    """Return the one-page style guide for embedding in prompts."""
    return """
**STYLE GUIDE**

**Tone & register:**
• Infer and mirror the tone from the Japanese (formality, directness, persuasion level, technicality).
• If the tone is ambiguous, default to neutral–professional.
• Do **not** add hype or soften/strengthen claims. Don’t “improve” style beyond what’s needed for natural US English.
• Preserve tags/placeholders exactly; keep bullet structure.

**Structure**
• Titles: Title Case, ≤ 10–12 words
• Bullets: fragments, not sentences. Capitalize first word. No full stops unless multiple sentences
• Parallelism: sibling bullets share the same grammatical form

**Punctuation & typography**  
• ASCII punctuation. Convert JP punctuation (、。・「」／) to EN equivalents (, . • "" /)
• Ranges use en dash: 5–10%; never hyphen for ranges
• Quotes: straight " " 
• Numbers/units: numerals everywhere; keep original numerals/percents

**Terminology**
• Enforce glossary; consistent choices (導入 → implementation, not "introduction")
• Keep product/brand names in source casing

**Overflow policy**
• If won't fit: condense by ~15% → (detail → Notes) → shrink (min 18pt title / 11pt body)
""".strip()

# Integration helpers for main translation pipeline
def normalize_translation_batch(translations: list, content_types: list = None) -> list:
    """Normalize a batch of translations with optional content type hints."""
    if content_types is None:
        content_types = [None] * len(translations)
    
    return [
        normalize_block(translation, content_type)
        for translation, content_type in zip(translations, content_types)
    ]

def apply_style_guide_to_prompt(base_prompt: str) -> str:
    """Add style guide to existing prompt."""
    style_guide = get_style_guide()
    return f"{base_prompt}\n\n{style_guide}"