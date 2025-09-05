import re

JP_TO_EN = str.maketrans({
    "\u3001": ", ",         # IDEOGRAPHIC COMMA (U+3001)
    "\u3002": ". ",         # IDEOGRAPHIC FULL STOP (U+3002)
    "\u300C": '"',          # LEFT CORNER BRACKET (U+300C)
    "\u300D": '"',          # RIGHT CORNER BRACKET (U+300D)
    "\uFF0F": "/",          # FULLWIDTH SOLIDUS (U+FF0F)
    "\u3000": " ",          # FULLWIDTH SPACE (U+3000)
    "\u30FB": "\u2022",     # KATAKANA MIDDLE DOT (U+30FB) → BULLET (U+2022)
    "\uFF1A": ":",          # FULLWIDTH COLON (U+FF1A)
    "\uFF1B": ";",          # FULLWIDTH SEMICOLON (U+FF1B)
})

def normalize_punct(s: str) -> str:
    s = (s or "").translate(JP_TO_EN)
    s = s.replace(" ,", ",").replace(" .", ".").replace(" %", "%").replace(" ;", ";").replace(" :", ":")
    # Hyphen digit ranges → en dash
    s = re.sub(r"(\d)\s*-\s*(\d)", r"\1\u2013\2", s)  # 5-10 → 5–10
    # JP tildes → en dash (with optional whitespace)
    s = re.sub(r"\s*[\u301C\uFF5E]\s*", "\u2013", s)   # 〜 / ～ → –
    # Normalise spaces around en dash
    s = re.sub(r"\s*\u2013\s*", "\u2013", s)          # 5 – 10 → 5–10
    s = re.sub(r"\s+", " ", s).strip()
    return s

def bullet_fragment(s: str) -> str:
    """Convert to bullet fragment: remove terminal punctuation, keep capitalization."""
    s = (s or "").strip()
    
    # Remove terminal punctuation (period, semicolon, colon) but preserve if multiple sentences
    sentence_count = len(re.findall(r'[.!?]+', s))
    if sentence_count <= 1:
        s = re.sub(r"[.;:\u3002\uFF1B\uFF1A]\s*$", "", s)  # also strip 。；：
    
    return s