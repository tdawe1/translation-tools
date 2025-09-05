import re

JP_TO_EN = str.maketrans({
    "\u3001": ", ",   # 、 comma
    "\u3002": ". ",   # 。 period
    "\u300C": '"',    # 「 open quote
    "\u300D": '"',    # 」 close quote
    "\uFF0F": "/",    # ／ full-width solidus
    "\u3000": " ",    # 　full-width space
    "\u30FB": "•",    # ・ middle dot → bullet (use sparingly)
    "\uFF1A": ":",    # ： full-width colon
    "\uFF1B": ";",    # ； full-width semicolon
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