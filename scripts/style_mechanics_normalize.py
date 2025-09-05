import re

JP_TO_EN = str.maketrans({"、": ", ", "。": ". ", "「": '"', "」": '"', "／": "/", "　": " "})

def normalize_punct(s: str) -> str:
    s = (s or "").translate(JP_TO_EN)
    s = s.replace(" ,", ",").replace(" .", ".").replace(" %", "%")
    s = re.sub(r"(\d)\s*-\s*(\d)", r"\1–\2", s)  # 5-10 → 5–10
    s = re.sub(r"\s*–\s*", "–", s)               # 5 – 10 → 5–10
    s = re.sub(r"\s+", " ", s).strip()
    return s

def bullet_fragment(s: str) -> str:
    """Convert to bullet fragment: remove terminal punctuation, keep capitalization."""
    s = s.strip()
    
    # Remove terminal punctuation (period, semicolon, colon) but preserve if multiple sentences
    sentence_count = len(re.findall(r'[.!?]+', s))
    if sentence_count <= 1:
        s = re.sub(r"[.;:]\s*$", "", s)
    
    return s