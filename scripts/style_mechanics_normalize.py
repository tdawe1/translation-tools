import re

JP_TO_EN = str.maketrans({"、": ", ", "。": ". ", "「": '"', "」": '"', "／": "/", "　": " "})

def normalize_punct(s: str) -> str:
    s = (s or "").translate(JP_TO_EN)
    s = s.replace(" ,", ",").replace(" .", ".").replace(" %", "%")
    s = re.sub(r"(\d)\s*-\s*(\d)", r"\1–\2", s)  # 5-10 → 5–10
    s = re.sub(r"\s+", " ", s).strip()
    return s

def bullet_fragment(s: str) -> str:
    # remove terminal . ; : for bullet fragments
    return re.sub(r"[.;:]\s*$", "", (s or "").strip())