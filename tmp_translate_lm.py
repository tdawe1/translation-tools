import json
import time
import requests
from pathlib import Path

import os

BASE = os.getenv("LM_BASE_URL", "http://192.168.0.25:1234/v1")
# Default to gemma-2-2bn-jpn for small tasks; override via LM_MODEL
MODEL = os.getenv("LM_MODEL", "gemma-2-2bn-jpn")

SRC = json.load(open("translation_cache_codex_cheap.json", "r", encoding="utf-8"))
keys = list(SRC.keys())
out: dict[str, str] = {}
session = requests.Session()


def translate_one(jp: str) -> str:
    system = "You are a professional Japanese-to-English translator. Return only the English translation as plain text."
    user = f"Translate to natural US English. Text: {jp}"
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "temperature": 0.2,
    }
    r = session.post(f"{BASE}/chat/completions", json=payload, timeout=60)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"].strip()

# Give LM Studio time to spin up the model
time.sleep(20)

for i, jp in enumerate(keys):
    last = None
    for attempt in range(6):
        try:
            out[jp] = translate_one(jp)
            break
        except Exception as e:
            last = e
            time.sleep(1 + attempt)
    else:
        out[jp] = SRC[jp]
        print(f"fallback at {i}: {jp[:20]} err={last}")
    if (i + 1) % 50 == 0:
        print("progress", i + 1)

Path("translation_cache_lm_direct.json").write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
wrapped = {k: {"translated": v, "font_scaling": 1.0} for k, v in out.items()}
Path("translations_full_lm_direct.json").write_text(json.dumps(wrapped, ensure_ascii=False, indent=2), encoding="utf-8")
print("done", len(out))
