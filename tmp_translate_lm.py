#!/usr/bin/env python3
"""
tmp_translate_lm.py

Helper script for direct LM Studio translation testing.
"""
import json
import time
import requests
from pathlib import Path
import os
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE = os.getenv("LM_BASE_URL", "http://localhost:1234/v1")
MODEL = os.getenv("LM_MODEL", "gemma-2-2bn-jpn")

def translate_one(jp: str, session: requests.Session) -> str:
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

    try:
        r = session.post(f"{BASE}/chat/completions", json=payload, timeout=60)
        r.raise_for_status()
        data = r.json()
        choices = data.get("choices") or []
        if not choices or "message" not in choices[0]:
            raise ValueError(f"Unexpected response schema: {data!r}")
        return choices[0]["message"]["content"].strip()
    except (requests.RequestException, ValueError, KeyError) as e:
        logger.warning(f"Translation attempt failed: {e}")
        raise

def main():
    src_path = Path("translation_cache_codex_cheap.json")
    if not src_path.exists():
        print("Source cache not found, skipping run.")
        return

    with open(src_path, "r", encoding="utf-8") as f:
        src = json.load(f)

    keys = list(src.keys())
    out: dict[str, str] = {}
    session = requests.Session()

    # Give LM Studio time to spin up the model
    print("Waiting 5s for model init...")
    time.sleep(5)

    for i, jp in enumerate(keys):
        last_error = None
        for attempt in range(6):
            try:
                out[jp] = translate_one(jp, session)
                break
            except Exception as e:
                last_error = e
                time.sleep(1 + attempt)
        else:
            # All retries failed
            if jp in src:
                # Use original value structure if possible, but here src values are dicts?
                # Let's assume src values are dicts {translated: ..., ...} based on other files
                # But the original script used SRC[jp] directly.
                # Let's just output a placeholder error if we can't get translation
                out[jp] = f"[FAILED: {last_error}]"
            print(f"fallback at {i}: {jp[:20]} err={last_error}")

        if (i + 1) % 50 == 0:
            print("progress", i + 1)

        # Save intermediate results
        if (i + 1) % 10 == 0:
             Path("translation_cache_lm_direct.json").write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")

    # Final save
    Path("translation_cache_lm_direct.json").write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")

    # Wrap for full catalog
    wrapped = {k: {"translated": v, "font_scaling": 1.0} for k, v in out.items()}
    Path("translations_full_lm_direct.json").write_text(json.dumps(wrapped, ensure_ascii=False, indent=2), encoding="utf-8")

    print("done", len(out))

if __name__ == "__main__":
    main()
