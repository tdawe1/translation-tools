#!/usr/bin/env python3
"""Translation batch processor for LM Studio."""
import json
import time
from pathlib import Path
import requests
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE = "http://localhost:1234/v1"
MODEL = "bartowski/Llama-3.3-70B-Instruct-GGUF"


def translate_one(jp: str, session: requests.Session) -> str:
    """Translate a single Japanese phrase to English."""
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": "You are a professional JP→EN translator."},
            {"role": "user", "content": f"Translate this to natural US English:\n{jp}"},
        ],
        "temperature": 0.1,
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
        logger.warning("Translation attempt failed: %s", e)
        raise


def main() -> None:
    """Main batch translation loop."""
    with open("translation_cache_codex_cheap.json", "r", encoding="utf-8") as f:
        src = json.load(f)

    keys = list(src.keys())
    out: dict[str, str] = {}
    session = requests.Session()

    print("Waiting 20s for model init...")
    time.sleep(20)

    for i, jp in enumerate(keys):
        last_error = None
        for attempt in range(6):
            try:
                out[jp] = translate_one(jp, session)
                break
            except Exception as e:  # noqa: BLE001
                last_error = e
                time.sleep(1 + attempt)
        else:
            logger.warning("Failed to translate after 6 attempts: %s... Error: %s", jp[:50], last_error)
            out[jp] = src[jp]

        if i % 50 == 0:
            print(f"Progress: {i}/{len(keys)}")

    Path("translation_cache_lm_direct.json").write_text(
        json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    full_format = {k: {"translated": v, "font_scaling": 1.0} for k, v in out.items()}
    Path("translations_full_lm_direct.json").write_text(
        json.dumps(full_format, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(f"Done: {len(out)} translations")


if __name__ == "__main__":
    main()
