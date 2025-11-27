#!/usr/bin/env python3
"""
run_lm_batch.py

Batch translate JP strings via an OpenAI-compatible LM Studio endpoint.
Defaults:
  LM_BASE_URL=http://localhost:1234/v1
"""
import json
import os
import time
from pathlib import Path
from textwrap import shorten
import requests
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE = os.getenv("LM_BASE_URL", "http://localhost:1234/v1")
# Default to gemma-2-2bn-jpn for small tasks; override via LM_MODEL
MODEL = os.getenv("LM_MODEL", "gemma-2-2bn-jpn")

SRC_PATH = Path("translation_cache_codex_cheap.json")
OUT_CACHE = Path("translation_cache_lm_jit.json")
STYLE_GUIDE_PATH = Path("STYLE_GUIDE.md")
GLOSSARY_PATH = Path("glossary.json")

# Try to load style guide for context
try:
    if STYLE_GUIDE_PATH.exists():
        text = STYLE_GUIDE_PATH.read_text(encoding="utf-8")
        STYLE_SNIPPET = shorten(text.replace("\n", " "), width=1200, placeholder="...")
    else:
        STYLE_SNIPPET = ""
except Exception as e:
    logger.debug("Failed to load style guide: %s", e)
    STYLE_SNIPPET = ""

# Try to load glossary for context
try:
    if GLOSSARY_PATH.exists():
        glossary_data = json.loads(GLOSSARY_PATH.read_text(encoding="utf-8"))
        entries = [f"{k}: {v}" for k, v in list(glossary_data.items())[:50]]
        GLOSSARY_SNIPPET = shorten(", ".join(entries), width=800, placeholder="...")
    else:
        GLOSSARY_SNIPPET = ""
except Exception as e:
    logger.debug("Failed to load glossary: %s", e)
    GLOSSARY_SNIPPET = ""

if not SRC_PATH.exists():
    raise SystemExit(f"Source cache not found: {SRC_PATH}")

def translate_batch(batch_items, session):
    """Call LM Studio with a batch of strings."""
    system = (
        "You are a professional Japanese-to-English translator. "
        "Return ONLY a JSON array of translated strings in the same order. "
        "No preamble, no markdown, no keys. "
        "Preserve formatting and numbers exactly."
    )
    if STYLE_SNIPPET:
        system += f"\nStyle Guide: {STYLE_SNIPPET}"
    if GLOSSARY_SNIPPET:
        system += f"\nGlossary: {GLOSSARY_SNIPPET}"

    user_content = json.dumps(batch_items, ensure_ascii=False)

    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": f"Translate these items to English:\n{user_content}"},
        ],
        "temperature": 0.1,
    }

    try:
        r = session.post(f"{BASE}/chat/completions", json=payload, timeout=120)
        r.raise_for_status()
        res = r.json()
        content = res["choices"][0]["message"]["content"].strip()

        # Strip markdown fences if present
        if content.startswith("```"):
            content = content.split("\n", 1)[1]
            if content.endswith("```"):
                content = content.rsplit("\n", 1)[0]

        return json.loads(content)
    except Exception as e:
        logger.warning(f"Batch failed: {e}")
        return None

def main():
    src = json.load(SRC_PATH.open(encoding="utf-8"))

    # Load existing output to resume
    out = {}
    if OUT_CACHE.exists():
        try:
            out = json.load(OUT_CACHE.open(encoding="utf-8"))
            print(f"Resuming with {len(out)} cached translations")
        except Exception:
            pass

    # Identify missing keys
    all_keys = list(src.keys())
    missing_keys = [k for k in all_keys if k not in out]

    print(f"Total items: {len(all_keys)}, Missing: {len(missing_keys)}")

    session = requests.Session()
    BATCH = 10

    # Give LM Studio time to spin up if needed
    try:
        requests.get(f"{BASE}/models", timeout=2)
    except Exception:
        print("Waiting for LM Studio...")
        time.sleep(5)

    for i in range(0, len(missing_keys), BATCH):
        batch = missing_keys[i : i + BATCH]

        # Retry loop
        success = False
        for attempt in range(3):
            arr = translate_batch(batch, session)
            if arr and isinstance(arr, list) and len(arr) == len(batch):
                for jp, en in zip(batch, arr, strict=True):
                    out[jp] = str(en).strip()
                success = True
                break
            time.sleep(2 + attempt)

        if not success:
            print(f"Failed batch starting at {i}, skipping...")
            # Fallback to original text to avoid holes? Or just skip?
            # For now, skip to allow retry later

        if (i // BATCH) % 5 == 0:
            print(f"Progress: {i}/{len(missing_keys)}")
            # Checkpoint
            OUT_CACHE.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")

    # Final write
    OUT_CACHE.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")

    # Also update the full catalog format
    wrapped = {k: {"translated": v, "font_scaling": 1.0} for k, v in out.items()}
    Path("translations_full_lm_jit.json").write_text(
        json.dumps(wrapped, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("Done.")

if __name__ == "__main__":
    main()
