"""
run_lm_batch.py

Batch translate JP strings via an OpenAI-compatible LM Studio endpoint.
Defaults:
  LM_BASE_URL=http://localhost:1234/v1
  LM_MODEL=gemma-2-2bn-jpn
  LM_BATCH_SIZE=15
  LM_TIMEOUT=120
  LM_START=0
  LM_END=9999

Input keys come from translation_cache_codex_cheap.json.
Outputs:
  translation_cache_lm_jit.json
  translations_full_lm_jit.json
"""

import ast
import json
import os
import sys
import time
from textwrap import shorten
from pathlib import Path

import requests

BASE = os.getenv("LM_BASE_URL", "http://localhost:1234/v1")
# Default to a known-good LM Studio model ID; override with LM_MODEL env.
MODEL = os.getenv("LM_MODEL", "gemma-2-2b-jpn-it-translate")
BATCH = int(os.getenv("LM_BATCH_SIZE", "15"))
TIMEOUT = int(os.getenv("LM_TIMEOUT", "120"))
START = int(os.getenv("LM_START", "0"))
END = int(os.getenv("LM_END", "9999"))
API_KEY = os.getenv("OPENAI_API_KEY", "")

# Resolve paths relative to repo root so the script works from any CWD.
ROOT = Path(__file__).resolve().parent.parent
SRC_PATH = ROOT / "translation_cache_codex_cheap.json"
OUT_CACHE = ROOT / "translation_cache_lm_jit.json"
OUT_FULL = ROOT / "translations_full_lm_jit.json"
STYLE_GUIDE_PATH = Path(os.getenv("LM_STYLE_GUIDE_FILE", ROOT / "STYLE_GUIDE.md"))
GLOSSARY_PATH = Path(os.getenv("LM_GLOSSARY_FILE", ROOT / "glossary.json"))

if not SRC_PATH.exists():
    raise SystemExit(f"Source cache not found: {SRC_PATH}")

src = json.load(SRC_PATH.open(encoding="utf-8"))
keys = list(src.keys())[START:END]
out: dict[str, str] = {}
s = requests.Session()
HEADERS = {}
if API_KEY:
    HEADERS["Authorization"] = f"Bearer {API_KEY}"

# Load a short style guide snippet to keep prompts compact
STYLE_SNIPPET = ""
try:
    if STYLE_GUIDE_PATH.exists():
        text = STYLE_GUIDE_PATH.read_text(encoding="utf-8")
        # keep it tight to avoid 400s
        STYLE_SNIPPET = shorten(text.replace("\n", " "), width=1200, placeholder="...")
except Exception:
    STYLE_SNIPPET = ""

# Load glossary terms (key -> value) to encourage consistent choices
GLOSSARY_SNIPPET = ""
try:
    if GLOSSARY_PATH.exists():
        gdata = json.load(GLOSSARY_PATH.open(encoding="utf-8"))
        if isinstance(gdata, dict):
            # take a compact subset
            subset = dict(list(gdata.items())[:80])
            GLOSSARY_SNIPPET = json.dumps(subset, ensure_ascii=False)
except Exception:
    GLOSSARY_SNIPPET = ""


def parse_array(text: str):
    text = text.strip()
    if text.startswith("```"):
        text = "\n".join(ln for ln in text.split("\n") if not ln.strip().startswith("```"))
    for fn in (json.loads, ast.literal_eval):
        try:
            arr = fn(text)
            if isinstance(arr, list):
                return arr
        except Exception:
            pass
    return None


# warmup
try:
    s.post(
        f"{BASE}/chat/completions",
        json={"model": MODEL, "messages": [{"role": "user", "content": "ping"}]},
        headers=HEADERS,
        timeout=30,
    )
except Exception:
    pass

for i in range(0, len(keys), BATCH):
    batch = keys[i : i + BATCH]
    arr = None
    last = None
    system_parts = [
        "Translate Japanese to natural US English.",
        "Preserve numbering, bullets, punctuation; do not drop content.",
        "Keep placeholders like OK/NG as-is; no new lines unless present.",
        "Return ONLY a JSON array of translations in the same order; no prose.",
    ]
    if STYLE_SNIPPET:
        system_parts.append(f"Style guidance: {STYLE_SNIPPET}")
    if GLOSSARY_SNIPPET:
        system_parts.append(f"Glossary (preferred terms): {GLOSSARY_SNIPPET}")
    system_msg = " ".join(system_parts)

    messages = [
        {
            "role": "system",
            "content": system_msg,
        },
        {"role": "user", "content": json.dumps({"strings": batch}, ensure_ascii=False)},
    ]
    body = {"model": MODEL, "messages": messages, "temperature": 0.2}

    for attempt in range(4):
        try:
            r = s.post(f"{BASE}/chat/completions", json=body, headers=HEADERS, timeout=TIMEOUT)
            r.raise_for_status()
            content = r.json()["choices"][0]["message"]["content"]
            arr = parse_array(content)
            if arr and len(arr) == len(batch):
                break
        except Exception as e:
            last = e
            time.sleep(1 + attempt)

    if not (arr and len(arr) == len(batch)):
        for jp in batch:
            out[jp] = src[jp]
        print(f"batch {i}-{i+len(batch)} fallback {last}", file=sys.stderr)
        continue

    for jp, en in zip(batch, arr):
        out[jp] = str(en).strip()

    if (i // BATCH + 1) % 5 == 0:
        print("progress", i + len(batch), "of", len(keys))

OUT_CACHE.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
wrapped = {k: {"translated": v, "font_scaling": 1.0} for k, v in out.items()}
OUT_FULL.write_text(json.dumps(wrapped, ensure_ascii=False, indent=2), encoding="utf-8")
print("done", len(out), "written to", OUT_CACHE)
