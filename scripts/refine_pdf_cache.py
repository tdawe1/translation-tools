#!/usr/bin/env python3
"""
refine_pdf_cache.py

Offline refinement tool for PDF translation caches.

Use this to clean up and upgrade specific JA→EN entries in a cache
for a given PDF, without making any new API calls.

Typical usage (for the Forever Digital Marketing Guide PDF):

  python scripts/refine_pdf_cache.py \
    --cache translation_cache_codex_max.json \
    --artifacts artifacts/692_translation.json \
    --out translation_cache_codex_max_refined.json

Then regenerate the PDF in cache-only mode:

  python scripts/translate_pdf.py \
    --in inputs/692001478ad3e_26c03878c6a011f08333d63d1c8134ec.pdf \
    --out outputs/692001478ad3e_26c03878c6a011f08333d63d1c8134ec_en_codex_max_refined.pdf \
    --cache translation_cache_codex_max_refined.json \
    --cache-only

The script:
  - Limits changes to Japanese segments actually present in the target PDF
    (via its artifacts JSON).
  - Applies safe mechanical fixes (quotes/punctuation).
  - Applies a small set of manual overrides for known awkward headings
    in the Forever Digital Marketing Guide.

You can extend MANUAL_OVERRIDES as you review the bilingual CSV.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Dict, Any, Set

try:
    # Reuse existing punctuation normalizer if available
    from style_mechanics_normalize import normalize_punct  # type: ignore
except ImportError:  # pragma: no cover - defensive fallback
    def normalize_punct(s: str) -> str:
        return s or ""


# Hand-tuned overrides for the Forever Digital Marketing Guide PDF.
# Keys are Japanese source strings, values are refined English.
MANUAL_OVERRIDES: Dict[str, str] = {
    # Front matter / TOC
    "目次": "Table of Contents",
    "フォーエバー デジタルマーケティングガイド": "Forever Digital Marketing Guide",
    "デジタル時代のソーシャルメディア活用術": "How to Use Social Media in the Digital Age",
    "収入・ライフスタイル表現と": "Income and Lifestyle Claims",
    "コンプライアンス": "Compliance",

    # Section headings
    "表現・投稿における注意点": "Points to Note for Expressions and Posts",
    "なぜ表現ルールが重要なのか？": "Why Are Wording Rules Important?",
    "製品紹介と表現のルール": "Rules for Product Introduction and Wording",
    "ウェブサイトのチェックポイント": "Website Checkpoints",
    "最後に": "In Closing",
    "既存のルールを守りましょう": "Let's Follow Existing Rules",

    # Pre‑posting checklist / questions
    "投稿の前にチェックしたい": "What to Check Before You Post",
    "つの質問": "Questions",
    "投稿前チェックリスト（": "Pre-posting Checklist (",
    "つの質問）": "Questions)",

    # “OK/NG” visual criteria heading
    "投稿ビジュアルの「": "Post visuals: “OK",
    "／": "/",
    "」判断基準": "NG” criteria",

    # Solicitation legality line
    "での勧誘は「正しく行えば合法」です": "Solicitation on SNS is legal if done correctly.",
    "SNSでの勧誘は「正しく行えば合法」です": "Solicitation on SNS is legal if done correctly.",

    # TOC fragment that caused "eachSNSFeatures..."
    "各": "",
    "の特徴と活用ポイント": "Features and Usage Points",
}


def apply_mechanical_fixes(text: str) -> str:
    """Apply safe, deterministic cleanup to an English string."""
    s = text or ""
    # Normalize odd quote patterns from some models
    s = s.replace("``", '"').replace("''", '"')
    # Normalize JP-style punctuation or spacing that slipped through
    s = normalize_punct(s)
    return s


def load_targets_from_artifacts(artifacts: Dict[str, Any]) -> Set[str]:
    """Collect the set of Japanese strings that appear in the target PDF."""
    targets: Set[str] = set()

    unique = artifacts.get("unique_texts")
    if isinstance(unique, list):
        targets.update(str(x) for x in unique)

    jtexts = artifacts.get("japanese_texts")
    if isinstance(jtexts, list):
        targets.update(str(x) for x in jtexts)

    mapping = artifacts.get("text_mapping")
    if isinstance(mapping, dict):
        targets.update(mapping.keys())

    return targets


def refine_cache(cache_path: Path, artifacts_path: Path, out_path: Path) -> None:
    cache_data = json.loads(cache_path.read_text(encoding="utf-8"))
    if not isinstance(cache_data, dict):
        raise ValueError(f"{cache_path} does not contain a JSON object")

    artifacts = json.loads(artifacts_path.read_text(encoding="utf-8"))
    targets = load_targets_from_artifacts(artifacts)

    updated = 0
    missing = 0

    for jp in sorted(targets):
        if jp not in cache_data:
            missing += 1
            continue

        original_en = cache_data[jp]
        new_en = original_en

        if jp in MANUAL_OVERRIDES:
            new_en = MANUAL_OVERRIDES[jp]
        else:
            new_en = apply_mechanical_fixes(original_en)

        if new_en != original_en:
            cache_data[jp] = new_en
            updated += 1

    out_path.write_text(json.dumps(cache_data, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Refined cache written to: {out_path}")
    print(f"  Updated entries: {updated}")
    print(f"  Missing (present in artifacts, absent in cache): {missing}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Refine JA→EN cache entries for a specific PDF (offline, no API calls)."
    )
    parser.add_argument(
        "--cache",
        required=True,
        help="Input translation cache JSON (e.g., translation_cache_codex_max.json)",
    )
    parser.add_argument(
        "--artifacts",
        required=True,
        help="Artifacts JSON for the target PDF (e.g., artifacts/692_translation.json)",
    )
    parser.add_argument(
        "--out",
        required=True,
        help="Output path for refined cache JSON",
    )

    args = parser.parse_args()

    cache_path = Path(args.cache)
    artifacts_path = Path(args.artifacts)
    out_path = Path(args.out)

    refine_cache(cache_path, artifacts_path, out_path)


if __name__ == "__main__":
    main()

