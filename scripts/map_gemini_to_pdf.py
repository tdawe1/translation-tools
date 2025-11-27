#!/usr/bin/env python3
"""
Helper to auto-fill a PDF translations JSON using a Gemini-translated DOCX.

Workflow (for this specific Forever PDF):
  1) Make sure you have:
       - PDF bilingual CSV from a Codex run
         e.g. outputs/6920..._en_codex_bilingual.csv
       - PDF extraction JSON with unique JP segments
         e.g. artifacts/692_translation.json
       - Gemini English DOCX
         e.g. "Gemini Export 26 November 2025 at 02_36_23 GMT.docx"
  2) Run:
       python scripts/map_gemini_to_pdf.py \\
         --template translations/692_from_gemini_translations.json \\
         --bilingual outputs/692001478ad3e_26c03878c6a011f08333d63d1c8134ec_en_codex_bilingual.csv \\
         --gemini "Gemini Export 26 November 2025 at 02_36_23 GMT.docx" \\
         --output translations/692_from_gemini_auto.json
  3) Review / tweak the auto-mapped JSON, then:
       pdf_env/bin/python scripts/apply_pdf_translation.py \\
         --input inputs/692001478ad3e_26c03878c6a011f08333d63d1c8134ec.pdf \\
         --output outputs/692001478ad3e_26c03878c6a011f08333d63d1c8134ec_en_gemini_auto.pdf \\
         --translations translations/692_from_gemini_auto.json \\
         --verbose --debug-matching

This script does NOT call any external APIs. It:
  - Uses the Codex bilingual CSV as JP -> rough-English anchor.
  - Splits the Gemini DOCX into per-page sentences/paragraphs.
  - For each JP segment, finds the closest Gemini sentence on the same page
    by string similarity against the Codex English.
  - Fills translated="" in the template JSON with Gemini text where confident,
    otherwise falls back to the Codex English.
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Ensure we import the external 'python-docx' library, not local 'scripts/docx'
import sys
import os

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR in sys.path:
    try:
        sys.path.remove(_SCRIPT_DIR)
    except ValueError:
        pass

from docx import Document


logger = logging.getLogger(__name__)


@dataclass
class GeminiSegment:
    """A candidate English segment from the Gemini DOCX."""

    page: Optional[int]
    text: str
    norm: str


def normalize_english(text: str) -> str:
    """Normalize English for fuzzy matching."""
    if not text:
        return ""
    # Drop bracketed indices like [12]
    text = re.sub(r"\[\d+\]", " ", text)
    # Lowercase
    text = text.lower()
    # Keep letters, digits and # (hashtags); collapse everything else to space
    text = re.sub(r"[^a-z0-9#]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def similarity_score(a: str, b: str) -> float:
    """Compute similarity between two normalized strings."""
    if not a or not b:
        return 0.0

    seq = SequenceMatcher(None, a, b)
    base = seq.ratio()

    tokens_a = set(a.split())
    tokens_b = set(b.split())
    if tokens_a and tokens_b:
        overlap = len(tokens_a & tokens_b)
        union = len(tokens_a | tokens_b)
        jaccard = overlap / union if union else 0.0
        return 0.7 * base + 0.3 * jaccard
    return base


def load_gemini_segments(docx_path: Path) -> List[GeminiSegment]:
    """Extract per-page Gemini English segments from DOCX."""
    logger.info("Loading Gemini DOCX: %s", docx_path)
    document = Document(docx_path)
    segments: List[GeminiSegment] = []

    current_page: Optional[int] = None

    for para in document.paragraphs:
        raw = para.text.strip()
        if not raw:
            continue

        # Page marker: "Page 1", "Page 2", ...
        m = re.match(r"^Page\s+(\d+)\s*$", raw)
        if m:
            try:
                current_page = int(m.group(1))
            except ValueError:
                current_page = None
            continue

        # Clean out inline indices like [1], [23]
        cleaned = re.sub(r"\[\d+\]", " ", raw)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        if not cleaned:
            continue

        # Full paragraph candidate
        norm_para = normalize_english(cleaned)
        if norm_para:
            segments.append(GeminiSegment(current_page, cleaned, norm_para))

        # Also add sentence-level candidates for finer matching
        for sent in re.split(r"(?<=[.!?])\s+", cleaned):
            sent = sent.strip()
            if len(sent) < 20:  # skip very short fragments
                continue
            norm_sent = normalize_english(sent)
            if norm_sent:
                segments.append(GeminiSegment(current_page, sent, norm_sent))

    logger.info("Collected %d Gemini segments", len(segments))
    return segments


def load_bilingual_jp_to_en(csv_path: Path) -> Tuple[Dict[str, str], Dict[str, int]]:
    """
    Load JP -> Codex English mapping and approximate page hints from bilingual CSV.

    Returns:
        jp_to_en:  japanese string -> english translation
        jp_to_page: japanese string -> first page where it appears
    """
    logger.info("Loading bilingual CSV: %s", csv_path)
    jp_to_en: Dict[str, str] = {}
    jp_to_page: Dict[str, int] = {}

    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            jp = row.get("japanese", "") or ""
            en = row.get("english", "") or ""
            if not jp:
                continue
            if jp not in jp_to_en and en:
                jp_to_en[jp] = en
            try:
                page_str = row.get("page")
                if page_str and jp not in jp_to_page:
                    jp_to_page[jp] = int(page_str)
            except (TypeError, ValueError):
                continue

    logger.info("Loaded %d JP->EN mappings from bilingual CSV", len(jp_to_en))
    return jp_to_en, jp_to_page


def auto_fill_translations(
    template_path: Path,
    bilingual_csv: Path,
    gemini_docx: Path,
    output_path: Path,
    min_sim_threshold: float = 0.55,
) -> None:
    """Main mapping routine."""
    jp_to_en, jp_to_page = load_bilingual_jp_to_en(bilingual_csv)
    segments = load_gemini_segments(gemini_docx)

    if not segments:
        raise SystemExit("No Gemini segments found; cannot proceed.")

    data = json.loads(template_path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise SystemExit(f"Template {template_path} must be a JSON list of objects.")

    total = 0
    filled_from_gemini = 0
    filled_from_codex = 0
    left_empty = 0

    for entry in data:
        if not isinstance(entry, dict):
            continue
        jp = entry.get("original", "")
        if not jp:
            continue
        total += 1

        # If already filled, respect it
        existing = entry.get("translated", "")
        if existing:
            continue

        codex_en = jp_to_en.get(jp, "")
        if not codex_en:
            # No anchor translation at all; leave for manual review
            left_empty += 1
            continue

        norm_src = normalize_english(codex_en)
        if not norm_src:
            # Degenerate translation; fall back directly
            entry["translated"] = codex_en
            filled_from_codex += 1
            continue

        # Restrict Gemini candidates by page when possible
        page_hint = jp_to_page.get(jp)
        candidate_pool = [
            seg for seg in segments
            if page_hint is None or seg.page is None or seg.page == page_hint
        ]
        if not candidate_pool:
            candidate_pool = segments

        best_seg: Optional[GeminiSegment] = None
        best_score = 0.0

        for seg in candidate_pool:
            score = similarity_score(norm_src, seg.norm)
            if score > best_score:
                best_score = score
                best_seg = seg

        if best_seg and best_score >= min_sim_threshold:
            entry["translated"] = best_seg.text
            filled_from_gemini += 1
            logger.debug(
                "JP '%s' -> Gemini '%s' (page %s, score=%.3f)",
                jp,
                best_seg.text,
                best_seg.page,
                best_score,
            )
        else:
            # Fallback: keep Codex English so the mapping is at least complete
            entry["translated"] = codex_en
            filled_from_codex += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    logger.info("Wrote auto-filled translations to %s", output_path)
    logger.info("Total JP entries: %s", total)
    logger.info("Filled from Gemini: %s", filled_from_gemini)
    logger.info("Fallback to Codex: %s", filled_from_codex)
    logger.info("Left empty (no anchor): %s", left_empty)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Auto-map Gemini DOCX English to PDF JP segments via Codex bilingual CSV.",
    )
    parser.add_argument(
        "--template",
        required=True,
        type=Path,
        help="Template translations JSON (with original/translated fields).",
    )
    parser.add_argument(
        "--bilingual",
        required=True,
        type=Path,
        help="Bilingual CSV for the PDF (page,japanese,english,...).",
    )
    parser.add_argument(
        "--gemini",
        required=True,
        type=Path,
        help="Gemini-exported DOCX with English translation.",
    )
    parser.add_argument(
        "--output",
        required=True,
        type=Path,
        help="Output JSON path for auto-filled translations.",
    )
    parser.add_argument(
        "--min-sim",
        type=float,
        default=0.55,
        help="Minimum similarity threshold to accept Gemini mapping (default: 0.55).",
    )
    return parser.parse_args()


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    args = parse_args()
    auto_fill_translations(
        template_path=args.template,
        bilingual_csv=args.bilingual,
        gemini_docx=args.gemini,
        output_path=args.output,
        min_sim_threshold=args.min_sim,
    )


if __name__ == "__main__":
    main()
