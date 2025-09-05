#!/usr/bin/env python3
"""
Estimate API costs to translate a PPTX deck (JA→EN) across multiple providers/models.
- Counts input tokens with tiktoken (rough approximation for non-OpenAI providers).
- Estimates output tokens from JP chars × expansion / avg_en_chars_per_token.
- Supports provider:model identifiers, e.g.:
    openai:gpt-5 | openai:gpt-5-mini | openai:gpt-4.1 (unset by default)
    anthropic:claude-opus-4.1 | anthropic:claude-sonnet-4 | anthropic:claude-haiku-3.5
    google:gemini-1.5-pro | google:gemini-1.5-flash | google:gemini-1.5-flash-8b | google:gemini-2.0-flash
    cohere:command | cohere:command-r-plus | cohere:command-light

Notes:
- OpenAI prompt-caching is modeled via cheaper 'in_cached' rate after the first call.
- Anthropic prompt-caching has separate write/read rates. Use --anthropic-cache-write to
  charge the first call's cached portion at 'cached_write' (instead of normal input).
- Gemini lists "context caching price" on their page; this script *does not* add an extra
  caching line-item for Gemini by default. You can model it by supplying custom pricing via --pricing.

Install:  pip install python-pptx tiktoken
"""

import argparse, math, zipfile, re, sys, os, json, html
from collections import defaultdict
# -------------------------
# Built-in pricing (USD per 1M tokens)
# You can override or extend with --pricing pricing.json
# -------------------------
PRICING = {
    # OpenAI — official API pricing page
    # https://openai.com/api/pricing/
    "openai:gpt-5":       {"in": 1.25, "in_cached": 0.125, "out": 10.00, "tokenizer": "o200k_base"},
    "openai:gpt-5-mini":  {"in": 0.25, "in_cached": 0.025, "out":  2.00, "tokenizer": "o200k_base"},
    "openai:gpt-5-nano":  {"in": 0.05, "in_cached": 0.005, "out":  0.40, "tokenizer": "o200k_base"},
    "openai:gpt-4o":      {"in": 2.50, "in_cached": 1.25,  "out": 10.00, "tokenizer": "o200k_base"},
    "openai:gpt-4o-mini": {"in": 0.60, "in_cached": 0.30,  "out":  2.40, "tokenizer": "o200k_base"},
    # GPT-4.1: OpenAI page currently shows FT pricing, not base per-token usage.
    # Leave unset by default; provide via --pricing if you have contracted rates.
    "openai:gpt-4.1":     {"in": None, "in_cached": None,  "out": None,  "tokenizer": "o200k_base"},

    # Anthropic — API pricing (per 1M tokens), plus prompt-caching write/read
    # https://www.anthropic.com/pricing
    "anthropic:claude-opus-4.1":  {"in": 15.00, "out": 75.00, "cached_write": 18.75, "cached_read": 1.50, "tokenizer": "o200k_base"},
    # Sonnet 4 has tiers; default to ≤200k token tier
    "anthropic:claude-sonnet-4":  {"in": 3.00,  "out": 15.00, "cached_write": 3.75,  "cached_read": 0.30, "tokenizer": "o200k_base"},
    "anthropic:claude-haiku-3.5": {"in": 0.80,  "out":  4.00, "cached_write": 1.00,  "cached_read": 0.08, "tokenizer": "o200k_base"},

    # Google — Gemini developer pricing
    # https://ai.google.dev/gemini-api/docs/pricing
    # Defaults use <=128k (or "Standard") rates where tiers exist.
    "google:gemini-1.5-pro":       {"in": 1.25,  "out": 5.00,  "tokenizer": "o200k_base"},
    "google:gemini-1.5-flash":     {"in": 0.075, "out": 0.30,  "tokenizer": "o200k_base"},
    "google:gemini-1.5-flash-8b":  {"in": 0.0375,"out": 0.15,  "tokenizer": "o200k_base"},
    "google:gemini-2.0-flash":     {"in": 0.10,  "out": 0.40,  "tokenizer": "o200k_base"},

    # Cohere — pricing
    # https://cohere.com/pricing
    "cohere:command":        {"in": 1.00, "out": 2.00,  "tokenizer": "o200k_base"},
    "cohere:command-r-plus": {"in": 2.50, "out": 10.00, "tokenizer": "o200k_base"},
    "cohere:command-light":  {"in": 0.30, "out": 0.60,  "tokenizer": "o200k_base"},
}

# (Optional) Reserved for future tokenizer alias mapping.

try:
    import tiktoken
except Exception:
    print("ERROR: tiktoken is required. Install with `pip install tiktoken`.", file=sys.stderr)
    raise

def encoding_for_model_key(model_key: str):
    cfg = PRICING.get(model_key, {})
    enc_name = cfg.get("tokenizer") or "o200k_base"
    try:
        return tiktoken.get_encoding(enc_name)
    except Exception:
        return tiktoken.get_encoding("o200k_base")

JP_RX = re.compile(r"[぀-ヿ㐀-鿿々〆ヵヶ]")
BR_RX = re.compile(r"(?:\r\n|\r|\n)")
def extract_text_blocks(pptx_path: str):
    """Return a list of visible text blocks (titles, bullets, notes) from the PPTX."""
    A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    blocks = []
    with zipfile.ZipFile(pptx_path, "r") as z:
        slide_names = sorted(
            n
            for n in z.namelist()
            if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        )
        for name in slide_names:
            xml = z.read(name).decode("utf-8", errors="ignore")
            # Preserve explicit line breaks from PPTX
            xml = re.sub(r"<a:br\s*/>", "\n", xml)
            texts = re.findall(r"<a:t>(.*?)</a:t>", xml, flags=re.S)
            if texts:
                s = BR_RX.sub("\n", "".join(html.unescape(t) for t in texts))
                parts = [p.strip() for p in re.split(r"\n{2,}", s) if p.strip()]
                blocks.extend(parts)

        note_names = sorted(
            n
            for n in z.namelist()
            if n.startswith("ppt/notesSlides/notesSlide") and n.endswith(".xml")
        )
        for name in note_names:
            xml = z.read(name).decode("utf-8", errors="ignore")
            xml = re.sub(r"<a:br\s*/>", "\n", xml)
            texts = re.findall(r"<a:t>(.*?)</a:t>", xml, flags=re.S)
            if texts:
                s = BR_RX.sub("\n", "".join(html.unescape(t) for t in texts))
                parts = [p.strip() for p in re.split(r"\n{2,}", s) if p.strip()]
                blocks.extend(parts)
        for name in note_names:
            xml = z.read(name).decode("utf-8", errors="ignore")
            texts = re.findall(rf"<a:t>(.*?)</a:t>", xml, flags=re.S)
            if texts:
                s = BR_RX.sub("
",  "".join(t.replace("&lt;","<").replace("&gt;",">").replace("&amp;","&") for t in texts))
                parts = [p.strip() for p in re.split(r"
{2,}", s) if p.strip()]
                blocks.extend(parts)
    return [b for b in blocks if b.strip()]
def count_tokens_for_blocks(blocks, model_key: str):
    enc = encoding_for_model_key(model_key)
    toks = 0
    jp_chars = 0
    for b in blocks:
        toks += len(enc.encode(b))
        jp_chars += sum(1 for _ in JP_RX.finditer(b))
    return toks, jp_chars
def compute_requests(n_blocks: int, batch_size: int) -> int:
    return max(1, math.ceil(n_blocks / max(1, batch_size)))

def estimate_output_tokens(total_jp_chars: int, expansion: float, avg_en_chars_per_token: float) -> int:
    en_chars = int(round(total_jp_chars * expansion))
    return int(round(en_chars / max(1e-9, avg_en_chars_per_token)))

def load_pricing_overrides(path: str):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    for k, v in data.items():
        base = {
            "in": v.get("in"),
            "in_cached": v.get("in_cached"),
            "out": v.get("out"),
            "cached_write": v.get("cached_write"),
            "cached_read": v.get("cached_read"),
            "tokenizer": v.get("tokenizer", "o200k_base"),
        }
        PRICING[k] = base

def split_cached_openai(total_in_tokens: int, prefix_tokens: int, n_reqs: int, no_cache: bool):
    """OpenAI: discounted cached input after the first call."""
    if no_cache or n_reqs <= 1 or prefix_tokens <= 0:
        return total_in_tokens, 0
    CACHE_THRESHOLD = 1024
    cached_per_call = max(0, prefix_tokens - CACHE_THRESHOLD)
    dynamic = max(0, total_in_tokens - prefix_tokens * n_reqs)
    uncached_total = dynamic + prefix_tokens  # first call's full prefix
    cached_total   = cached_per_call * (n_reqs - 1)
    return uncached_total, cached_total

def split_cached_anthropic(total_in_tokens: int, prefix_tokens: int, n_reqs: int, no_cache: bool, charge_write: bool):
    """
    Anthropic exposes separate caching prices (write/read).
    We approximate:
      - first call: prefix portion is 'cached_write' if charge_write else billed as regular input
      - subsequent calls: cached portion billed at 'cached_read'
    """
    if no_cache or n_reqs <= 1 or prefix_tokens <= 0:
        return {"uncached_in": total_in_tokens, "cached_read": 0, "cached_write": 0}
    CACHE_THRESHOLD = 1024
    cache_portion = max(0, prefix_tokens - CACHE_THRESHOLD)
    dynamic = max(0, total_in_tokens - prefix_tokens * n_reqs)
    uncached_in = dynamic + (0 if charge_write else prefix_tokens)  # if charge_write, don't also count the prefix as regular input
    cached_write = cache_portion if charge_write else 0
    cached_read  = cache_portion * (n_reqs - 1)
    return {"uncached_in": uncached_in, "cached_read": cached_read, "cached_write": cached_write}

def estimate_cost(model_key: str, parts: dict):
    """
    parts may contain:
      - uncached_in, cached_in, out   (OpenAI-like)
      - uncached_in, cached_read, cached_write, out  (Anthropic-like)
    """
    cfg = PRICING.get(model_key, {})
    per_m = 1_000_000.0
    if not cfg or (cfg.get("in") is None and cfg.get("out") is None and cfg.get("in_cached") is None
                   and cfg.get("cached_read") is None and cfg.get("cached_write") is None):
        raise ValueError(f"No pricing configured for {model_key}. Use --pricing to supply rates.")

    cost = 0.0
    # OpenAI-style fields
    if "cached_in" in parts:
        cost += (parts["uncached_in"] / per_m) * (cfg["in"] or 0.0)
        cost += (parts["cached_in"]   / per_m) * (cfg.get("in_cached") or cfg.get("in") or 0.0)
    # Anthropic-style fields
    if "cached_read" in parts or "cached_write" in parts:
        # Uncached input (dynamic + maybe first-call prefix if not charged as write)
        cost += (parts.get("uncached_in", 0) / per_m) * (cfg["in"] or 0.0)
        # Cache read/write
        if parts.get("cached_read", 0):
            cost += (parts["cached_read"] / per_m) * (cfg.get("cached_read") or cfg.get("in") or 0.0)
        if parts.get("cached_write", 0):
            cost += (parts["cached_write"] / per_m) * (cfg.get("cached_write") or cfg.get("in") or 0.0)
    # Outputs (common)
    cost += (parts.get("out", 0) / per_m) * (cfg.get("out") or 0.0)
    return cost

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx", help="Path to PPTX")
    ap.add_argument("--producer", default="openai:gpt-5", help="Provider:model for Producer")
    ap.add_argument("--reviewer", default="openai:gpt-5-mini", help="Provider:model for Reviewer")
    ap.add_argument("--batch-size", type=int, default=16, help="Blocks per request")
    ap.add_argument("--prefix-file", default=None, help="File with the repeated prefix (system+schema+glossary)")
    ap.add_argument("--expansion", type=float, default=1.45, help="EN char expansion over JA (e.g., 1.3–1.6)")
    ap.add_argument("--avg-en-chars-per-token", type=float, default=4.0, help="Avg English chars per token (~4)")
    ap.add_argument("--also", nargs="*", default=[], help="Also show Producer costs for these provider:models")
    ap.add_argument("--no-cache", action="store_true", help="Ignore prompt caching entirely")
    ap.add_argument("--anthropic-cache-write", action="store_true",
                    help="Charge first-call cached prefix at Anthropic 'cached_write' price instead of regular input")
    ap.add_argument("--pricing", default=None, help="Path to pricing JSON to override/add models")
    args = ap.parse_args()

    if args.pricing and os.path.exists(args.pricing):
        load_pricing_overrides(args.pricing)

    models_to_check = [args.producer, args.reviewer] + args.also
    for m in models_to_check:
        if m not in PRICING:
            print(f"ERROR: unknown model '{m}'. Known: {', '.join(PRICING)}", file=sys.stderr)
            sys.exit(2)

    blocks = extract_text_blocks(args.pptx)
    if not blocks:
        print("No text blocks found. Is this PPTX empty or image-only?", file=sys.stderr)
        sys.exit(1)

    # Input tokens per model (tokenizers can differ)
    prod_in_tokens, prod_jp_chars = count_tokens_for_blocks(blocks, args.producer)
    rev_in_tokens,  _             = count_tokens_for_blocks(blocks, args.reviewer)

    n_reqs = compute_requests(len(blocks), args.batch_size)

    # Prefix tokens (for caching math)
    prefix_tokens = {}
    if args.prefix_file and os.path.exists(args.prefix_file):
        text = open(args.prefix_file, "r", encoding="utf-8").read()
        for m in [args.producer, args.reviewer] + args.also:
            enc = encoding_for_model_key(m)
            prefix_tokens[m] = len(enc.encode(text))
    else:
        for m in [args.producer, args.reviewer] + args.also:
            prefix_tokens[m] = 0

    # Output token estimate from JA chars
    out_tokens_est = estimate_output_tokens(prod_jp_chars, args.expansion, args.avg_en_chars_per_token)

    # Compute Producer input cost parts by provider
    def input_parts(model_key, total_in):
        provider = model_key.split(":")[0]
        pf = prefix_tokens.get(model_key, 0)
        if provider == "openai":
            unc, cac = split_cached_openai(total_in, pf, n_reqs, args.no_cache)
            return {"uncached_in": unc, "cached_in": cac}
        elif provider == "anthropic":
            return split_cached_anthropic(total_in, pf, n_reqs, args.no_cache, args.anthropic_cache_write)
        else:
            # default: no caching math
            return {"uncached_in": total_in}

    prod_parts = input_parts(args.producer, prod_in_tokens)
    prod_parts["out"] = out_tokens_est

    rev_parts  = input_parts(args.reviewer, rev_in_tokens)
    # reviewer assumed to return small JSON diagnostics only (treat as ~0 output)
    rev_parts["out"] = 0

    # Cost calculations
    try:
        prod_cost = estimate_cost(args.producer, prod_parts)
        rev_cost  = estimate_cost(args.reviewer,  rev_parts)
    except ValueError as e:
        print(f"Pricing error: {e}", file=sys.stderr); sys.exit(2)

    # Summary
    def money(x): return f"${x:,.4f}"
    print(f"Deck: {os.path.basename(args.pptx)}")
    print(f"Blocks detected: {len(blocks)}  |  Requests: {n_reqs} (batch-size={args.batch_size})")
    print(f"JP chars: {prod_jp_chars:,}  |  Producer input tokens: {prod_in_tokens:,}  |  Est. EN output tokens: {out_tokens_est:,}")
    if args.prefix_file:
        print(f"Prefix file: {args.prefix_file}")
        print(f"  Prefix tokens (Producer): {prefix_tokens[args.producer]:,}")
        print(f"  Prefix tokens (Reviewer): {prefix_tokens[args.reviewer]:,}")
    print("=== Producer & Reviewer cost ===")
    print(f"Producer: {args.producer:28} → {money(prod_cost)}  ({prod_parts})")
    print(f"Reviewer: {args.reviewer:28} → {money(rev_cost)}  ({rev_parts})")
    print(f"TOTAL: {money(prod_cost + rev_cost)}")
    if args.anthropic_cache_write:
        print("Note: Anthropic first-call prefix billed at 'cached_write' rate.")

    if args.also:
        print("=== Alternatives for Producer (same tokens/caching) ===")
        for m in args.also:
            parts = input_parts(m, prod_in_tokens); parts["out"] = out_tokens_est
            try:
                alt = estimate_cost(m, parts)
            except ValueError as e:
                print(f"{m:28} → ERROR: {e}")
                continue
            print(f"{m:28} → {money(alt)}  ({parts})")

if __name__ == "__main__":
    main()