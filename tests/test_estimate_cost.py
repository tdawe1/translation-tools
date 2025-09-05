"""
Test suite for estimate_cost module functions.

Test framework: pytest
- We use pytest fixtures (monkeypatch, tmp_path, capsys) and idiomatic asserts.
- External dependency 'tiktoken' is mocked via sys.modules to ensure importability.

These tests exercise:
- Tokenizer selection and fallbacks
- PPTX text extraction (slides + notes) and HTML entity decoding
- Token counting and JP character detection
- Batch/request computation and output token estimation
- OpenAI- and Anthropic-style prompt caching split logic
- Cost estimation math for multiple providers
- Pricing overrides loader behavior
- CLI error path (unknown model) with proper exit code and stderr message
"""
# ruff: noqa: S101

from __future__ import annotations

import importlib
import io
import json
import math
import sys
import types
import zipfile
from typing import Optional

import pytest


# ---------- Helpers & fixtures ----------


def _install_fake_tiktoken():
    """
    Install a lightweight fake 'tiktoken' module into sys.modules before importing
    the module-under-test. This prevents the real dependency from being required.
    """
    fake = types.ModuleType("tiktoken")

    class FakeEncoding:
        def __init__(self, name: str = "fake"):
            self.name = name

        def encode(self, s: str):
            # Simple deterministic "tokenization": 1 token per character
            # (sufficient for verifying arithmetic in tests).
            return [ord(ch) for ch in s]

    def get_encoding(name: str):
        fake.last_requested = name
        return FakeEncoding(name)

    fake.get_encoding = get_encoding  # type: ignore[attr-defined]
    fake.last_requested = None  # type: ignore[attr-defined]
    sys.modules["tiktoken"] = fake
    return fake


@pytest.fixture
def mod():
    """
    Import the module under test after installing a fake 'tiktoken'.
    We try a few common locations to remain robust to repo layout.
    """
    fake = _install_fake_tiktoken()
    candidates = [
        "estimate_cost",
        "src.estimate_cost",
        "app.estimate_cost",
        "scripts.estimate_cost",
        "tools.estimate_cost",
    ]
    # Ensure fresh import
    for name in candidates:
        sys.modules.pop(name, None)

    module = None
    last_err: Optional[BaseException] = None
    for name in candidates:
        try:
            module = importlib.import_module(name)
            break
        except ModuleNotFoundError as e:
            last_err = e
            continue

    if module is None:
        raise ImportError from last_err

    # Attach reference to fake tiktoken to the module for convenience in some assertions
    module._fake_tiktoken = fake  # type: ignore[attr-defined]
    return module


@pytest.fixture
def reset_pricing(mod):
    """Snapshot and restore global PRICING dict to avoid test cross-talk."""
    import copy

    orig = copy.deepcopy(mod.PRICING)
    try:
        yield
    finally:
        mod.PRICING.clear()
        mod.PRICING.update(orig)


def make_min_pptx(tmp_path, *, slide_texts, notes_texts=()):
    """
    Create a minimal PPTX-like ZIP file containing slide and notes XML with <a:t> tags.
    Returns the path to the created file.
    """
    pptx_path = tmp_path / "sample.pptx"
    with zipfile.ZipFile(pptx_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        # Slides
        for i, parts in enumerate(slide_texts, start=1):
            xml = "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
            for t in parts:
                # Include HTML entities to verify decoding
                xml += f"<a:t>{t}</a:t>"
            xml += "</p:sld>"
            z.writestr(f"ppt/slides/slide{i}.xml", xml)

        # Notes
        for i, parts in enumerate(notes_texts, start=1):
            xml = "<p:notes xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
            for t in parts:
                xml += f"<a:t>{t}</a:t>"
            xml += "</p:notes>"
            z.writestr(f"ppt/notesSlides/notesSlide{i}.xml", xml)
    return pptx_path


# ---------- Tests: tokenizer and encoding ----------


def test_encoding_for_model_key_uses_configured_tokenizer(mod, reset_pricing):
    _ = reset_pricing
    # Add a dummy model with a custom tokenizer alias
    mod.PRICING["dummy:model"] = {"in": 1.0, "out": 2.0, "tokenizer": "my_tokenizer"}
    enc = mod.encoding_for_model_key("dummy:model")
    assert hasattr(enc, "encode")
    # Ensure our fake tiktoken was asked for the right tokenizer name
    assert getattr(sys.modules["tiktoken"], "last_requested", None) == "my_tokenizer"


def test_encoding_for_model_key_falls_back_when_unknown_tokenizer(mod, reset_pricing):
    _ = reset_pricing
    # Force tiktoken.get_encoding to raise and verify fallback to "o200k_base"
    tk = sys.modules["tiktoken"]

    def raising_get_encoding(name: str):
        if name != "o200k_base":
            raise RuntimeError("unknown encoding")  # noqa
        return type("E", (), {"encode": lambda _self, s: list(s)})()

    tk.get_encoding = raising_get_encoding  # type: ignore[assignment]
    enc = mod.encoding_for_model_key("openai:gpt-5")  # configured to "o200k_base"
    assert hasattr(enc, "encode")


# ---------- Tests: PPTX extraction & counting ----------


def test_extract_text_blocks_reads_slides_and_notes_and_decodes_entities(mod, tmp_path):
    pptx = make_min_pptx(
        tmp_path,
        slide_texts=[
            ["こんにちは", " ", "&lt;tag&gt;", " &amp; ", "World"],
            ["Line2"],
        ],
        notes_texts=[["Note1"], ["Note2 &amp; more"]],
    )
    blocks = mod.extract_text_blocks(str(pptx))
    assert isinstance(blocks, list) and len(blocks) >= 1

    combined = "\n".join(blocks)
    # Entities decoded
    assert "<tag>" in combined
    assert "&" in combined
    # Slide and note content present
    assert "こんにちは" in combined
    assert "World" in combined
    assert "Line2" in combined
    assert "Note1" in combined
    assert "Note2" in combined


def test_count_tokens_for_blocks_counts_jp_chars_and_tokens(mod):
    blocks = ["abcこんにちはde", "ノート"]
    toks, jp_chars = mod.count_tokens_for_blocks(blocks, "openai:gpt-5")
    # Our fake encoding: tokens == total characters across blocks
    assert toks == sum(len(b) for b in blocks)
    # Cross-check JP char count using the module's own regex
    expected_jp = sum(1 for b in blocks for _ in mod.JP_RX.finditer(b))
    assert jp_chars == expected_jp


# ---------- Tests: small pure functions ----------


@pytest.mark.parametrize(
    "n_blocks,batch,expected",
    [
        (0, 0, 1),
        (1, 1, 1),
        (16, 16, 1),
        (17, 16, 2),
        (100, 16, 7),
        (100, -5, 1),  # batch clamped to >=1
    ],
)
def test_compute_requests_various(mod, n_blocks, batch, expected):
    assert mod.compute_requests(n_blocks, batch) == expected


@pytest.mark.parametrize(
    "jp_chars,expansion,avg_chars_per_token,expected",
    [
        (100, 1.5, 5.0, 30),   # 150/5
        (101, 1.5, 4.0, 38),   # round(151.5)=152; 152/4 = 38
        (0,   2.0, 4.0, 0),
    ],
)
def test_estimate_output_tokens_rounding(mod, jp_chars, expansion, avg_chars_per_token, expected):
    assert mod.estimate_output_tokens(jp_chars, expansion, avg_chars_per_token) == expected


# ---------- Tests: caching split logic ----------


def test_split_cached_openai_happy_path(mod):
    total_in = 10_000
    prefix = 2_000
    n_reqs = 3
    uncached, cached = mod.split_cached_openai(total_in, prefix, n_reqs, no_cache=False)
    # Manual expectations:
    # CACHE_THRESHOLD=1024 -> cached_per_call = 2000-1024=976
    # dynamic = total - prefix*n = 10000 - 6000 = 4000
    # uncached_total = dynamic + prefix(first call) = 4000 + 2000 = 6000
    # cached_total = 976 * (3-1) = 1952
    assert (uncached, cached) == (6000, 1952)


def test_split_cached_openai_no_cache_or_single_request(mod):
    assert mod.split_cached_openai(10000, 2000, 1, no_cache=False) == (10000, 0)
    assert mod.split_cached_openai(10000, 2000, 3, no_cache=True) == (10000, 0)
    # Prefix below threshold -> no cached portion
    assert mod.split_cached_openai(10000, 1024, 3, no_cache=False) == (10000, 0)


def test_split_cached_anthropic_write_vs_no_write(mod):
    total_in = 10_000
    prefix = 2_000
    n_reqs = 3
    # charge_write=True
    parts = mod.split_cached_anthropic(total_in, prefix, n_reqs, no_cache=False, charge_write=True)
    # cache_portion = 976, dynamic=4000
    assert parts == {"uncached_in": 4000, "cached_read": 1952, "cached_write": 976}
    # charge_write=False
    parts2 = mod.split_cached_anthropic(total_in, prefix, n_reqs, no_cache=False, charge_write=False)
    assert parts2 == {"uncached_in": 6000, "cached_read": 1952, "cached_write": 0}


# ---------- Tests: cost estimation ----------


def test_estimate_cost_openai_style_math(mod, reset_pricing):
    _ = reset_pricing
    parts = {"uncached_in": 6000, "cached_in": 1952, "out": 2500}
    total = mod.estimate_cost("openai:gpt-5", parts)
    # Expected:
    # uncached_in: 6000/1e6 * 1.25 = 0.0075
    # cached_in:   1952/1e6 * 0.125 = 0.000244
    # out:         2500/1e6 * 10.00 = 0.025
    expected = 0.0075 + 0.000244 + 0.025
    assert math.isclose(total, expected, rel_tol=1e-9, abs_tol=1e-12)


def test_estimate_cost_anthropic_style_math(mod, reset_pricing):
    _ = reset_pricing
    parts = {"uncached_in": 4000, "cached_read": 1952, "cached_write": 976, "out": 0}
    total = mod.estimate_cost("anthropic:claude-sonnet-4", parts)
    # Expected:
    # uncached: 4000/1e6 * 3.00   = 0.012
    # read:     1952/1e6 * 0.30   = 0.0005856
    # write:     976/1e6 * 3.75   = 0.00366
    expected = 0.012 + 0.0005856 + 0.00366
    assert math.isclose(total, expected, rel_tol=1e-9, abs_tol=1e-12)


def test_estimate_cost_raises_for_unpriced_model(mod):
    # PRICING entry where all relevant prices are None -> should raise
    with pytest.raises(ValueError):
        mod.estimate_cost("openai:gpt-4.1", {"uncached_in": 1000, "out": 1000})


def test_load_pricing_overrides_adds_and_updates_models(mod, tmp_path, reset_pricing):
    _ = reset_pricing
    pricing = {
        "openai:new-mini": {"in": 0.11, "in_cached": 0.01, "out": 0.9, "tokenizer": "o200k_base"},
        "acme:ultra": {"in": 2.0, "out": 5.0, "cached_read": 0.3, "cached_write": 1.1, "tokenizer": "acme_tok"},
    }
    pfile = tmp_path / "pricing.json"
    pfile.write_text(json.dumps(pricing), encoding="utf-8")

    mod.load_pricing_overrides(str(pfile))

    assert "openai:new-mini" in mod.PRICING
    assert "acme:ultra" in mod.PRICING
    assert mod.PRICING["acme:ultra"]["tokenizer"] == "acme_tok"

    # Ensure encoding_for_model_key requests the tokenizer we just set
    enc = mod.encoding_for_model_key("acme:ultra")
    assert hasattr(enc, "encode")
    assert getattr(sys.modules["tiktoken"], "last_requested", None) == "acme_tok"


# ---------- Tests: CLI error path ----------


def test_main_exits_with_error_on_unknown_model(mod, monkeypatch, capsys, tmp_path):
    # Unknown reviewer model should trigger early exit (before PPTX parsing).
    dummy_pptx = tmp_path / "does_not_matter.pptx"
    # No need to create the file; model validation happens first.
    argv = [
        "estimate_cost.py",
        str(dummy_pptx),
        "--producer", "openai:gpt-5",
        "--reviewer", "acme:nonexistent",
    ]
    monkeypatch.setattr(sys, "argv", argv)
    with pytest.raises(SystemExit) as e:
        mod.main()
    assert e.value.code == 2
    err = capsys.readouterr().err
    assert "unknown model 'acme:nonexistent'" in err