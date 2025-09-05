# ruff: noqa: S101
# NOTE: Test framework: prefers pytest if available; also compatible with unittest discovery.
# These tests focus on functions defined in style_checker: create_style_checker_prompt,
# check_title_case_violations, check_bullet_punctuation, check_glossary_violations,
# check_banned_phrases, check_punctuation_errors, analyze_parallelism, check_tone_drift,
# run_style_check, apply_style_fixes, and model_style_check.

import json
import re
import types
import importlib
import sys

import pytest  # pytest is widely used; if not present, replace with unittest equivalents.
from typing import ClassVar

# Ensure the module under test can be imported consistently across repo layouts.
# Try common locations for style_checker.
_STYLE_CHECKER = None
_IMPORT_ERRORS = []

for mod_name in ("style_checker", "src.style_checker", "app.style_checker"):
    try:
        _STYLE_CHECKER = importlib.import_module(mod_name)
        break
    except ImportError as e:  # capture import error to help debugging
        _IMPORT_ERRORS.append((mod_name, repr(e)))

if _STYLE_CHECKER is None:
    raise ImportError(f"Could not import style_checker module. Tried: {_IMPORT_ERRORS}")  # noqa: TRY003

sc = _STYLE_CHECKER


def _install_style_normalize_stub(monkeypatch,
                                  small_words=None,
                                  banned_phrases=None,
                                  title_case_impl=None,
                                  guide="1) Titles use Title Case\n2) Bullets omit terminal punctuation"):
    """
    Install a stub style_normalize module to isolate tests from external dependencies.
    """
    stub = types.SimpleNamespace()
    stub.SMALL_WORDS = small_words or {"a", "an", "the", "on", "in", "to", "and", "or", "for", "of"}
    stub.BANNED_PHRASES = banned_phrases or {
        "state-of-the-art": "advanced",
        "industry-leading": "leading",
        "cutting-edge": "advanced"
    }

    def _default_title_case(s: str) -> str:
        # Very simple title-case for testing: capitalize all words except small words.
        words = re.split(r'(\s+)', s or "")
        out = []
        for w in words:
            if not w.strip():
                out.append(w)
            else:
                lw = w.lower()
                if lw in stub.SMALL_WORDS:
                    out.append(lw)
                else:
                    out.append(lw[:1].upper() + lw[1:])
        return "".join(out).strip()

    stub.title_case = title_case_impl or _default_title_case
    stub.get_style_guide = lambda: guide

    # Inject stub into sys.modules so 'from style_normalize import ...' resolves.
    monkeypatch.setitem(sys.modules, "style_normalize", stub)
    # Reload module under test so it binds to the stubbed module attributes
    importlib.reload(sc)
    return stub


class TestCreateStyleCheckerPrompt:
    def test_prompt_includes_style_guide_and_sections(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        prompt = sc.create_style_checker_prompt(
            glossary={"導入": "Implementation", "システム": "System"},
            deck_tone={"voice": "confident", "formality": "neutral"}
        )
        assert "You are a style reviewer for marketing slide translations" in prompt
        assert "Key glossary terms" in prompt
        # Ensure terms are serialized as requested
        assert '"導入" → "Implementation"' in prompt
        # Deck tone should be pretty-printed JSON
        assert '"voice": "confident"' in prompt
        assert '"formality": "neutral"' in prompt

    def test_prompt_without_optional_sections(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        prompt = sc.create_style_checker_prompt()
        assert "Key glossary terms" not in prompt
        assert "Deck Tone Profile" not in prompt


class TestTitleCaseViolations:
    def test_detects_non_title_case_in_short_title(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        texts = ["unlock the next wave of growth"]
        out = sc.check_title_case_violations(texts)
        assert len(out) == 1
        assert out[0]["index"] == 0
        assert out[0]["issue"] == "not in Title Case"
        assert "suggested_fix" in out[0]
        assert out[0]["suggested_fix"] != texts[0]

    def test_skips_sentences_or_long_lines(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        texts = ["This is a full sentence.", "This is a long line " + "word " * 20]
        out = sc.check_title_case_violations(texts)
        assert out == []


class TestBulletPunctuation:
    def test_flags_terminal_punctuation_for_bullets(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        texts = ["[li-1] Increase revenue.", "• Expand to APAC;"]
        out = sc.check_bullet_punctuation(texts)
        assert {v["index"] for v in out} == {0, 1}
        # Suggested fix should remove trailing punctuation
        assert out[0]["suggested_fix"].endswith("Increase revenue")
        assert out[1]["suggested_fix"].endswith("Expand to APAC")

    def test_does_not_flag_multi_sentence_bullet(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        texts = ["• Expand to APAC. Start with Japan."]
        out = sc.check_bullet_punctuation(texts)
        assert out == []


class TestGlossaryViolations:
    def test_detects_glossary_mismatches(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        glossary = {"導入": "Implementation", "システム": "System"}
        # Uses an alternative "deployment" for 導入; should be flagged
        texts = ["Seamless deployment across regions"]
        out = sc.check_glossary_violations(texts, glossary)
        assert len(out) == 1
        assert out[0]["index"] == 0
        assert out[0]["term"] == "導入"
        assert out[0]["expected"] == "Implementation"
        assert out[0]["found"] == "deployment"

    def test_returns_empty_when_no_glossary(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        out = sc.check_glossary_violations(["Any text"], {})
        assert out == []


class TestBannedPhrases:
    def test_flags_banned_phrases_case_insensitive(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch, banned_phrases={"Cutting-edge": "advanced"})
        texts = ["Our cutting-EDGE solution is reliable"]
        out = sc.check_banned_phrases(texts)
        assert len(out) == 1
        assert out[0]["phrase"].lower() == "cutting-edge"
        assert out[0]["suggested"] == "advanced"
        assert "context" in out[0]


class TestPunctuationErrors:
    def test_detects_japanese_punctuation(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        texts = ["高性能、低コスト。", "「Smart」ソリューション"]
        out = sc.check_punctuation_errors(texts)
        originals = {(v["index"], v["original"]) for v in out}
        # Expect to see 、 and 。 and 「 and 」
        assert (0, "、") in originals
        assert (0, "。") in originals
        assert (1, "「") in originals
        assert (1, "」") in originals


class TestAnalyzeParallelism:
    def test_mixed_verb_forms_are_flagged(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        texts = [
            "[li-1] Launching in Q4",   # gerund
            "[li-2] Reduce churn",      # verb base
            "[li-3] Expanding margins", # gerund
        ]
        out = sc.analyze_parallelism(texts)
        assert len(out) == 1
        issue = out[0]
        assert issue["issue"] == "inconsistent verb forms"
        assert len(issue["indices"]) == 3

    def test_groups_less_than_three_not_flagged(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        texts = [
            "[li-1] Launching in Q4",
            "[li-2] Expanding margins",
        ]
        assert sc.analyze_parallelism(texts) == []


class TestRunStyleCheckAndFixes:
    def test_run_style_check_structure_and_apply_fixes(self, monkeypatch):
        # Stub out style_normalize and tone drift to isolate tests.
        _install_style_normalize_stub(monkeypatch)

        class FakeClient:
            # check_tone_drift should not be called if deck_tone is None; still define minimal shape
            class chat:
                class completions:
                    @staticmethod
                    def create(**_kwargs):
                        raise RuntimeError

        translations = [
            "unlock the next wave of growth",       # title case violation
            "[li-1] Increase revenue.",             # bullet punctuation
            "Seamless deployment across regions",   # glossary violation for 導入->Implementation
            "Our state-of-the-art solution",        # banned phrase
            "高性能、低コスト。"                        # punctuation errors
        ]
        glossary = {"導入": "Implementation"}
        diagnostics = sc.run_style_check(FakeClient(), translations, glossary=glossary, deck_tone=None)

        # Validate structure
        assert "style" in diagnostics
        style = diagnostics["style"]
        for key in ("title_case_violations", "bullet_terminal_punctuation",
                    "parallelism_issues", "glossary_violations",
                    "banned_phrases", "punctuation_errors"):
            assert key in style

        assert diagnostics.get("tone_flags", {}) == {}

        # Now apply fixes and verify deterministic changes
        fixed = sc.apply_style_fixes(translations, diagnostics)
        # Title case should be applied in index 0 (simplistic stub)
        assert fixed[0] != translations[0]
        # Bullet punctuation removed
        assert fixed[1].endswith("Increase revenue")
        assert not fixed[1].rstrip().endswith(".")
        # Glossary replacement (deployment -> Implementation)
        assert "Implementation" in fixed[2]
        assert "deployment" not in fixed[2]
        # Banned phrase replaced
        assert "state-of-the-art" not in fixed[3]
        # JP punctuation converted
        assert "、" not in fixed[4] and "。" not in fixed[4]

    def test_apply_style_fixes_bounds_checks(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        translations = ["Keep as-is"]
        diagnostics = {
            "style": {
                "title_case_violations": [{"index": 5, "suggested_fix": "X"}],
                "bullet_terminal_punctuation": [{"index": -1, "suggested_fix": "X"}],
                "banned_phrases": [{"index": 99, "phrase": "foo", "suggested": "bar"}],
                "punctuation_errors": [{"index": 42, "original": "。", "correct": ". "}],
                "glossary_violations": [{"index": 100, "found": "old", "expected": "new"}],
                "parallelism_issues": []
            }
        }
        fixed = sc.apply_style_fixes(translations, diagnostics)
        assert fixed == translations  # out-of-range indices are ignored safely


class TestToneDrift:
    def test_tone_drift_returns_empty_when_no_deck_tone(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)
        class FakeClient:
            pass
        out = sc.check_tone_drift(FakeClient(), ["A"], deck_tone=None)
        assert out == {}

    def test_tone_drift_success_path_parses_json(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)

        class FakeCreate:
            @staticmethod
            def create(**_kwargs):
                class Msg:
                    content = json.dumps({
                        "tone_flags": {
                            "added_hype": ["industry-leading"],
                            "softened_claims": [],
                            "over_formalized": False,
                            "over_casual": False,
                            "deviation_from_deck_profile": []
                        }
                    })
                class Choice:
                    message = Msg()
                class Resp:
                    choices: ClassVar[list] = [Choice()]
                return Resp()

        class FakeClient:
            class chat:
                completions = FakeCreate()

        out = sc.check_tone_drift(FakeClient(), ["A"], deck_tone={"voice": "calm"})
        assert "tone_flags" in out
        assert out["tone_flags"]["added_hype"] == ["industry-leading"]


class TestModelStyleCheck:
    def test_model_style_check_success_and_fallback(self, monkeypatch):
        _install_style_normalize_stub(monkeypatch)

        # Success path
        class RespSuccess:
            def output_text(self):
                return json.dumps({"style": {"ok": True}})
            choices: ClassVar[list] = []

        class ClientSuccess:
            class responses:
                @staticmethod
                def create(**_kwargs):
                    return RespSuccess()

        out = sc.model_style_check(ClientSuccess(), ["Title one", "Title two"], glossary=None, deck_tone=None)
        assert out == {"style": {"ok": True}}

        # Fallback path (simulate no content returned) should call run_style_check
        class RespEmpty:
            output_text = None
            choices: ClassVar[list] = []

        class ClientEmpty:
            class responses:
                @staticmethod
                def create(**_kwargs):
                    return RespEmpty()

        result = sc.model_style_check(ClientEmpty(), ["lowercase title"], glossary=None, deck_tone=None)
        # Should return structure from run_style_check (style dict present)
        assert "style" in result
        assert isinstance(result["style"], dict)