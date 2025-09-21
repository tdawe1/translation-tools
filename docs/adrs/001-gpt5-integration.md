# ADR 001: Integration of GPT-5 Adapter into Translation Scripts

## Status
Proposed

## Context
The current translation pipeline in `scripts/translate_pptx_inplace.py` and `scripts/translate_pdf.py` relies on direct calls to OpenAI's API using GPT-4o models for Japanese-to-English translation. With the anticipated release of GPT-5, which promises enhanced translation accuracy, contextual understanding, and efficiency, we need to integrate support for it. This integration must maintain backward compatibility, allow flexible model selection, and ensure no performance degradation or increased error rates. The goal is to abstract the AI interaction to enable easy extension to future models or providers without refactoring core logic.

Key requirements:
- Support CLI flag `--adapter gpt5` for explicit selection.
- Environment variable `OPENAI_ADAPTER=gpt5` for configuration (with defaults to `gpt4o`).
- Fallback to GPT-4o on GPT-5 errors (e.g., rate limits, unavailability).
- Preserve existing batching, caching, and style processing logic.
- No performance regression (e.g., translation time within 10% of current).

## Decision
Adopt an **adapter pattern** to abstract all OpenAI API interactions. This involves:
- Creating a base `TranslationAdapter` interface defining the translation contract.
- Implementing concrete adapters for `Gpt4oAdapter` (existing/default) and `Gpt5Adapter` (new).
- Replacing direct OpenAI calls in the scripts with adapter invocations at specific integration points.
- Configuring adapter selection via CLI flag or environment variable, with `gpt4o` as default.
- Implementing fallback logic in the scripts: attempt GPT-5, revert to GPT-4o on failure.

This approach ensures modularity, testability, and minimal changes to existing code (~50-100 LOC additions per script).

### Configuration
- **Default:** `gpt4o` (maps to model `'gpt-4o-2024-08-06'`).
- **CLI Flag:** `--adapter gpt5` (or `--adapter gpt4o`); parsed via `argparse`.
- **Environment Variable:** `OPENAI_ADAPTER=gpt5` takes precedence over CLI; falls back to default if invalid.
- Factory function `get_adapter(name: str) -> TranslationAdapter` in a new `adapters/factory.py` module to instantiate based on config.

### Integration Points
1. **scripts/translate_pptx_inplace.py:**
   - Locate the batch translation loop (likely in `_process_batches` or similar function, around lines where `openai.ChatCompletion.create` or `client.chat.completions.create` is called).
   - Replace direct API call with `adapter.translate_batch(batch_items, prompt_template)`.
   - Ensure cache keys include adapter name to avoid mixing translations.

2. **scripts/translate_pdf.py:**
   - Similar to PPTX: In the translation phase after text extraction (e.g., `_translate_extracted_texts` function).
   - Replace OpenAI calls with adapter method.
   - Adjust for PDF-specific prompts if needed (pass via `prompt_template`).

3. **Shared Logic:**
   - If there's a shared translation module (e.g., `core/translation.py`), integrate there first for reuse.
   - Update cost estimation in `tools/estimate_cost.py` and `tools/estimate_cost_pdf.py` to account for adapter/model token rates.

### Fallback Mechanism
- In script main loop: `try: results = adapter.translate_batch(...) except AdapterError: fallback_adapter = get_adapter('gpt4o'); results = fallback_adapter.translate_batch(...)`.
- Log fallback events with details (error type, batch size).
- Limit fallbacks to once per job to avoid infinite loops.

### Error Handling
- Adapters raise custom `AdapterError` on API failures (e.g., rate limits via `openai.RateLimitError`).
- Implement retry logic in adapters: 3 attempts with exponential backoff (base 1s, max 60s, jitter enabled) before raising error.

## Alternatives Considered
1. **Direct Model Switching:** Pass model name as a parameter to OpenAI calls.
   - Pros: Minimal code changes (~10 LOC).
   - Cons: Tight coupling to OpenAI; hard to extend to non-OpenAI providers; no uniform error/fallback handling.
   - Rejected: Lacks future-proofing.

2. **Configuration File:** Use YAML/JSON for model configs.
   - Pros: Centralized config.
   - Cons: Adds file I/O overhead; overkill for simple selection.
   - Rejected: CLI/env sufficient for ops flexibility.

3. **Full Service Layer:** Extract translation to a dedicated microservice.
   - Pros: Scalable for high load.
   - Cons: Increases complexity/timeline; not needed for current batch processing.
   - Rejected: Premature optimization.

Chosen adapter pattern balances simplicity, extensibility, and the requirement for fallbacks.

## Consequences
### Positive
- **Flexibility:** Easy to add adapters for other models (e.g., GPT-5 variants, Anthropic) or providers.
- **Maintainability:** Core scripts unchanged; logic isolated in adapters.
- **Testability:** Mock adapters for unit/integration tests without real API calls.
- **No Rework:** Existing cache, batching, and post-processing remain intact.

### Negative
- **Added Abstraction:** Slight indirection may confuse new developers (mitigate with docs).
- **Potential Overhead:** Minimal (~5-10ms per batch call); monitor in tests.
- **Dependency:** Relies on OpenAI SDK stability; version pin to `openai>=1.0.0,<2.0.0`.

## Technical Specifications

### Interfaces
Create `adapters/base.py`:
```python
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional
from openai import OpenAIError  # Or custom exceptions

class AdapterError(Exception):
    """Base exception for adapter failures."""
    pass

class TranslationAdapter(ABC):
    def __init__(self, model: str, api_key: Optional[str] = None):
        self.model = model
        self.client = openai.OpenAI(api_key=api_key or os.getenv('OPENAI_API_KEY'))

    @abstractmethod
    def translate_batch(
        self,
        items: List[Dict[str, str]],  # e.g., [{'text': 'Japanese text', 'context': 'slide title'}]
        prompt_template: str,  # Base prompt, e.g., "Translate to natural English: {text}"
        max_tokens: int = 4096,
        temperature: float = 0.3
    ) -> List[Dict[str, Any]]:  # e.g., [{'original': '...', 'translated': '...', 'confidence': 0.95}]
        """
        Translate a batch of items using the adapter's model.
        
        Raises:
            AdapterError: On API or parsing failures.
        """
        pass

    def _make_api_call(self, messages: List[Dict], **kwargs) -> Dict:
        """Internal: Handle retries and OpenAI call."""
        # Implement retry logic here
        for attempt in range(3):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    **kwargs
                )
                return response.choices[0].message
            except OpenAIError as e:
                if attempt == 2 or not self._should_retry(e):
                    raise AdapterError(f"Translation failed: {e}")
                time.sleep(2 ** attempt + random.uniform(0, 1))  # Backoff with jitter
        raise AdapterError("Max retries exceeded")
```

Concrete implementations (`adapters/gpt4o.py`, `adapters/gpt5.py`):
- Format messages as JSON-structured prompts for consistent parsing.
- Parse response to extract translations, handling GPT-5's potentially richer outputs (e.g., confidence scores if available).

Factory in `adapters/factory.py`:
```python
def get_adapter(name: str) -> TranslationAdapter:
    mapping = {
        'gpt4o': Gpt4oAdapter('gpt-4o-2024-08-06'),
        'gpt5': Gpt5Adapter('gpt-5-preview'),  # Update with actual model ID
    }
    if name not in mapping:
        raise ValueError(f"Unknown adapter: {name}")
    return mapping[name]
```

### Checklists
#### Implementation Checklist
- [ ] Create `adapters/` directory with base.py, factory.py, gpt4o.py, gpt5.py.
- [ ] Update argparse in both scripts to include `--adapter` (str, default=None).
- [ ] Add env var parsing: `adapter_name = os.getenv('OPENAI_ADAPTER') or args.adapter or 'gpt4o'`.
- [ ] Replace OpenAI calls with `adapter.translate_batch` at identified points.
- [ ] Implement fallback logic in main translation loop.
- [ ] Update logging to include adapter name and fallback events (e.g., `logger.info(f"Using adapter: {adapter_name}")`).
- [ ] Adjust `tools/estimate_cost.py` for GPT-5 pricing (stub if unavailable).

#### Verification Checklist
- [ ] Run `python scripts/translate_pptx_inplace.py --adapter gpt5 sample.pptx` (mock API); verify logs show GPT-5.
- [ ] Set `OPENAI_ADAPTER=gpt4o` and run without flag; confirm default.
- [ ] Simulate error in GPT-5 adapter; ensure fallback to GPT-4o and job completes.
- [ ] Measure perf: Time translation of 10-slide PPTX/PDF; delta <10%.
- [ ] Test rate limit: Mock 429 error; verify retry (3x) then fallback.
- [ ] Cache compatibility: Translations from different adapters don't conflict (key includes adapter name).

## Risks and Mitigations
- **API Rate Limits (High):** GPT-5 may have lower RPM/TPM.  
  *Mitigation:* Built-in retries with backoff; fallback to GPT-4o; monitor via logs. Add circuit breaker if failures >20% (threshold=5 calls, timeout=60s).

- **Model Unavailability/Cost (Medium):** GPT-5 not released or higher cost.  
  *Mitigation:* Stub adapter for testing; update cost estimators; env var for budget caps.

- **API Changes (Medium):** GPT-5 introduces breaking changes in responses.  
  *Mitigation:* Versioned adapters; strict JSON parsing with schema validation; unit tests for response handling.

- **Performance Regression (Low):** Slower inference or larger outputs.  
  *Mitigation:* Benchmark tests; optimize prompts for GPT-5; cap max_tokens.

- **Security (Low):** API key exposure unchanged.  
  *Mitigation:* No new secrets; continue using env vars.

If risks materialize (e.g., high fallback rate), propose phased rollout: optional flag first, default later.

## Test Plan
### Scope
- Unit: Adapter logic, factory, error raising.
- Integration: Full script runs with mocked/real API.
- Performance: Timing benchmarks.
- E2E: Translation quality parity (manual review or BLEU score).

### Tests to Add
1. **tests/test_adapters.py:**
   - `test_factory_selection`: Assert correct adapter instantiated.
   - `test_translate_batch_mock`: Mock OpenAI, verify input/output shapes.
   - `test_fallback_error`: Raise error in GPT-5, confirm switch.
   - `test_retry_logic`: Simulate 429, assert backoff delays.

2. **Update Existing:**
   - `tests/test_translate_pptx.py`: Parametrize with adapter flag; assert translated output.
   - `tests/test_translate_pdf.py`: Similar, test PDF-specific batches.
   - Add perf test: `pytest --benchmark` for timing.

3. **Manual/Integration:**
   - Run with real API on small sample: Verify GPT-5 quality > GPT-4o (subjective review).
   - Smoke test: `make test-all` passes post-integration.
   - Load test: 5 concurrent jobs; no >20% failure rate.

### Tools
- `pytest` with `pytest-mock` for API mocking.
- `pytest-benchmark` for perf.
- Coverage >90% for new code.

Run tests in CI via updated `tests/run_integration_tests.py`.

## Next Steps for Implementer
1. Read current OpenAI call sites in scripts (use `grep -r "openai\." scripts/`).
2. Implement adapters as specced.
3. Integrate minimally; commit small diffs.
4. Run checklists; address issues.
5. Propose PR with test results.

This ADR provides a clear, testable path to GPT-5 integration without ambiguity.