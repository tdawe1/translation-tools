# ADR 003: Extend Document Abstraction Layer to Support DOCX and XLSX Formats

## Status
Proposed

## Context

The current translation pipeline supports PPTX and PDF formats through specialized extractors and appliers in `scripts/` (e.g., `pptx_extractor.py`, `translate_pdf.py`). These handle text extraction, batch translation via the shared `core/models/` batching and caching system, and text replacement while preserving layout/formatting. To extend support to DOCX (Microsoft Word) and XLSX (Microsoft Excel) documents without duplicating logic, we need an abstraction layer that reuses the existing translation engine (>80% code reuse). This involves defining interfaces for extraction (text with layout context) and application (replacing translated text), integrating with the Phase 1 `batch_translate` adapter, and ensuring backward compatibility for PPTX/PDF.

Key inputs:
- Existing `core/` modules: `models/` for batching/caching (e.g., `batch_translate` function handles JSON-structured inputs/outputs).
- Dependencies: `python-docx` (for DOCX) and `openpyxl` (for XLSX), pinned to minimal versions to avoid bloat.
- Non-functional: Low capacity (no heavy libs like LibreOffice); performance parity with existing formats; no breaking changes to PPTX/PDF pipelines.

Success criteria:
- >80% code reuse in translation engine (measured by shared batching/caching logic).
- Tests validate sample DOCX/XLSX extraction and full translation cycles.
- No performance regression (e.g., extraction/apply times within 10% of PPTX equivalents on sample docs).

## Decision

We will introduce a document abstraction layer in `core/document/` with base classes for extraction and application:
- `DocumentExtractor`: Abstract base for format-specific text extraction, returning structured data (e.g., list of dicts with `id`, `text`, `layout` keys) compatible with `batch_translate`.
- `DocumentApplier`: Abstract base for applying translations back to the document, using the same structured data format.
- Format-specific implementations: `PptxExtractor`, `PdfExtractor` (refactored from existing scripts), `DocxExtractor`, `XlsxExtractor`.
- Integration: New CLI/entrypoints (e.g., `translate_docx.py`, `translate_xlsx.py`) will instantiate the appropriate extractor/applier, feed into `core/models/batch_translate`, and handle caching.

This design hooks into the existing `batch_translate` adapter from Phase 1 by standardizing the input/output schema (JSON-serializable batches of text segments with metadata). Existing PPTX/PDF scripts will be refactored to use these base classes without altering their public APIs, ensuring no breakage.

Dependencies:
- Pin `python-docx==1.1.2` and `openpyxl==3.1.2` in `requirements.txt` (minimal, lightweight libs; no additional heavy deps).

Phased rollout:
1. Define interfaces and refactor PPTX/PDF to use them (internal, no user impact).
2. Implement DOCX/XLSX extractors/appliers.
3. Add tests and verify no perf regression.

## Alternatives Considered

1. **Monolithic Script per Format** (Status Quo Extension):
   - Pros: Simple, no abstraction overhead.
   - Cons: Duplicates batching/caching logic (<50% reuse); harder to maintain/extend (e.g., future formats); violates DRY principle.
   - Why not: Fails reuse goal (>80%); increases tech debt.

2. **Unified Parser with Heavy Lib (e.g., LibreOffice UNO)**:
   - Pros: Single interface for all Office formats.
   - Cons: High capacity overhead (requires Java bridge, ~500MB+); platform-dependent; violates low-capacity constraint.
   - Why not: Exceeds minimal deps; potential perf regression on low-resource envs.

3. **External Service (e.g., Google Docs API)**:
   - Pros: Offloads processing; handles multiple formats.
   - Cons: Introduces external dependency/latency; auth complexity; not offline-capable; breaks self-contained pipeline.
   - Why not: Misaligns with local, low-latency goals.

4. **Minimal Abstraction with Protocol (Dataclasses/Protocols)**:
   - Pros: Lighter than ABCs; type-safe.
   - Cons: Less explicit for inheritance-based refactoring of existing code; harder to enforce structure.
   - Why not: ABCs better for clear extension points and testing mocks.

Chosen: ABC-based abstraction balances reuse, maintainability, and constraints.

## Consequences

### Positive
- **High Reuse**: Core translation logic (batching, caching, prompting) shared across formats; extraction/apply become thin adapters (~20% new code).
- **Extensibility**: Easy to add future formats (e.g., ODT) by implementing base classes.
- **Maintainability**: Centralized interfaces reduce duplication; standardized schema simplifies integration.
- **Testability**: Base classes enable mocking for unit tests; E2E tests reuse existing fixtures.
- **Performance**: Lightweight deps ensure no regression; structured data minimizes serialization overhead.

### Negative
- **Refactoring Overhead**: Initial effort to wrap PPTX/PDF (~1-2 days); risk of introducing bugs (mitigated by tests).
- **Format-Specific Quirks**: DOCX tables/headers, XLSX formulas may require custom handling (mitigated by fallbacks to plain text).
- **Dependency Bloat**: Adding two libs increases install size (~10MB); pinned versions limit this.

### Neutral
- **Schema Evolution**: Fixed input/output format may constrain future changes (e.g., richer layout metadata); can evolve via versioning in `core/models/`.

## Interfaces and Stubs

### Core Schema (Shared Across Formats)
All extractors/appliers use this JSON-compatible structure for segments:
```python
from typing import List, Dict, Any, Optional
from dataclasses import dataclass

@dataclass
class TranslationSegment:
    id: str  # Unique identifier (e.g., "slide_1_para_2")
    original_text: str  # Japanese text to translate
    translated_text: Optional[str] = None  # Filled post-translation
    metadata: Dict[str, Any] = None  # Format-specific (e.g., {"font_size": 12, "position": (x,y)} for layout)
    context: Optional[str] = None  # Surrounding text for better translation (e.g., slide title)
```

Batch input to `batch_translate`: `List[TranslationSegment]`.

### Base Classes (in `core/document/base.py`)

```python
from abc import ABC, abstractmethod
from typing import List
from .schema import TranslationSegment  # Assume schema.py defines above

class DocumentExtractor(ABC):
    """Extracts translatable segments from a document while preserving layout context."""
    
    @abstractmethod
    def extract(self, file_path: str) -> List[TranslationSegment]:
        """Extract segments from the document at file_path.
        
        Returns:
            List of segments ready for batch translation.
        Raises:
            ValueError: If file format unsupported or extraction fails.
        """
        pass
    
    @abstractmethod
    def get_document_type(self) -> str:
        """Returns the supported document type (e.g., 'pptx', 'docx')."""
        pass

class DocumentApplier(ABC):
    """Applies translated segments back to the original document."""
    
    @abstractmethod
    def apply(self, file_path: str, segments: List[TranslationSegment], output_path: str) -> None:
        """Replace original_text with translated_text in segments.
        
        Preserves layout/formatting as much as possible.
        Args:
            file_path: Path to original document.
            segments: List with translated_text populated.
            output_path: Where to save the translated document.
        Raises:
            ValueError: If application fails (e.g., text too long for layout).
        """
        pass
    
    @abstractmethod
    def get_document_type(self) -> str:
        """Returns the supported document type."""
        pass
```

### Example Stubs for New Formats (in `core/document/docx.py` and `xlsx.py`)

For DOCX (using python-docx):
```python
from docx import Document
from .base import DocumentExtractor, DocumentApplier, TranslationSegment

class DocxExtractor(DocumentExtractor):
    def get_document_type(self) -> str:
        return 'docx'
    
    def extract(self, file_path: str) -> List[TranslationSegment]:
        doc = Document(file_path)
        segments = []
        for i, para in enumerate(doc.paragraphs):
            if self._is_japanese(para.text):  # Placeholder for lang detection
                segments.append(TranslationSegment(
                    id=f"para_{i}",
                    original_text=para.text.strip(),
                    metadata={"style": para.style.name, "level": para._element.xpath('.//w:pPr/w:pStyle/@w:val')}
                ))
        # Handle tables similarly...
        return segments

class DocxApplier(DocumentApplier):
    def get_document_type(self) -> str:
        return 'docx'
    
    def apply(self, file_path: str, segments: List[TranslationSegment], output_path: str) -> None:
        doc = Document(file_path)
        # Map segments back to paragraphs/tables and replace text
        # Preserve runs/formatting where possible
        doc.save(output_path)
```

For XLSX (similar structure, focusing on cell text; skip formulas):
```python
from openpyxl import load_workbook
# ... similar stubs, extracting from ws.cell(row, col).value if Japanese
```

### Integration with batch_translate
In new scripts (e.g., `scripts/translate_docx.py`):
```python
from core.document import DocxExtractor, DocxApplier
from core.models import batch_translate  # Phase 1 adapter

def translate_docx(input_path: str, output_path: str, model: str):
    extractor = DocxExtractor()
    segments = extractor.extract(input_path)
    translated_segments = batch_translate(segments, model=model)  # Hooks into shared engine
    applier = DocxApplier()
    applier.apply(input_path, translated_segments, output_path)
```

Refactor existing PPTX/PDF similarly, e.g., `PptxExtractor` wrapping `pptx_extractor.py` logic.

## Implementation Checklist

### Phase 1: Abstraction and Refactor (1-2 days)
- [ ] Define `core/document/` with `base.py` and `schema.py`.
- [ ] Implement `PptxExtractor` and `PptxApplier` (wrap existing logic; >90% reuse).
- [ ] Implement `PdfExtractor` and `PdfApplier` (similar wrap).
- [ ] Update `scripts/translate_pptx_inplace.py` and `translate_pdf.py` to use new classes (no API changes).
- [ ] Pin deps in `requirements.txt`: `python-docx==1.1.2`, `openpyxl==3.1.2`.
- [ ] Verify PPTX/PDF pipelines unchanged (run existing tests).

### Phase 2: New Formats (2-3 days)
- [ ] Implement `DocxExtractor/Applier` (handle paragraphs, tables, headers; fallback to plain text for complex elements like images/footnotes).
- [ ] Implement `XlsxExtractor/Applier` (cell-by-cell; skip non-text/formulas; preserve cell styles).
- [ ] Add CLI entrypoints: `translate_docx.py`, `translate_xlsx.py` (mirror existing args: `--model`, `--offline`).
- [ ] Integrate caching: Use segment `id` as cache key in `core/models/`.
- [ ] Measure code reuse: Diff new vs. existing; target >80% shared (batching/caching).

### Phase 3: Verification (1 day)
- [ ] Run perf benchmarks: Time extraction/apply on sample docs (DOCX/XLSX vs. PPTX); ensure <10% regression.
- [ ] Audit for breakage: Full E2E on PPTX/PDF samples.
- [ ] Static analysis: Type check, lint new code.

## Verification Checklist
- [ ] Unit: Extraction yields valid segments (mock files; assert len(segments) >0, all have `id`/`text`).
- [ ] Unit: Application replaces text without corrupting doc (compare output diffs).
- [ ] Integration: Full translate cycle on samples (input → extract → batch_translate(mock) → apply → validate output has English text).
- [ ] Perf: Benchmark suite shows no regression (e.g., via `time` or `pytest-benchmark`).
- [ ] Compatibility: Existing scripts pass all tests unchanged.
- [ ] Deps: `pip check` passes; no unpinned versions.

## Test Plan

### Unit Tests (in `tests/core/test_document.py`)
- Mock file paths; use `unittest.mock` for deps (e.g., patch `docx.Document`).
- Test `extract`: Input sample DOCX bytes → output segments with metadata.
- Test `apply`: Input segments with translations → output doc with replaced text (validate via re-extraction).
- Coverage: >90% for base/new classes; include edge cases (empty docs, non-JP text).

### Integration Tests (in `tests/integration/test_docx_translation.py`)
- Use real sample files (e.g., `tests/samples/sample.docx` with JP text).
- Mock `batch_translate` to return fixed translations (e.g., {"こんにちは": "Hello"}).
- Assert: Post-translation doc contains English; layout preserved (e.g., para count unchanged).
- Offline mode: Test cache hit (stub cache dict).

### E2E Tests (in `tests/e2e/test_new_formats.py`)
- Full pipeline: `python scripts/translate_docx.py sample.docx output.docx --model mock`.
- Use mocked OpenAI (via `pytest` fixtures from existing tests).
- Validate: Output opens in LibreOffice/Excel; text translated; no crashes.
- Perf: Run with `pytest --benchmark-compare`; assert times < threshold.

### Mocking Strategy
- For `python-docx/openpyxl`: Patch constructors to return mock objects with iterable elements.
- For `batch_translate`: Fixture returning segments with `translated_text` swapped.
- Samples: Generate via `tests/create_sample_docs.py` (extend existing for DOCX/XLSX).

## Risks and Mitigations

1. **Format Quirks (High)**: DOCX tracked changes/fields, XLSX merged cells/formulas may break extraction/apply.
   - Mitigation: Fallback to plain text extraction (lose layout); log warnings; phased support (start with simple docs). Test on diverse samples (e.g., tables, lists).

2. **Performance Regression (Medium)**: New libs slower than ZIP-based PPTX.
   - Mitigation: Benchmark early; optimize (e.g., lazy loading in openpyxl); cap doc size in CLI (e.g., --max-pages). If >10% slower, profile and fallback to external tools only if needed.

3. **Dependency Conflicts (Low)**: Pinned versions clash with existing (e.g., lxml in python-docx).
   - Mitigation: `pip check` in CI; use virtualenv isolation. If conflict, vendor or find alternatives (but avoid).

4. **Schema Rigidity (Low)**: Fixed segment format limits rich features (e.g., XLSX hyperlinks).
   - Mitigation: Version schema (e.g., v1 plain, v2 rich); optional `metadata` fields. Prototype with samples before finalizing.

5. **Testing Gaps (Medium)**: Hard to mock complex docs.
   - Mitigation: Generate programmatic samples; use real open-source JP docs for E2E. Require 100% test pass rate before merge.

Escalation: If quirks block >50% use cases, prototype with stakeholder review or defer to Phase 4 with richer libs.
