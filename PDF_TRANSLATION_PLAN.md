# PDF Translation Implementation Plan

## Executive Summary

This plan outlines the implementation of PDF translation capability for the existing Japanese-to-English translation pipeline. The solution will preserve original formatting while handling text expansion and reusing existing infrastructure.

## 1. Architecture Overview

### 1.1 System Integration
```
Existing Pipeline: PPTX → XML Extract → Batch Translate → Replace → Style → Audit
New PDF Pipeline:   PDF → Text Extract → Batch Translate → Layout Adjust → Style → Audit
```

### 1.2 Core Components
1. **PDF Extractor** - Text and layout extraction
2. **PDF Back-Projector** - Text replacement with layout preservation  
3. **Layout Engine** - Handle text expansion and formatting
4. **PDF Audit Tools** - Quality assessment specific to PDFs

## 2. Technical Approach

### 2.1 Library Selection
- **Primary**: `PyMuPDF (fitz)` - Best for text replacement and layout preservation
- **Secondary**: `pdfplumber` - Fallback for complex text extraction
- **OCR**: `pytesseract` - For image-based PDFs (future enhancement)

### 2.2 Processing Strategy
1. **Extract**: Preserve text position, font, and formatting information
2. **Translate**: Use existing batch translation system
3. **Adjust**: Handle text expansion with font scaling and layout optimization
4. **Replace**: Insert translated text while maintaining original structure
5. **Audit**: Verify translation quality and layout integrity

## 3. Implementation Phases

### Phase 1: MVP Text-Only Translation (1-2 weeks)
- Basic PDF text extraction
- Integration with existing translation pipeline
- Simple text output (bilingual format)
- Basic audit capabilities

### Phase 2: Layout Preservation (2-3 weeks)  
- Text position mapping
- Font scaling for text expansion
- Basic layout preservation
- PDF output generation

### Phase 3: Advanced Features (1-2 weeks)
- Multi-column layout support
- Table preservation
- Header/footer handling
- Image text extraction (OCR)

### Phase 4: Full Integration (1 week)
- CLI integration
- Makefile updates
- CI/CD pipeline updates
- Documentation updates

## 4. Component Breakdown

### 4.1 PDF Extractor (`scripts/extract_pdf.py`)
```python
class PDFExtractor:
    def extract_text_blocks(self, pdf_path: str) -> List[TextBlock]
    def extract_layout_info(self, pdf_path: str) -> LayoutInfo
    def get_page_dimensions(self, pdf_path: str) -> PageDimensions
```

### 4.2 PDF Back-Projector (`scripts/apply_pdf_translation.py`)
```python
class PDFBackProjector:
    def replace_text(self, pdf_path: str, translations: Dict[str, str]) -> str
    def adjust_font_sizes(self, pdf_path: str, expansion_factor: float) -> None
    def preserve_layout(self, pdf_path: str, layout_info: LayoutInfo) -> None
```

### 4.3 Layout Engine (`scripts/pdf_layout_engine.py`)
```python
class PDFLayoutEngine:
    def calculate_expansion_factor(self, jp_text: str, en_text: str) -> float
    def optimize_font_sizes(self, text_blocks: List[TextBlock]) -> None
    def handle_overflow(self, text_blocks: List[TextBlock]) -> None
```

### 4.4 PDF Audit Tools (`scripts/audit_pdf.py`)
```python
class PDFAuditor:
    def count_residual_jp(self, pdf_path: str) -> int
    def check_layout_integrity(self, original: str, translated: str) -> bool
    def generate_audit_report(self, pdf_path: str) -> AuditReport
```

## 5. Integration Points

### 5.1 Reuse Existing Components
- `translate_batch()` - Core translation logic
- `load_cache()` / `save_cache()` - Translation caching
- `glossary.json` - Terminology consistency
- `style_checker.py` - Style validation
- `estimate_cost.py` - Cost calculation

### 5.2 New CLI Commands
```bash
# Basic PDF translation
python scripts/translate_pdf.py --in input.pdf --out output_en.pdf

# With layout preservation
python scripts/translate_pdf.py --in input.pdf --out output_en.pdf --preserve-layout

# Cost estimation for PDF
python tools/estimate_cost_pdf.py input.pdf
```

### 5.3 Makefile Updates
```makefile
# Add to existing Makefile
estimate-pdf:
	@python tools/estimate_cost_pdf.py inputs/sample.pdf

translate-pdf:
	@python scripts/translate_pdf.py --in inputs/sample.pdf --out outputs/sample_en.pdf
```

## 6. Risk Assessment

### 6.1 Technical Risks
- **High**: Complex PDF layouts may not preserve perfectly
- **Medium**: Text expansion may cause overflow in constrained layouts
- **Low**: Integration with existing translation pipeline

### 6.2 Mitigation Strategies
- **Progressive Enhancement**: Start with simple layouts, add complexity
- **Graceful Degradation**: Fall back to text-only for complex PDFs
- **Comprehensive Testing**: Test with various PDF types and layouts

## 7. Testing Strategy

### 7.1 Test Categories
- **Unit Tests**: Individual component functionality
- **Integration Tests**: End-to-end PDF translation
- **Layout Tests**: Formatting preservation verification
- **Performance Tests**: Processing speed and memory usage

### 7.2 Test Data
- Simple text PDFs
- Multi-column layouts
- Tables and forms
- Image-heavy PDFs
- Mixed content types

## 8. Success Criteria

### 8.1 Functional Requirements
- [ ] Extract text from standard PDFs with 95% accuracy
- [ ] Preserve basic layout formatting (position, font size)
- [ ] Handle text expansion with automatic font scaling
- [ ] Integrate with existing caching and glossary systems
- [ ] Generate audit reports for PDF translations

### 8.2 Quality Requirements
- [ ] No loss of original content during translation
- [ ] Layout integrity maintained for 90% of standard PDFs
- [ ] Font scaling prevents text overflow in 95% of cases
- [ ] Residual Japanese character detection < 2%

## 9. Resource Requirements

### 9.1 Development Resources
- **Lead Developer**: 1 person
- **Testing/QA**: 1 person (part-time)
- **Timeline**: 6-8 weeks total

### 9.2 Infrastructure
- Additional Python packages: `PyMuPDF`, `pdfplumber`
- Test PDF samples in various formats
- CI/CD pipeline updates

## 10. Next Steps

1. **Approve Plan**: Review and finalize implementation approach
2. **Create Branch**: Set up feature branch for development
3. **Setup Environment**: Install required PDF libraries
4. **Implement Phase 1**: Build MVP text-only translation
5. **Test and Iterate**: Validate with sample PDFs
6. **Proceed to Phases**: Implement layout preservation and advanced features

---

This plan provides a comprehensive roadmap for adding PDF translation capability while maintaining the existing system's quality and reliability standards.