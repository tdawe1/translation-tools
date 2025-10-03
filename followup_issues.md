# Follow-up Issues for PDF Translation Pipeline

## 1. PDF Back-Projector Test Issues

### Problem
- `test_extract_text_blocks` and `test_full_processing_workflow` in `tests/test_apply_pdf_translation.py` fail because they try to open non-existent PDF files
- Tests are not properly mocking the PyMuPDF `fitz.open()` function

### Solution
- Properly mock `fitz.open()` to return mock document objects instead of trying to open real files
- Ensure all file system access is mocked in tests

## 2. Estimate Cost Test Issues

### Problem
- `test_compute_requests_various[100--5-1]` in `tests/test_estimate_cost.py` fails because the function doesn't properly clamp negative batch sizes
- `test_split_cached_openai_no_cache_or_single_request` has incorrect logic for cache splitting

### Solution
- Fix `compute_requests` function to properly clamp batch sizes to >= 1
- Review and fix the cache splitting logic in `split_cached_openai`

## 3. PDF Extractor Test Issues

### Problem
- `test_block_type_classification` fails because table detection logic incorrectly classifies text as "title" instead of "table"
- `test_extractor_initialization` fails because `use_fallback` is not being initialized correctly
- `test_extract_with_mock_pdf` fails because it tries to access a non-existent file

### Solution
- Fix the block type classification logic in `PDFExtractor._classify_block_type`
- Correct the initialization of `use_fallback` parameter
- Properly mock file system access in the extraction test

## 4. PDF Integration Test Issues

### Problem
- Multiple tests fail because statistics are not being updated correctly in the orchestrator
- `test_audit_report_generation` fails because `save_report_json` function doesn't exist in `translate_pdf.py`

### Solution
- Fix statistics updating logic in `PDFTranslationOrchestrator.translate_pdf`
- Add the missing `save_report_json` function or properly mock it in tests

## 5. Style Checker Test Issues

### Problem
- `test_prompt_includes_style_guide_and_sections` and `test_prompt_without_optional_sections` fail due to incorrect function signatures for `create_style_checker_prompt`
- `test_mixed_verb_forms_are_flagged` fails because the parallelism detection logic is not working correctly

### Solution
- Fix function signatures in `style_checker.py` to match test expectations
- Review and fix the parallelism detection logic in `analyze_parallelism`

## Priority Recommendations

1. **High Priority**: PDF Back-Projector and PDF Extractor test fixes - These are core components of the PDF translation pipeline
2. **Medium Priority**: Estimate Cost and Integration test fixes - These affect the accuracy of cost estimation and overall pipeline functionality
3. **Low Priority**: Style Checker test fixes - These are less critical for the core translation functionality