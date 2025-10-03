# CI/CD Integration for PDF Translation Tests

This document describes the integration testing setup for the PDF translation system.

## Test Suite Overview

### Quality Metrics Tests (`tests/test_pdf_quality_metrics.py`)
- **Residual Japanese Threshold**: Enforces ≤2% residual Japanese content
- **Layout Integrity**: Requires ≥95% layout preservation
- **Cache Efficiency**: Validates ≥80% cache hit rate
- **Translation Completeness**: Ensures ≥85% content translation
- **Overall Quality**: Maintains ≥80% overall quality score
- **Performance Benchmarks**: Validates processing time ≤5 seconds for test data
- **Error Handling**: Tests file size limits and edge cases

### Integration Tests (`tests/test_pdf_integration.py`)
- **End-to-End Translation**: Tests complete pipeline from extraction to audit
- **Cache Effectiveness**: Validates cache performance and hit rates
- **Error Scenarios**: Handles corrupted files, empty files, missing translations
- **Layout Handling**: Supports multi-column and mixed content layouts
- **Output Generation**: Validates CSV and audit report generation

### Sample Data (`tests/data/`)
- **simple_japanese.txt**: Basic Japanese business document
- **multi_column_japanese.txt**: Multi-column layout with structured content
- **mixed_content_japanese.txt**: Mixed content with tables, numbers, and formatting

## Quality Metric Thresholds

| Metric | Threshold | Purpose |
|--------|-----------|---------|
| Residual Japanese | ≤2% | Ensure complete translation |
| Layout Integrity | ≥95% | Preserve original formatting |
| Cache Hit Rate | ≥80% | Optimize performance and cost |
| Translation Completeness | ≥85% | Ensure full content translation |
| Overall Quality | ≥80% | Maintain overall translation quality |
| Processing Time | ≤5s | Performance benchmark for test data |

## Running Tests

### Local Development
```bash
# Run all quality metrics tests
python -m pytest tests/test_pdf_quality_metrics.py -v

# Run specific test categories
python -m pytest tests/test_pdf_quality_metrics.py::TestPDFQualityMetrics -v
python -m pytest tests/test_pdf_quality_metrics.py::TestPDFIntegrationValidation -v

# Run integration tests (with mocked dependencies)
python -m pytest tests/test_pdf_integration.py::TestQualityMetricsEnforcement -v
```

### CI/CD Pipeline
```yaml
# Example GitHub Actions workflow
name: PDF Translation Tests

on:
  push:
    paths: ['tests/**', 'scripts/**']
  pull_request:
    paths: ['tests/**', 'scripts/**']

jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      
      - name: Set up Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.11'
      
      - name: Install dependencies
        run: |
          pip install pytest
          pip install -r requirements_pdf.txt
      
      - name: Run quality metrics tests
        run: python -m pytest tests/test_pdf_quality_metrics.py -v
      
      - name: Run integration tests
        run: python -m pytest tests/test_pdf_integration.py::TestQualityMetricsEnforcement -v
      
      - name: Generate test report
        run: |
          python -m pytest tests/ --tb=short --junitxml=test-results.xml
```

## Test Data Management

### Sample Data Generation
```bash
# Generate sample PDF/text files
python tests/create_sample_pdfs.py
```

### Cache Management
```bash
# Clean test cache
rm -f translation_cache.json

# Backup test cache
cp translation_cache.json translation_cache.json.backup
```

### Test Environment Setup
```bash
# Set up test environment
export OPENAI_API_KEY=test_key
export PYTHONPATH=/path/to/project/root

# Install test dependencies
pip install pytest pytest-cov pytest-mock
```

## Quality Enforcement

### Pre-commit Hooks
```yaml
# .pre-commit-config.yaml
repos:
  - repo: local
    hooks:
      - id: pdf-quality-tests
        name: Run PDF quality tests
        entry: python -m pytest tests/test_pdf_quality_metrics.py
        language: python
        pass_filenames: false
        always_run: true
```

### Continuous Integration
The CI pipeline enforces:
- All quality metric tests must pass
- Integration tests must pass with mocked dependencies
- Code coverage requirements (≥80%)
- Performance benchmarks must be met

### Release Criteria
Before releasing PDF translation features:
1. All quality metrics tests pass
2. Integration tests pass with real PDF files
3. Performance benchmarks are met
4. Security scans pass
5. Documentation is updated

## Monitoring and Alerting

### Test Metrics Collection
- Test execution time
- Pass/fail rates
- Quality metric scores
- Performance benchmarks

### Alerting Conditions
- Quality metrics below threshold
- Test failures in CI
- Performance degradation
- Cache efficiency drops

## Test Maintenance

### Adding New Quality Metrics
1. Update `test_pdf_quality_metrics.py` with new test cases
2. Define clear thresholds and pass/fail criteria
3. Update documentation
4. Add to CI pipeline

### Updating Sample Data
1. Regenerate sample files with `create_sample_pdfs.py`
2. Validate Japanese content quality
3. Ensure files cover edge cases
4. Update documentation

### Test Environment
- Python 3.11+
- Required PDF libraries (PyMuPDF, pypdf, etc.)
- Test fixtures and mock data
- Sample Japanese PDF/text files

## Troubleshooting

### Common Issues
1. **Import Errors**: Ensure all dependencies are installed
2. **Missing Test Data**: Run `create_sample_pdfs.py` to generate samples
3. **Mock Failures**: Check mock setup in integration tests
4. **Performance Issues**: Verify system meets requirements

### Debug Commands
```bash
# Debug individual test
python -m pytest tests/test_pdf_quality_metrics.py::TestPDFQualityMetrics::test_residual_japanese_threshold -v -s

# Check test coverage
python -m pytest tests/ --cov=scripts --cov-report=html

# Run tests with verbose output
python -m pytest tests/ -v --tb=long
```

## Future Enhancements

### Planned Improvements
- Real PDF file testing (not just mocked)
- Integration with actual OpenAI API for end-to-end testing
- Performance load testing
- Cross-platform compatibility testing
- Accessibility testing for translated documents

### Test Data Expansion
- More complex PDF layouts
- Different Japanese content types
- Edge case documents
- Large file handling
- Corrupted file recovery