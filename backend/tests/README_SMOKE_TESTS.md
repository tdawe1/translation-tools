# Smoke Tests for Translation Pipeline Backend

This directory contains comprehensive smoke tests for the authentication and job submission workflow. These tests verify the entire end-to-end functionality of the translation pipeline API.

## Test Categories

### 1. Authentication Flow (`TestAuthenticationFlow`)
- User registration with validation
- Login with JWT token generation
- Token refresh functionality
- Protected endpoint access
- Error handling for invalid credentials

### 2. Job Submission Workflow (`TestJobSubmissionWorkflow`)
- PPTX file upload and translation
- PDF file upload and translation
- Job status tracking
- File download after completion
- Job history and pagination
- Job search functionality

### 3. Error Scenarios (`TestErrorScenarios`)
- Invalid file types
- Missing authentication
- Job not found errors
- Expired tokens
- Rate limiting
- Large file handling

### 4. Integration Points (`TestIntegrationPoints`)
- API endpoint structure verification
- CORS headers
- Error response formats
- Models and formats endpoints
- Job statistics

### 5. Performance Tests (`TestPerformance`)
- Concurrent job creation
- Job list performance with large datasets
- (Marked with `@pytest.mark.slow`)

## Running the Tests

### Quick Start
```bash
# Run all smoke tests
./run_smoke_tests.sh

# Run specific category
./run_smoke_tests.sh --auth      # Authentication only
./run_smoke_tests.sh --jobs      # Job workflow only
./run_smoke_tests.sh --errors    # Error scenarios only
./run_smoke_tests.sh --integration  # Integration tests only
```

### Using pytest directly
```bash
# Run all smoke tests
pytest tests/test_smoke_workflow.py -v

# Run specific test class
pytest tests/test_smoke_workflow.py::TestAuthenticationFlow -v

# Run with coverage
pytest tests/test_smoke_workflow.py --cov=app --cov-report=term-missing

# Run performance tests
pytest tests/test_smoke_workflow.py -m slow
```

## Test Dependencies

The tests require the following additional packages:
```bash
pip install pytest pytest-cov pytest-mock python-pptx reportlab
```

## Test Fixtures

The `test_workflow_fixtures.py` file provides comprehensive fixtures for:
- Sample PPTX files with Japanese text
- Sample PDF files with Japanese text
- Mock translation responses
- Mock job database entries
- Test data generators

## Key Test Features

### 1. Realistic Test Data
- Japanese text content
- Proper PPTX/PDF file structure
- Various formatting scenarios

### 2. Comprehensive API Coverage
- All `/api/auth/*` endpoints
- All `/api/translate/*` endpoints
- All `/api/jobs/*` endpoints
- SSE endpoint verification

### 3. Error Handling
- Validates proper HTTP status codes
- Checks error message formats
- Verifies security measures

### 4. Performance Considerations
- Tests for concurrent access
- Large file handling
- Response time validation

## Test Output Example

```
🔥 Running Smoke Tests for Translation Pipeline Backend
======================================================

🚀 Starting smoke test suite...

📋 Running all smoke tests...
----------------------------------------
============================= test session starts ==============================
...
collected 42 items

tests/test_smoke_workflow.py::TestAuthenticationFlow::test_user_registration_success PASSED
tests/test_smoke_workflow.py::TestAuthenticationFlow::test_user_registration_duplicate_email PASSED
...
tests/test_smoke_workflow.py::TestJobSubmissionWorkflow::test_pptx_translation_job_complete_workflow PASSED
...
tests/test_smoke_workflow.py::TestErrorScenarios::test_invalid_file_type PASSED
...
tests/test_smoke_workflow.py::TestIntegrationPoints::test_api_endpoints_structure PASSED
...

========================= 42 passed in 15.23s ==========================

✅ all smoke tests passed
```

## Continuous Integration

These smoke tests are designed to run quickly in CI environments:
- Average execution time: < 30 seconds
- No external API dependencies (all mocked)
- In-memory database for fast setup/teardown
- Minimal resource requirements

## Adding New Tests

When adding new smoke tests:

1. Follow the existing class structure
2. Use the provided fixtures for test data
3. Mock external dependencies
4. Include both success and error cases
5. Add appropriate markers (smoke, slow, etc.)

Example:
```python
@pytest.mark.smoke
def test_new_feature(client, auth_headers):
    """Test new feature smoke test"""
    # Arrange
    test_data = create_test_data()

    # Act
    response = client.post("/api/new-feature", json=test_data, headers=auth_headers)

    # Assert
    assert response.status_code == 200
    assert response.json()["success"] is True
```

## Debugging Failed Tests

For debugging failed tests:

1. Run with verbose output: `pytest -v`
2. Use `--tb=long` for full tracebacks
3. Run single test: `pytest tests/test_smoke_workflow.py::TestClass::test_method`
4. Use breakpoints with `pytest --pdb`

## Test Environment

The tests use a controlled environment:
- SQLite in-memory database
- Mocked OpenAI API responses
- Temporary upload/output directories
- Test-specific configuration values

See `conftest.py` for complete environment setup.