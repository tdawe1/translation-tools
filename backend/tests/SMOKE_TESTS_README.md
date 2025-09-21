# Backend API Smoke Tests

This directory contains comprehensive smoke tests for the Translation Pipeline Backend API. These tests validate that all core API functionality works correctly end-to-end.

## Test Files

### 1. `test_smoke_comprehensive.py`
The main comprehensive smoke test suite that covers:
- **User Authentication**
  - User registration (`POST /api/auth/register`)
  - User login (`POST /api/auth/login`)
  - Token refresh (`POST /api/auth/refresh`)
  - User logout (`POST /api/auth/logout`)
  - Get current user (`GET /api/auth/me`)

- **Translation Jobs**
  - Create PPTX translation job (`POST /api/translate`)
  - Create PDF translation job (`POST /api/translate`)
  - List jobs with filtering (`GET /api/jobs`)
  - Get job details (`GET /api/jobs/{job_id}`)
  - Cancel job (`POST /api/jobs/{job_id}/cancel`)
  - Retry failed job (`POST /api/jobs/{job_id}/retry`)
  - Delete job (`DELETE /api/jobs/{job_id}`)

- **Job Management**
  - Job search with filters (`POST /api/jobs/search`)
  - Job statistics (`GET /api/jobs/statistics`)
  - Queue status (`GET /api/jobs/queue`)
  - Job logs (`GET /api/jobs/{job_id}/logs`)
  - Bulk operations (`POST /api/jobs/bulk/cancel`, `POST /api/jobs/bulk/retry`)
  - Job export (`GET /api/jobs/export`)

- **Utility Endpoints**
  - Get translation models (`GET /api/translate/models`)
  - Get supported formats (`GET /api/translate/formats`)
  - Health check (`GET /health`)

- **Error Handling**
  - Unauthorized access
  - Invalid tokens
  - Invalid job IDs
  - Malformed requests
  - User isolation (users can't access other users' jobs)

### 2. `test_smoke_simple.py`
Simplified smoke tests covering the basic workflow:
- User registration and login
- Job creation and status checking
- Basic error scenarios

### 3. `test_main.py`
Basic API tests including:
- Health check
- Authentication
- Protected endpoints

## Running the Tests

### Using the Test Runner (Recommended)

The `run_smoke_tests.py` script provides an easy way to run the smoke tests:

```bash
# Run all smoke tests
python run_smoke_tests.py

# Run only simple tests
python run_smoke_tests.py --simple

# Run comprehensive tests only
python run_smoke_tests.py --comprehensive
```

### Using pytest Directly

```bash
# Run all smoke tests
python -m pytest tests/test_smoke_comprehensive.py tests/test_smoke_simple.py -v

# Run specific test file
python -m pytest tests/test_smoke_comprehensive.py -v

# Run specific test class
python -m pytest tests/test_smoke_comprehensive.py::TestAuthenticationEndpoints -v

# Run specific test method
python -m pytest tests/test_smoke_comprehensive.py::TestAuthenticationEndpoints::test_user_registration -v
```

## Test Environment

The tests use a dedicated test environment with:
- In-memory SQLite database
- Test-specific upload/output directories
- Mocked OpenAI API calls (no real API costs)
- Test-specific configuration values

### Environment Variables

The tests automatically configure the following environment variables:
- `DEBUG`: `true`
- `SECRET_KEY`: Test-specific key
- `OPENAI_API_KEY`: Mock key for testing
- `DATABASE_URL`: `sqlite:///:memory:`
- `UPLOAD_DIR`: `test_uploads`
- `OUTPUT_DIR`: `test_outputs`
- `ACCESS_TOKEN_EXPIRE_MINUTES`: `5` (test override)
- `REFRESH_TOKEN_EXPIRE_DAYS`: `1` (test override)

## Test Fixtures

The tests use several fixtures to set up the test environment:

- `client`: FastAPI TestClient with test configuration
- `test_db`: In-memory SQLite database session
- `auth_headers`: Authentication headers for protected endpoints
- `sample_pptx_file`: Creates a minimal PPTX file with Japanese text
- `sample_pdf_file`: Creates a minimal PDF file
- `mock_translation_service`: Mocks the translation service to avoid API calls

## Test Data

Tests create realistic test data:
- PPTX files with Japanese text content
- PDF files with basic structure
- User accounts with proper authentication
- Translation jobs with various parameters

## Continuous Integration

These smoke tests are designed to run in CI/CD pipelines to verify API functionality before deployment. They:
- Use no external dependencies (except the test database)
- Complete quickly (typically under 30 seconds)
- Provide clear pass/fail results
- Test all critical API endpoints

## Troubleshooting

### Common Issues

1. **Import Errors**: Ensure you're running from the backend directory
   ```bash
   cd backend
   python -m pytest tests/
   ```

2. **Database Errors**: The tests use an in-memory database that's created fresh for each test

3. **Permission Errors**: Test directories are automatically created and cleaned up

4. **Missing Dependencies**: Install test dependencies
   ```bash
   pip install pytest pytest-asyncio httpx
   ```

### Running Tests Manually

To run tests without pytest, you can import and run the test classes directly:

```python
from tests.test_smoke_comprehensive import TestAuthenticationEndpoints
import unittest

# Create test suite
suite = unittest.TestLoader().loadTestsFromTestCase(TestAuthenticationEndpoints)
runner = unittest.TextTestRunner(verbosity=2)
result = runner.run(suite)
```

## Contributing

When adding new smoke tests:
1. Follow the existing test structure and naming conventions
2. Test both success and error scenarios
3. Use the provided fixtures where appropriate
4. Mock external services to avoid dependencies
5. Ensure tests run quickly and reliably
6. Add documentation for complex test scenarios