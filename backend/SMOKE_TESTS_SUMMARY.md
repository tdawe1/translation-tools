# Backend API Smoke Tests - Summary

## Overview

I have created comprehensive smoke tests for the Translation Pipeline Backend API that validate all core functionality works correctly end-to-end.

## Created Files

### 1. `tests/test_smoke_comprehensive.py` (NEW)
A comprehensive test suite covering:
- **Authentication Endpoints** (16 tests)
  - User registration, login, token refresh, logout
  - Get current user info
  - Invalid credentials handling

- **Translation Job Management** (20+ tests)
  - Create PPTX/PDF translation jobs
  - List jobs with filtering and pagination
  - Get job details and logs
  - Cancel, retry, and delete jobs
  - Bulk operations
  - Job statistics and export

- **Protected Endpoints** (4 tests)
  - Verify authentication requirements
  - Invalid token handling

- **Error Scenarios** (8 tests)
  - Invalid job IDs
  - Invalid parameters
  - User isolation (access control)

- **Utility Endpoints** (4 tests)
  - Health check
  - Get models and formats

**Total: 50+ test methods**

### 2. `tests/test_smoke_simple.py` (Existing)
Main smoke test file with 4 test functions:
- `test_complete_translation_workflow`: Full PPTX translation workflow
- `test_pdf_translation_workflow`: PDF translation with page ranges
- `test_error_handling_scenarios`: Various error cases
- `test_bulk_operations`: Bulk job cancellation

### 3. `run_smoke_tests.py` (Updated)
Enhanced test runner script with:
- Support for running simple, comprehensive, or all tests
- Better error handling and output
- Command-line argument support
- Automatic virtual environment detection

### 4. `check_api_health.py` (NEW)
Quick health check script that:
- Tests basic API connectivity
- Checks key endpoints without running full tests
- Provides quick pass/fail feedback
- Useful for CI/CD pipelines

### 5. `tests/SMOKE_TESTS_README.md` (NEW)
Comprehensive documentation covering:
- Test structure and organization
- Running instructions
- Test environment setup
- Troubleshooting guide

## Key Features

### Test Coverage
- All authentication flows (register, login, refresh, logout)
- Complete job lifecycle (create → monitor → cancel/delete)
- Both PPTX and PDF translation workflows
- Error handling and edge cases
- User access control
- Bulk operations
- Statistics and reporting

### Test Environment
- In-memory SQLite database (isolated per test)
- Test-specific upload/output directories
- Mocked OpenAI API calls (no costs)
- Proper fixture management

### Sample Data Generation
- Minimal but valid PPTX files with Japanese text
- Basic PDF files for testing
- Realistic test scenarios

## Usage

### Running All Tests
```bash
# From backend directory
python run_smoke_tests.py

# Or with pytest directly
python -m pytest tests/test_smoke_comprehensive.py -v
```

### Running Specific Tests
```bash
# Only authentication tests
python -m pytest tests/test_smoke_comprehensive.py::TestAuthenticationEndpoints -v

# Only job management tests
python -m pytest tests/test_smoke_comprehensive.py::TestJobManagementEndpoints -v

# Single test method
python -m pytest tests/test_smoke_comprehensive.py::TestJobManagementEndpoints::test_create_pptx_translation_job -v
```

### Quick Health Check
```bash
# Check if API is responding
python check_api_health.py

# Check specific API URL
python check_api_health.py --url http://localhost:8000
```

## Endpoints Tested

### Authentication (6 endpoints)
- `POST /api/auth/register` - User registration
- `POST /api/auth/login` - User login
- `POST /api/auth/refresh` - Token refresh
- `POST /api/auth/logout` - User logout
- `GET /api/auth/me` - Get current user
- `POST /api/auth/api-keys` - Create API key

### Translation (4 endpoints)
- `POST /api/translate` - Create translation job (PPTX/PDF)
- `GET /api/translate/models` - Get available models
- `GET /api/translate/formats` - Get supported formats
- `POST /api/translate/pptx` - PPTX-specific translation

### Jobs (12+ endpoints)
- `GET /api/jobs` - List jobs with filtering/pagination
- `POST /api/jobs/search` - Advanced job search
- `GET /api/jobs/{job_id}` - Get job details
- `GET /api/jobs/{job_id}/logs` - Get job logs
- `POST /api/jobs/{job_id}/cancel` - Cancel job
- `POST /api/jobs/{job_id}/retry` - Retry failed job
- `DELETE /api/jobs/{job_id}` - Delete job
- `GET /api/jobs/statistics` - Job statistics
- `GET /api/jobs/queue` - Queue status
- `GET /api/jobs/export` - Export job data
- `POST /api/jobs/bulk/cancel` - Bulk cancel
- `POST /api/jobs/bulk/retry` - Bulk retry
- `GET /{job_id}/download` - Download result

### Health (1 endpoint)
- `GET /health` - Health check

## Integration Notes

These smoke tests are designed to:
1. Run quickly (typically under 30 seconds)
2. Use no external dependencies (except test database)
3. Provide clear pass/fail results
4. Test all critical API paths
5. Validate the API works end-to-end

They can be integrated into:
- CI/CD pipelines
- Pre-deployment checks
- Development workflow
- Production monitoring

## Dependencies

The tests use the existing test fixtures from `conftest.py` for database and directory setup. All required dependencies should be installed in the virtual environment.