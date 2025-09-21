# Smoke Test Coverage

This document describes the end-to-end smoke tests for the translation pipeline backend API.

## Test Files

### 1. `test_smoke_simple.py`
Main end-to-end test file covering the complete user workflow with mocked services.

### 2. `test_main.py`
Basic API functionality tests (already exists).

### 3. `test_smoke_e2e.py`
Extended end-to-end test file with additional scenarios (may need fixes for import issues).

## Test Categories

### Test Functions in `test_smoke_simple.py`

#### `test_complete_translation_workflow`
- **User Registration**: POST `/api/auth/register`
- **User Login**: POST `/api/auth/login`
- **Token-based Authentication**: Using Bearer tokens
- **PPTX Translation**: Complete workflow from upload to job completion
  - POST `/api/translate` with PPTX file
  - GET `/api/jobs/{job_id}` for status checking
  - GET `/api/jobs` for listing jobs
- **Job Management**:
  - POST `/api/jobs/{job_id}/cancel` for job cancellation
  - Status verification after cancellation
- **Job Statistics**: GET `/api/jobs/statistics`
  - Total jobs count
  - Status distribution

#### `test_pdf_translation_workflow`
- **PDF Translation**: Complete workflow with PDF files
  - POST `/api/translate` with PDF file
  - Page range parameter testing (e.g., "1-5")
- **Job Search**: POST `/api/jobs/search`
  - Search by filename
  - Status filtering
  - Pagination

#### `test_error_handling_scenarios`
- **Unauthorized Access**: Missing or invalid tokens
- **Invalid Login**: Wrong credentials
- **Job Not Found**: Accessing non-existent job IDs
- **Invalid File Types**: Uploading unsupported file formats

#### `test_bulk_operations`
- **Multiple Job Submission**: Creating several jobs
- **Bulk Cancellation**: POST `/api/jobs/bulk/cancel`
  - Multiple job IDs in single request
  - Success/failure reporting for each job

### Additional Tests in `test_main.py`
- **Health Check**: GET `/health`
- **Authentication**: Basic user registration and login
- **API Models**: GET `/api/translate/models`
- **File Formats**: GET `/api/translate/formats`
- **Environment Configuration**: Test settings validation
- **Directory Setup**: Test directory creation and management

## Mocked Services

To avoid external dependencies and costs:
- **Translation Service**: Mocked to return immediate "completed" status
- **OpenAI API**: Not actually called
- **File Processing**: Minimal test files created programmatically

## Running the Tests

### Option 1: Using the smoke test runner
```bash
cd backend
python run_smoke_tests.py
```

### Option 2: Using pytest directly
```bash
cd backend
pytest tests/test_smoke_e2e.py -v
```

### Option 3: Running specific test categories
```bash
# Run only authentication tests
pytest tests/test_smoke_e2e.py::TestUserAuthWorkflow -v

# Run only translation workflow tests
pytest tests/test_smoke_e2e.py::TestTranslationJobWorkflow -v

# Run tests with specific markers
pytest tests/ -m "e2e" -v
```

## Test Environment

The tests use:
- In-memory SQLite database
- Temporary upload and output directories
- Mock environment variables
- Sample PPTX and PDF files created on-the-fly

## Endpoints Tested

### Authentication
- `POST /api/auth/register`
- `POST /api/auth/login`

### Translation
- `POST /api/translate`
- `GET /api/translate/models`
- `GET /api/translate/formats`

### Jobs
- `GET /api/jobs`
- `POST /api/jobs/search`
- `GET /api/jobs/{job_id}`
- `POST /api/jobs/{job_id}/cancel`
- `POST /api/jobs/{job_id}/retry`
- `DELETE /api/jobs/{job_id}`
- `GET /api/jobs/statistics`
- `GET /api/jobs/queue`
- `GET /api/jobs/export`
- `POST /api/jobs/bulk/cancel`
- `POST /api/jobs/bulk/retry`
- `GET /api/jobs/{job_id}/logs`

### Health
- `GET /health` (from test_main.py)

## Success Criteria

All tests should pass with:
- No external API calls made
- All endpoints responding with correct status codes
- Proper authentication and authorization
- Correct job lifecycle management
- Error handling for edge cases