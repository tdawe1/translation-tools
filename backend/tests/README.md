# Backend Test Fixtures

This directory contains pytest fixtures for backend testing that ensure tests run with a clean, isolated environment.

## Environment Setup

The test environment is automatically configured through:

1. **`.env.test` file**: Contains test-specific environment variables
2. **`conftest.py`**: Sets up test fixtures and environment variables

## Key Fixtures

### `clean_test_environment` (autouse)
- Runs before and after each test
- Cleans up test database file
- Cleans up test directories (test_uploads, test_outputs)
- Ensures each test starts with a clean slate

### `test_upload_dir`
- Creates a temporary upload directory for each test
- Automatically cleaned up after the test
- Overrides settings.UPLOAD_DIR temporarily

### `test_output_dir`
- Creates a temporary output directory for each test
- Automatically cleaned up after the test
- Overrides settings.OUTPUT_DIR temporarily

### `test_db`
- Creates an in-memory SQLite database for each test
- Provides a clean database session
- Automatically creates/drops tables
- Uses dependency injection to override the app's database

### `client`
- Provides a FastAPI TestClient instance
- Automatically uses test settings

### `auth_headers`
- Creates a test user and returns authentication headers
- Useful for testing protected endpoints

### `mock_openai`
- Mocks OpenAI API calls to avoid real API calls during testing
- Returns predefined mock responses

## Running Tests

```bash
# Run all tests
cd backend
python -m pytest tests/

# Run with verbose output
python -m pytest tests/ -v

# Run specific test file
python -m pytest tests/test_main.py -v

# Run specific test
python -m pytest tests/test_main.py::test_health_check -v

# Use the provided script
./run_tests.sh
```

## Test Environment Variables

The test environment uses the following special values:

- `DEBUG`: true
- `SECRET_KEY`: test-secret-key-for-pytest-testing-only
- `OPENAI_API_KEY`: mock-sk-for-testing
- `DATABASE_URL`: sqlite:///./test_translation_pipeline.db
- `UPLOAD_DIR`: test_uploads
- `OUTPUT_DIR`: test_outputs
- `LOG_LEVEL`: WARNING

## Best Practices

1. **Use fixtures**: Always use the provided fixtures instead of creating your own test clients or directories
2. **Test isolation**: Each test runs in complete isolation thanks to the autouse `clean_test_environment` fixture
3. **Mock external services**: Use the `mock_openai` fixture to avoid real API calls
4. **Database testing**: Use the `test_db` fixture for any database operations
5. **Authentication**: Use the `auth_headers` fixture for testing protected endpoints

## Adding New Tests

When adding new test files, place them in the `tests/` directory with the prefix `test_`. The fixtures in `conftest.py` will be automatically available.