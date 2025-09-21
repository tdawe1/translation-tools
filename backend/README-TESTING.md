# Backend Test Environment Setup

This document describes the pytest test environment setup for the backend.

## Overview

The test environment is designed to:
- Run tests in isolation with predictable settings
- Use in-memory SQLite for database tests
- Create temporary directories for file operations
- Mock external API calls (OpenAI)
- Clean up resources after tests complete

## Test Configuration Files

### `.env.test`
Contains test-specific environment variables:
- `DEBUG=true` - Enable debug mode
- `SECRET_KEY=test-secret-key-for-pytest-testing-only-32-chars-long` - Test secret key
- `OPENAI_API_KEY=mock-sk-for-testing` - Mock API key that won't make real calls
- `DATABASE_URL=sqlite:///:memory:` - Use in-memory SQLite
- `UPLOAD_DIR=test_uploads` and `OUTPUT_DIR=test_outputs` - Test directories
- Shorter token expiration times for testing

### `conftest.py`
Contains pytest fixtures that set up the test environment:

#### Fixtures:
- `clean_test_environment` - Cleans up test directories before/after each test
- `test_upload_dir` - Creates temporary upload directory
- `test_output_dir` - Creates temporary output directory
- `test_db` - Creates in-memory SQLite database session
- `client` - TestClient with database override
- `auth_headers` - Creates authentication headers for protected endpoints
- `admin_auth_headers` - Creates admin authentication headers
- `mock_openai` - Mocks OpenAI API calls

### `pytest.ini`
Pytest configuration:
- Test discovery patterns
- Output formatting
- Warning filters
- Custom markers

## Running Tests

```bash
# Run all tests
cd backend
python -m pytest

# Run specific test file
python -m pytest tests/test_main.py

# Run with verbose output
python -m pytest -v

# Run specific test
python -m pytest tests/test_main.py::test_health_check

# Run with coverage
python -m pytest --cov=app
```

## Test Structure

Tests are organized by:
- `test_main.py` - Main application tests
- Tests should use fixtures for database, auth, and API mocking
- Each test runs in isolation with clean database

## Key Patterns

### Database Tests
```python
def test_database_operation(test_db):
    # Use the test_db fixture for database operations
    from app.database.database import User

    user = User(...)
    test_db.add(user)
    test_db.commit()
```

### API Tests
```python
def test_api_endpoint(client):
    # Use client fixture for API calls
    response = client.get("/api/endpoint")
    assert response.status_code == 200
```

### Authenticated API Tests
```python
def test_protected_endpoint(client, auth_headers):
    # Use auth_headers fixture for authenticated requests
    response = client.get("/api/protected", headers=auth_headers)
    assert response.status_code == 200
```

## Known Issues

1. **SQLAlchemy Connection Pooling**: Tests may encounter connection issues when mixing different database engines. Always use the `test_db` fixture for database operations.

2. **Async/Sync Mixing**: The application uses async endpoints but tests run synchronously. This is handled by the TestClient.

3. **File Operations**: Test files are automatically cleaned up by the `clean_test_environment` fixture.

## Debugging Tests

To debug test failures:
1. Run with `-v` flag for verbose output
2. Use `--tb=short` or `--tb=long` for traceback control
3. Add print statements or use pdb for debugging
4. Check test isolation - ensure tests don't depend on each other