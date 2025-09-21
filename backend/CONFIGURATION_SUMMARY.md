# Configuration Refactoring Summary

## Overview
The backend configuration has been refactored to support safe defaults for development and testing modes while maintaining strict requirements for production environments.

## Key Changes

### 1. Safe Defaults for DEBUG Mode
When `DEBUG=true`:
- **SECRET_KEY**: Uses a stable development key `"dev-secret-key-32-characters-long-for-local-dev"`
- **OPENAI_API_KEY**: Uses a placeholder key `"debug-placeholder-key-invalid-for-production"`
- Clear warnings are logged when using defaults

### 2. Production Validation
When `DEBUG=false` (production):
- **SECRET_KEY**: Must be provided and at least 32 characters long
- **OPENAI_API_KEY**: Must be a valid API key (not the placeholder)
- Validation errors are raised for missing or invalid values

### 3. Environment Variables
The configuration now properly documents all expected environment variables:

#### Core Settings
- `DEBUG`: Enable development mode (default: false)
- `SECRET_KEY`: Secret key for JWT tokens (required in production)
- `APP_NAME`: Application name (default: "Translation Pipeline API")
- `VERSION`: Application version (default: "1.0.0")

#### API Settings
- `API_PREFIX`: URL prefix for API endpoints (default: "/api")
- `ALLOWED_ORIGINS`: CORS allowed origins (default: localhost ports 3000, 3001)

#### File Storage
- `UPLOAD_DIR`: Directory for uploaded files (default: "uploads")
- `OUTPUT_DIR`: Directory for translated files (default: "outputs")
- `MAX_FILE_SIZE`: Maximum upload size in bytes (default: 100MB)

#### OpenAI Integration
- `OPENAI_API_KEY`: OpenAI API key (required for translation)
- `OPENAI_MODEL`: Default model (default: "gpt-4o-2024-08-06")
- `OPENAI_TEMPERATURE`: Translation temperature (default: 0.6)

#### Authentication
- `ACCESS_TOKEN_EXPIRE_MINUTES`: JWT token expiration (default: 30)
- `REFRESH_TOKEN_EXPIRE_DAYS`: Refresh token expiration (default: 7)

#### Infrastructure
- `REDIS_URL`: Redis connection for job queue (default: "redis://localhost:6379")

## Usage Examples

### Development Mode
```bash
# Start with safe defaults
DEBUG=true python -m uvicorn app.main:app --reload

# Or set specific values
DEBUG=true \
SECRET_KEY=your-dev-secret \
OPENAI_API_KEY=sk-your-api-key \
python -m uvicorn app.main:app --reload
```

### Production Mode
```bash
# All required values must be set
DEBUG=false \
SECRET_KEY=your-secure-32-character-secret \
OPENAI_API_KEY=sk-your-production-api-key \
ALLOWED_ORIGINS='["https://yourdomain.com"]' \
python -m uvicorn app.main:app
```

## Helper Methods
The configuration class provides useful helper methods:

- `is_openai_configured()`: Check if OpenAI API is properly configured
- `is_redis_configured()`: Check if Redis is using non-default settings
- `get_environment_info()`: Get information about current environment and warnings

## Testing
The configuration automatically loads `.env.test` when running under pytest. Test environment should set `PYTEST_RUNNING=true` to use test configuration.

## Security Notes
- Development defaults are intentionally obvious and should never be used in production
- The development SECRET_KEY is stable across restarts for session consistency
- All configuration values can be overridden by environment variables
- Environment files (.env, .env.test) are never committed to version control