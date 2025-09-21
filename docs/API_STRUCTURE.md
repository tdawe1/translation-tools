# API Structure Documentation

## Base URL
```
http://localhost:8000
```

## Overview
The Translation Pipeline API provides endpoints for authentication, document translation, job management, and real-time updates via Server-Sent Events (SSE).

## Endpoints

### Health Check
- **GET** `/health` - Health check endpoint
  - No authentication required
  - Returns service status and configuration

### Authentication (`/api/auth`)
All endpoints require authentication except login and register.

- **POST** `/api/auth/login` - User login
  - Body: `{ username, password }`
  - Returns: JWT access token

- **POST** `/api/auth/register` - User registration
  - Body: `{ username, email, password }`
  - Returns: JWT access token

- **GET** `/api/auth/me` - Get current user info
  - Requires: Bearer token
  - Returns: User information

- **POST** `/api/auth/refresh` - Refresh access token
  - Requires: Refresh token
  - Returns: New access token

- **POST** `/api/auth/logout` - User logout
  - Requires: Bearer token

- **POST** `/api/auth/google/callback` - Google OAuth callback
  - Handles Google OAuth2 authentication

- **GET** `/api/auth/api-keys` - List API keys
  - Requires: Bearer token

- **POST** `/api/auth/api-keys` - Create new API key
  - Requires: Bearer token

- **DELETE** `/api/auth/api-keys/{keyId}` - Delete API key
  - Requires: Bearer token

### Translation (`/api/translate`)
All endpoints require authentication.

- **POST** `/api/translate/translate` - Create translation job
  - Requires: Bearer token, file upload
  - Query params: `file_type`, `model`, `temperature`, `offline`, `pages`, `auto_fit`
  - Returns: Job details

- **GET** `/api/translate/translate/models` - List available models
  - Requires: Bearer token
  - Returns: Model information

- **GET** `/api/translate/translate/formats` - List supported formats
  - Requires: Bearer token
  - Returns: Format information

- **POST** `/api/translate/translate/pptx` - PPTX-specific translation
  - Requires: Bearer token, file upload
  - Specialized for PowerPoint presentations

### Jobs (`/api/jobs`)
All endpoints require authentication.

- **GET** `/api/jobs/jobs` - List jobs with pagination
  - Requires: Bearer token
  - Query params: `page`, `page_size`, `status`, `file_type`, `search`, etc.
  - Returns: Paginated job list

- **POST** `/api/jobs/jobs/search` - Advanced job search
  - Requires: Bearer token
  - Body: Search filters
  - Returns: Filtered job list

- **GET** `/api/jobs/jobs/{job_id}` - Get job details
  - Requires: Bearer token
  - Returns: Job details and logs

- **POST** `/api/jobs/jobs/{job_id}/cancel` - Cancel a job
  - Requires: Bearer token
  - Returns: Success message

- **POST** `/api/jobs/jobs/{job_id}/retry` - Retry a failed job
  - Requires: Bearer token
  - Returns: New job ID

- **GET** `/api/jobs/jobs/statistics` - Get job statistics
  - Requires: Bearer token
  - Query params: `days`
  - Returns: Job statistics

- **GET** `/api/jobs/jobs/queue` - Get queue status
  - Requires: Bearer token
  - Returns: Current queue information

- **GET** `/api/jobs/jobs/{job_id}/logs` - Get job logs
  - Requires: Bearer token
  - Query params: `limit`
  - Returns: Job logs

- **GET** `/api/jobs/jobs/export` - Export job data
  - Requires: Bearer token
  - Query params: `format` (csv/json)
  - Returns: Export data

- **DELETE** `/api/jobs/jobs/{job_id}` - Delete a job
  - Requires: Bearer token
  - Returns: Success message

- **GET** `/api/jobs/{job_id}/download` - Download translated file
  - Requires: Bearer token
  - Returns: File download

- **POST** `/api/jobs/submit` - Submit job (stub endpoint)
  - Returns: Stub response

### Bulk Operations
- **POST** `/api/jobs/jobs/bulk/cancel` - Cancel multiple jobs
- **POST** `/api/jobs/jobs/bulk/retry` - Retry multiple jobs

### Server-Sent Events (`/api/sse`)
- **GET** `/api/sse/subscribe` - Subscribe to job updates
  - Requires: Bearer token
  - Query params: `job_id` (optional)
  - Returns: SSE stream for real-time updates

## Authentication
All protected endpoints require a Bearer token in the Authorization header:
```
Authorization: Bearer <jwt_token>
```

## Error Responses
- **401 Unauthorized** - Invalid or missing authentication
- **403 Forbidden** - Insufficient permissions
- **404 Not Found** - Endpoint or resource not found
- **422 Unprocessable Entity** - Invalid request body or parameters
- **500 Internal Server Error** - Server error

## File Upload
Translation endpoints accept file uploads with the following constraints:
- Maximum file size: 100MB
- Supported formats: PPTX, PDF
- Content-Type: Must match file type

## SSE Implementation
- Endpoint: `/api/sse/subscribe?token=<jwt>&job_id=<optional>`
- Events: job_status_update, job_progress, job_complete, job_failed
- Auto-reconnect with last event ID
- Heartbeat every 10 seconds

## Frontend Integration
The frontend expects the following base URLs:
- Auth endpoints: `/api/auth/*`
- Translation endpoints: `/api/translate/*`
- Jobs endpoints: `/api/jobs/*`
- SSE endpoint: `/api/sse/*`

## Environment Configuration
- `DEBUG` - Enable development mode
- `SECRET_KEY` - JWT signing key (32+ characters)
- `ALLOWED_ORIGINS` - CORS allowed origins
- `OPENAI_API_KEY` - OpenAI API key
- `DATABASE_URL` - Database connection string
- `REDIS_URL` - Redis connection for job queue