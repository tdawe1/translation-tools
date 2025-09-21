# Backend API Audit

## Current Structure

### Main Application (backend/app/main.py)
- **FastAPI App**: Initialized with title "Translation Pipeline API".
- **Middleware**:
  - CORS middleware configured for origins `["http://localhost:3000", "http://localhost:3001"]`, allowing all methods and headers, with credentials.
- **Router Mounting**:
  - `auth.router` mounted at `/api/auth` with tags `["auth"]`.
  - `translate.router` mounted at `/api` with tags `["translate"]`.
  - `jobs.router` mounted at `/api` with tags `["jobs"]`.
  - `sse.router` mounted at `/api` with tags `["sse"]`.
- **Direct Endpoints** (defined in main.py, outside routers):
  - `GET /health`: Health check returning `{"status": "healthy"}`.
  - `POST /upload`: Uploads PPTX/PDF files, returns file_id, filename, path. Stores in `uploads/` directory.
  - `POST /translate`: Starts translation job in background, returns job_id and status "started". Uses simple script invocation for PPTX/PDF.
  - `GET /jobs/{job_id}`: Retrieves job status from in-memory `jobs` dict.
  - `GET /jobs`: Lists all jobs from in-memory dict.
  - `GET /jobs/{job_id}/download`: Downloads completed job output file.

### API Routers (backend/app/api/)

#### Auth Router (auth.py, mounted at `/api/auth`)
- Authentication and user management endpoints.
- Endpoints:
  - `POST /register`: Register new user (response: User model).
  - `POST /login`: User login, returns access/refresh tokens (response: Token model).
  - `POST /refresh`: Refresh access token using refresh token.
  - `POST /logout`: Revoke refresh token.
  - `GET /me`: Get current user info (requires auth).
  - `GET /google/auth-url`: Get Google OAuth URL.
  - `POST /google/callback`: Handle Google OAuth callback, returns tokens and user.
  - `POST /api-keys`: Create API key (requires auth).
  - `GET /api-keys`: List user's API keys.
  - `DELETE /api-keys/{key_id}`: Revoke API key.
- Full paths: e.g., `/api/auth/login`, `/api/auth/api-keys` (note: `/api-keys` is relative to `/api/auth`, so `/api/auth/api-keys`).

#### Translate Router (translate.py, mounted at `/api`)
- Translation job creation and configuration.
- Endpoints:
  - `POST /translate`: Upload and start translation job for PPTX/PDF (params: file_type, model, etc.; response: JobResponse).
  - `GET /translate/models`: List available models (e.g., gpt-4o-2024-08-06).
  - `GET /translate/formats`: List supported formats and options.
- Full paths: `/api/translate`, `/api/translate/models`, `/api/translate/formats`.

#### Jobs Router (jobs.py, mounted at `/api`)
- Job management, listing, and operations.
- Endpoints:
  - `GET /jobs`: List jobs with pagination/filtering.
  - `POST /jobs/search`: Advanced job search.
  - `GET /jobs/{job_id}`: Get job details including logs.
  - `POST /jobs/{job_id}/cancel`: Cancel a job.
  - `POST /jobs/{job_id}/retry`: Retry a failed job.
  - `POST /jobs/bulk/cancel`: Bulk cancel jobs.
  - `POST /jobs/bulk/retry`: Bulk retry jobs.
  - `GET /jobs/statistics`: Get job statistics.
  - `GET /jobs/queue`: Get queue status.
  - `GET /jobs/{job_id}/logs`: Get job logs.
  - `GET /jobs/export`: Export jobs data (CSV/JSON).
  - `DELETE /jobs/{job_id}`: Delete completed/failed job.
- Full paths: e.g., `/api/jobs`, `/api/jobs/{job_id}`.

#### SSE Router (sse.py, mounted at `/api`)
- Server-Sent Events for real-time updates.
- Endpoints:
  - `GET /sse/subscribe`: SSE stream for job updates (params: job_id optional).
- Full path: `/api/sse/subscribe`.

### Middleware and Dependencies
- CORS is applied globally in main.py.
- Security: HTTPBearer for most endpoints in routers; API key support in auth.
- Database: SQLAlchemy sessions via Depends(get_db) in auth/jobs.
- No additional middleware (e.g., rate limiting, logging) observed in main.py.

### Overlaps/Duplications
- `/translate` (POST): Defined both in main.py (simple version) and `/api/translate` (advanced in router).
- `/jobs` (GET) and `/jobs/{job_id}` (GET): Defined in main.py (in-memory simple) and in jobs router (DB-backed advanced).

## Gaps

### Expected vs. Actual Endpoints
- **/api/auth/login**: Exists (POST /api/auth/login).
- **/api/translate/pptx**: Does not exist exactly. Closest is `POST /api/translate` which handles both PPTX and PDF via `file_type` parameter. No dedicated PPTX endpoint; translation type is inferred from upload.
- **/api/jobs/submit**: Does not exist. Job submission happens via `POST /api/translate` (which creates the job). No dedicated `/api/jobs` POST for submitting a job ID or reference. Jobs router focuses on management (list, get, cancel), not creation.

### Other Gaps
- **Upload Endpoint**: Direct `POST /upload` in main.py is simple but not mounted under `/api`. No equivalent in routers; uploads are handled inline in `/api/translate`.
- **Download Endpoint**: `GET /jobs/{job_id}/download` is only in main.py (simple file response). Jobs router has no download; assumes output file access via job details.
- **Job Creation**: No explicit `/api/jobs/create` or `/api/jobs/submit`. Translation jobs are created via translate endpoint.
- **Health Check**: `/health` is root-level, not under `/api`.
- **Missing Features**:
  - No dedicated endpoint for PDF-specific translation (e.g., `/api/translate/pdf`).
  - No bulk translation submission (e.g., `/api/translate/bulk`).
  - No endpoint for cache management or offline job submission.
  - Potential gap in idempotency for job creation (mentioned in env, but not enforced in endpoints).
- **Security/Consistency**: Direct endpoints in main.py lack auth dependencies (e.g., no Depends(security)), making them unauthenticated. Routers enforce auth.
- **API Versioning**: No prefix like `/v1`; all under `/api`.
- **Error Handling**: Inconsistent; main.py uses simple HTTPException, routers use more structured responses.
- **Documentation**: No explicit OpenAPI tags beyond basic; could use more descriptions.

### Verification Notes
- Analyzed via code inspection (read tools).
- No local uvicorn run performed as code inspection suffices for static audit. To verify dynamically:
  - Run `uvicorn backend.app.main:app --reload` (assumes in backend dir).
  - Access `/docs` for Swagger UI to list all endpoints.
  - Direct endpoints appear at root; router endpoints under `/api`.
- In-memory `jobs` dict in main.py conflicts with DB-backed jobs in routers (e.g., jobs.py uses SQLite).

## Recommendations

### Mounting and Structure
- **Already Mounted**: All routers are properly included in main.py with appropriate prefixes. No changes needed for mounting.
- **Consolidate Direct Endpoints**: Migrate `/upload`, `/translate`, `/jobs`, `/jobs/{job_id}`, `/jobs/{job_id}/download` from main.py into appropriate routers:
  - Move upload logic to a new `/api/files` router or integrate into translate/jobs.
  - Deprecate simple `/translate` in favor of `/api/translate`.
  - Add `GET /api/jobs/{job_id}/download` to jobs router for consistency.
  - Move `/health` to `/api/health` or keep as is.
- **Add Missing Endpoints**:
  - Implement `POST /api/jobs/submit` in jobs router for explicit job creation (e.g., reference existing upload).
  - Add `POST /api/translate/pptx` and `POST /api/translate/pdf` as wrappers around `/api/translate` for type-specific paths.
  - Consider `POST /api/files/upload` for standalone uploads with auth.
- **Resolve Duplications**: Remove overlapping endpoints from main.py after verifying router versions work (e.g., switch to DB-backed jobs).
- **Enhance Middleware**:
  - Add global auth middleware or rate limiting (e.g., slowapi).
  - Ensure all endpoints (including direct ones) require authentication via Depends.
- **API Design Improvements**:
  - Introduce versioning: Mount routers under `/api/v1`.
  - Standardize responses: Use consistent JobResponse models across endpoints.
  - Add idempotency: Support `Idempotency-Key` header for `/api/translate` and `/api/jobs/submit`.
  - Document gaps in OpenAPI: Add deprecated tags for main.py endpoints.
- **Migration Plan**:
  1. Create new endpoints in routers to match direct ones.
  2. Update frontend/backend integrations to use `/api` paths.
  3. Remove direct endpoints from main.py.
  4. Test with uvicorn and curl (e.g., `curl http://localhost:8000/api/docs`).
- **Prepare for Task 4**: This audit identifies overlaps and gaps. Task 4 can focus on refactoring main.py to delegate to routers, adding missing endpoints like `/api/jobs/submit`, and cleaning up duplications.

## Audit Script
A bash script `scripts/audit_api.sh` has been added to grep for endpoint definitions:
- Grep for `@app.` and `@router.` decorators.
- List all paths and methods.
- Run: `./scripts/audit_api.sh` to output endpoint summary.
