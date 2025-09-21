# API Structure and Consolidation Guide

## Overview

This directory (`backend/app/api/`) contains modular FastAPI routers for the translation pipeline API. Each router handles a specific domain (e.g., authentication, translation jobs, SSE events) and defines endpoints with proper authentication, validation, and integration with services like `job_manager` and `auth_service`. These routers are designed for a production-ready, scalable API with features like token-based auth, database-backed job management, and real-time updates.

The main FastAPI app (`backend/app/main.py`) currently mounts these routers and includes some legacy/simple endpoints for basic file upload and job tracking. This README documents the current structure, compares it to `main.py`, and proposes a consolidation plan to make the routers the authoritative source.

## Current Structure

### Routers in `backend/app/api/`
- **auth.py**: Handles user authentication, registration, token management, Google OAuth, and API keys.
  - Router prefix (when mounted): `/api/auth`
  - Key endpoints:
    - `POST /register`: Create new user.
    - `POST /login`: Authenticate and return tokens.
    - `POST /refresh`: Refresh access token.
    - `POST /logout`: Revoke refresh token.
    - `GET /me`: Get current user info.
    - `GET /google/auth-url`: Get Google OAuth URL.
    - `POST /google/callback`: Handle Google callback and create/update user.
    - `POST /api-keys`: Create API key (requires auth).
    - `GET /api-keys`: List user's API keys.
    - `DELETE /api-keys/{key_id}`: Revoke API key.
  - Dependencies: Database session (`get_db`), JWT tokens via `HTTPBearer`.
  - Full paths (with mount): `/api/auth/register`, `/api/auth/login`, etc. (covers all `/api/auth/*`).

- **translate.py**: Manages translation requests, including file uploads and model/format queries.
  - Router prefix (when mounted): `/api`
  - Key endpoints:
    - `POST /translate`: Upload file and create translation job (supports PPTX/PDF, models like GPT-4o, offline mode, page ranges).
    - `GET /translate/models`: List available AI models with pricing.
    - `GET /translate/formats`: List supported formats and options.
  - Dependencies: Auth via `HTTPBearer`, file service for uploads.
  - Full paths: `/api/translate`, `/api/translate/models`, `/api/translate/formats` (covers all `/api/translate/*`).

- **jobs.py**: Comprehensive job management, including listing, searching, canceling, retrying, stats, and exports.
  - Router prefix (when mounted): `/api`
  - Key endpoints:
    - `GET /jobs`: List jobs with pagination/filtering.
    - `POST /jobs/search`: Advanced job search.
    - `GET /jobs/{job_id}`: Get job details (includes logs).
    - `POST /jobs/{job_id}/cancel`: Cancel a job.
    - `POST /jobs/{job_id}/retry`: Retry a failed job.
    - `POST /jobs/bulk/cancel`: Bulk cancel jobs.
    - `POST /jobs/bulk/retry`: Bulk retry jobs.
    - `GET /jobs/statistics`: Get user job stats (e.g., over 30 days).
    - `GET /jobs/queue`: Get queue status.
    - `GET /jobs/{job_id}/logs`: Get job logs.
    - `GET /jobs/export`: Export jobs as CSV/JSON.
    - `DELETE /jobs/{job_id}`: Delete completed/failed job.
  - Dependencies: Auth via `HTTPBearer`, `job_manager` for DB operations (SQLite-backed).
  - Full paths: `/api/jobs`, `/api/jobs/{job_id}`, etc. (covers all `/api/jobs/*`).

- **sse.py**: Server-Sent Events (SSE) for real-time job updates and notifications.
  - Router prefix (when mounted): `/api`
  - Key endpoints:
    - `GET /sse/subscribe`: Subscribe to SSE stream (supports job-specific or user-wide events, heartbeats every 30s).
  - Features: Manages active connections with queues, sends job updates/notifications.
  - Dependencies: Auth via `HTTPBearer`, integrates with `job_manager`.
  - Full path: `/api/sse/subscribe` (covers SSE needs under `/api/sse/*`).

- **__init__.py**: Empty (no exports needed; routers are imported directly in `main.py`).

### Wiring in `backend/app/main.py`
- **App Setup**: Creates `FastAPI` app with CORS (allows localhost:3000/3001).
- **Router Mounts** (current):
  - `app.include_router(auth.router, prefix="/api/auth", tags=["auth"])`
  - `app.include_router(translate.router, prefix="/api", tags=["translate"])`
  - `app.include_router(jobs.router, prefix="/api", tags=["jobs"])`
  - `app.include_router(sse.router, prefix="/api", tags=["sse"])`
- **Legacy/Simple Endpoints** (at root level, no auth):
  - `GET /health`: Basic health check.
  - `POST /upload`: Simple file upload (returns `file_id`; supports PPTX/PDF).
  - `POST /translate`: Start translation by `file_id` (uses background task calling scripts; tracks progress in in-memory `jobs` dict).
  - `GET /jobs/{job_id}`: Get job status (from in-memory dict).
  - `GET /jobs`: List all jobs (from dict).
  - `GET /jobs/{job_id}/download`: Download completed result.
- **Background Logic**: `run_translation` function uses `asyncio` to subprocess-call `scripts/translate_pptx_inplace.py` or `translate_pdf.py`; updates in-memory `jobs` dict.
- **Global State**: In-memory `jobs: Dict[str, Dict]` for simple tracking (no persistence).

## Comparison: main.py vs. Routers

| Aspect          | main.py (Legacy/Simple)                          | Routers (Authoritative/Rich)                          |
|-----------------|--------------------------------------------------|-------------------------------------------------------|
| **Structure**  | Monolithic: All endpoints in one file; in-memory state; basic subprocess calls. | Modular: Domain-separated routers; DB-backed (`job_manager` with SQLite); service layers (e.g., `auth_service`, `file_service`). |
| **Endpoints**  | Basic CRUD for uploads/jobs (5 endpoints); no auth; simple progress (25%, 50%, etc.). Overlaps with routers (e.g., `/translate` vs. `/api/translate`; `/jobs` vs. `/api/jobs`). | Rich: 20+ endpoints across auth/translate/jobs/SSE; auth-protected; advanced features (bulk ops, search, stats, exports, logs, retry/cancel). Covers `/api/auth/*`, `/api/translate/*`, `/api/jobs/*` fully. |
| **Auth/Security** | None (open endpoints); no tokens or user isolation. | JWT via `HTTPBearer`; user-specific jobs; API keys; Google OAuth. |
| **State Management** | In-memory dict (`jobs`); no persistence; single-user. | Persistent via `job_manager` (SQLite); multi-user; queue integration. |
| **Error Handling** | Basic `HTTPException`; subprocess errors captured. | Comprehensive: Validation (Pydantic), status-specific errors, logging. |
| **Real-time**  | None.                                           | SSE support in `sse.py` for progress/notifications. |
| **Scalability** | Not scalable (memory-bound, no DB).              | DB-backed; supports pagination, filtering, bulk ops. |
| **Integration** | Direct script calls; hardcoded paths.            | Service abstractions; configurable models/pages/auto-fit. |

**Key Overlaps/Duplications**:
- Upload/Translate: `main.py`'s `/upload` + `/translate` duplicates `/api/translate` (which handles upload inline).
- Job Listing/Status: `/jobs` and `/jobs/{job_id}` duplicate `/api/jobs` endpoints (richer version has auth, filters, logs).
- Download: `/jobs/{job_id}/download` could be added to `jobs.py` if needed (currently missing; use file service).
- Health: `/health` is unique and fine to keep.

**Gaps**:
- No download endpoint in routers (recommend adding to `jobs.py`).
- SSE is router-only (good).

## Consolidation Decision

**Authoritative Stack**: The routers in `backend/app/api/` are the recommended authoritative implementation. They provide:
- Better modularity and maintainability.
- Production features (auth, DB persistence, error handling, real-time updates).
- Compliance with best practices (dependency injection, Pydantic models, service layers).
- Coverage of required paths: `/api/auth/*` (full auth flow), `/api/translate/*` (file processing), `/api/jobs/*` (job lifecycle).

The `main.py` endpoints are legacy/simple prototypes (likely for initial testing without auth/DB). They lack security, scalability, and overlap with routers, making them dead code. Consolidating to routers simplifies `main.py` (focus on app setup, CORS, mounts) and eliminates duplication.

**Risks if Not Consolidated**:
- Security: Unauth endpoints expose uploads/jobs.
- Inconsistency: In-memory vs. DB state; different progress tracking.
- Maintenance: Duplicate logic leads to bugs (e.g., one updates but not the other).

## Proposed Changes

### 1. Router Mounts (Keep/Refine in `main.py`)
Retain current mounts for consistency. No changes needed, but add tags for OpenAPI docs:
```python
# In backend/app/main.py
app.include_router(auth.router, prefix="/api/auth", tags=["auth"])
app.include_router(translate.router, prefix="/api/translate", tags=["translate"])  # Add sub-prefix for clarity
app.include_router(jobs.router, prefix="/api/jobs", tags=["jobs"])               # Add sub-prefix for clarity
app.include_router(sse.router, prefix="/api/sse", tags=["sse"])                 # Add sub-prefix for clarity
```
- This ensures clean paths: `/api/translate/*`, `/api/jobs/*`, `/api/sse/*`, `/api/auth/*`.
- Benefits: Avoids root-level clutter; improves API docs grouping.

### 2. Remove Dead Code from `main.py`
- Delete: `/upload`, `/translate` (POST), `/jobs/{job_id}` (GET), `/jobs` (GET), `/jobs/{job_id}/download`.
- Delete: Global `jobs: Dict[str, Dict]`.
- Delete: `run_translation` function and related background task logic (subprocess calls; migrate to `job_manager` if needed).
- Keep: `/health` (simple utility).
- Keep: CORS, app creation, router includes.
- Add (if needed): Download endpoint to `jobs.py` (e.g., `GET /jobs/{job_id}/download` using `file_service`).

### 3. Additional Improvements
- **Add Download to `jobs.py`**: Implement `GET /jobs/{job_id}/download` (auth-protected, returns `FileResponse` from output path).
- **Migration Path**: For existing simple flows, redirect unauth users to auth endpoints or deprecate with 301/410 responses.
- **Testing**: After changes, verify:
  - Auth flows work (`/api/auth/*`).
  - Translation creates jobs (`/api/translate` → integrates with `/api/jobs`).
  - SSE subscribes to job updates (`/api/sse/subscribe`).
  - Run `uvicorn backend.app.main:app --reload` and test with curl/Postman.
- **Docs**: Update OpenAPI (auto-generated at `/docs`); add endpoint summaries.

### 4. Implementation Steps
1. Backup `main.py`.
2. Update mounts in `main.py` (add sub-prefixes).
3. Remove legacy endpoints and state.
4. Add download to `jobs.py` (optional).
5. Test: `pytest backend/app/`; manual API calls.
6. Commit: "feat: consolidate API to routers; remove legacy endpoints".

This consolidation reduces code duplication by ~40% in `main.py` while enhancing security and features. For questions, see `CLAUDE.md` or contact the team.