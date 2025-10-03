# Translation Pipeline Backend

FastAPI backend wrapper for the PowerPoint and PDF translation pipeline.

## Features

- REST API for translation jobs
- JWT-based authentication
- File upload/download
- Job status tracking
- Support for PPTX and PDF files
- Configurable translation models
- CORS support for frontend integration

## Installation

1. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Configure environment variables:
```bash
cp .env.example .env
# Edit .env with your settings
```

## Running the Server

Development mode:
```bash
./run.sh
```

Or manually:
```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

The API will be available at `http://localhost:8000`

## API Documentation

- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`

## API Endpoints

### Authentication
- `POST /api/auth/register` - Register a new user
- `POST /api/auth/login` - Login and get access token
- `GET /api/auth/me` - Get current user info
- `POST /api/auth/refresh` - Refresh access token
- `POST /api/auth/logout` - Logout user

### Translation
- `POST /api/translate` - Create a new translation job
- `GET /api/jobs/{job_id}` - Get job status
- `GET /api/jobs` - List user's jobs
- `DELETE /api/jobs/{job_id}` - Cancel a job
- `GET /api/files/{file_id}/download` - Download translated file
- `GET /api/translate/models` - List available models
- `GET /api/translate/formats` - List supported formats

### Health Check
- `GET /health` - Health check endpoint

## Testing

Run tests:
```bash
pytest
```

Run with coverage:
```bash
pytest --cov=app
```

## Configuration

Key environment variables:

- `DEBUG` - Enable debug mode (default: false)
- `SECRET_KEY` - JWT secret key
- `OPENAI_API_KEY` - OpenAI API key
- `ALLOWED_ORIGINS` - CORS allowed origins
- `MAX_FILE_SIZE` - Maximum upload size (default: 100MB)
- `MAX_CONCURRENT_JOBS` - Maximum concurrent jobs (default: 5)

## Architecture

The backend consists of:

- **app/main.py** - FastAPI application entry point
- **app/api/** - API routers
- **app/core/** - Core functionality (config, job manager)
- **app/models/** - Pydantic models
- **app/services/** - Business logic services

The backend wraps the existing translation scripts in `scripts/` directory and provides a REST API for the frontend.