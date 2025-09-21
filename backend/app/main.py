#!/usr/bin/env python3
"""
Minimal FastAPI app for DOCX translation integration tests.
"""

from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import tempfile
import uuid
from pathlib import Path
import json
from typing import Dict, List, Optional
import shutil

# Create app
app = FastAPI(
    title="DOCX Translation API",
    description="API for translating DOCX documents",
    version="0.1.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory job storage for testing
jobs: Dict[str, Dict] = {}

# Create artifacts directory
ARTIFACTS_DIR = Path("artifacts")
ARTIFACTS_DIR.mkdir(exist_ok=True)

@app.get("/healthz")
async def health_check():
    """Health check endpoint."""
    return {"status": "healthy"}

@app.get("/readyz")
async def readiness_check():
    """Readiness check endpoint."""
    return {"status": "ready"}

@app.post("/api/translate")
async def translate_document(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    model: str = "gpt-4",
    source_lang: str = "auto",
    target_lang: str = "en",
    glossary_id: Optional[str] = None
):
    """Upload and translate a DOCX document."""

    # Validate file type
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Only DOCX files are supported")

    # Validate file size (50MB limit)
    max_size = 50 * 1024 * 1024
    file_size = 0
    contents = b""

    # Read file in chunks to check size
    for chunk in file.file:
        file_size += len(chunk)
        if file_size > max_size:
            raise HTTPException(status_code=413, detail="File too large (max 50MB)")
        contents += chunk

    # Reset file pointer
    file.file.seek(0)

    # Create job ID
    job_id = str(uuid.uuid4())

    # Create job directory
    job_dir = Path("uploads") / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    # Save uploaded file
    input_path = job_dir / "test.docx"
    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # Create job info
    job_info = {
        "id": job_id,
        "status": "pending",
        "input_file": str(input_path),
        "output_file": str(job_dir / "translated_test.docx"),
        "model": model,
        "source_lang": source_lang,
        "target_lang": target_lang,
        "glossary_id": glossary_id,
        "created_at": "2025-09-21T00:00:00Z",
        "updated_at": "2025-09-21T00:00:00Z",
        "progress": 0,
        "segments_translated": 0,
        "total_segments": 0,
        "words_translated": 0,
        "total_words": 0,
        "error": None,
        "artifacts": {}
    }

    # Store job
    jobs[job_id] = job_info

    # Start background translation (mock)
    background_tasks.add_task(mock_translate, job_id)

    return {"job_id": job_id, "status": "pending"}

@app.get("/api/jobs/{job_id}")
async def get_job_status(job_id: str):
    """Get job status."""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")

    return jobs[job_id]

@app.get("/api/jobs/{job_id}/download")
async def download_result(job_id: str):
    """Download translated document."""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")

    job = jobs[job_id]

    if job["status"] != "completed":
        raise HTTPException(status_code=400, detail="Translation not completed")

    output_path = Path(job["output_file"])
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="Output file not found")

    return FileResponse(
        path=str(output_path),
        filename=f"translated_{job_id}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/api/glossaries")
async def list_glossaries():
    """List available glossaries."""
    return {"glossaries": []}

@app.get("/v1/config")
async def get_config():
    """Get configuration."""
    return {
        "max_file_size": 50 * 1024 * 1024,
        "supported_formats": [".docx"],
        "default_model": "gpt-4",
        "available_models": ["gpt-4", "gpt-3.5-turbo"]
    }

async def mock_translate(job_id: str):
    """Mock translation function."""
    import asyncio

    job = jobs[job_id]

    # Update status to processing
    job["status"] = "processing"
    job["updated_at"] = "2025-09-21T00:00:01Z"

    # Simulate translation progress
    for i in range(1, 6):
        await asyncio.sleep(0.1)
        job["progress"] = i * 20
        job["updated_at"] = f"2025-09-21T00:00:0{i}Z"

    # Copy input to output (mock translation)
    input_path = Path(job["input_file"])
    output_path = Path(job["output_file"])
    shutil.copy2(input_path, output_path)

    # Create audit JSON
    audit_path = Path(input_path.parent) / "audit.json"
    audit_data = {
        "job_id": job_id,
        "segments": [
            {
                "id": "1",
                "text": "これはテストです。",
                "translation": "This is a test.",
                "metadata": {"position": "0:0"}
            }
        ],
        "total_segments": 1,
        "words_translated": 4,
        "total_words": 4
    }
    with open(audit_path, "w", encoding="utf-8") as f:
        json.dump(audit_data, f, indent=2, ensure_ascii=False)

    # Update job to completed
    job["status"] = "completed"
    job["progress"] = 100
    job["segments_translated"] = 1
    job["total_segments"] = 1
    job["words_translated"] = 4
    job["total_words"] = 4
    job["updated_at"] = "2025-09-21T00:00:06Z"
    job["artifacts"] = {"audit": str(audit_path)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)