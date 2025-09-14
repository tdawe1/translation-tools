from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import os
import uuid
import json
from typing import Dict, List
import subprocess
import asyncio

app = FastAPI(title="Translation Pipeline API")

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "http://localhost:3001"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Job storage
jobs: Dict[str, Dict] = {}

@app.get("/health")
async def health_check():
    return {"status": "healthy"}

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(('.pptx', '.pdf')):
        raise HTTPException(status_code=400, detail="Only PPTX and PDF files are supported")

    file_id = str(uuid.uuid4())
    uploads_dir = os.path.join(os.path.dirname(__file__), "uploads")
    os.makedirs(uploads_dir, exist_ok=True)
    file_path = os.path.join(uploads_dir, f"{file_id}_{file.filename}")

    with open(file_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    return {"file_id": file_id, "filename": file.filename, "path": file_path}

@app.post("/translate")
async def translate_file(
    file_id: str,
    filename: str,
    model: str = "gpt-4o",
    background_tasks: BackgroundTasks = None
):
    job_id = str(uuid.uuid4())

    jobs[job_id] = {
        "id": job_id,
        "status": "pending",
        "filename": filename,
        "model": model,
        "progress": 0,
        "created_at": "2024-01-01T00:00:00Z"
    }

    background_tasks.add_task(run_translation, job_id, file_id, filename, model)

    return {"job_id": job_id, "status": "started"}

async def run_translation(job_id: str, file_id: str, filename: str, model: str):
    try:
        jobs[job_id]["status"] = "running"

        # Find the uploaded file
        upload_path = os.path.join(os.path.dirname(__file__), "uploads")
        input_file = None
        if os.path.exists(upload_path):
            for f in os.listdir(upload_path):
                if f.startswith(file_id):
                    input_file = os.path.join(upload_path, f)
                    break

        if not input_file:
            jobs[job_id]["status"] = "failed"
            jobs[job_id]["error"] = "File not found"
            return

        # Update progress
        jobs[job_id]["progress"] = 25

        # Run translation script
        results_dir = os.path.join(os.path.dirname(__file__), "results")
        os.makedirs(results_dir, exist_ok=True)
        output_file = os.path.join(results_dir, f"{job_id}_{filename}")

        # Get the project root directory (parent of backend)
        base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

        if filename.lower().endswith('.pptx'):
            cmd = [
                "python", f"{base_dir}/scripts/translate_pptx_inplace.py",
                "--in", input_file,
                "--out", output_file,
                "--model", model
            ]
        else:
            cmd = [
                "python", f"{base_dir}/scripts/translate_pdf.py",
                "--in", input_file,
                "--out", output_file,
                "--model", model
            ]

        jobs[job_id]["progress"] = 50

        process = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE
        )

        stdout, stderr = await process.communicate()

        jobs[job_id]["progress"] = 90

        if process.returncode == 0:
            jobs[job_id]["status"] = "completed"
            jobs[job_id]["progress"] = 100
            jobs[job_id]["output_file"] = output_file
        else:
            jobs[job_id]["status"] = "failed"
            jobs[job_id]["error"] = stderr.decode()

    except Exception as e:
        jobs[job_id]["status"] = "failed"
        jobs[job_id]["error"] = str(e)

@app.get("/jobs/{job_id}")
async def get_job_status(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    return jobs[job_id]

@app.get("/jobs")
async def list_jobs():
    return list(jobs.values())

@app.get("/jobs/{job_id}/download")
async def download_result(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")

    job = jobs[job_id]
    if job["status"] != "completed" or "output_file" not in job:
        raise HTTPException(status_code=400, detail="Job not completed")

    if not os.path.exists(job["output_file"]):
        raise HTTPException(status_code=404, detail="Result file not found")

    return FileResponse(
        job["output_file"],
        filename=job["filename"],
        media_type='application/octet-stream'
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)