from fastapi import APIRouter, Depends, HTTPException, status, Query, Body
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from typing import Optional, Dict, Any, List
from datetime import datetime, timedelta
import io
import sqlite3
import logging

from ..models.job import TranslationRequest, JobResponse
from ..core.job_manager import job_manager
from ..services.auth_service import auth_service
from pydantic import BaseModel, Field

logger = logging.getLogger(__name__)

router = APIRouter()
security = HTTPBearer()

class JobSearchRequest(BaseModel):
    """Request model for job search"""
    search: Optional[str] = Field(None, description="Search in file names and messages")
    status: Optional[str] = Field(None, description="Filter by status")
    file_type: Optional[str] = Field(None, description="Filter by file type")
    date_from: Optional[str] = Field(None, description="Start date (YYYY-MM-DD)")
    date_to: Optional[str] = Field(None, description="End date (YYYY-MM-DD)")
    sort_by: str = Field("created_at", description="Sort field")
    sort_order: str = Field("desc", description="Sort order (asc/desc)")
    page: int = Field(1, ge=1, description="Page number")
    page_size: int = Field(20, ge=1, le=100, description="Page size")

class BulkJobRequest(BaseModel):
    """Request model for bulk operations"""
    job_ids: List[str] = Field(..., description="List of job IDs")

class JobStatisticsResponse(BaseModel):
    """Response model for job statistics"""
    total_jobs: int
    status_counts: Dict[str, int]
    average_duration_minutes: float
    total_cost: float
    daily_stats: List[Dict[str, Any]]
    file_type_distribution: Dict[str, int]
    period_days: int

@router.get("/jobs", response_model=Dict[str, Any])
async def list_jobs(
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=100),
    status: Optional[str] = Query(None),
    file_type: Optional[str] = Query(None),
    search: Optional[str] = Query(None),
    sort_by: str = Query("created_at"),
    sort_order: str = Query("desc"),
    date_from: Optional[str] = Query(None),
    date_to: Optional[str] = Query(None),
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """List jobs with filtering and pagination"""
    user_id = auth_service.verify_token(credentials.credentials)

    filters = {}
    if status:
        filters["status"] = status
    if file_type:
        filters["file_type"] = file_type
    if search:
        filters["search"] = search
    if date_from:
        filters["date_from"] = date_from + "T00:00:00"
    if date_to:
        filters["date_to"] = date_to + "T23:59:59"

    skip = (page - 1) * page_size
    jobs, total = await job_manager.list_jobs(
        user_id=user_id,
        skip=skip,
        limit=page_size,
        filters=filters,
        sort_by=sort_by,
        sort_order=sort_order
    )

    return {
        "jobs": [job.dict() for job in jobs],
        "pagination": {
            "page": page,
            "page_size": page_size,
            "total": total,
            "pages": (total + page_size - 1) // page_size
        }
    }

@router.post("/jobs/search", response_model=Dict[str, Any])
async def search_jobs(
    request: JobSearchRequest,
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Search jobs with advanced filters"""
    user_id = auth_service.verify_token(credentials.credentials)

    filters = {}
    if request.search:
        filters["search"] = request.search
    if request.status:
        filters["status"] = request.status
    if request.file_type:
        filters["file_type"] = request.file_type
    if request.date_from:
        filters["date_from"] = request.date_from + "T00:00:00"
    if request.date_to:
        filters["date_to"] = request.date_to + "T23:59:59"

    skip = (request.page - 1) * request.page_size
    jobs, total = await job_manager.list_jobs(
        user_id=user_id,
        skip=skip,
        limit=request.page_size,
        filters=filters,
        sort_by=request.sort_by,
        sort_order=request.sort_order
    )

    return {
        "jobs": [job.dict() for job in jobs],
        "pagination": {
            "page": request.page,
            "page_size": request.page_size,
            "total": total,
            "pages": (total + request.page_size - 1) // request.page_size
        }
    }

@router.get("/jobs/{job_id}", response_model=Dict[str, Any])
async def get_job_details(
    job_id: str,
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Get detailed job information including logs"""
    user_id = auth_service.verify_token(credentials.credentials)

    job = await job_manager.get_job(job_id)
    if not job or job.user_id != user_id:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Job not found"
        )

    logs = await job_manager.get_job_logs(job_id)

    return {
        "job": job.dict(),
        "logs": logs
    }

@router.post("/jobs/{job_id}/cancel", response_model=Dict[str, str])
async def cancel_job(
    job_id: str,
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Cancel a specific job"""
    user_id = auth_service.verify_token(credentials.credentials)

    job = await job_manager.get_job(job_id)
    if not job or job.user_id != user_id:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Job not found"
        )

    success = await job_manager.cancel_job(job_id)
    if not success:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Job cannot be cancelled"
        )

    return {"message": "Job cancelled successfully"}

@router.post("/jobs/{job_id}/retry", response_model=Dict[str, str])
async def retry_job(
    job_id: str,
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Retry a failed job"""
    user_id = auth_service.verify_token(credentials.credentials)

    job = await job_manager.get_job(job_id)
    if not job or job.user_id != user_id:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Job not found"
        )

    if job.status != "failed":
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Only failed jobs can be retried"
        )

    new_job = await job_manager.create_job(
        user_id=user_id,
        input_file=job.input_file,
        request=job.request
    )

    return {"message": "Job retried", "job_id": new_job.id}

@router.post("/jobs/bulk/cancel", response_model=Dict[str, Any])
async def cancel_jobs_bulk(
    request: BulkJobRequest,
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Cancel multiple jobs"""
    user_id = auth_service.verify_token(credentials.credentials)

    results = await job_manager.cancel_jobs(user_id, request.job_ids)

    cancelled_count = sum(1 for success in results.values() if success)
    return {
        "message": f"Cancelled {cancelled_count} of {len(request.job_ids)} jobs",
        "results": results
    }

@router.post("/jobs/bulk/retry", response_model=Dict[str, Any])
async def retry_jobs_bulk(
    request: BulkJobRequest,
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Retry multiple failed jobs"""
    user_id = auth_service.verify_token(credentials.credentials)

    retried_jobs = await job_manager.retry_jobs(user_id, request.job_ids)

    return {
        "message": f"Retried {len(retried_jobs)} jobs",
        "retried_job_ids": retried_jobs
    }

@router.get("/jobs/statistics", response_model=JobStatisticsResponse)
async def get_job_statistics(
    days: int = Query(30, ge=1, le=365),
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Get job statistics for the user"""
    user_id = auth_service.verify_token(credentials.credentials)

    stats = await job_manager.get_job_statistics(user_id, days)
    return JobStatisticsResponse(**stats)

@router.get("/jobs/queue", response_model=Dict[str, Any])
async def get_queue_status(
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Get current queue status"""
    user_id = auth_service.verify_token(credentials.credentials)

    # Get current job counts
    jobs, _ = await job_manager.list_jobs(user_id, limit=10000)

    status_counts = {}
    for job in jobs:
        status_counts[job.status] = status_counts.get(job.status, 0) + 1

    # Get active jobs
    active_jobs = [job for job in jobs if job.status in ["pending", "running"]]

    return {
        "status_counts": status_counts,
        "active_jobs": len(active_jobs),
        "total_jobs": len(jobs)
    }

@router.get("/jobs/{job_id}/logs", response_model=List[Dict[str, Any]])
async def get_job_logs(
    job_id: str,
    limit: int = Query(100, ge=1, le=1000),
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Get logs for a specific job"""
    user_id = auth_service.verify_token(credentials.credentials)

    job = await job_manager.get_job(job_id)
    if not job or job.user_id != user_id:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Job not found"
        )

    logs = await job_manager.get_job_logs(job_id)
    return logs[:limit]

@router.get("/jobs/export")
async def export_jobs(
    format: str = Query("csv", regex="^(csv|json)$"),
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Export job data"""
    user_id = auth_service.verify_token(credentials.credentials)

    try:
        data = await job_manager.export_job_report(user_id, format)

        filename = f"translation_jobs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{format}"

        if format == "csv":
            return {
                "data": data,
                "filename": filename,
                "media_type": "text/csv"
            }
        else:
            return {
                "data": data,
                "filename": filename,
                "media_type": "application/json"
            }

    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to export jobs: {str(e)}"
        )

@router.delete("/jobs/{job_id}", response_model=Dict[str, str])
async def delete_job(
    job_id: str,
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Delete a job (admin only or own job)"""
    user_id = auth_service.verify_token(credentials.credentials)

    job = await job_manager.get_job(job_id)
    if not job or job.user_id != user_id:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Job not found"
        )

    # Only allow deletion of completed or failed jobs
    if job.status not in ["completed", "failed", "cancelled"]:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Cannot delete active jobs"
        )

    # Remove from database
    try:
        conn = sqlite3.connect(str(job_manager.db_path))
        cursor = conn.cursor()

        cursor.execute("DELETE FROM job_logs WHERE job_id = ?", [job_id])
        cursor.execute("DELETE FROM jobs WHERE id = ? AND user_id = ?", [job_id, user_id])

        conn.commit()
        conn.close()

        # Remove from memory
        if job_id in job_manager.jobs:
            del job_manager.jobs[job_id]

        return {"message": "Job deleted successfully"}

    except Exception as e:
        logger.error(f"Failed to delete job: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
        detail="Failed to delete job"
    )

@router.post("/submit")
async def submit_job():
    """Stub endpoint for job submission"""
    return {"job_id": "stub-123", "status": "queued"}
