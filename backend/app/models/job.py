from pydantic import BaseModel, Field
from typing import Optional, Dict, Any, List
from datetime import datetime
from enum import Enum

class JobStatus(str, Enum):
    PENDING = "pending"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"

class TranslationRequest(BaseModel):
    file_type: str = Field(..., description="Type of file (pptx or pdf)")
    model: str = Field(default="gpt-4o-2024-08-06", description="OpenAI model to use")
    temperature: float = Field(default=0.6, ge=0.0, le=2.0, description="Model temperature")
    offline: bool = Field(default=False, description="Use offline translation cache")
    pages: Optional[str] = Field(None, description="Page range for PDF (e.g., '1-10')")
    auto_fit: str = Field(default="norm", description="Auto-fit mode (norm, shape, none)")

class Job(BaseModel):
    id: str = Field(..., description="Unique job identifier")
    user_id: str = Field(..., description="User ID who created the job")
    status: JobStatus = Field(..., description="Current job status")
    input_file: str = Field(..., description="Input file path")
    output_file: Optional[str] = Field(None, description="Output file path")
    request: TranslationRequest = Field(..., description="Translation request parameters")
    progress: float = Field(default=0.0, ge=0.0, le=100.0, description="Progress percentage")
    message: Optional[str] = Field(None, description="Status message")
    error: Optional[str] = Field(None, description="Error message if failed")
    created_at: datetime = Field(..., description="Job creation timestamp")
    started_at: Optional[datetime] = Field(None, description="Job start timestamp")
    completed_at: Optional[datetime] = Field(None, description="Job completion timestamp")
    metadata: Dict[str, Any] = Field(default_factory=dict, description="Additional job metadata")
    estimated_cost: Optional[float] = Field(None, description="Estimated translation cost")
    actual_cost: Optional[float] = Field(None, description="Actual translation cost")
    quality_metrics: Optional[Dict[str, Any]] = Field(None, description="Quality assessment metrics")

    class Config:
        json_encoders = {
            datetime: lambda v: v.isoformat()
        }

class JobResponse(BaseModel):
    job: Job
    message: str = Field(default="Job created successfully")