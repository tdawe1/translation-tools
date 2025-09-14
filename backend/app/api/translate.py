from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from typing import Optional, Dict, Any

from ..models.job import TranslationRequest, JobResponse
from ..core.job_manager import job_manager
from ..services.auth_service import auth_service
from ..services.file_service import file_service

router = APIRouter()
security = HTTPBearer()

@router.post("/translate", response_model=JobResponse)
async def create_translation_job(
    file: UploadFile = File(...),
    file_type: str = "pptx",
    model: str = "gpt-4o-2024-08-06",
    temperature: float = 0.6,
    offline: bool = False,
    pages: Optional[str] = None,
    auto_fit: str = "norm",
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """Create a new translation job"""
    # Verify token and get user ID
    user_id = auth_service.verify_token(credentials.credentials)

    # Validate file type
    if file_type not in ["pptx", "pdf"]:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="file_type must be either 'pptx' or 'pdf'"
        )

    # Validate auto_fit mode
    if auto_fit not in ["norm", "shape", "none"]:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="auto_fit must be one of: norm, shape, none"
        )

    # Save uploaded file
    try:
        input_file_path = await file_service.save_upload_file(file, user_id)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to save file: {str(e)}"
        )

    # Create translation request
    request = TranslationRequest(
        file_type=file_type,
        model=model,
        temperature=temperature,
        offline=offline,
        pages=pages,
        auto_fit=auto_fit
    )

    # Create job
    job = await job_manager.create_job(
        user_id=user_id,
        input_file=input_file_path,
        request=request
    )

    return JobResponse(job=job)

@router.get("/translate/models")
async def list_available_models(
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """List available translation models"""
    # Verify token
    auth_service.verify_token(credentials.credentials)

    # Return available models
    models = [
        {
            "id": "gpt-4o-2024-08-06",
            "name": "GPT-4o (Latest)",
            "description": "Most capable model, best for complex translations",
            "pricing": "$5.00 / 1M input tokens, $15.00 / 1M output tokens"
        },
        {
            "id": "gpt-4o-mini",
            "name": "GPT-4o Mini",
            "description": "Cost-effective option for simpler translations",
            "pricing": "$0.15 / 1M input tokens, $0.60 / 1M output tokens"
        },
        {
            "id": "gpt-5",
            "name": "GPT-5",
            "description": "Most advanced model (if available)",
            "pricing": "$15.00 / 1M input tokens, $75.00 / 1M output tokens"
        }
    ]

    return {"models": models}

@router.get("/translate/formats")
async def list_supported_formats(
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """List supported file formats and their options"""
    # Verify token
    auth_service.verify_token(credentials.credentials)

    formats = {
        "pptx": {
            "name": "PowerPoint Presentation",
            "extensions": [".pptx"],
            "max_size": "100MB",
            "options": {
                "model": ["gpt-4o-2024-08-06", "gpt-4o-mini", "gpt-5"],
                "temperature": {"min": 0.0, "max": 2.0, "default": 0.6},
                "offline": True,
                "auto_fit": ["norm", "shape", "none"]
            }
        },
        "pdf": {
            "name": "PDF Document",
            "extensions": [".pdf"],
            "max_size": "100MB",
            "options": {
                "model": ["gpt-4o-2024-08-06", "gpt-4o-mini", "gpt-5"],
                "temperature": {"min": 0.0, "max": 2.0, "default": 0.6},
                "offline": True,
                "pages": "Page range (e.g., '1-10', '1,3,5-7')",
                "auto_fit": ["norm", "shape", "none"]
            }
        }
    }

    return {"formats": formats}