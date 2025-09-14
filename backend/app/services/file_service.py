import os
import uuid
import shutil
from pathlib import Path
from typing import Optional
from fastapi import HTTPException, UploadFile, status
from ..core.config import settings

class FileService:
    def __init__(self):
        # Ensure directories exist
        os.makedirs(settings.UPLOAD_DIR, exist_ok=True)
        os.makedirs(settings.OUTPUT_DIR, exist_ok=True)

    async def save_upload_file(
        self,
        file: UploadFile,
        user_id: str
    ) -> str:
        """Save uploaded file and return file path"""
        # Check file size
        file.file.seek(0, 2)  # Seek to end
        file_size = file.file.tell()
        file.file.seek(0)  # Reset position

        if file_size > settings.MAX_FILE_SIZE:
            raise HTTPException(
                status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                detail=f"File size exceeds maximum limit of {settings.MAX_FILE_SIZE} bytes"
            )

        # Check file extension
        allowed_extensions = {".pptx", ".pdf"}
        file_ext = Path(file.filename).suffix.lower()
        if file_ext not in allowed_extensions:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"File type {file_ext} not supported. Supported types: {', '.join(allowed_extensions)}"
            )

        # Generate unique filename
        file_id = str(uuid.uuid4())
        safe_filename = f"{file_id}_{file.filename}"
        file_path = Path(settings.UPLOAD_DIR) / safe_filename

        # Save file
        try:
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            return str(file_path)
        except Exception as e:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Failed to save file: {str(e)}"
            )

    async def get_file_path(self, file_id: str) -> Optional[str]:
        """Get file path by file ID"""
        # Search in both upload and output directories
        for directory in [settings.UPLOAD_DIR, settings.OUTPUT_DIR]:
            for file_path in Path(directory).glob(f"*{file_id}*"):
                return str(file_path)
        return None

    async def delete_file(self, file_path: str) -> bool:
        """Delete a file"""
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
            return False
        except Exception:
            return False

    async def cleanup_old_files(self, days: int = 7):
        """Clean up old files"""
        import time
        current_time = time.time()
        cutoff_time = current_time - (days * 24 * 60 * 60)

        for directory in [settings.UPLOAD_DIR, settings.OUTPUT_DIR]:
            for file_path in Path(directory).glob("*"):
                if file_path.is_file() and file_path.stat().st_mtime < cutoff_time:
                    try:
                        file_path.unlink()
                    except Exception as e:
                        print(f"Failed to delete {file_path}: {e}")