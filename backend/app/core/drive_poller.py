import os
import logging
from datetime import datetime, timedelta
from pydantic import Field
from ..models.job import JobResponse  # Assuming
from typing import Dict, Any

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from apscheduler.schedulers.asyncio import AsyncIOScheduler

from ..api.jobs import create_drive_job, ManifestRequest
from ..api.websocket import manager

logger = logging.getLogger(__name__)

class DrivePoller:
    def __init__(self):
        self.service = None
        self.scheduler = AsyncIOScheduler()
        self.processed_files = set()  # In-memory for low-capacity
        self.last_poll = None
        self.monitored_folder_id = os.getenv('DRIVE_MONITORED_FOLDER_ID')
        if not self.monitored_folder_id:
            raise ValueError("DRIVE_MONITORED_FOLDER_ID env var required")
        self.setup_drive_service()

    def setup_drive_service(self):
        """Setup Google Drive service with OAuth credentials from env."""
        token_info = {
            'client_id': os.getenv('GOOGLE_DRIVE_CLIENT_ID'),
            'client_secret': os.getenv('GOOGLE_DRIVE_CLIENT_SECRET'),
            'refresh_token': os.getenv('GOOGLE_DRIVE_REFRESH_TOKEN'),
            'token_uri': 'https://oauth2.googleapis.com/token',
        }
        creds = Credentials.from_authorized_user_info(token_info, scopes=['https://www.googleapis.com/auth/drive.readonly'])
        self.service = build('drive', 'v3', credentials=creds)

    async def poll_drive(self):
        """Poll Drive for new files and enqueue jobs."""
        try:
            logger.info("Starting Drive poll")
            self.last_poll = datetime.utcnow()

            # List files in monitored folder, modified since last poll
            query = f"'{self.monitored_folder_id}' in parents and trashed=false"
            if self.last_poll:
                query += f" and modifiedTime > '{self.last_poll.isoformat()}Z'"
            results = self.service.files().list(q=query, fields="files(id, name, modifiedTime)").execute()
            files = results.get('files', [])

            for file in files:
                file_id = file['id']
                if file_id in self.processed_files:
                    continue

                # Create manifest
                manifest = {
                    "type": "drive",
                    "drive_file_id": file_id,
                    "name": file['name'],
                    "modified_time": file['modifiedTime'],
                    "idempotency_key": f"drive_{file_id}"  # Simple key
                }

                # Validate manifest (basic)
                if not self._validate_manifest(manifest):
                    logger.warning(f"Invalid manifest for file {file_id}")
                    continue

                # Enqueue job
                job_request = ManifestRequest(**manifest)
                job = await create_drive_job(job_request)
                if job:
                    self.processed_files.add(file_id)
                    logger.info(f"Enqueued drive job for {file['name']}: {job.id}")

            logger.info(f"Drive poll completed, processed {len(files)} files")
        except HttpError as e:
            logger.error(f"Drive API error: {e}")
        except Exception as e:
            logger.error(f"Poll error: {e}")

    def _validate_manifest(self, manifest: Dict[str, Any]) -> bool:
        """Basic manifest validation."""
        required = ["type", "drive_file_id", "name", "idempotency_key"]
        return manifest.get("type") == "drive" and all(manifest.get(k) for k in required)

    def start(self):
        """Start the scheduler."""
        self.scheduler.add_job(self.poll_drive, 'interval', minutes=5)
        self.scheduler.start()
        logger.info("Drive poller started")

    def stop(self):
        """Stop the scheduler."""
        self.scheduler.shutdown()
        logger.info("Drive poller stopped")
