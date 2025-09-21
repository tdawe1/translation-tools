import pytest
import os
from unittest.mock import Mock, patch, AsyncMock
from datetime import datetime, timedelta

from backend.app.core.drive_poller import DrivePoller

@pytest.fixture
def mock_env(monkeypatch):
    monkeypatch.setenv("DRIVE_MONITORED_FOLDER_ID", "test_folder_id")
    monkeypatch.setenv("GOOGLE_DRIVE_CLIENT_ID", "test_client_id")
    monkeypatch.setenv("GOOGLE_DRIVE_CLIENT_SECRET", "test_secret")
    monkeypatch.setenv("GOOGLE_DRIVE_REFRESH_TOKEN", "test_token")

@pytest.mark.asyncio
async def test_drive_poller_initialization(mock_env):
    with patch("google.oauth2.credentials.Credentials.from_authorized_user_info") as mock_creds:
        with patch("googleapiclient.discovery.build") as mock_build:
            poller = DrivePoller()
            assert poller.monitored_folder_id == "test_folder_id"
            mock_creds.assert_called_once()
            mock_build.assert_called_once_with("drive", "v3", credentials=mock_creds.return_value)

@pytest.mark.asyncio
async def test_poll_drive_no_new_files(mock_env):
    poller = DrivePoller()
    poller.service = Mock()
    poller.service.files.return_value.list.return_value.execute.return_value = {"files": []}
    poller.processed_files = set(["old_file"])

    with patch.object(poller, "_validate_manifest") as mock_validate:
        mock_validate.return_value = True
        with patch("backend.app.api.jobs.create_drive_job") as mock_create:
            mock_create.return_value = None
            await poller.poll_drive()
            mock_create.assert_not_called()

@pytest.mark.asyncio
async def test_poll_drive_new_file(mock_env):
    poller = DrivePoller()
    poller.service = Mock()
    test_file = {"id": "new_file_id", "name": "test.pptx", "modifiedTime": "2023-01-01T00:00:00Z"}
    poller.service.files.return_value.list.return_value.execute.return_value = {"files": [test_file]}
    poller.last_poll = datetime.utcnow() - timedelta(hours=1)
    poller.processed_files = set()

    with patch.object(poller, "_validate_manifest") as mock_validate:
        mock_validate.return_value = True
        with patch("backend.app.api.jobs.create_drive_job") as mock_create:
            mock_job = Mock()
            mock_job.id = "test_job_id"
            mock_create.return_value = mock_job
            await poller.poll_drive()
            mock_create.assert_called_once()
            assert "new_file_id" in poller.processed_files

@pytest.mark.asyncio
async def test_poll_drive_already_processed(mock_env):
    poller = DrivePoller()
    poller.service = Mock()
    test_file = {"id": "processed_file_id", "name": "test.pptx", "modifiedTime": "2023-01-01T00:00:00Z"}
    poller.service.files.return_value.list.return_value.execute.return_value = {"files": [test_file]}
    poller.processed_files = {"processed_file_id"}

    with patch("backend.app.api.jobs.create_drive_job") as mock_create:
        await poller.poll_drive()
        mock_create.assert_not_called()

@pytest.mark.asyncio
async def test_scheduler_start_stop(mock_env):
    poller = DrivePoller()
    with patch.object(poller.scheduler, "add_job") as mock_add:
        with patch.object(poller.scheduler, "start") as mock_start:
            poller.start()
            mock_add.assert_called_once()
            mock_start.assert_called_once()
    poller.stop()
    poller.scheduler.shutdown.assert_called_once()

@pytest.mark.asyncio
async def test_manifest_validation(mock_env):
    poller = DrivePoller()
    valid_manifest = {
        "type": "drive",
        "drive_file_id": "123",
        "name": "test.pptx",
        "idempotency_key": "key123"
    }
    assert poller._validate_manifest(valid_manifest) == True

    invalid_manifest = {"type": "invalid"}
    assert poller._validate_manifest(invalid_manifest) == False