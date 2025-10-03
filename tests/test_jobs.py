import pytest
from unittest.mock import AsyncMock, patch, MagicMock
from fastapi.testclient import TestClient
from backend.app.main import app
from backend.app.core.job_manager import job_manager
from backend.app.api.websocket import manager, batch_progress
from backend.app.models.job import TranslationRequest
from fastapi import HTTPException

client = TestClient(app)

@pytest.fixture
def mock_job_manager(mocker):
    mocker.patch.object(job_manager, 'create_job', new_callable=AsyncMock)
    mocker.patch.object(job_manager, 'get_job', new_callable=AsyncMock)
    job_manager.create_job.return_value = MagicMock(id='test-job', dict=lambda: {'id': 'test-job', 'status': 'pending'})
    job_manager.get_job.return_value = None
    yield job_manager

@pytest.fixture
def mock_manager(mocker):
    mocker.patch.object(manager, 'broadcast_batch_update', new_callable=AsyncMock)
    mocker.patch.object(manager, 'broadcast_drive_update', new_callable=AsyncMock)
    yield manager

def test_create_single_drive_job(mock_job_manager, mock_manager):
    manifest = {
        "type": "drive",
        "drive_file_id": "test123",
        "name": "test.pptx",
        "idempotency_key": "single-key-123"
    }
    response = client.post("/api/jobs", json=manifest)
    assert response.status_code == 200
    data = response.json()
    assert data['id'] == 'test-job'
    mock_job_manager.create_job.assert_called_once()
    mock_manager.broadcast_drive_update.assert_called_once_with('test-job', 'queued', manifest)

def test_create_batch_manifest(mock_job_manager, mock_manager):
    manifest = {
        "type": "batch",
        "idempotency_key": "batch-key-123",
        "jobs": [
            {
                "input": "drive://id1",
                "file_type": "pptx",
                "model": "gpt-4o-mini"
            },
            {
                "input": "drive://id2",
                "file_type": "pdf",
                "model": "gpt-4o-mini",
                "pages": "1-10"
            }
        ]
    }
    with patch('backend.app.api.jobs.batch_progress') as mock_bp:
        mock_bp.__setitem__.return_value = None  # mock set
        response = client.post("/api/jobs", json=manifest)
        assert response.status_code == 200
        data = response.json()
        assert data['batch_id'] == 'batch-key-123'
        assert data['num_jobs'] == 2
        assert len(data['job_ids']) == 2
        assert mock_job_manager.create_job.call_count == 2
        mock_manager.broadcast_batch_update.assert_called_once_with('batch-key-123', 'queued', 0.0)

def test_batch_reuse(mock_job_manager, mock_manager):
    manifest = {
        "type": "batch",
        "idempotency_key": "reuse-key",
        "jobs": [{"input": "drive://id1", "file_type": "pptx", "model": "gpt-4o-mini"}]
    }
    # First creation
    response1 = client.post("/api/jobs", json=manifest)
    assert response1.status_code == 200
    assert response1.json()['status'] == 'queued'
    # Second, reuse
    response2 = client.post("/api/jobs", json=manifest)
    assert response2.status_code == 200
    assert response2.json()['status'] == 'reused'

def test_invalid_manifest():
    invalid = {"type": "invalid"}
    response = client.post("/api/jobs", json=invalid)
    assert response.status_code == 400
