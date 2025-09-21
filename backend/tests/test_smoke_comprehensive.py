"""
Comprehensive smoke tests for the Translation Pipeline Backend API.

This test file validates the core API functionality end-to-end, including:
- User authentication (register, login, token refresh)
- Translation job management (create, list, get details, cancel)
- Error handling and edge cases
- Both PPTX and PDF translation workflows

Run with: python -m pytest tests/test_smoke_comprehensive.py -v
"""

import pytest
import tempfile
import os
import time
import json
from pathlib import Path
from unittest.mock import patch, AsyncMock
import uuid

# Test data
TEST_USER = {
    "email": "smoketest@example.com",
    "password": "SmokeTest123!",
    "full_name": "Smoke Test User"
}

TEST_ADMIN = {
    "email": "adminsmoke@example.com",
    "password": "AdminSmoke123!",
    "full_name": "Admin Smoke User"
}


@pytest.fixture
def sample_pptx_file():
    """Create a sample PPTX file with Japanese text for testing"""
    with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as f:
        pptx_path = f.name

    # Create a minimal valid PPTX file
    import zipfile
    from xml.etree.ElementTree import Element, SubElement, tostring

    with zipfile.ZipFile(pptx_path, 'w') as zf:
        # Create [Content_Types].xml
        content_types = Element('Types', xmlns='http://schemas.openxmlformats.org/package/2006/content-types')
        SubElement(content_types, 'Default', Extension='rels', ContentType='application/vnd.openxmlformats-package.relationships+xml')
        SubElement(content_types, 'Default', Extension='xml', ContentType='application/xml')
        SubElement(content_types, 'Override', PartName='/ppt/presentation.xml', ContentType='application/vnd.openxmlformats-presentationml.presentation.main+xml')
        zf.writestr('[Content_Types].xml', tostring(content_types))

        # Create _rels/.rels
        rels = Element('Relationships', xmlns='http://schemas.openxmlformats.org/package/2006/relationships')
        SubElement(rels, 'Relationship', Id='rId1', Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', Target='ppt/presentation.xml')
        zf.writestr('_rels/.rels', tostring(rels))

        # Create presentation with slides
        pres = Element('p:presentation', {
            'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        })

        # Add slide master
        sldMasterIdLst = SubElement(pres, 'p:sldMasterIdLst')
        SubElement(sldMasterIdLst, 'p:sldMasterId', {'id': '2147483648', 'r:id': 'rId1'})

        # Add slide
        sldIdLst = SubElement(pres, 'p:sldIdLst')
        SubElement(sldIdLst, 'p:sldId', {'id': '256', 'r:id': 'rId2'})

        zf.writestr('ppt/presentation.xml', tostring(pres))

        # Create slide with Japanese text
        slide = Element('p:sld', {
            'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        })
        cSld = SubElement(slide, 'p:cSld')
        spTree = SubElement(cSld, 'p:spTree')

        # Add text shape
        sp = SubElement(spTree, 'p:sp')
        txBody = SubElement(sp, 'p:txBody')
        bodyPr = SubElement(txBody, 'a:bodyPr')
        lstStyle = SubElement(txBody, 'a:lstStyle')
        p = SubElement(txBody, 'a:p')
        r = SubElement(p, 'a:r')
        t = SubElement(r, 'a:t')
        t.text = "これは日本語のテキストです。"

        zf.writestr('ppt/slides/slide1.xml', tostring(slide))

    yield pptx_path

    # Cleanup
    if os.path.exists(pptx_path):
        os.unlink(pptx_path)


@pytest.fixture
def sample_pdf_file():
    """Create a sample PDF file for testing"""
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as f:
        pdf_path = f.name

    # Create a minimal PDF file
    with open(pdf_path, 'wb') as f:
        f.write(b'%PDF-1.4\n')
        f.write(b'1 0 obj\n')
        f.write(b'<<\n')
        f.write(b'/Type /Catalog\n')
        f.write(b'/Pages 2 0 R\n')
        f.write(b'>>\n')
        f.write(b'endobj\n')
        f.write(b'2 0 obj\n')
        f.write(b'<<\n')
        f.write(b'/Type /Pages\n')
        f.write(b'/Kids [3 0 R]\n')
        f.write(b'/Count 1\n')
        f.write(b'>>\n')
        f.write(b'endobj\n')
        f.write(b'3 0 obj\n')
        f.write(b'<<\n')
        f.write(b'/Type /Page\n')
        f.write(b'/Parent 2 0 R\n')
        f.write(b'/MediaBox [0 0 612 792]\n')
        f.write(b'>>\n')
        f.write(b'endobj\n')
        f.write(b'xref\n')
        f.write(b'0 4\n')
        f.write(b'0000000000 65535 f \n')
        f.write(b'0000000009 00000 n \n')
        f.write(b'0000000058 00000 n \n')
        f.write(b'0000000115 00000 n \n')
        f.write(b'trailer\n')
        f.write(b'<<\n')
        f.write(b'/Size 4\n')
        f.write(b'/Root 1 0 R\n')
        f.write(b'>>\n')
        f.write(b'startxref\n')
        f.write(b'174\n')
        f.write(b'%%EOF\n')

    yield pdf_path

    # Cleanup
    if os.path.exists(pdf_path):
        os.unlink(pdf_path)


@pytest.fixture
def mock_translation_service():
    """Mock the translation service to avoid actual API calls"""
    with patch('app.services.translation_service.TranslationService') as mock:
        instance = mock.return_value
        instance.translate_document = AsyncMock(return_value={
            "status": "completed",
            "output_file": "/test/path/output.pptx",
            "cost": 0.50,
            "duration_seconds": 30,
            "tokens_used": {
                "input": 1000,
                "output": 500
            }
        })
        yield instance


class TestAuthenticationEndpoints:
    """Test authentication-related endpoints"""

    def test_user_registration(self, client):
        """Test user registration endpoint"""
        # Test successful registration
        response = client.post("/api/auth/register", json=TEST_USER)
        assert response.status_code == 200

        data = response.json()
        assert data["email"] == TEST_USER["email"]
        assert data["full_name"] == TEST_USER["full_name"]
        assert "id" in data
        assert data["is_active"] is True

        # Test duplicate email registration
        response = client.post("/api/auth/register", json=TEST_USER)
        assert response.status_code == 400
        assert "already registered" in response.json()["detail"]

    def test_user_login(self, client):
        """Test user login endpoint"""
        # First register a user
        client.post("/api/auth/register", json=TEST_USER)

        # Test successful login
        login_data = {
            "email": TEST_USER["email"],
            "password": TEST_USER["password"]
        }
        response = client.post("/api/auth/login", json=login_data)
        assert response.status_code == 200

        data = response.json()
        assert "access_token" in data
        assert "refresh_token" in data
        assert data["token_type"] == "bearer"
        assert "expires_in" in data

        # Test invalid credentials
        invalid_login = {
            "email": TEST_USER["email"],
            "password": "wrongpassword"
        }
        response = client.post("/api/auth/login", json=invalid_login)
        assert response.status_code == 401

        # Test non-existent user
        nonexistent_login = {
            "email": "nonexistent@example.com",
            "password": "anypassword"
        }
        response = client.post("/api/auth/login", json=nonexistent_login)
        assert response.status_code == 401

    def test_token_refresh(self, client):
        """Test token refresh endpoint"""
        # Register and login to get refresh token
        client.post("/api/auth/register", json=TEST_USER)
        login_response = client.post("/api/auth/login", json={
            "email": TEST_USER["email"],
            "password": TEST_USER["password"]
        })
        refresh_token = login_response.json()["refresh_token"]

        # Test token refresh
        response = client.post("/api/auth/refresh", params={"refresh_token": refresh_token})
        assert response.status_code == 200

        data = response.json()
        assert "access_token" in data
        assert data["token_type"] == "bearer"
        assert "expires_in" in data

        # Test invalid refresh token
        response = client.post("/api/auth/refresh", params={"refresh_token": "invalid_token"})
        assert response.status_code == 401

    def test_get_current_user(self, client):
        """Test getting current user info"""
        # Register and login
        client.post("/api/auth/register", json=TEST_USER)
        login_response = client.post("/api/auth/login", json={
            "email": TEST_USER["email"],
            "password": TEST_USER["password"]
        })
        access_token = login_response.json()["access_token"]
        headers = {"Authorization": f"Bearer {access_token}"}

        # Test getting user info
        response = client.get("/api/auth/me", headers=headers)
        assert response.status_code == 200

        data = response.json()
        assert data["email"] == TEST_USER["email"]
        assert data["full_name"] == TEST_USER["full_name"]
        assert "id" in data

    def test_logout(self, client):
        """Test user logout"""
        # Register and login
        client.post("/api/auth/register", json=TEST_USER)
        login_response = client.post("/api/auth/login", json={
            "email": TEST_USER["email"],
            "password": TEST_USER["password"]
        })
        refresh_token = login_response.json()["refresh_token"]

        # Test logout
        response = client.post("/api/auth/logout", params={"refresh_token": refresh_token})
        assert response.status_code == 200
        assert response.json()["message"] == "Successfully logged out"

        # Test using the same refresh token should fail
        response = client.post("/api/auth/refresh", params={"refresh_token": refresh_token})
        assert response.status_code == 401


class TestProtectedEndpoints:
    """Test that protected endpoints require authentication"""

    def test_unauthorized_access(self, client):
        """Test accessing protected endpoints without authentication"""
        protected_endpoints = [
            ("/api/jobs", "GET"),
            ("/api/translate/models", "GET"),
            ("/api/translate/formats", "GET"),
            ("/api/jobs/statistics", "GET"),
        ]

        for endpoint, method in protected_endpoints:
            if method == "GET":
                response = client.get(endpoint)
            else:
                response = client.post(endpoint)

            assert response.status_code == 403

    def test_invalid_token(self, client):
        """Test accessing endpoints with invalid token"""
        headers = {"Authorization": "Bearer invalid_token"}

        response = client.get("/api/jobs", headers=headers)
        assert response.status_code == 401


class TestTranslationEndpoints:
    """Test translation job endpoints"""

    def get_auth_headers(self, client, user_data=TEST_USER):
        """Helper method to get authentication headers"""
        # Register and login
        client.post("/api/auth/register", json=user_data)
        login_response = client.post("/api/auth/login", json={
            "email": user_data["email"],
            "password": user_data["password"]
        })
        access_token = login_response.json()["access_token"]
        return {"Authorization": f"Bearer {access_token}"}

    def test_get_translation_models(self, client):
        """Test getting available translation models"""
        headers = self.get_auth_headers(client)

        response = client.get("/api/translate/models", headers=headers)
        assert response.status_code == 200

        data = response.json()
        assert "models" in data
        assert len(data["models"]) > 0

        # Check model structure
        for model in data["models"]:
            assert "id" in model
            assert "name" in model
            assert "description" in model
            assert "pricing" in model

    def test_get_supported_formats(self, client):
        """Test getting supported file formats"""
        headers = self.get_auth_headers(client)

        response = client.get("/api/translate/formats", headers=headers)
        assert response.status_code == 200

        data = response.json()
        assert "formats" in data
        assert "pptx" in data["formats"]
        assert "pdf" in data["formats"]

        # Check format structure
        for format_name, format_info in data["formats"].items():
            assert "name" in format_info
            assert "extensions" in format_info
            assert "max_size" in format_info
            assert "options" in format_info

    def test_create_pptx_translation_job(self, client, sample_pptx_file, mock_translation_service):
        """Test creating a PPTX translation job"""
        headers = self.get_auth_headers(client)

        # Test with file upload
        with open(sample_pptx_file, 'rb') as f:
            files = {"file": ("test.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {
                "file_type": "pptx",
                "model": "gpt-4o-mini",
                "temperature": 0.6,
                "offline": False,
                "auto_fit": "norm"
            }

            response = client.post("/api/translate", files=files, data=data, headers=headers)
            assert response.status_code == 200

            job_data = response.json()
            assert "job" in job_data
            job = job_data["job"]
            assert "id" in job
            assert job["status"] == "pending"
            assert job["file_type"] == "pptx"
            assert job["user_id"] is not None

            return job["id"]

    def test_create_pdf_translation_job(self, client, sample_pdf_file, mock_translation_service):
        """Test creating a PDF translation job"""
        headers = self.get_auth_headers(client)

        # Test with file upload and pages parameter
        with open(sample_pdf_file, 'rb') as f:
            files = {"file": ("test.pdf", f, "application/pdf")}
            data = {
                "file_type": "pdf",
                "model": "gpt-4o-2024-08-06",
                "temperature": 0.7,
                "offline": False,
                "pages": "1-5",
                "auto_fit": "shape"
            }

            response = client.post("/api/translate", files=files, data=data, headers=headers)
            assert response.status_code == 200

            job_data = response.json()
            assert "job" in job_data
            job = job_data["job"]
            assert job["file_type"] == "pdf"
            assert job["request"]["pages"] == "1-5"
            assert job["request"]["auto_fit"] == "shape"

            return job["id"]

    def test_create_job_invalid_file_type(self, client, sample_pptx_file):
        """Test creating job with invalid file type"""
        headers = self.get_auth_headers(client)

        with open(sample_pptx_file, 'rb') as f:
            files = {"file": ("test.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {
                "file_type": "invalid",  # Invalid file type
                "model": "gpt-4o-mini"
            }

            response = client.post("/api/translate", files=files, data=data, headers=headers)
            assert response.status_code == 400
            assert "file_type must be either" in response.json()["detail"]

    def test_create_job_without_auth(self, client, sample_pptx_file):
        """Test creating job without authentication"""
        with open(sample_pptx_file, 'rb') as f:
            files = {"file": ("test.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {"file_type": "pptx"}

            response = client.post("/api/translate", files=files, data=data)
            assert response.status_code == 403


class TestJobManagementEndpoints:
    """Test job management endpoints"""

    def setup_job(self, client, sample_pptx_file, mock_translation_service):
        """Helper method to create a test job"""
        headers = self.get_auth_headers(client)

        with open(sample_pptx_file, 'rb') as f:
            files = {"file": ("test.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {"file_type": "pptx"}

            response = client.post("/api/translate", files=files, data=data, headers=headers)
            return response.json()["job"]["id"], headers

    def get_auth_headers(self, client, user_data=TEST_USER):
        """Helper method to get authentication headers"""
        # Register and login
        client.post("/api/auth/register", json=user_data)
        login_response = client.post("/api/auth/login", json={
            "email": user_data["email"],
            "password": user_data["password"]
        })
        access_token = login_response.json()["access_token"]
        return {"Authorization": f"Bearer {access_token}"}

    def test_list_jobs(self, client, sample_pptx_file, mock_translation_service):
        """Test listing jobs"""
        # Create multiple jobs
        job_ids = []
        for i in range(3):
            job_id, headers = self.setup_job(client, sample_pptx_file, mock_translation_service)
            job_ids.append(job_id)

        headers = self.get_auth_headers(client)

        # Test listing all jobs
        response = client.get("/api/jobs", headers=headers)
        assert response.status_code == 200

        data = response.json()
        assert "jobs" in data
        assert "pagination" in data
        assert len(data["jobs"]) >= 3

        # Check pagination info
        pagination = data["pagination"]
        assert "page" in pagination
        assert "page_size" in pagination
        assert "total" in pagination
        assert "pages" in pagination

        # Test filtering by status
        response = client.get("/api/jobs?status=pending", headers=headers)
        assert response.status_code == 200
        filtered_jobs = response.json()["jobs"]
        assert all(job["status"] == "pending" for job in filtered_jobs)

        # Test filtering by file type
        response = client.get("/api/jobs?file_type=pptx", headers=headers)
        assert response.status_code == 200
        filtered_jobs = response.json()["jobs"]
        assert all(job["file_type"] == "pptx" for job in filtered_jobs)

        # Test pagination
        response = client.get("/api/jobs?page=1&page_size=2", headers=headers)
        assert response.status_code == 200
        data = response.json()
        assert len(data["jobs"]) <= 2

    def test_get_job_details(self, client, sample_pptx_file, mock_translation_service):
        """Test getting job details"""
        job_id, headers = self.setup_job(client, sample_pptx_file, mock_translation_service)

        # Test getting job details
        response = client.get(f"/api/jobs/{job_id}", headers=headers)
        assert response.status_code == 200

        data = response.json()
        assert "job" in data
        assert "logs" in data

        job = data["job"]
        assert job["id"] == job_id
        assert "created_at" in job
        assert "updated_at" in job
        assert "request" in job
        assert "status" in job

        # Test getting non-existent job
        fake_id = str(uuid.uuid4())
        response = client.get(f"/api/jobs/{fake_id}", headers=headers)
        assert response.status_code == 404

    def test_cancel_job(self, client, sample_pptx_file, mock_translation_service):
        """Test cancelling a job"""
        job_id, headers = self.setup_job(client, sample_pptx_file, mock_translation_service)

        # Test cancelling job
        response = client.post(f"/api/jobs/{job_id}/cancel", headers=headers)
        assert response.status_code == 200
        assert response.json()["message"] == "Job cancelled successfully"

        # Verify job is cancelled
        response = client.get(f"/api/jobs/{job_id}", headers=headers)
        assert response.json()["job"]["status"] == "cancelled"

        # Test cancelling already cancelled job
        response = client.post(f"/api/jobs/{job_id}/cancel", headers=headers)
        assert response.status_code == 400

    def test_job_search(self, client, sample_pptx_file, mock_translation_service):
        """Test job search functionality"""
        # Create a job
        job_id, headers = self.setup_job(client, sample_pptx_file, mock_translation_service)

        # Test search
        search_data = {
            "search": "test.pptx",
            "status": "pending",
            "file_type": "pptx",
            "page": 1,
            "page_size": 10
        }

        response = client.post("/api/jobs/search", json=search_data, headers=headers)
        assert response.status_code == 200

        data = response.json()
        assert "jobs" in data
        assert "pagination" in data

        # Test search with date range
        from datetime import datetime, timedelta
        today = datetime.now().strftime("%Y-%m-%d")
        tomorrow = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")

        search_data["date_from"] = today
        search_data["date_to"] = tomorrow

        response = client.post("/api/jobs/search", json=search_data, headers=headers)
        assert response.status_code == 200

    def test_job_statistics(self, client, sample_pptx_file, mock_translation_service):
        """Test job statistics endpoint"""
        # Create multiple jobs
        for i in range(3):
            self.setup_job(client, sample_pptx_file, mock_translation_service)

        headers = self.get_auth_headers(client)

        # Test getting statistics
        response = client.get("/api/jobs/statistics", headers=headers)
        assert response.status_code == 200

        stats = response.json()
        assert "total_jobs" in stats
        assert "status_counts" in stats
        assert "average_duration_minutes" in stats
        assert "total_cost" in stats
        assert "daily_stats" in stats
        assert "file_type_distribution" in stats
        assert "period_days" in stats

        # Test with custom period
        response = client.get("/api/jobs/statistics?days=7", headers=headers)
        assert response.status_code == 200
        assert response.json()["period_days"] == 7

    def test_queue_status(self, client, sample_pptx_file, mock_translation_service):
        """Test queue status endpoint"""
        # Create some jobs
        for i in range(2):
            self.setup_job(client, sample_pptx_file, mock_translation_service)

        headers = self.get_auth_headers(client)

        response = client.get("/api/jobs/queue", headers=headers)
        assert response.status_code == 200

        data = response.json()
        assert "status_counts" in data
        assert "active_jobs" in data
        assert "total_jobs" in data

    def test_job_logs(self, client, sample_pptx_file, mock_translation_service):
        """Test getting job logs"""
        job_id, headers = self.setup_job(client, sample_pptx_file, mock_translation_service)

        # Test getting job logs
        response = client.get(f"/api/jobs/{job_id}/logs", headers=headers)
        assert response.status_code == 200

        logs = response.json()
        assert isinstance(logs, list)

        # Test with limit
        response = client.get(f"/api/jobs/{job_id}/logs?limit=5", headers=headers)
        assert response.status_code == 200
        assert len(response.json()) <= 5

    def test_bulk_operations(self, client, sample_pptx_file, mock_translation_service):
        """Test bulk job operations"""
        # Create multiple jobs
        job_ids = []
        for i in range(3):
            job_id, _ = self.setup_job(client, sample_pptx_file, mock_translation_service)
            job_ids.append(job_id)

        headers = self.get_auth_headers(client)

        # Test bulk cancel
        bulk_request = {"job_ids": job_ids}
        response = client.post("/api/jobs/bulk/cancel", json=bulk_request, headers=headers)
        assert response.status_code == 200

        results = response.json()
        assert "results" in results
        assert "message" in results
        assert len(results["results"]) == 3

        # Verify all jobs are cancelled
        for job_id in job_ids:
            response = client.get(f"/api/jobs/{job_id}", headers=headers)
            assert response.json()["job"]["status"] == "cancelled"

    def test_job_export(self, client, sample_pptx_file, mock_translation_service):
        """Test job export functionality"""
        # Create a job
        self.setup_job(client, sample_pptx_file, mock_translation_service)

        headers = self.get_auth_headers(client)

        # Test CSV export
        response = client.get("/api/jobs/export?format=csv", headers=headers)
        assert response.status_code == 200

        data = response.json()
        assert "data" in data
        assert "filename" in data
        assert "media_type" in data
        assert data["filename"].endswith(".csv")
        assert data["media_type"] == "text/csv"

        # Test JSON export
        response = client.get("/api/jobs/export?format=json", headers=headers)
        assert response.status_code == 200

        data = response.json()
        assert data["filename"].endswith(".json")
        assert data["media_type"] == "application/json"

    def test_delete_job(self, client, sample_pptx_file, mock_translation_service):
        """Test deleting a job"""
        job_id, headers = self.setup_job(client, sample_pptx_file, mock_translation_service)

        # Cancel the job first (can only delete completed/failed/cancelled jobs)
        client.post(f"/api/jobs/{job_id}/cancel", headers=headers)

        # Test deleting job
        response = client.delete(f"/api/jobs/{job_id}", headers=headers)
        assert response.status_code == 200
        assert response.json()["message"] == "Job deleted successfully"

        # Verify job is deleted
        response = client.get(f"/api/jobs/{job_id}", headers=headers)
        assert response.status_code == 404

    def test_retry_job(self, client, sample_pptx_file):
        """Test retrying a failed job"""
        from app.core.job_manager import job_manager

        job_id, headers = self.setup_job(client, sample_pptx_file)

        # Manually mark job as failed
        import asyncio
        asyncio.run(job_manager.update_job_status(job_id, "failed"))

        # Test retrying job
        response = client.post(f"/api/jobs/{job_id}/retry", headers=headers)
        assert response.status_code == 200

        data = response.json()
        assert "message" in data
        assert "job_id" in data

        # Verify new job is created
        new_job_id = data["job_id"]
        response = client.get(f"/api/jobs/{new_job_id}", headers=headers)
        assert response.status_code == 200
        assert response.json()["job"]["status"] == "pending"

    def test_access_other_user_job(self, client, sample_pptx_file):
        """Test that users cannot access other users' jobs"""
        # Create job with first user
        job_id, headers1 = self.setup_job(client, sample_pptx_file)

        # Create second user
        user2_data = {
            "email": "user2@example.com",
            "password": "User2Pass123!",
            "full_name": "User Two"
        }
        client.post("/api/auth/register", json=user2_data)
        login_response = client.post("/api/auth/login", json={
            "email": user2_data["email"],
            "password": user2_data["password"]
        })
        headers2 = {"Authorization": f"Bearer {login_response.json()['access_token']}"}

        # Try to access first user's job with second user's credentials
        response = client.get(f"/api/jobs/{job_id}", headers=headers2)
        assert response.status_code == 404


class TestErrorScenarios:
    """Test various error scenarios"""

    def get_auth_headers(self, client, user_data=TEST_USER):
        """Helper method to get authentication headers"""
        client.post("/api/auth/register", json=user_data)
        login_response = client.post("/api/auth/login", json={
            "email": user_data["email"],
            "password": user_data["password"]
        })
        access_token = login_response.json()["access_token"]
        return {"Authorization": f"Bearer {access_token}"}

    def test_invalid_job_id_format(self, client):
        """Test with invalid job ID format"""
        headers = self.get_auth_headers(client)

        response = client.get("/api/jobs/invalid-id", headers=headers)
        assert response.status_code == 422  # Validation error

    def test_invalid_pagination_parameters(self, client):
        """Test with invalid pagination parameters"""
        headers = self.get_auth_headers(client)

        # Test negative page
        response = client.get("/api/jobs?page=-1", headers=headers)
        assert response.status_code == 422

        # Test page size too large
        response = client.get("/api/jobs?page_size=1000", headers=headers)
        assert response.status_code == 422

    def test_invalid_search_parameters(self, client):
        """Test with invalid search parameters"""
        headers = self.get_auth_headers(client)

        search_data = {
            "page": -1,  # Invalid
            "page_size": 1000  # Invalid
        }

        response = client.post("/api/jobs/search", json=search_data, headers=headers)
        assert response.status_code == 422

    def test_bulk_operations_with_invalid_job_ids(self, client):
        """Test bulk operations with invalid job IDs"""
        headers = self.get_auth_headers(client)

        bulk_request = {"job_ids": ["invalid-id"]}

        response = client.post("/api/jobs/bulk/cancel", json=bulk_request, headers=headers)
        assert response.status_code == 422


class TestHealthEndpoint:
    """Test health check endpoint"""

    def test_health_check(self, client):
        """Test the health check endpoint"""
        response = client.get("/health")
        assert response.status_code == 200

        data = response.json()
        assert "status" in data
        assert data["status"] == "healthy"
        assert "version" in data
        assert "name" in data
        assert "openai_configured" in data
        assert "redis_configured" in data


# Standalone test runner
if __name__ == "__main__":
    import sys

    # Add the parent directory to Python path
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

    # Run the tests
    pytest.main([__file__, "-v", "--tb=short"])