import pytest
import tempfile
import os
import time
import json
from pathlib import Path
from unittest.mock import patch, AsyncMock
from fastapi.testclient import TestClient
from app.main import app
from app.core.job_manager import job_manager

client = TestClient(app)


@pytest.fixture
def mock_translation_service():
    """Mock the translation service to avoid actual API calls"""
    with patch('app.services.translation_service.TranslationService') as mock:
        # Mock the translation method to return immediately with "completed" status
        instance = mock.return_value
        instance.translate_document = AsyncMock(return_value={
            "status": "completed",
            "output_file": "/test/path/output.pptx",
            "cost": 0.50,
            "duration_seconds": 30
        })
        yield instance




@pytest.fixture
def sample_pptx_file():
    """Create a sample PPTX file for testing"""
    # Create a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        pptx_path = os.path.join(temp_dir, "test.pptx")

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

            # Create ppt/presentation.xml with some slides
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

            # Create minimal slide
            slide = Element('p:sld', {
                'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
            })
            cSld = SubElement(slide, 'p:cSld')
            spTree = SubElement(cSld, 'p:spTree')
            nvGrpSpPr = SubElement(spTree, 'p:nvGrpSpPr')
            grpSpPr = SubElement(spTree, 'p:grpSpPr')

            # Add a text shape with Japanese content
            sp = SubElement(spTree, 'p:sp')
            nvSpPr = SubElement(sp, 'p:nvSpPr')
            cnvSpPr = SubElement(nvSpPr, 'p:cNvSpPr')
            spPr = SubElement(sp, 'p:spPr')
            txBody = SubElement(sp, 'p:txBody')
            bodyPr = SubElement(txBody, 'a:bodyPr')
            lstStyle = SubElement(txBody, 'a:lstStyle')
            p = SubElement(txBody, 'a:p')
            r = SubElement(p, 'a:r')
            rPr = SubElement(r, 'a:rPr', {'lang': 'ja-JP'})
            t = SubElement(r, 'a:t')
            t.text = "これはテストです。"

            zf.writestr('ppt/slides/slide1.xml', tostring(slide))

        yield pptx_path


@pytest.fixture
def sample_pdf_file():
    """Create a sample PDF file for testing"""
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, "test.pdf")

        # Create a minimal PDF file
        with open(pdf_path, 'wb') as f:
            # Minimal PDF header and structure
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
            f.write(b'/Resources <<\n')
            f.write(b'/Font <<\n')
            f.write(b'/F1 4 0 R\n')
            f.write(b'>>\n')
            f.write(b'>>\n')
            f.write(b'/MediaBox [0 0 612 792]\n')
            f.write(b'/Contents 5 0 R\n')
            f.write(b'>>\n')
            f.write(b'endobj\n')
            f.write(b'4 0 obj\n')
            f.write(b'<<\n')
            f.write(b'/Type /Font\n')
            f.write(b'/Subtype /Type1\n')
            f.write(b'/BaseFont /Helvetica\n')
            f.write(b'>>\n')
            f.write(b'endobj\n')
            f.write(b'5 0 obj\n')
            f.write(b'<<\n')
            f.write(b'/Length 44\n')
            f.write(b'>>\n')
            f.write(b'stream\n')
            f.write(b'BT\n')
            f.write(b'/F1 18 Tf\n')
            f.write(b'72 720 Td\n')
            f.write('(PDF \xe3\x83\x86\xe3\x82\xb9\xe3\x83\x88) Tj\n'.encode('latin1'))
            f.write(b'ET\n')
            f.write(b'endstream\n')
            f.write(b'endobj\n')
            f.write(b'xref\n')
            f.write(b'0 6\n')
            f.write(b'0000000000 65535 f \n')
            f.write(b'0000000009 00000 n \n')
            f.write(b'0000000058 00000 n \n')
            f.write(b'0000000115 00000 n \n')
            f.write(b'0000000246 00000 n \n')
            f.write(b'0000000320 00000 n \n')
            f.write(b'0000000421 00000 n \n')
            f.write(b'trailer\n')
            f.write(b'<<\n')
            f.write(b'/Size 6\n')
            f.write(b'/Root 1 0 R\n')
            f.write(b'>>\n')
            f.write(b'startxref\n')
            f.write(b'578\n')
            f.write(b'%%EOF\n')

        yield pdf_path


class TestUserAuthWorkflow:
    """Test user authentication workflow"""

    def test_complete_user_auth_flow(self, client):
        """Test complete user registration and login flow"""
        # 1. Register a new user
        user_data = {
            "email": "workflow@example.com",
            "password": "workflow123",
            "full_name": "Workflow Test User"
        }

        response = client.post("/api/auth/register", json=user_data)
        assert response.status_code == 200
        data = response.json()
        assert data["email"] == user_data["email"]
        assert data["full_name"] == user_data["full_name"]
        assert "id" in data

        # 2. Login with the user
        login_data = {
            "email": "workflow@example.com",
            "password": "workflow123"
        }

        response = client.post("/api/auth/login", json=login_data)
        assert response.status_code == 200
        data = response.json()
        assert "access_token" in data
        assert data["token_type"] == "bearer"

        # 3. Access protected endpoint with token
        headers = {"Authorization": f"Bearer {data['access_token']}"}
        response = client.get("/api/translate/models", headers=headers)
        assert response.status_code == 200
        assert "models" in response.json()

        # 4. Test invalid login
        invalid_login = {
            "email": "workflow@example.com",
            "password": "wrongpassword"
        }
        response = client.post("/api/auth/login", json=invalid_login)
        assert response.status_code == 401


class TestTranslationJobWorkflow:
    """Test translation job submission and monitoring workflow"""

    def test_pptx_translation_workflow(self, sample_pptx_file, auth_headers, mock_translation_service):
        """Test complete PPTX translation workflow"""
        # 1. Submit translation job
        with open(sample_pptx_file, 'rb') as f:
            files = {"file": ("test.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {
                "file_type": "pptx",
                "model": "gpt-4o-mini",
                "temperature": 0.6,
                "offline": False
            }

            response = client.post("/api/translate", files=files, data=data, headers=auth_headers)
            assert response.status_code == 200

            job_data = response.json()
            assert "job" in job_data
            job_id = job_data["job"]["id"]
            assert job_data["job"]["status"] == "pending"
            assert job_data["job"]["file_type"] == "pptx"

        # 2. Check job status
        response = client.get(f"/api/jobs/{job_id}", headers=auth_headers)
        assert response.status_code == 200
        job_details = response.json()
        assert job_details["job"]["id"] == job_id

        # 3. List all jobs
        response = client.get("/api/jobs", headers=auth_headers)
        assert response.status_code == 200
        jobs_list = response.json()
        assert "jobs" in jobs_list
        assert len(jobs_list["jobs"]) >= 1

        # 4. Check job exists in the list
        job_found = any(job["id"] == job_id for job in jobs_list["jobs"])
        assert job_found, "Submitted job not found in jobs list"

        # 5. Test job filtering
        response = client.get("/api/jobs?status=pending", headers=auth_headers)
        assert response.status_code == 200
        filtered_jobs = response.json()
        assert all(job["status"] == "pending" for job in filtered_jobs["jobs"])

    def test_pdf_translation_workflow(self, sample_pdf_file, auth_headers, mock_translation_service):
        """Test complete PDF translation workflow"""
        # 1. Submit PDF translation job
        with open(sample_pdf_file, 'rb') as f:
            files = {"file": ("test.pdf", f, "application/pdf")}
            data = {
                "file_type": "pdf",
                "model": "gpt-4o-2024-08-06",
                "temperature": 0.7,
                "offline": False,
                "pages": "1-5"
            }

            response = client.post("/api/translate", files=files, data=data, headers=auth_headers)
            assert response.status_code == 200

            job_data = response.json()
            job_id = job_data["job"]["id"]
            assert job_data["job"]["file_type"] == "pdf"
            assert job_data["job"]["request"]["pages"] == "1-5"

        # 2. Test job search functionality
        search_data = {
            "search": "test.pdf",
            "status": "pending",
            "page": 1,
            "page_size": 10
        }

        response = client.post("/api/jobs/search", json=search_data, headers=auth_headers)
        assert response.status_code == 200
        search_results = response.json()
        assert "jobs" in search_results
        assert "pagination" in search_results

    async def test_job_management_workflow(self, sample_pptx_file, auth_headers):
        """Test job management operations (cancel, retry, delete)"""
        # 1. Submit a job
        with open(sample_pptx_file, 'rb') as f:
            files = {"file": ("test.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {"file_type": "pptx"}

            response = client.post("/api/translate", files=files, data=data, headers=auth_headers)
            job_id = response.json()["job"]["id"]

        # 2. Cancel the job
        response = client.post(f"/api/jobs/{job_id}/cancel", headers=auth_headers)
        assert response.status_code == 200
        assert response.json()["message"] == "Job cancelled successfully"

        # 3. Verify job status changed
        response = client.get(f"/api/jobs/{job_id}", headers=auth_headers)
        assert response.json()["job"]["status"] == "cancelled"

        # 4. Submit another job for retry test
        with open(sample_pptx_file, 'rb') as f:
            files = {"file": ("test2.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {"file_type": "pptx"}

            response = client.post("/api/translate", files=files, data=data, headers=auth_headers)
            job_id2 = response.json()["job"]["id"]

        # 5. Manually mark job as failed for retry test
        # This would normally be done by the job processor
        await job_manager.update_job_status(job_id2, "failed")

        # 6. Retry the failed job
        response = client.post(f"/api/jobs/{job_id2}/retry", headers=auth_headers)
        assert response.status_code == 200
        assert "job_id" in response.json()

        # 7. Delete completed/cancelled jobs
        response = client.delete(f"/api/jobs/{job_id}", headers=auth_headers)
        assert response.status_code == 200
        assert response.json()["message"] == "Job deleted successfully"

        # 8. Verify job is deleted
        response = client.get(f"/api/jobs/{job_id}", headers=auth_headers)
        assert response.status_code == 404


class TestJobStatisticsWorkflow:
    """Test job statistics and reporting workflow"""

    def test_job_statistics_flow(self, sample_pptx_file, auth_headers):
        """Test retrieving job statistics"""
        # Submit multiple jobs
        for i in range(3):
            with open(sample_pptx_file, 'rb') as f:
                files = {"file": (f"test{i}.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
                data = {"file_type": "pptx"}

                client.post("/api/translate", files=files, data=data, headers=auth_headers)

        # Get job statistics
        response = client.get("/api/jobs/statistics", headers=auth_headers)
        assert response.status_code == 200

        stats = response.json()
        assert "total_jobs" in stats
        assert "status_counts" in stats
        assert "file_type_distribution" in stats
        assert "period_days" in stats

        # Get queue status
        response = client.get("/api/jobs/queue", headers=auth_headers)
        assert response.status_code == 200

        queue_status = response.json()
        assert "status_counts" in queue_status
        assert "active_jobs" in queue_status
        assert "total_jobs" in queue_status

    def test_job_export_workflow(self, sample_pptx_file, auth_headers):
        """Test job data export functionality"""
        # Submit a job first
        with open(sample_pptx_file, 'rb') as f:
            files = {"file": ("export_test.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {"file_type": "pptx"}

            client.post("/api/translate", files=files, data=data, headers=auth_headers)

        # Export jobs as CSV
        response = client.get("/api/jobs/export?format=csv", headers=auth_headers)
        assert response.status_code == 200

        export_data = response.json()
        assert "data" in export_data
        assert "filename" in export_data
        assert export_data["filename"].endswith(".csv")

        # Export jobs as JSON
        response = client.get("/api/jobs/export?format=json", headers=auth_headers)
        assert response.status_code == 200

        export_data = response.json()
        assert "data" in export_data
        assert export_data["filename"].endswith(".json")


class TestBulkOperationsWorkflow:
    """Test bulk job operations"""

    def test_bulk_cancel_workflow(self, sample_pptx_file, auth_headers):
        """Test cancelling multiple jobs at once"""
        job_ids = []

        # Submit multiple jobs
        for i in range(3):
            with open(sample_pptx_file, 'rb') as f:
                files = {"file": (f"bulk{i}.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
                data = {"file_type": "pptx"}

                response = client.post("/api/translate", files=files, data=data, headers=auth_headers)
                job_ids.append(response.json()["job"]["id"])

        # Cancel all jobs
        bulk_request = {"job_ids": job_ids}
        response = client.post("/api/jobs/bulk/cancel", json=bulk_request, headers=auth_headers)
        assert response.status_code == 200

        results = response.json()
        assert "results" in results
        assert all(success for success in results["results"].values())

    async def test_bulk_retry_workflow(self, sample_pptx_file, auth_headers):
        """Test retrying multiple failed jobs"""
        job_ids = []

        # Submit and fail multiple jobs
        for i in range(2):
            with open(sample_pptx_file, 'rb') as f:
                files = {"file": (f"bulk_fail{i}.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
                data = {"file_type": "pptx"}

                response = client.post("/api/translate", files=files, data=data, headers=auth_headers)
                job_id = response.json()["job"]["id"]
                job_ids.append(job_id)

                # Manually mark as failed
                await job_manager.update_job_status(job_id, "failed")

        # Retry all failed jobs
        bulk_request = {"job_ids": job_ids}
        response = client.post("/api/jobs/bulk/retry", json=bulk_request, headers=auth_headers)
        assert response.status_code == 200

        results = response.json()
        assert "retried_job_ids" in results
        assert len(results["retried_job_ids"]) == 2


class TestErrorHandlingWorkflow:
    """Test error handling in various scenarios"""

    def test_invalid_file_upload(self, auth_headers):
        """Test uploading invalid file types"""
        # Create a text file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
            f.write("This is not a valid document file")
            txt_path = f.name

        try:
            with open(txt_path, 'rb') as f:
                files = {"file": ("invalid.txt", f, "text/plain")}
                data = {"file_type": "pptx"}  # Wrong file type

                response = client.post("/api/translate", files=files, data=data, headers=auth_headers)
                # The service should still accept it but may fail during processing
                assert response.status_code == 200
        finally:
            os.unlink(txt_path)

    def test_unauthorized_access(self):
        """Test accessing endpoints without authentication"""
        # Try to access protected endpoint without token
        response = client.get("/api/jobs")
        assert response.status_code == 403

        # Try to submit translation without auth
        with tempfile.NamedTemporaryFile(suffix='.pptx') as f:
            files = {"file": ("test.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {"file_type": "pptx"}

            response = client.post("/api/translate", files=files, data=data)
            assert response.status_code == 403

    def test_job_not_found(self, auth_headers):
        """Test accessing non-existent job"""
        fake_job_id = "00000000-0000-0000-0000-000000000000"
        response = client.get(f"/api/jobs/{fake_job_id}", headers=auth_headers)
        assert response.status_code == 404

    def test_invalid_job_operations(self, sample_pptx_file, auth_headers):
        """Test invalid job operations"""
        # Submit a job
        with open(sample_pptx_file, 'rb') as f:
            files = {"file": ("test.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {"file_type": "pptx"}

            response = client.post("/api/translate", files=files, data=data, headers=auth_headers)
            job_id = response.json()["job"]["id"]

        # Try to retry a non-failed job
        response = client.post(f"/api/jobs/{job_id}/retry", headers=auth_headers)
        assert response.status_code == 400

        # Try to delete an active job
        response = client.delete(f"/api/jobs/{job_id}", headers=auth_headers)
        assert response.status_code == 400


# Run the tests
if __name__ == "__main__":
    pytest.main([__file__, "-v"])