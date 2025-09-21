import pytest
import tempfile
import os
from pathlib import Path
from unittest.mock import patch, AsyncMock

def test_complete_translation_workflow(client, test_upload_dir, test_output_dir):
    """Test the complete translation workflow from registration to job completion"""

    # 1. Register a user
    user_data = {
        "email": "smoketest@example.com",
        "password": "smoketest123!",
        "full_name": "Smoke Test User"
    }
    response = client.post("/api/auth/register", json=user_data)
    assert response.status_code == 200
    user = response.json()
    assert user["email"] == user_data["email"]

    # 2. Login to get token
    login_data = {
        "email": "smoketest@example.com",
        "password": "smoketest123!"
    }
    response = client.post("/api/auth/login", json=login_data)
    assert response.status_code == 200
    token = response.json()["access_token"]
    headers = {"Authorization": f"Bearer {token}"}

    # 3. Create a sample PPTX file
    pptx_path = os.path.join(test_upload_dir, "smoke_test.pptx")
    create_minimal_pptx(pptx_path)

    # 4. Mock the translation service to avoid actual API calls
    with patch('app.services.translation_service.TranslationService') as mock:
        instance = mock.return_value
        instance.translate_document = AsyncMock(return_value={
            "status": "completed",
            "output_file": os.path.join(test_output_dir, "output.pptx"),
            "cost": 0.50,
            "duration_seconds": 30
        })

        # 5. Submit translation job
        with open(pptx_path, 'rb') as f:
            files = {"file": ("smoke_test.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {
                "file_type": "pptx",
                "model": "gpt-4o-mini",
                "temperature": 0.6
            }

            response = client.post("/api/translate", files=files, data=data, headers=headers)
            assert response.status_code == 200

            job_data = response.json()
            job_id = job_data["job"]["id"]
            assert job_data["job"]["status"] == "pending"
            assert job_data["job"]["request"]["file_type"] == "pptx"

    # 6. Check job status
    response = client.get(f"/api/jobs/{job_id}", headers=headers)
    assert response.status_code == 200
    job_details = response.json()
    assert job_details["job"]["id"] == job_id

    # 7. List all jobs
    response = client.get("/api/jobs", headers=headers)
    assert response.status_code == 200
    jobs_list = response.json()
    assert len(jobs_list["jobs"]) >= 1

    # 8. Test job cancellation
    response = client.post(f"/api/jobs/{job_id}/cancel", headers=headers)
    assert response.status_code == 200

    # 9. Verify job was cancelled
    response = client.get(f"/api/jobs/{job_id}", headers=headers)
    assert response.json()["job"]["status"] == "cancelled"

    # 10. Test statistics endpoint
    response = client.get("/api/jobs/statistics", headers=headers)
    assert response.status_code == 200
    stats = response.json()
    assert "total_jobs" in stats
    assert "status_counts" in stats


def test_pdf_translation_workflow(client, test_upload_dir, test_output_dir):
    """Test PDF translation workflow"""

    # 1. Register and login
    user_data = {
        "email": "pdfsmoke@example.com",
        "password": "pdfsmoke123",
        "full_name": "PDF Smoke User"
    }
    client.post("/api/auth/register", json=user_data)

    login_data = {
        "email": "pdfsmoke@example.com",
        "password": "pdfsmoke123"
    }
    response = client.post("/api/auth/login", json=login_data)
    token = response.json()["access_token"]
    headers = {"Authorization": f"Bearer {token}"}

    # 2. Create a sample PDF file
    pdf_path = os.path.join(test_upload_dir, "smoke_test.pdf")
    create_minimal_pdf(pdf_path)

    # 3. Submit PDF translation with pages parameter
    with open(pdf_path, 'rb') as f:
        files = {"file": ("smoke_test.pdf", f, "application/pdf")}
        data = {
            "file_type": "pdf",
            "model": "gpt-4o-2024-08-06",
            "temperature": 0.7,
            "pages": "1-5"
        }

        response = client.post("/api/translate", files=files, data=data, headers=headers)
        assert response.status_code == 200

        job_data = response.json()
        job_id = job_data["job"]["id"]
        assert job_data["job"]["file_type"] == "pdf"
        assert job_data["job"]["request"]["pages"] == "1-5"

    # 4. Test job search
    search_data = {
        "search": "smoke_test.pdf",
        "status": "pending",
        "page": 1,
        "page_size": 10
    }

    response = client.post("/api/jobs/search", json=search_data, headers=headers)
    assert response.status_code == 200
    search_results = response.json()
    assert "jobs" in search_results
    assert len(search_results["jobs"]) >= 0


def test_error_handling_scenarios(client):
    """Test various error handling scenarios"""

    # 1. Test unauthorized access
    response = client.get("/api/jobs")
    assert response.status_code == 403

    # 2. Test invalid login
    login_data = {
        "email": "nonexistent@example.com",
        "password": "wrongpassword"
    }
    response = client.post("/api/auth/login", json=login_data)
    assert response.status_code == 401

    # 3. Register a user for further tests
    user_data = {
        "email": "errors@example.com",
        "password": "errors123",
        "full_name": "Error Test User"
    }
    client.post("/api/auth/register", json=user_data)

    login_data = {
        "email": "errors@example.com",
        "password": "errors123"
    }
    response = client.post("/api/auth/login", json=login_data)
    token = response.json()["access_token"]
    headers = {"Authorization": f"Bearer {token}"}

    # 4. Test accessing non-existent job
    response = client.get("/api/jobs/00000000-0000-0000-0000-000000000000", headers=headers)
    assert response.status_code == 404

    # 5. Test invalid file type
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
        f.write("This is not a valid document")
        txt_path = f.name

    try:
        with open(txt_path, 'rb') as f:
            files = {"file": ("invalid.txt", f, "text/plain")}
            data = {"file_type": "pptx"}  # Wrong file type declared

            response = client.post("/api/translate", files=files, data=data, headers=headers)
            # Should accept it but will fail during processing
            assert response.status_code == 200
    finally:
        os.unlink(txt_path)


def test_bulk_operations(client, test_upload_dir):
    """Test bulk job operations"""

    # 1. Register and login
    user_data = {
        "email": "bulk@example.com",
        "password": "bulk123",
        "full_name": "Bulk Test User"
    }
    client.post("/api/auth/register", json=user_data)

    login_data = {
        "email": "bulk@example.com",
        "password": "bulk123"
    }
    response = client.post("/api/auth/login", json=login_data)
    token = response.json()["access_token"]
    headers = {"Authorization": f"Bearer {token}"}

    job_ids = []

    # 2. Submit multiple jobs
    for i in range(3):
        pptx_path = os.path.join(test_upload_dir, f"bulk_test_{i}.pptx")
        create_minimal_pptx(pptx_path)

        with open(pptx_path, 'rb') as f:
            files = {"file": (f"bulk_test_{i}.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
            data = {"file_type": "pptx"}

            response = client.post("/api/translate", files=files, data=data, headers=headers)
            job_ids.append(response.json()["job"]["id"])

    # 3. Bulk cancel
    bulk_request = {"job_ids": job_ids}
    response = client.post("/api/jobs/bulk/cancel", json=bulk_request, headers=headers)
    assert response.status_code == 200

    results = response.json()
    assert "results" in results
    assert len(results["results"]) == 3


def create_minimal_pptx(file_path):
    """Create a minimal valid PPTX file for testing"""
    import zipfile
    from xml.etree.ElementTree import Element, SubElement, tostring

    with zipfile.ZipFile(file_path, 'w') as zf:
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

        # Create minimal presentation
        pres = Element('p:presentation', {
            'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        })

        sldMasterIdLst = SubElement(pres, 'p:sldMasterIdLst')
        SubElement(sldMasterIdLst, 'p:sldMasterId', {'id': '2147483648', 'r:id': 'rId1'})

        sldIdLst = SubElement(pres, 'p:sldIdLst')
        SubElement(sldIdLst, 'p:sldId', {'id': '256', 'r:id': 'rId2'})

        zf.writestr('ppt/presentation.xml', tostring(pres))


def create_minimal_pdf(file_path):
    """Create a minimal PDF file for testing"""
    with open(file_path, 'wb') as f:
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
        f.write(b'xref\n')
        f.write(b'0 3\n')
        f.write(b'0000000000 65535 f \n')
        f.write(b'0000000009 00000 n \n')
        f.write(b'0000000058 00000 n \n')
        f.write(b'trailer\n')
        f.write(b'<<\n')
        f.write(b'/Size 3\n')
        f.write(b'/Root 1 0 R\n')
        f.write(b'>>\n')
        f.write(b'startxref\n')
        f.write(b'109\n')
        f.write(b'%%EOF\n')