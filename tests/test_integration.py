#!/usr/bin/env python3
"""
Integration tests for the DOCX translation pipeline.

Tests the complete flow from API upload to translation completion.
"""

import asyncio
import contextlib
import csv
import json
import os
import shutil
import tempfile
import time
import uuid
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
import httpx
from fastapi.testclient import TestClient

# Set test environment variables before importing
os.environ["DEBUG"] = "true"
os.environ["SECRET_KEY"] = "test-secret-key-not-for-production"
os.environ["OPENAI_API_KEY"] = "test-api-key-not-for-production"

# Add backend and project root to path
import sys
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))
sys.path.insert(0, str(project_root / "backend"))

from backend.app.main import app
from backend.app.core.config import settings


@pytest.fixture
def client():
    """Test client for FastAPI app."""
    return TestClient(app)


@pytest.fixture
def sample_docx_content():
    """Create a minimal DOCX file for testing."""
    # Since python-docx might not be available, create a minimal valid DOCX
    # DOCX files are ZIP archives with XML content
    import zipfile
    import xml.etree.ElementTree as ET

    # Create temporary DOCX
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
        # Create minimal DOCX structure
        with zipfile.ZipFile(tmp.name, 'w') as zf:
            # Create document.xml
            doc_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>これはテストです。</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>翻訳が必要な文章です。</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This is English text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t xml:space="preserve">Japanese: </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>日本語のテキスト</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>'''
            zf.writestr('word/document.xml', doc_xml)

            # Minimal required files
            zf.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')

            zf.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')

            zf.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')

        yield tmp.name

    # Cleanup
    os.unlink(tmp.name)


@pytest.fixture
def mock_openai_client():
    """Mock OpenAI client for testing."""
    mock_client = AsyncMock()

    # Mock chat completion response
    mock_response = MagicMock()
    mock_response.choices = [
        MagicMock(
            message=MagicMock(
                content=json.dumps({
                    "translations": [
                        {"original": "これはテストです。", "translated": "This is a test."},
                        {"original": "翻訳が必要な文章です。", "translated": "This is a sentence that needs translation."},
                        {"original": "日本語のテキスト", "translated": "Japanese text"}
                    ]
                })
            )
        )
    ]

    mock_client.chat.completions.create.return_value = mock_response
    return mock_client


class TestTranslationIntegration:
    """Integration tests for the complete translation pipeline."""

    def test_complete_translation_flow(self, client, sample_docx_content, mock_openai_client):
        """Test the complete flow from upload to download."""
        with patch('openai.AsyncOpenAI', return_value=mock_openai_client):
            # Step 1: Upload file for translation
            with open(sample_docx_content, 'rb') as f:
                upload_response = client.post(
                    "/api/translate",
                    files={"file": ("test.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                    data={
                        "model": "gpt-4o-2024-08-06",
                        "temperature": 0.6,
                        "batch_size": 10
                    }
                )

            # Verify upload response
            assert upload_response.status_code == 202
            job_data = upload_response.json()
            assert "job_id" in job_data
            assert job_data["status"] == "processing"
            assert job_data["input_file"] == "test.docx"

            job_id = job_data["job_id"]

            # Step 2: Poll for completion (simulate background task)
            max_wait = 10  # seconds
            start_time = time.time()

            while time.time() - start_time < max_wait:
                status_response = client.get(f"/api/translate/{job_id}")
                assert status_response.status_code == 200

                status_data = status_response.json()
                if status_data["status"] == "completed":
                    break
                elif status_data["status"] == "failed":
                    pytest.fail(f"Translation failed: {status_data.get('error', 'Unknown error')}")

                time.sleep(0.5)

            # Verify completion
            final_status = client.get(f"/api/translate/{job_id}").json()
            assert final_status["status"] == "completed"
            assert final_status["progress"] == 100
            assert "segments_translated" in final_status

            # Step 3: Download translated file
            download_response = client.get(f"/api/translate/{job_id}/download")
            assert download_response.status_code == 200
            assert download_response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            assert "attachment" in download_response.headers["content-disposition"]

    def test_invalid_file_type(self, client):
        """Test rejection of non-DOCX files."""
        # Create a text file
        with tempfile.NamedTemporaryFile(suffix='.txt', mode='w', delete=False) as tmp:
            tmp.write("This is not a DOCX file")
            tmp.flush()

            with open(tmp.name, 'rb') as f:
                response = client.post(
                    "/api/translate",
                    files={"file": ("test.txt", f, "text/plain")}
                )

            assert response.status_code == 400
            assert "Only DOCX files are supported" in response.json()["detail"]

            os.unlink(tmp.name)

    def test_file_too_large(self, client):
        """Test rejection of files over size limit."""
        # Create a large file (simulate 51MB)
        large_content = b"x" * (51 * 1024 * 1024)

        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
            tmp.write(large_content)
            tmp.flush()

            with open(tmp.name, 'rb') as f:
                response = client.post(
                    "/api/translate",
                    files={"file": ("large.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
                )

            assert response.status_code == 413
            assert "File size exceeds 50MB limit" in response.json()["detail"]

            os.unlink(tmp.name)

    def test_job_not_found(self, client):
        """Test 404 for non-existent job."""
        response = client.get("/api/translate/non-existent-job")
        assert response.status_code == 404
        assert "Job not found" in response.json()["detail"]

    def test_download_before_completion(self, client, sample_docx_content):
        """Test download attempt before translation completes."""
        # Start a translation job
        with open(sample_docx_content, 'rb') as f:
            upload_response = client.post(
                "/api/translate/translate",
                files={"file": ("test.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
            )

        job_id = upload_response.json()["job_id"]

        # Try to download immediately
        download_response = client.get(f"/api/translate/{job_id}/download")
        assert download_response.status_code == 400
        assert "Translation not completed" in download_response.json()["detail"]

    def test_invalid_glossary_format(self, client, sample_docx_content):
        """Test rejection of invalid glossary JSON."""
        with open(sample_docx_content, 'rb') as f:
            response = client.post(
                "/api/translate",
                files={"file": ("test.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                data={
                    "glossary": "invalid json {"
                }
            )

        assert response.status_code == 400
        assert "Invalid glossary JSON" in response.json()["detail"]

    @patch('backend.app.api.translate.run_translation_job')
    def test_translation_error_handling(self, mock_logger, client, sample_docx_content):
        """Test error handling during translation."""
        # Mock the translation to fail
        with patch('backend.translation_orchestrator.orchestrator.translate_document') as mock_translate:
            mock_translate.side_effect = Exception("Translation failed")

            with open(sample_docx_content, 'rb') as f:
                response = client.post(
                    "/api/translate",
                    files={"file": ("test.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
                )

            # Should still return 202 (accepted) but job will fail in background
            assert response.status_code == 202
            job_id = response.json()["job_id"]

            # Wait for background task to fail
            time.sleep(1)

            # Check job status
            status_response = client.get(f"/api/translate/translate/{job_id}")
            status_data = status_response.json()
            assert status_data["status"] == "failed"
            assert "error" in status_data

            # Verify error was logged
            mock_logger.error.assert_called()

    def test_concurrent_jobs(self, client, sample_docx_content, mock_openai_client):
        """Test handling of concurrent translation jobs."""
        with patch('openai.AsyncOpenAI', return_value=mock_openai_client):
            # Start multiple jobs concurrently
            job_ids = []
            for i in range(3):
                with open(sample_docx_content, 'rb') as f:
                    response = client.post(
                        "/api/translate",
                        files={"file": (f"test{i}.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
                    )
                    assert response.status_code == 202
                    job_ids.append(response.json()["job_id"])

            # All jobs should have unique IDs
            assert len(set(job_ids)) == 3

            # Wait for all jobs to complete
            for job_id in job_ids:
                max_wait = 10
                start_time = time.time()

                while time.time() - start_time < max_wait:
                    status_response = client.get(f"/api/translate/{job_id}")
                    status_data = status_response.json()

                    if status_data["status"] in ["completed", "failed"]:
                        break

                    time.sleep(0.5)

                final_status = client.get(f"/api/translate/{job_id}").json()
                assert final_status["status"] == "completed"

    def test_job_directory_cleanup(self, client, sample_docx_content, tmp_path):
        """Test that job directories are created and cleaned up properly."""
        # Override upload directory for testing
        original_upload_dir = settings.UPLOAD_DIR
        settings.UPLOAD_DIR = str(tmp_path)

        try:
            with patch('openai.AsyncOpenAI') as mock_client_class:
                mock_client = AsyncMock()
                mock_client.chat.completions.create.return_value = MagicMock(
                    choices=[MagicMock(message=MagicMock(
                        content=json.dumps({"translations": [
                            {"original": "これはテストです。", "translated": "This is a test."}
                        ]})
                    ))]
                )
                mock_client_class.return_value = mock_client

                with open(sample_docx_content, 'rb') as f:
                    response = client.post(
                        "/api/translate",
                        files={"file": ("test.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
                    )

                job_id = response.json()["job_id"]

                # Wait for completion
                time.sleep(2)

                # Verify job directory exists and contains expected files
                job_dir = tmp_path / job_id
                assert job_dir.exists()
                assert (job_dir / "job_info.json").exists()
                assert any(job_dir.glob("*.docx"))  # Output file

        finally:
            # Restore original upload directory
            settings.UPLOAD_DIR = original_upload_dir


def test_translate_docx_cli_end_to_end(monkeypatch, tmp_path):
    """Run the translate_docx CLI and assert outputs are generated with translated content."""
    from backend.translation_orchestrator import TranslationResult
    from scripts.docx_adapter import DocxAdapter
    from scripts import translate_docx

    fixture_doc = (Path(__file__).parent / "fixtures" / "cli_sample.docx").resolve()
    input_doc = tmp_path / "cli_sample.docx"
    shutil.copy(fixture_doc, input_doc)
    output_doc = tmp_path / "cli_output.docx"

    monkeypatch.setenv("OPENAI_API_KEY", "test-key")
    monkeypatch.setenv("OPENAI_MODEL", "test-model")
    monkeypatch.chdir(tmp_path)

    async def fake_translate_document(*, input_path, output_path, **kwargs):
        adapter = DocxAdapter()
        input_path = str(input_path)
        output_path = str(output_path)
        segments = adapter.extract_segments(input_path)
        translated_count = 0
        total_words = 0

        for seg in segments:
            total_words += seg.word_count
            if seg.has_japanese:
                seg.text = f"Translated {seg.text}"
                translated_count += 1

        adapter.apply_translations(input_path, segments, output_path=output_path)

        return TranslationResult(
            output_path=output_path,
            segments_translated=translated_count,
            total_segments=len(segments),
            words_translated=sum(len(seg.text.split()) for seg in segments),
            total_words=total_words,
            cache_hits=0,
            processing_time=0.5,
            warnings=[],
            artifacts={}
        )

    monkeypatch.setattr(translate_docx.orchestrator, "translate_document", fake_translate_document)

    argv = [
        "translate_docx.py",
        "--in", str(input_doc),
        "--out", str(output_doc),
        "--model", "test-model",
        "--bilingual-csv",
        "--json-audit",
        "--no-backup",
        "--no-cache",
    ]
    monkeypatch.setattr(sys, "argv", argv)

    asyncio.run(translate_docx.main())

    assert output_doc.exists()

    csv_path = tmp_path / f"{output_doc.stem}_bilingual.csv"
    audit_path = tmp_path / f"{output_doc.stem}_audit.json"

    assert csv_path.exists(), "Bilingual CSV should be created"
    assert audit_path.exists(), "Audit JSON should be created"

    with open(csv_path, newline='', encoding='utf-8') as csv_file:
        rows = list(csv.reader(csv_file))
    assert any("Translated" in row[2] for row in rows[1:]), "Translated text should be present in CSV"

    audit_payload = json.loads(audit_path.read_text(encoding='utf-8'))
    assert audit_payload["segments"], "Audit report should contain segment details"
    for segment in audit_payload["segments"]:
        assert "metadata" in segment

    stray_outputs = list(tmp_path.glob("*_translated.docx"))
    assert not stray_outputs, f"Unexpected translated artifacts created: {stray_outputs}"


@pytest.mark.asyncio
async def test_orchestrator_integration():
    """Test the orchestrator's integration with the translation system."""
    from backend.translation_orchestrator import orchestrator

    # Test with real DOCX adapter
    test_docx = Path(__file__).parent / "fixtures" / "test_japanese.docx"

    if not test_docx.exists():
        pytest.skip("Test fixture not available")

    # Mock the orchestrator's batch translation helper. Prefer the DOCX hook when
    # available, falling back to the PPTX helper in older branches.
    with contextlib.ExitStack() as stack:
        try:
            mock_translate = stack.enter_context(
                patch(
                    'backend.translation_orchestrator.orchestrator._call_batch_translation',
                    new_callable=AsyncMock,
                )
            )
        except (AttributeError, ModuleNotFoundError):
            mock_translate = stack.enter_context(
                patch(
                    'scripts.translate_pptx_inplace.translate_batch',
                    new_callable=AsyncMock,
                )
            )

        mock_translate.return_value = ["This is a translated text."]

        result = await orchestrator.translate_document(
            input_path=str(test_docx),
            output_path=str(test_docx.parent / "output.docx"),
            model="gpt-4o-2024-08-06",
            batch_size=10,
            temperature=0.6
        )

        assert result.segments_translated > 0
        assert result.total_segments > 0
        assert result.processing_time > 0
        assert os.path.exists(result.output_path)


def test_smoke_translate_docx(monkeypatch, tmp_path):
    """Test the smoke_translate_docx CLI integration."""
    from scripts import smoke_translate_docx
    import subprocess

    fixture_doc = (Path(__file__).parent / "fixtures" / "cli_sample.docx").resolve()
    if not fixture_doc.exists():
        pytest.skip("CLI sample fixture not available")

    input_doc = tmp_path / "smoke_input.docx"
    shutil.copy(fixture_doc, input_doc)
    output_doc = tmp_path / "smoke_output.docx"

    monkeypatch.setenv("OPENAI_API_KEY", "test-key")
    monkeypatch.chdir(tmp_path)

    # Mock translation similar to CLI test
    async def fake_translate(*args, **kwargs):
        from backend.translation_orchestrator import TranslationResult
        from scripts.docx_adapter import DocxAdapter
        adapter = DocxAdapter()
        input_path = kwargs['input_path']
        output_path = kwargs['output_path']
        segments = adapter.extract_segments(input_path)
        translated_count = sum(1 for seg in segments if seg.has_japanese)
        for seg in segments:
            if seg.has_japanese:
                seg.text = f"Mock translated: {seg.text}"
        adapter.apply_translations(input_path, segments, output_path=output_path)
        return TranslationResult(
            output_path=output_path,
            segments_translated=translated_count,
            total_segments=len(segments),
            words_translated=100,
            total_words=100,
            cache_hits=0,
            processing_time=0.1,
            warnings=[],
            artifacts={}
        )

    with patch('scripts.translate_docx.orchestrator.translate_document', fake_translate):
        # Run smoke CLI
        cmd = ['python', '-m', 'scripts.smoke_translate_docx', '--input', str(input_doc), '--output', str(output_doc)]
        result = subprocess.run(cmd, capture_output=True, text=True)
        assert result.returncode == 0, f"Smoke test failed: {result.stderr}"

        # Verify output exists and parity would be high (since structure preserved)
        assert output_doc.exists()

        # Check samples collected
        smoke_dir = tmp_path / "smoke_out"
        assert smoke_dir.exists()
        assert len(list(smoke_dir.glob("*.docx"))) >= 2  # input and output samples

        # Verify translated content (simple check)
        from scripts.docx_adapter import DocxAdapter
        adapter = DocxAdapter()
        translated_segments = adapter.extract_segments(str(output_doc))
        assert any("Mock translated" in seg.text for seg in translated_segments)

if __name__ == "__main__":
    # Run tests directly
    pytest.main([__file__, "-v"])
