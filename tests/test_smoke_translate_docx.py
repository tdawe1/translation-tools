import pytest
from pathlib import Path
import subprocess
import tempfile
import zipfile
import sys
from unittest.mock import patch, MagicMock

# Add scripts to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

def test_smoke_translate_docx_success():
    """Test smoke_translate_docx.py runs successfully with valid input."""
    # Use real fixture instead of creating mock
    input_file = Path("tests/fixtures/simple_japanese.docx")

    if not input_file.exists():
        pytest.skip("Fixture simple_japanese.docx not found")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_file = Path(tmpdir) / "output.docx"

        result = subprocess.run([
            'python', 'scripts/smoke_translate_docx.py',
            '--input', str(input_file),
            '--output', str(output_file)
        ], capture_output=True, text=True)

        # Handle dependency on translate_docx.py gracefully
        if "No module named" in result.stderr or "translate_docx.py" in result.stderr:
            pytest.skip("Skipping until translate_docx.py is available (PR #14)")

        if result.returncode == 0:
            assert "Smoke test PASSED" in result.stdout
            assert Path(output_file).exists()
        else:
            # Print actual error for debugging
            print(f"STDOUT: {result.stdout}")
            print(f"STDERR: {result.stderr}")
            pytest.fail(f"Smoke test failed with return code {result.returncode}")

def test_smoke_translate_docx_failure():
    """Test smoke_translate_docx.py fails on invalid input."""
    result = subprocess.run([
        'python', 'scripts/smoke_translate_docx.py',
        '--input', '/nonexistent/input.docx',
        '--output', '/tmp/output.docx'
    ], capture_output=True, text=True)

    assert result.returncode != 0
    assert "Error: Input file not found" in result.stdout

def test_get_structure_xml():
    """Test get_structure_xml function extracts XML correctly."""
    # Import here to handle missing defusedxml gracefully
    try:
        from smoke_translate_docx import get_structure_xml
    except ImportError:
        pytest.skip("smoke_translate_docx module not available")

    # Use a real fixture for more reliable testing
    fixture_path = Path("tests/fixtures/simple_japanese.docx")
    if fixture_path.exists():
        xml_str = get_structure_xml(fixture_path)
        assert isinstance(xml_str, str)
        assert len(xml_str) > 0
        # Should contain XML structure elements
        assert '<' in xml_str and '>' in xml_str
    else:
        # Fallback to mock if fixture not available
        with tempfile.NamedTemporaryFile(suffix='.docx') as tmp:
            with zipfile.ZipFile(tmp.name, 'w') as zf:
                zf.writestr('word/document.xml',
                           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>')
                zf.writestr('[Content_Types].xml',
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
                zf.writestr('_rels/.rels',
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')

            xml_str = get_structure_xml(Path(tmp.name))
            assert '<w:document>' in xml_str
            assert 'Test' not in xml_str  # Text content should be removed

def test_compute_parity():
    """Test compute_parity function returns similarity ratio."""
    # Import here to handle missing defusedxml gracefully
    try:
        from smoke_translate_docx import compute_parity
    except ImportError:
        pytest.skip("smoke_translate_docx module not available")

    # Use real fixtures if available
    fixture1 = Path("tests/fixtures/simple_japanese.docx")
    fixture2 = Path("tests/fixtures/sample_japanese.docx")

    if fixture1.exists() and fixture2.exists():
        # Test with real fixtures
        ratio = compute_parity(fixture1, fixture2)
        assert isinstance(ratio, float)
        assert 0.0 <= ratio <= 1.0

        # Test identical files
        ratio_identical = compute_parity(fixture1, fixture1)
        assert ratio_identical == 1.0
    else:
        # Fallback to mock testing
        with tempfile.NamedTemporaryFile(suffix='.docx') as tmp1, \
             tempfile.NamedTemporaryFile(suffix='.docx') as tmp2:

            # Create identical files
            doc_content = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>'
            for tmp_file in [tmp1, tmp2]:
                with zipfile.ZipFile(tmp_file.name, 'w') as zf:
                    zf.writestr('word/document.xml', doc_content)
                    # Add required DOCX files
                    zf.writestr('[Content_Types].xml',
                               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
                    zf.writestr('_rels/.rels',
                               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')

            ratio = compute_parity(Path(tmp1.name), Path(tmp2.name))
            assert ratio == 1.0  # Identical files

            # Create different file
            with zipfile.ZipFile(tmp2.name, 'w') as zf:
                zf.writestr('word/document.xml',
                           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Different</w:t></w:r></w:p></w:body></w:document>')
                zf.writestr('[Content_Types].xml',
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
                zf.writestr('_rels/.rels',
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')

            ratio = compute_parity(Path(tmp1.name), Path(tmp2.name))
            assert 0.0 <= ratio < 1.0

def test_smoke_with_min_parity():
    """Test smoke test with minimum parity threshold."""
    # Use real fixture if available
    input_file = Path("tests/fixtures/simple_japanese.docx")

    if not input_file.exists():
        pytest.skip("Fixture simple_japanese.docx not found")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_file = Path(tmpdir) / "output.docx"

        result = subprocess.run([
            'python', 'scripts/smoke_translate_docx.py',
            '--input', str(input_file),
            '--output', str(output_file),
            '--min-parity', '0.90'  # Lower threshold
        ], capture_output=True, text=True)

        # Handle dependency on translate_docx.py gracefully
        if "No module named" in result.stderr or "translate_docx.py" in result.stderr:
            pytest.skip("Skipping until translate_docx.py is available (PR #14)")

        if result.returncode == 0:
            assert "Smoke test PASSED" in result.stdout
            assert "XML structure parity:" in result.stdout
        else:
            # Print actual error for debugging
            print(f"STDOUT: {result.stdout}")
            print(f"STDERR: {result.stderr}")
            pytest.fail(f"Smoke test failed with return code {result.returncode}")

def test_smoke_artifact_collection():
    """Test that smoke test collects artifacts correctly."""
    input_file = Path("tests/fixtures/simple_japanese.docx")

    if not input_file.exists():
        pytest.skip("Fixture simple_japanese.docx not found")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_file = Path(tmpdir) / "output.docx"

        result = subprocess.run([
            'python', 'scripts/smoke_translate_docx.py',
            '--input', str(input_file),
            '--output', str(output_file)
        ], capture_output=True, text=True)

        # Handle dependency on translate_docx.py gracefully
        if "No module named" in result.stderr or "translate_docx.py" in result.stderr:
            pytest.skip("Skipping until translate_docx.py is available (PR #14)")

        if result.returncode == 0:
            # Check if artifacts directory was mentioned
            assert "Samples collected in" in result.stdout
            assert "tmp/smoke_out" in result.stdout
        else:
            # Print actual error for debugging
            print(f"STDOUT: {result.stdout}")
            print(f"STDERR: {result.stderr}")
            pytest.fail(f"Smoke test failed with return code {result.returncode}")