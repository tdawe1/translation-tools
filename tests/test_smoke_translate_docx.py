import pytest
from pathlib import Path
import subprocess
import tempfile
from unittest.mock import patch, MagicMock

def test_smoke_translate_docx_success():
    """Test smoke_translate_docx.py runs successfully with valid input."""
    with tempfile.TemporaryDirectory() as tmpdir:
        input_file = Path(tmpdir) / "input.docx"
        output_file = Path(tmpdir) / "output.docx"
        
        # Create a minimal input file (mock)
        with open(input_file, 'w') as f:
            f.write("Mock input")
        
        result = subprocess.run([
            'python', 'scripts/smoke_translate_docx.py',
            '--input', str(input_file),
            '--output', str(output_file)
        ], capture_output=True, text=True)
        
        assert result.returncode == 0
        assert "Smoke test PASSED" in result.stdout
        assert Path(output_file).exists()

def test_smoke_translate_docx_failure():
    """Test smoke_translate_docx.py fails on invalid input."""
    result = subprocess.run([
        'python', 'scripts/smoke_translate_docx.py',
        '--input', '/nonexistent/input.docx',
        '--output', '/tmp/output.docx'
    ], capture_output=True, text=True)
    
    assert result.returncode != 0
    assert "Error: Input file not found" in result.stderr

def test_get_structure_xml():
    """Test get_structure_xml function extracts XML correctly."""
    from scripts.smoke_translate_docx import get_structure_xml
    
    with tempfile.NamedTemporaryFile(suffix='.docx') as tmp:
        # Mock a simple DOCX
        with zipfile.ZipFile(tmp.name, 'w') as zf:
            zf.writestr('word/document.xml', '<w:document><w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>')
            zf.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types><Default Extension="xml" ContentType="application/xml"/></Types>')
            zf.writestr('_rels/.rels', '<?xml version="1.0"?><Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        
        xml_str = get_structure_xml(Path(tmp.name))
        assert '<w:document>' in xml_str
        assert 'Test' not in xml_str  # Text removed

def test_compute_parity():
    """Test compute_parity function returns similarity ratio."""
    from scripts.smoke_translate_docx import compute_parity, get_structure_xml
    
    with tempfile.NamedTemporaryFile(suffix='.docx') as tmp1, tempfile.NamedTemporaryFile(suffix='.docx') as tmp2:
        # Identical files
        with zipfile.ZipFile(tmp1.name, 'w') as zf:
            zf.writestr('word/document.xml', '<w:document><w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>')
        with zipfile.ZipFile(tmp2.name, 'w') as zf:
            zf.writestr('word/document.xml', '<w:document><w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>')
        
        ratio = compute_parity(Path(tmp1.name), Path(tmp2.name))
        assert ratio == 1.0  # Identical

        # Different files
        with zipfile.ZipFile(tmp2.name, 'w') as zf:
            zf.writestr('word/document.xml', '<w:document><w:body><w:p><w:r><w:t>Different</w:t></w:r></w:p></w:body></w:document>')
        
        ratio = compute_parity(Path(tmp1.name), Path(tmp2.name))
        assert 0.0 < ratio < 1.0