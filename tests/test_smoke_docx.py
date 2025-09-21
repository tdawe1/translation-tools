#!/usr/bin/env python3
"""
Tests for DOCX smoke script (HX-1.3).
"""

import pytest
import subprocess
from pathlib import Path

def test_smoke_docx_runs_successfully(tmp_path: Path):
    """Test smoke script exits 0 and generates artifacts."""
    fixture = "tests/fixtures/simple_japanese.docx"  # Assume exists
    output_dir = tmp_path / "artifacts"
    
    result = subprocess.run([
        "python", "scripts/smoke_docx.py",
        "--fixture", fixture,
        "--output-dir", str(output_dir),
        "--task-id", "test"
    ], capture_output=True, text=True)
    
    assert result.returncode == 0
    assert "Smoke passed" in result.stdout
    task_dir = output_dir / "test"
    assert (task_dir / "simple_japanese_audit.json").exists()
    assert (task_dir / "simple_japanese_bilingual.csv").exists()
    assert (task_dir / "smoke.log").exists()

def test_smoke_fails_on_missing_fixture(tmp_path: Path):
    """Test smoke fails if fixture missing."""
    result = subprocess.run([
        "python", "scripts/smoke_docx.py",
        "--fixture", "/nonexistent.docx",
        "--output-dir", str(tmp_path)
    ], capture_output=True, text=True)
    
    assert result.returncode != 0
    assert "Fixture not found" in result.stderr