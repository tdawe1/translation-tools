#!/usr/bin/env python3
"""
Basic test to ensure audit JSON generation serializes segment metadata dicts
and writes a valid JSON file.
"""
import json
from pathlib import Path

from backend.document_adapter import DocumentMetadata, Segment, SegmentType
from scripts.translate_docx import generate_audit_report

import pytest
from jsonschema import ValidationError, validate


def test_generate_audit_report_serialization(tmp_path: Path):
    # Prepare minimal segments with dict metadata
    segs = [
        Segment(
            id="word_document.xml_0_0",
            text="テスト",
            segment_type=SegmentType.PARAGRAPH,
            file_path="word/document.xml",
            paragraph_index=0,
            run_index=0,
            metadata={"bold": True, "italic": False, "size": 12.0}
        )
    ]

    # Minimal document metadata
    meta = DocumentMetadata(
        file_path="/tmp/test.docx",
        format="docx",
        segment_count=1,
        custom_properties={"title": "Test"}
    )

    out = tmp_path / "audit.json"
    generate_audit_report(
        input_file="/tmp/test.docx",
        output_file="/tmp/out.docx",
        segments=segs,
        metadata=meta,
        processing_time=0.1,
        cache_stats={"hits": 0, "misses": 1},
        output_path=str(out)
    )

    assert out.exists(), "Audit JSON not written"
    data = json.loads(out.read_text(encoding="utf-8"))
    # Metadata dict is serialized and accessible
    assert data["segments"][0]["metadata"]["bold"] is True

    # Validate against schema
    schema_path = Path(__file__).parent.parent / "schemas" / "audit_v1.schema.json"
    assert schema_path.exists(), "Schema file must exist for validation test"
    schema = json.load(open(schema_path, "r", encoding="utf-8"))
    validate(instance=data, schema=schema)


def test_audit_validation_fails_invalid_structure(tmp_path: Path):
    """Test that invalid audit JSON raises ValidationError."""
    # Create an invalid report missing required fields
    invalid_data = {
        "translation_info": {},  # missing required properties
        "segments": []  # ok, but overall invalid
    }

    out = tmp_path / "invalid_audit.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(invalid_data, f, indent=2)

    data = json.loads(out.read_text(encoding="utf-8"))

    schema_path = Path(__file__).parent.parent / "schemas" / "audit_v1.schema.json"
    schema = json.load(open(schema_path, "r", encoding="utf-8"))

    with pytest.raises(ValidationError):
        validate(instance=data, schema=schema)