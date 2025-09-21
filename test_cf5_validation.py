#!/usr/bin/env python3
"""
Simple test to verify CF-5 schema validation implementation.
"""
import json
import tempfile
from pathlib import Path
from jsonschema import validate, ValidationError

def test_schema_validation():
    """Test that the audit schema validation works."""

    # Load schema
    schema_path = Path("schemas/audit_v1.schema.json")
    if not schema_path.exists():
        print("❌ Schema file not found")
        return False

    with open(schema_path, "r") as f:
        schema = json.load(f)

    # Test 1: Valid audit JSON should pass
    valid_audit = {
        "translation_info": {
            "input_file": "/tmp/test.docx",
            "output_file": "/tmp/out.docx",
            "timestamp": "2023-09-21T12:00:00Z",
            "processing_time_seconds": 1.5,
            "model": "gpt-4o"
        },
        "document_metadata": {
            "title": "Test Document",
            "author": "Test Author",
            "paragraph_count": 10,
            "table_count": 2,
            "has_headers": True,
            "has_footers": False,
            "has_footnotes": False
        },
        "translation_stats": {
            "total_segments": 10,
            "segments_translated": 10,
            "cache_hits": 2,
            "cache_misses": 8
        },
        "segments": [
            {
                "id": "seg_1",
                "file_path": "word/document.xml",
                "paragraph_index": 0,
                "run_index": 0,
                "metadata": {
                    "bold": True,
                    "italic": False,
                    "color": "FF0000"
                }
            }
        ]
    }

    try:
        validate(instance=valid_audit, schema=schema)
        print("✅ Valid audit JSON passes validation")
    except ValidationError as e:
        print(f"❌ Valid audit JSON failed validation: {e}")
        return False

    # Test 2: Invalid audit JSON should fail
    invalid_audit = {
        "translation_info": {},  # Missing required fields
        "segments": []
    }

    try:
        validate(instance=invalid_audit, schema=schema)
        print("❌ Invalid audit JSON should have failed validation")
        return False
    except ValidationError:
        print("✅ Invalid audit JSON correctly fails validation")

    return True

def test_ci_workflow():
    """Test that CI workflow can validate audit JSON."""

    # Simulate the CI validation step
    import subprocess
    import sys

    try:
        # Try to run the validation command from CI
        result = subprocess.run([
            sys.executable, "-c", """
import json
from pathlib import Path
from jsonschema import validate

# Create minimal test data
test_data = {
    "translation_info": {
        "input_file": "/tmp/test.docx",
        "output_file": "/tmp/out.docx",
        "timestamp": "2023-09-21T12:00:00Z",
        "processing_time_seconds": 1.0,
        "model": "gpt-4o"
    },
    "document_metadata": {"title": "Test"},
    "translation_stats": {
        "total_segments": 1,
        "segments_translated": 1
    },
    "segments": [{"id": "test_1"}]
}

# Load and validate schema
schema_path = Path("schemas/audit_v1.schema.json")
schema = json.load(open(schema_path))
validate(instance=test_data, schema=schema)
print("CI validation: ✅")
"""
        ], capture_output=True, text=True, cwd=".")

        if result.returncode == 0 and "✅" in result.stdout:
            print("✅ CI workflow validation works")
            return True
        else:
            print(f"❌ CI workflow validation failed: {result.stderr}")
            return False

    except Exception as e:
        print(f"❌ CI workflow validation error: {e}")
        return False

if __name__ == "__main__":
    print("Testing CF-5 Schema Validation Implementation")
    print("=" * 50)

    success = True
    success &= test_schema_validation()
    success &= test_ci_workflow()

    print("=" * 50)
    if success:
        print("✅ All CF-5 validation tests passed")
    else:
        print("❌ Some CF-5 validation tests failed")
        sys.exit(1)