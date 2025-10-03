"""Test file to verify pytest fixtures work correctly"""

import os
import pytest
from pathlib import Path


def test_environment_variables_set():
    """Test that environment variables are properly set for testing"""
    assert os.environ.get("PYTEST_RUNNING") == "1"
    assert os.environ.get("DEBUG") == "true"
    assert os.environ.get("SECRET_KEY") == "test-secret-key-for-pytest-testing-only-32-chars-long"
    assert os.environ.get("OPENAI_API_KEY") == "mock-sk-for-testing"


def test_test_upload_dir_fixture(test_upload_dir):
    """Test that test_upload_dir fixture creates a temporary directory"""
    import tempfile
    import shutil

    # Check that the directory exists
    assert os.path.exists(test_upload_dir)
    assert os.path.isdir(test_upload_dir)

    # Check that it's a temporary directory (should be in /tmp)
    assert test_upload_dir.startswith(tempfile.gettempdir())

    # Check that it's writable
    test_file = os.path.join(test_upload_dir, "test.txt")
    with open(test_file, 'w') as f:
        f.write("test")
    assert os.path.exists(test_file)

    # The fixture should clean up automatically after the test


def test_test_output_dir_fixture(test_output_dir):
    """Test that test_output_dir fixture creates a temporary directory"""
    import tempfile

    # Check that the directory exists
    assert os.path.exists(test_output_dir)
    assert os.path.isdir(test_output_dir)

    # Check that it's a temporary directory
    assert test_output_dir.startswith(tempfile.gettempdir())

    # The fixture should clean up automatically after the test


def test_clean_test_environment_autouse():
    """Test that clean_test_environment fixture runs automatically"""
    # This fixture is autouse, so it should run before this test
    # Check that test directories exist
    assert os.path.exists("test_uploads")
    assert os.path.exists("test_outputs")
    assert os.path.isdir("test_uploads")
    assert os.path.isdir("test_outputs")


def test_database_isolation(test_db):
    """Test that each test gets a fresh database session"""
    # The test_db fixture provides an in-memory database
    # We can test basic database operations
    from app.database.database import User
    from sqlalchemy import select
    from datetime import datetime
    import uuid

    # Create a user
    user = User(
        id=str(uuid.uuid4()),
        email="fixture_test@example.com",
        full_name="Fixture Test User",
        hashed_password="hashed_password",
        is_active=True,
        created_at=datetime.utcnow()
    )
    test_db.add(user)
    test_db.commit()

    # Query the user
    result = test_db.execute(select(User).where(User.email == "fixture_test@example.com"))
    found_user = result.scalar_one()
    assert found_user is not None
    assert found_user.full_name == "Fixture Test User"


def test_client_fixture(client):
    """Test that the client fixture provides a working test client"""
    # Test a simple endpoint
    response = client.get("/health")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "healthy"


def test_settings_import():
    """Test that settings can be imported and have test values"""
    from app.core.config import settings

    # Check that settings have been loaded with test values
    assert settings.DEBUG is True
    assert "test" in settings.UPLOAD_DIR
    assert "test" in settings.OUTPUT_DIR
    assert settings.SECRET_KEY == "test-secret-key-for-pytest-testing-only-32-chars-long"
    assert settings.OPENAI_API_KEY == "mock-sk-for-testing"
    # Note: DATABASE_URL may be overridden by test_db fixture to use in-memory SQLite