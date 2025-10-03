import pytest
from app.core.config import settings

def test_health_check(client):
    """Test the health check endpoint"""
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json()["status"] == "healthy"

def test_register_user(client):
    """Test user registration"""
    user_data = {
        "email": "test@example.com",
        "password": "testpassword123",
        "full_name": "Test User"
    }
    response = client.post("/api/auth/register", json=user_data)
    assert response.status_code == 200
    data = response.json()
    assert data["email"] == user_data["email"]
    assert data["full_name"] == user_data["full_name"]
    assert "id" in data

def test_login_user(client):
    """Test user login"""
    # First register a user
    user_data = {
        "email": "login@example.com",
        "password": "loginpassword123",
        "full_name": "Login User"
    }
    client.post("/api/auth/register", json=user_data)

    # Then login
    login_data = {
        "email": "login@example.com",
        "password": "loginpassword123"
    }
    response = client.post("/api/auth/login", json=login_data)
    assert response.status_code == 200
    data = response.json()
    assert "access_token" in data
    assert data["token_type"] == "bearer"

def test_protected_endpoint_without_token(client):
    """Test accessing protected endpoint without token"""
    response = client.get("/api/jobs")
    assert response.status_code == 403

def test_get_translation_models(client, auth_headers):
    """Test getting available translation models"""
    # Access protected endpoint
    response = client.get("/api/translate/models", headers=auth_headers)
    assert response.status_code == 200
    data = response.json()
    assert "models" in data
    assert len(data["models"]) > 0

def test_get_supported_formats(client, auth_headers):
    """Test getting supported file formats"""
    # Access protected endpoint
    response = client.get("/api/translate/formats", headers=auth_headers)
    assert response.status_code == 200
    data = response.json()
    assert "formats" in data
    assert "pptx" in data["formats"]
    assert "pdf" in data["formats"]

def test_environment_configuration():
    """Test that test environment is properly configured"""
    # Check that we're using test configuration
    assert settings.DEBUG is True
    assert "test" in settings.UPLOAD_DIR or "test" in settings.OUTPUT_DIR
    assert settings.OPENAI_API_KEY == "mock-sk-for-testing"
    assert settings.SECRET_KEY == "test-secret-key-for-pytest-testing-only-32-chars-long"
    assert settings.DATABASE_URL.startswith("sqlite")
    assert settings.LOG_LEVEL == "WARNING"

def test_directory_setup(test_upload_dir, test_output_dir):
    """Test that test directories are properly created and accessible"""
    import os
    # Check that directories exist
    assert os.path.exists(test_upload_dir)
    assert os.path.exists(test_output_dir)
    # Check that they are directories
    assert os.path.isdir(test_upload_dir)
    assert os.path.isdir(test_output_dir)
    # Check that they are writable
    test_file = os.path.join(test_upload_dir, "test_write.txt")
    with open(test_file, 'w') as f:
        f.write("test")
    assert os.path.exists(test_file)
    os.remove(test_file)


def test_test_database_isolation(test_db):
    """Test that each test gets a clean database"""
    # Test that we can create and query data
    from app.database.database import User
    from sqlalchemy import select
    from datetime import datetime
    import uuid

    # Create a test user
    user = User(
        id=str(uuid.uuid4()),
        email="dbtest@example.com",
        full_name="DB Test User",
        hashed_password="hashed_password_here",
        is_active=True,
        created_at=datetime.utcnow()
    )
    test_db.add(user)
    test_db.commit()

    # Query the user
    result = test_db.execute(select(User).where(User.email == "dbtest@example.com"))
    assert result.scalar_one() is not None


def test_settings_configuration():
    """Test that all test-specific settings are properly configured"""
    # Check feature flags
    assert settings.ENABLE_STYLE_CHECKING is True
    assert settings.ENABLE_EXPANSION_POLICY is True
    assert settings.ENABLE_FORMATTING_PROFILE is True

    # Check job settings (from .env.test)
    assert settings.JOB_TIMEOUT == 60
    assert settings.MAX_CONCURRENT_JOBS == 2

    # Check auth settings (from .env.test)
    assert settings.ACCESS_TOKEN_EXPIRE_MINUTES == 5  # Test override
    assert settings.REFRESH_TOKEN_EXPIRE_DAYS == 1  # Test override

    # Check rate limiting (from .env.test)
    assert settings.RATE_LIMIT_REQUESTS == 1000
    assert settings.RATE_LIMIT_WINDOW == 60


@pytest.mark.smoke
def test_smoke_tests_available():
    """Test that smoke tests are discoverable by pytest"""
    # This test ensures the smoke test marker is properly configured
    assert True