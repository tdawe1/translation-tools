import pytest
from fastapi.testclient import TestClient
from app.main import app
from app.core.config import settings

client = TestClient(app)

def test_health_check():
    """Test the health check endpoint"""
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json()["status"] == "healthy"
    assert "timestamp" in response.json()

def test_register_user():
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

def test_login_user():
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

def test_protected_endpoint_without_token():
    """Test accessing protected endpoint without token"""
    response = client.get("/api/jobs")
    assert response.status_code == 403

def test_get_translation_models():
    """Test getting available translation models"""
    # First login to get token
    user_data = {
        "email": "models@example.com",
        "password": "modelspassword123",
        "full_name": "Models User"
    }
    client.post("/api/auth/register", json=user_data)

    login_data = {
        "email": "models@example.com",
        "password": "modelspassword123"
    }
    login_response = client.post("/api/auth/login", json=login_data)
    token = login_response.json()["access_token"]

    # Access protected endpoint
    headers = {"Authorization": f"Bearer {token}"}
    response = client.get("/api/translate/models", headers=headers)
    assert response.status_code == 200
    data = response.json()
    assert "models" in data
    assert len(data["models"]) > 0

def test_get_supported_formats():
    """Test getting supported file formats"""
    # First login to get token
    user_data = {
        "email": "formats@example.com",
        "password": "formatspassword123",
        "full_name": "Formats User"
    }
    client.post("/api/auth/register", json=user_data)

    login_data = {
        "email": "formats@example.com",
        "password": "formatspassword123"
    }
    login_response = client.post("/api/auth/login", json=login_data)
    token = login_response.json()["access_token"]

    # Access protected endpoint
    headers = {"Authorization": f"Bearer {token}"}
    response = client.get("/api/translate/formats", headers=headers)
    assert response.status_code == 200
    data = response.json()
    assert "formats" in data
    assert "pptx" in data["formats"]
    assert "pdf" in data["formats"]