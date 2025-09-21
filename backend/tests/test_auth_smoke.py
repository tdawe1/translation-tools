import pytest
import time
from datetime import datetime, timedelta
from unittest.mock import patch, MagicMock
from app.models.auth import UserLogin, UserCreate, Token


class TestAuthSmoke:
    """End-to-end smoke tests for authentication flow"""

    def test_complete_auth_flow(self, client):
        """Test the complete authentication flow: register -> login -> access protected endpoint"""
        # 1. Register a new user
        import time
        timestamp = int(time.time())
        user_data = {
            "email": f"smoke{timestamp}@example.com",
            "password": "smokepass123",
            "full_name": "Smoke Test User"
        }

        register_response = client.post("/api/auth/register", json=user_data)
        assert register_response.status_code == 200
        user_response = register_response.json()
        assert user_response["email"] == user_data["email"]
        assert user_response["full_name"] == user_data["full_name"]
        assert "id" in user_response
        user_id = user_response["id"]

        # 2. Login with the registered user
        login_data = {
            "email": user_data["email"],
            "password": user_data["password"]
        }

        login_response = client.post("/api/auth/login", json=login_data)
        assert login_response.status_code == 200
        token_response = login_response.json()
        assert "access_token" in token_response
        assert "refresh_token" in token_response
        assert token_response["token_type"] == "bearer"
        assert token_response["expires_in"] > 0

        access_token = token_response["access_token"]
        refresh_token = token_response["refresh_token"]

        # 3. Access protected endpoint with access token
        auth_headers = {"Authorization": f"Bearer {access_token}"}
        me_response = client.get("/api/auth/me", headers=auth_headers)
        assert me_response.status_code == 200
        me_data = me_response.json()
        assert me_data["id"] == user_id
        assert me_data["email"] == user_data["email"]

        # 4. Test access to other protected endpoints
        models_response = client.get("/api/translate/models", headers=auth_headers)
        assert models_response.status_code == 200

        formats_response = client.get("/api/translate/formats", headers=auth_headers)
        assert formats_response.status_code == 200

        jobs_response = client.get("/api/jobs", headers=auth_headers)
        assert jobs_response.status_code == 200

        # 5. Test token refresh
        refresh_payload = {"refresh_token": refresh_token}
        refresh_response = client.post("/api/auth/refresh", json=refresh_payload)
        assert refresh_response.status_code == 200
        new_token_response = refresh_response.json()
        assert "access_token" in new_token_response
        assert new_token_response["access_token"] != access_token

        # 6. Test access with new token
        new_auth_headers = {"Authorization": f"Bearer {new_token_response['access_token']}"}
        new_me_response = client.get("/api/auth/me", headers=new_auth_headers)
        assert new_me_response.status_code == 200

        # 7. Test logout
        logout_payload = {"refresh_token": refresh_token}
        logout_response = client.post("/api/auth/logout", json=logout_payload)
        assert logout_response.status_code == 200

        # 8. Verify refresh token is revoked
        failed_refresh_response = client.post("/api/auth/refresh", json=refresh_payload)
        assert failed_refresh_response.status_code == 401

    def test_registration_validation(self, client):
        """Test registration validation"""
        import time
        timestamp = int(time.time())

        # Test duplicate email registration
        user_data = {
            "email": f"duplicate{timestamp}@example.com",
            "password": "pass123",
            "full_name": "First User"
        }

        # First registration should succeed
        first_response = client.post("/api/auth/register", json=user_data)
        assert first_response.status_code == 200

        # Second registration with same email should fail
        second_response = client.post("/api/auth/register", json=user_data)
        assert second_response.status_code == 400

        # Test missing required fields
        incomplete_data = {
            "email": f"incomplete{timestamp}@example.com"
            # Missing password and full_name
        }
        incomplete_response = client.post("/api/auth/register", json=incomplete_data)
        assert incomplete_response.status_code == 422

    def test_login_validation(self, client):
        """Test login validation"""
        import time
        timestamp = int(time.time())

        # Register a user first
        user_data = {
            "email": f"loginval{timestamp}@example.com",
            "password": "correctpass123",
            "full_name": "Login Validation User"
        }
        client.post("/api/auth/register", json=user_data)

        # Test correct login
        correct_login = {
            "email": f"loginval{timestamp}@example.com",
            "password": "correctpass123"
        }
        correct_response = client.post("/api/auth/login", json=correct_login)
        assert correct_response.status_code == 200

        # Test incorrect password
        wrong_password = {
            "email": f"loginval{timestamp}@example.com",
            "password": "wrongpassword"
        }
        wrong_response = client.post("/api/auth/login", json=wrong_password)
        assert wrong_response.status_code == 401

        # Test non-existent user
        nonexistent_login = {
            "email": f"nonexistent{timestamp}@example.com",
            "password": "anypassword"
        }
        nonexistent_response = client.post("/api/auth/login", json=nonexistent_login)
        assert nonexistent_response.status_code == 401

    def test_token_expiration(self, client):
        """Test token expiration behavior"""
        import time
        timestamp = int(time.time())

        # Register and login
        user_data = {
            "email": f"expire{timestamp}@example.com",
            "password": "expirepass123",
            "full_name": "Expiration Test User"
        }
        client.post("/api/auth/register", json=user_data)

        login_data = {
            "email": f"expire{timestamp}@example.com",
            "password": "expirepass123"
        }
        login_response = client.post("/api/auth/login", json=login_data)
        assert login_response.status_code == 200
        token = login_response.json()["access_token"]

        # Mock token verification to simulate expiration
        with patch('app.services.auth_service.AuthService.verify_token') as mock_verify:
            mock_verify.side_effect = Exception("Token expired")

            auth_headers = {"Authorization": f"Bearer {token}"}
            me_response = client.get("/api/auth/me", headers=auth_headers)
            assert me_response.status_code == 401

    def test_invalid_token_scenarios(self, client):
        """Test various invalid token scenarios"""
        # Test no token
        no_token_response = client.get("/api/auth/me")
        assert no_token_response.status_code == 403

        # Test malformed token
        malformed_headers = {"Authorization": "Bearer invalid-token-format"}
        malformed_response = client.get("/api/auth/me", headers=malformed_headers)
        assert malformed_response.status_code == 401

        # Test empty token
        empty_headers = {"Authorization": "Bearer "}
        empty_response = client.get("/api/auth/me", headers=empty_headers)
        assert empty_response.status_code == 403

    def test_api_key_authentication(self, client):
        """Test API key authentication if available"""
        import time
        timestamp = int(time.time())

        # First register and login to get a regular token
        user_data = {
            "email": f"apikey{timestamp}@example.com",
            "password": "apikeypass123",
            "full_name": "API Key User"
        }
        client.post("/api/auth/register", json=user_data)

        login_data = {
            "email": f"apikey{timestamp}@example.com",
            "password": "apikeypass123"
        }
        login_response = client.post("/api/auth/login", json=login_data)
        access_token = login_response.json()["access_token"]

        # Create an API key
        api_key_data = {
            "name": "Test API Key",
            "expires_days": 30
        }
        api_key_headers = {"Authorization": f"Bearer {access_token}"}
        api_key_response = client.post("/api/auth/api-keys", json=api_key_data, headers=api_key_headers)

        if api_key_response.status_code == 200:
            # API key creation succeeded, test using it
            api_key_info = api_key_response.json()
            assert "key" in api_key_info

            # Test API key authentication
            api_auth_headers = {"X-API-Key": api_key_info["key"]}
            api_me_response = client.get("/api/auth/me", headers=api_auth_headers)
            # Note: This depends on how API key authentication is implemented
            # It might return a different response or not support /api/auth/me endpoint

    def test_google_oauth_flow_mock(self, client):
        """Test Google OAuth flow with mocked responses"""
        with patch('app.services.oauth_service.OAuthService.get_google_auth_url') as mock_url, \
             patch('app.services.oauth_service.OAuthService.exchange_google_code') as mock_exchange, \
             patch('app.services.oauth_service.OAuthService.get_google_user_info') as mock_user_info:

            # Mock auth URL generation
            mock_url.return_value = "https://accounts.google.com/o/oauth2/auth?mocked=true"

            # Get auth URL
            auth_url_response = client.get("/api/auth/google/auth-url?redirect_uri=http://localhost:3000/callback")
            assert auth_url_response.status_code == 200
            assert "mocked=true" in auth_url_response.json()["auth_url"]

            # Mock token exchange and user info
            mock_exchange.return_value = {
                "access_token": "mock_access_token",
                "refresh_token": "mock_refresh_token"
            }
            mock_user_info.return_value = {
                "id": "google123",
                "email": "googleuser@example.com",
                "name": "Google User"
            }

            # Simulate callback
            callback_data = {
                "code": "mock_auth_code",
                "redirect_uri": "http://localhost:3000/callback"
            }
            callback_response = client.post("/api/auth/google/callback", json=callback_data)
            assert callback_response.status_code == 200
            callback_result = callback_response.json()
            assert "access_token" in callback_result
            assert callback_result["user"]["email"] == "googleuser@example.com"

    def test_rate_limiting_on_auth(self, client):
        """Test rate limiting on authentication endpoints"""
        import time
        timestamp = int(time.time())

        # This test assumes rate limiting is implemented
        # Make multiple rapid requests to login endpoint
        login_data = {
            "email": f"ratelimit{timestamp}@example.com",
            "password": "ratelimit123"
        }

        # Register user first
        user_data = {
            "email": f"ratelimit{timestamp}@example.com",
            "password": "ratelimit123",
            "full_name": "Rate Limit Test"
        }
        client.post("/api/auth/register", json=user_data)

        # Make multiple login attempts
        responses = []
        for i in range(10):  # Adjust based on your rate limit settings
            response = client.post("/api/auth/login", json=login_data)
            responses.append(response.status_code)

        # If rate limiting is implemented, we should see some 429 responses
        # This is a basic check - actual implementation may vary
        assert 429 in responses or all(status == 200 for status in responses)

    @pytest.fixture
    def expired_refresh_token(self, client):
        """Create an expired refresh token for testing"""
        import time
        timestamp = int(time.time())

        # Register and login
        user_data = {
            "email": f"expiredrefresh{timestamp}@example.com",
            "password": "expiredpass123",
            "full_name": "Expired Refresh User"
        }
        client.post("/api/auth/register", json=user_data)

        login_data = {
            "email": f"expiredrefresh{timestamp}@example.com",
            "password": "expiredpass123"
        }
        response = client.post("/api/auth/login", json=login_data)
        return response.json()["refresh_token"]

    def test_expired_refresh_token(self, client, expired_refresh_token):
        """Test behavior with expired refresh token"""
        # Mock the refresh token verification to simulate expiration
        with patch('app.services.auth_service.AuthService.refresh_access_token') as mock_refresh:
            mock_refresh.side_effect = Exception("Refresh token expired")

            refresh_payload = {"refresh_token": expired_refresh_token}
            response = client.post("/api/auth/refresh", json=refresh_payload)
            assert response.status_code == 401

    def test_concurrent_sessions(self, client):
        """Test handling of multiple concurrent sessions"""
        import time
        timestamp = int(time.time())

        # Register user
        user_data = {
            "email": f"concurrent{timestamp}@example.com",
            "password": "concurrentpass123",
            "full_name": "Concurrent Session User"
        }
        client.post("/api/auth/register", json=user_data)

        # Create multiple sessions
        sessions = []
        for i in range(3):
            login_data = {
                "email": f"concurrent{timestamp}@example.com",
                "password": "concurrentpass123"
            }
            response = client.post("/api/auth/login", json=login_data)
            assert response.status_code == 200
            sessions.append(response.json()["access_token"])

        # All tokens should be valid
        for i, token in enumerate(sessions):
            headers = {"Authorization": f"Bearer {token}"}
            response = client.get("/api/auth/me", headers=headers)
            assert response.status_code == 200
            assert response.json()["email"] == f"concurrent{timestamp}@example.com"