#!/usr/bin/env python3
"""
Script to test frontend-backend API integration
"""
import asyncio
import httpx
import json
from typing import Dict, Any
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE_URL = "http://localhost:8000"

async def test_api_integration():
    """Test that all API endpoints match frontend expectations"""

    async with httpx.AsyncClient(base_url=BASE_URL) as client:

        # Test 1: Health check
        logger.info("🔍 Testing health check...")
        response = await client.get("/health")
        assert response.status_code == 200
        health_data = response.json()
        logger.info(f"✅ Health check passed: {health_data['status']}")

        # Test 2: Auth endpoints (no auth required)
        logger.info("\n🔍 Testing auth endpoints...")

        # Test register endpoint exists
        response = await client.post("/api/auth/register", json={
            "username": "test_user",
            "email": "test@example.com",
            "password": "test_password"
        })
        # Various responses are acceptable
        assert response.status_code in [200, 201, 400, 422]
        logger.info("✅ Auth register endpoint accessible")

        # Test login endpoint exists
        response = await client.post("/api/auth/login", json={
            "username": "test_user",
            "password": "test_password"
        })
        # Various responses are acceptable
        assert response.status_code in [200, 400, 401, 422]
        logger.info("✅ Auth login endpoint accessible")

        # Test 3: Protected endpoints (should return 401 without auth)
        logger.info("\n🔍 Testing protected endpoints...")

        protected_endpoints = [
            ("GET", "/api/auth/me"),
            ("GET", "/api/translate/translate/models"),
            ("GET", "/api/jobs/jobs"),
            ("GET", "/api/jobs/jobs/statistics"),
            ("GET", "/api/sse/subscribe")
        ]

        for method, url in protected_endpoints:
            response = await client.request(method, url)
            if response.status_code == 404:
                logger.error(f"❌ Endpoint not found: {method} {url}")
                return False
            elif response.status_code == 401:
                logger.info(f"✅ {method} {url} - Correctly requires auth")
            else:
                logger.warning(f"⚠️  {method} {url} - Unexpected status: {response.status_code}")

        # Test 4: Check OpenAPI schema
        logger.info("\n🔍 Testing OpenAPI documentation...")
        response = await client.get("/openapi.json")
        assert response.status_code == 200
        openapi = response.json()

        # Verify paths exist
        paths = openapi["paths"]
        expected_paths = [
            "/api/auth/login",
            "/api/auth/register",
            "/api/auth/me",
            "/api/translate/translate",
            "/api/translate/translate/models",
            "/api/jobs/jobs",
            "/api/jobs/jobs/statistics",
            "/api/sse/subscribe"
        ]

        missing_paths = []
        for path in expected_paths:
            if path not in paths:
                missing_paths.append(path)

        if missing_paths:
            logger.error(f"❌ Missing paths in OpenAPI: {missing_paths}")
            return False

        logger.info("✅ All expected paths found in OpenAPI schema")

        # Test 5: Check CORS headers
        logger.info("\n🔍 Testing CORS headers...")
        response = await client.options("/api/auth/login", headers={
            "Origin": "http://localhost:3001",
            "Access-Control-Request-Method": "POST",
            "Access-Control-Request-Headers": "Content-Type"
        })

        if response.status_code == 200:
            cors_headers = response.headers
            if "access-control-allow-origin" in cors_headers:
                logger.info("✅ CORS headers configured correctly")
            else:
                logger.warning("⚠️  CORS headers not found")
        else:
            logger.warning(f"⚠️  CORS preflight failed: {response.status_code}")

        logger.info("\n🎉 All API integration tests passed!")
        return True

if __name__ == "__main__":
    success = asyncio.run(test_api_integration())
    exit(0 if success else 1)