#!/usr/bin/env python3
"""
Script to verify all API endpoints are accessible
"""
import asyncio
import httpx
import json
from typing import Dict, List, Any
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE_URL = "http://localhost:8000"

# Test endpoints
ENDPOINTS = {
    # Health check (no auth required)
    "GET /health": {
        "method": "GET",
        "url": "/health",
        "auth_required": False
    },

    # Auth endpoints (no auth required for login/register)
    "POST /api/auth/login": {
        "method": "POST",
        "url": "/api/auth/login",
        "auth_required": False,
        "json": {"username": "test", "password": "test"}
    },
    "POST /api/auth/register": {
        "method": "POST",
        "url": "/api/auth/register",
        "auth_required": False,
        "json": {"username": "test", "email": "test@example.com", "password": "test"}
    },
    "GET /api/auth/me": {
        "method": "GET",
        "url": "/api/auth/me",
        "auth_required": True
    },

    # Translate endpoints (auth required)
    "POST /api/translate/translate": {
        "method": "POST",
        "url": "/api/translate/translate",
        "auth_required": True,
        # Note: Would normally include file upload
    },
    "GET /api/translate/translate/models": {
        "method": "GET",
        "url": "/api/translate/translate/models",
        "auth_required": True
    },

    # Jobs endpoints (auth required)
    "GET /api/jobs/jobs": {
        "method": "GET",
        "url": "/api/jobs/jobs",
        "auth_required": True
    },
    "GET /api/jobs/jobs/statistics": {
        "method": "GET",
        "url": "/api/jobs/jobs/statistics",
        "auth_required": True
    },

    # SSE endpoint (auth required)
    "GET /api/sse/subscribe": {
        "method": "GET",
        "url": "/api/sse/subscribe?job_id=test",
        "auth_required": True
    }
}

async def test_endpoint(client: httpx.AsyncClient, name: str, config: Dict[str, Any]) -> Dict[str, Any]:
    """Test a single endpoint"""
    result = {
        "endpoint": name,
        "method": config["method"],
        "url": config["url"],
        "status": "unknown",
        "status_code": None,
        "error": None
    }

    try:
        # Prepare headers
        headers = {}
        if config["auth_required"]:
            # Use a dummy token for testing
            headers["Authorization"] = "Bearer dummy_token"

        # Make request
        response = await client.request(
            method=config["method"],
            url=config["url"],
            headers=headers,
            json=config.get("json"),
            timeout=5.0
        )

        result["status_code"] = response.status_code

        # Check if endpoint exists (404 vs other errors)
        if response.status_code == 404:
            result["status"] = "not_found"
            result["error"] = "Endpoint not found"
        elif config["auth_required"] and response.status_code == 401:
            result["status"] = "accessible"  # 401 is expected for dummy token
        elif not config["auth_required"] and response.status_code == 422:
            result["status"] = "accessible"  # 422 for invalid request body
        elif 200 <= response.status_code < 300:
            result["status"] = "accessible"
        else:
            result["status"] = "error"
            result["error"] = f"HTTP {response.status_code}"

    except httpx.ConnectError:
        result["status"] = "connection_error"
        result["error"] = "Failed to connect to server"
    except Exception as e:
        result["status"] = "error"
        result["error"] = str(e)

    return result

async def main():
    """Main verification function"""
    logger.info("🔍 Verifying API endpoints...")

    async with httpx.AsyncClient(base_url=BASE_URL) as client:
        # Test all endpoints
        tasks = []
        for name, config in ENDPOINTS.items():
            task = test_endpoint(client, name, config)
            tasks.append(task)

        results = await asyncio.gather(*tasks)

    # Report results
    logger.info("\n📊 Endpoint Verification Results:")
    logger.info("=" * 60)

    accessible = 0
    not_found = 0
    errors = 0

    for result in results:
        status_icon = {
            "accessible": "✅",
            "not_found": "❌",
            "error": "⚠️",
            "connection_error": "🔥"
        }.get(result["status"], "❓")

        logger.info(f"{status_icon} {result['method']} {result['url']}")
        if result["status_code"]:
            logger.info(f"   Status: {result['status_code']}")
        if result["error"]:
            logger.info(f"   Error: {result['error']}")

        if result["status"] == "accessible":
            accessible += 1
        elif result["status"] == "not_found":
            not_found += 1
        else:
            errors += 1

    logger.info("\n" + "=" * 60)
    logger.info(f"✅ Accessible: {accessible}")
    logger.info(f"❌ Not Found: {not_found}")
    logger.info(f"⚠️ Errors: {errors}")

    if not_found > 0:
        logger.error("\n🚨 Some endpoints are not accessible!")
        return False

    logger.info("\n🎉 All endpoints are properly mounted!")
    return True

if __name__ == "__main__":
    success = asyncio.run(main())
    exit(0 if success else 1)