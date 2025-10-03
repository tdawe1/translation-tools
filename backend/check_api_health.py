#!/usr/bin/env python3
"""
Quick API health check script.

This script performs a quick check of the API endpoints to verify they are responding.
It's useful for verifying the API is running after deployment or during development.

Usage:
    python check_api_health.py
    python check_api_health.py --url http://localhost:8000
"""

import sys
import argparse
import requests
import json
import time

# Default API URL
DEFAULT_URL = "http://localhost:8000"

def check_endpoint(url, endpoint, method="GET", expected_status=200, auth_token=None):
    """Check a single endpoint"""
    full_url = f"{url}{endpoint}"
    headers = {}
    if auth_token:
        headers["Authorization"] = f"Bearer {auth_token}"

    try:
        if method == "GET":
            response = requests.get(full_url, headers=headers, timeout=5)
        elif method == "POST":
            response = requests.post(full_url, headers=headers, timeout=5)
        else:
            return False, f"Unsupported method: {method}"

        if response.status_code == expected_status:
            return True, f"✓ {endpoint} - {response.status_code}"
        else:
            return False, f"✗ {endpoint} - {response.status_code} (expected {expected_status})"

    except requests.exceptions.RequestException as e:
        return False, f"✗ {endpoint} - Request failed: {e}"

def main():
    parser = argparse.ArgumentParser(description="Check API health")
    parser.add_argument("--url", default=DEFAULT_URL, help=f"API base URL (default: {DEFAULT_URL})")
    args = parser.parse_args()

    base_url = args.url.rstrip('/')
    print(f"Checking API health at: {base_url}")
    print("=" * 50)

    # Check endpoints
    endpoints = [
        ("/health", "GET", 200),
        ("/", "GET", 200),
        ("/api/translate/models", "GET", 403),  # Should require auth
        ("/api/jobs", "GET", 403),  # Should require auth
        ("/api/auth/register", "POST", 422),  # Should fail without data
    ]

    all_passed = True
    auth_token = None

    # First, try to register a test user and get a token
    test_user = {
        "email": f"healthcheck{int(time.time())}@example.com",
        "password": "HealthCheck123!",
        "full_name": "Health Check User"
    }

    try:
        # Register user
        response = requests.post(f"{base_url}/api/auth/register", json=test_user, timeout=5)
        if response.status_code == 200:
            # Login to get token
            login_data = {
                "email": test_user["email"],
                "password": test_user["password"]
            }
            response = requests.post(f"{base_url}/api/auth/login", json=login_data, timeout=5)
            if response.status_code == 200:
                auth_token = response.json().get("access_token")
                print("✓ Obtained authentication token")
    except:
        pass

    # Check each endpoint
    for endpoint, method, expected_status in endpoints:
        success, message = check_endpoint(base_url, endpoint, method, expected_status, auth_token)
        print(message)
        if not success:
            all_passed = False

    # If we have auth, check a few more endpoints
    if auth_token:
        auth_endpoints = [
            ("/api/translate/models", "GET", 200),
            ("/api/translate/formats", "GET", 200),
            ("/api/jobs", "GET", 200),
        ]

        print("\nChecking authenticated endpoints:")
        for endpoint, method, expected_status in auth_endpoints:
            success, message = check_endpoint(base_url, endpoint, method, expected_status, auth_token)
            print(message)
            if not success:
                all_passed = False

    print("\n" + "=" * 50)
    if all_passed:
        print("✓ API health check passed")
        return 0
    else:
        print("✗ API health check failed")
        return 1

if __name__ == "__main__":
    sys.exit(main())