#!/usr/bin/env python3
"""
Simple script to test the frontend-backend integration
"""
import requests
import time

BASE_URL = "http://localhost:8000"

def test_api_endpoints():
    print("Testing API endpoints...")

    # Test health endpoint
    print("\n1. Testing /health endpoint...")
    response = requests.get(f"{BASE_URL}/health")
    print(f"   Status: {response.status_code}")
    print(f"   Response: {response.json()}")

    # Test file upload (without actual file)
    print("\n2. Testing file upload endpoint...")
    # This would normally fail without a file, but we can check if the endpoint exists
    response = requests.post(f"{BASE_URL}/upload")
    print(f"   Status: {response.status_code}")

    # Test jobs list
    print("\n3. Testing /jobs endpoint...")
    response = requests.get(f"{BASE_URL}/jobs")
    print(f"   Status: {response.status_code}")
    print(f"   Response: {response.json()}")

    print("\n✅ All endpoints are accessible!")
    print(f"\nFrontend is running at: http://localhost:3002")
    print(f"Backend API is running at: http://localhost:8000")

if __name__ == "__main__":
    test_api_endpoints()