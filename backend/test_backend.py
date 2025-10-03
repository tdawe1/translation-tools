#!/usr/bin/env python3
"""
Simple test script to verify the backend is working
"""
import requests
import time

def test_backend():
    base_url = "http://localhost:8000"

    # Test health endpoint
    print("Testing health endpoint...")
    try:
        response = requests.get(f"{base_url}/health")
        print(f"Health check: {response.json()}")
    except Exception as e:
        print(f"Failed to connect to backend: {e}")
        return False

    # Test file upload
    print("\nTesting file upload...")
    # Create a dummy file for testing
    with open("test.txt", "w") as f:
        f.write("This is a test file")

    try:
        with open("test.txt", "rb") as f:
            files = {"file": f}
            response = requests.post(f"{base_url}/upload", files=files)
            print(f"Upload response: {response.json()}")
    except Exception as e:
        print(f"Upload failed: {e}")
        return False
    finally:
        # Clean up
        import os
        if os.path.exists("test.txt"):
            os.remove("test.txt")

    print("\nBackend is working correctly!")
    return True

if __name__ == "__main__":
    test_backend()