#!/usr/bin/env python3
"""
Drive Poller CLI - Demo Version

Polls for new files (demo: posts mock manifests every 60 seconds).

In production, integrate with Google Drive API to watch for changes.

Requires: requests, python-dotenv
"""

import os
import time
import requests
from dotenv import load_dotenv
import json

load_dotenv()

API_URL = os.getenv("BACKEND_URL", "http://localhost:8000/api")

def post_mock_manifest():
    """Post a mock batch manifest to the API."""
    timestamp = int(time.time())
    manifest = {
        "type": "batch",
        "idempotency_key": f"poll-demo-{timestamp}",
        "jobs": [
            {
                "input": f"drive://demo-pptx-{timestamp}.pptx",
                "file_type": "pptx",
                "model": "gpt-4o-mini"
            },
            {
                "input": f"drive://demo-pdf-{timestamp}.pdf",
                "file_type": "pdf",
                "model": "gpt-4o-mini",
                "pages": "1-5"
            }
        ]
    }
    try:
        response = requests.post(
            f"{API_URL}/jobs",
            json=manifest,
            headers={"Content-Type": "application/json"},
            timeout=10
        )
        if response.status_code == 200:
            print(f"✅ Submitted mock manifest {timestamp}: {response.json()}")
        else:
            print(f"❌ Failed to submit manifest {timestamp}: {response.status_code} - {response.text}")
    except requests.RequestException as e:
        print(f"❌ Network error submitting manifest {timestamp}: {e}")

if __name__ == "__main__":
    print("🚀 Starting Drive Poller CLI (Demo Mode)")
    print("📡 Posting mock batch manifests every 60 seconds to", API_URL)
    print("⏹️  Press Ctrl+C to stop\n")
    try:
        while True:
            post_mock_manifest()
            time.sleep(60)
    except KeyboardInterrupt:
        print("\n👋 Stopped Drive Poller CLI")
