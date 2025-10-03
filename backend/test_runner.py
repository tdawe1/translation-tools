#!/usr/bin/env python3
"""
Simple test runner for backend tests
"""
import os
import sys
import subprocess

def main():
    # Set environment variables for testing
    env = os.environ.copy()
    env.update({
        "DEBUG": "true",
        "SECRET_KEY": "test-secret-key-for-pytest-testing-only-32-chars-long",
        "OPENAI_API_KEY": "mock-sk-for-testing",
        "DATABASE_URL": "sqlite:///./test_translation_pipeline.db",
        "LOG_LEVEL": "WARNING",
        "UPLOAD_DIR": "test_uploads",
        "OUTPUT_DIR": "test_outputs",
        "PYTHONPATH": "."
    })

    # Run pytest
    cmd = [sys.executable, "-m", "pytest", "tests/test_main.py", "-v", "--tb=short"]

    try:
        result = subprocess.run(cmd, env=env, cwd=".", check=True)
        return result.returncode
    except subprocess.CalledProcessError as e:
        print(f"Tests failed with exit code: {e.returncode}")
        return e.returncode

if __name__ == "__main__":
    sys.exit(main())