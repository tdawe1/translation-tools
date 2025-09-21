#!/usr/bin/env python3
"""
Run smoke tests for the backend API
"""
import subprocess
import sys
from pathlib import Path

def main():
    """Run smoke tests"""
    # Change to backend directory
    backend_dir = Path(__file__).parent.parent
    os.chdir(backend_dir)

    # Run pytest with smoke marker
    cmd = [
        sys.executable,
        "-m",
        "pytest",
        "tests/",
        "-v",
        "-m",
        "smoke",
        "--tb=short"
    ]

    print("🔥 Running smoke tests...")
    print(f"Command: {' '.join(cmd)}")
    print()

    result = subprocess.run(cmd, cwd=backend_dir)

    if result.returncode == 0:
        print("\n✅ All smoke tests passed!")
    else:
        print("\n❌ Some smoke tests failed!")
        sys.exit(result.returncode)

if __name__ == "__main__":
    import os
    main()