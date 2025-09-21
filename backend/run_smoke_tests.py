#!/usr/bin/env python3
"""
Comprehensive smoke test runner for the translation pipeline backend.

This script runs all end-to-end smoke tests to ensure the API is working correctly.
Tests are run with mocked translation services to avoid API costs.

Usage:
    python run_smoke_tests.py           # Run all smoke tests
    python run_smoke_tests.py --simple  # Run only simple tests
    python run_smoke_tests.py --comprehensive  # Run comprehensive tests
"""

import subprocess
import sys
import os
import argparse
from pathlib import Path

# Get the path to the virtual environment's python
VENV_PYTHON = Path(__file__).parent / "backend_venv" / "bin" / "python"
PYTHON_CMD = sys.executable  # Use current python if venv doesn't exist

if VENV_PYTHON.exists():
    PYTHON_CMD = str(VENV_PYTHON)

def run_smoke_tests(test_type="all"):
    """Run the smoke tests"""
    print("=" * 60)
    print("Running Translation Pipeline Smoke Tests")
    print("=" * 60)
    print(f"Test type: {test_type}")
    print(f"Python executable: {PYTHON_CMD}")
    print()

    # Change to backend directory
    backend_dir = Path(__file__).parent
    os.chdir(backend_dir)

    # Build pytest command based on test type
    cmd = [PYTHON_CMD, "-m", "pytest", "-v", "--tb=short", "--color=yes"]

    # Set environment variables for the subprocess
    env = os.environ.copy()
    env["PYTHONPATH"] = str(backend_dir)

    # Use tests directory in parent directory (project root)
    tests_dir = Path("../tests")

    if test_type == "simple":
        cmd.extend([
            str(tests_dir / "test_smoke_simple.py"),
            str(tests_dir / "test_main.py"),
            "-m", "not comprehensive"
        ])
    elif test_type == "comprehensive":
        comprehensive_file = tests_dir / "test_smoke_comprehensive.py"
        cmd.extend([
            str(comprehensive_file),
            "-m", "comprehensive" if comprehensive_file.exists() and "comprehensive" in open(comprehensive_file).read() else ""
        ])
    else:  # all
        cmd.extend([
            str(tests_dir / "test_smoke_simple.py"),
            str(tests_dir / "test_smoke_comprehensive.py"),
            str(tests_dir / "test_main.py")
        ])

    # Add common options
    cmd.extend(["--disable-warnings"])

    print(f"Running command: {' '.join(cmd)}")
    print()

    try:
        result = subprocess.run(cmd, env=env, check=True)
        print("\n" + "=" * 60)
        print("✅ All smoke tests passed!")
        print("=" * 60)
        return 0
    except subprocess.CalledProcessError as e:
        print("\n" + "=" * 60)
        print("❌ Smoke tests failed!")
        print(f"Exit code: {e.returncode}")
        print("=" * 60)
        return e.returncode

def main():
    parser = argparse.ArgumentParser(description="Run smoke tests for Translation Pipeline Backend")
    parser.add_argument("--simple", action="store_true", help="Run only simple smoke tests")
    parser.add_argument("--comprehensive", action="store_true", help="Run comprehensive smoke tests")
    args = parser.parse_args()

    if args.comprehensive:
        test_type = "comprehensive"
    elif args.simple:
        test_type = "simple"
    else:
        test_type = "all"

    exit_code = run_smoke_tests(test_type)
    sys.exit(exit_code)


if __name__ == "__main__":
    main()