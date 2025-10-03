#!/bin/bash
# Script to run backend tests with the pytest fixtures

echo "Running backend tests with pytest fixtures..."
echo "=============================================="

# Navigate to backend directory
cd backend

# Set PYTHONPATH to include app directory
export PYTHONPATH=.

# Run pytest with verbose output
python -m pytest tests/ -v --tb=short

echo ""
echo "Test run completed."