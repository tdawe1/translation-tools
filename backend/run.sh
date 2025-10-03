#!/bin/bash

# Start the FastAPI backend

# Set environment variables
export PYTHONPATH="${PYTHONPATH}:$(pwd)"

# Install dependencies if needed
if [ ! -d "venv" ]; then
    python -m venv venv
fi

source venv/bin/activate
pip install -r requirements.txt

# Create .env file if it doesn't exist
if [ ! -f ".env" ]; then
    cp .env.example .env
    echo "Created .env file. Please edit it with your configuration."
fi

# Start the server
echo "Starting FastAPI server on http://localhost:8000"
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload