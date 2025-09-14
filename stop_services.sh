#!/bin/bash

echo "Stopping Translation Pipeline Services..."

# Stop backend
pkill -f "backend/app/main.py" 2>/dev/null
echo "✅ Backend stopped"

# Stop frontend
pkill -f "next dev" 2>/dev/null
echo "✅ Frontend stopped"

echo "🛑 All services stopped"