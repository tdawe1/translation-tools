#!/bin/bash

echo "Starting Translation Pipeline Services..."

# Check if backend is running
if curl -s http://localhost:8000/health > /dev/null 2>&1; then
    echo "✅ Backend is already running on http://localhost:8000"
else
    echo "🚀 Starting backend..."
    ./backend/backend_venv/bin/python backend/app/main.py > backend.log 2>&1 &
    sleep 3
    if curl -s http://localhost:8000/health > /dev/null 2>&1; then
        echo "✅ Backend started successfully"
    else
        echo "❌ Failed to start backend"
        exit 1
    fi
fi

# Check if frontend is running
if curl -s http://localhost:3000 > /dev/null 2>&1; then
    echo "✅ Frontend is already running on http://localhost:3000"
else
    echo "🚀 Starting frontend..."
    cd frontend && npm run dev > ../frontend.log 2>&1 &
    sleep 5
    if curl -s http://localhost:3000 > /dev/null 2>&1; then
        echo "✅ Frontend started successfully"
    else
        echo "❌ Failed to start frontend"
        exit 1
    fi
fi

echo ""
echo "🎉 All services are running!"
echo "   Frontend: http://localhost:3000"
echo "   Backend API: http://localhost:8000"
echo "   API Docs: http://localhost:8000/docs"
echo ""
echo "To stop services, run: ./stop_services.sh"