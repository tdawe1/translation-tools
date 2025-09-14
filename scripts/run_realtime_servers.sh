#!/bin/bash

# Script to start real-time progress tracking servers

echo "🚀 Starting Real-time Translation Progress Servers"
echo "================================================"

# Check if Python is available
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is required but not installed"
    exit 1
fi

# Check if required packages are installed
echo "📦 Checking dependencies..."
python3 -c "import websockets, aiohttp, aiohttp_cors" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  Installing missing dependencies..."
    pip install websockets aiohttp aiohttp-cors
fi

# Start the servers
echo "🔌 Starting WebSocket server on port 8081..."
echo "📡 Starting SSE fallback server on port 8082..."
echo ""
echo "Press Ctrl+C to stop the servers"
echo ""

# Run the server manager
python3 scripts/start_realtime_servers.py