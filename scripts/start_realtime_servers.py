#!/usr/bin/env python3
"""
start_realtime_servers.py

Script to start both WebSocket and SSE fallback servers for real-time updates.
"""

import asyncio
import logging
import signal
import sys
from translation_websocket_server import TranslationWebSocketServer
from sse_fallback_server import SSEFallbackServer

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class RealtimeServerManager:
    """Manages both WebSocket and SSE servers"""

    def __init__(self):
        self.ws_server = None
        self.sse_server = None
        self.running = False

    async def start(self):
        """Start both servers"""
        logger.info("Starting real-time servers...")

        # Start WebSocket server
        self.ws_server = TranslationWebSocketServer(host="localhost", port=8081)
        ws_task = asyncio.create_task(self._run_ws_server())

        # Start SSE fallback server
        self.sse_server = SSEFallbackServer(host="localhost", port=8082)
        sse_task = asyncio.create_task(self._run_sse_server())

        self.running = True

        logger.info("Both servers started successfully")
        logger.info("WebSocket server: ws://localhost:8081")
        logger.info("SSE fallback: http://localhost:8082/sse")

        # Wait for both servers
        await asyncio.gather(ws_task, sse_task)

    async def _run_ws_server(self):
        """Run WebSocket server"""
        try:
            await self.ws_server.start_server()
        except asyncio.CancelledError:
            logger.info("WebSocket server stopped")
        except Exception as e:
            logger.error(f"WebSocket server error: {e}")

    async def _run_sse_server(self):
        """Run SSE server"""
        try:
            await self.sse_server.start_server()
        except asyncio.CancelledError:
            logger.info("SSE server stopped")
        except Exception as e:
            logger.error(f"SSE server error: {e}")

    async def stop(self):
        """Stop all servers"""
        logger.info("Stopping servers...")
        self.running = False

        # Cancel all tasks
        for task in asyncio.all_tasks():
            if task is not asyncio.current_task():
                task.cancel()

        logger.info("All servers stopped")

    def handle_signal(self, signum, frame):
        """Handle shutdown signals"""
        logger.info(f"Received signal {signum}")
        asyncio.create_task(self.stop())

async def main():
    """Main entry point"""
    manager = RealtimeServerManager()

    # Setup signal handlers
    signal.signal(signal.SIGINT, manager.handle_signal)
    signal.signal(signal.SIGTERM, manager.handle_signal)

    try:
        await manager.start()
    except KeyboardInterrupt:
        await manager.stop()

if __name__ == "__main__":
    asyncio.run(main())