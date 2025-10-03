#!/usr/bin/env python3
"""
sse_fallback_server.py

HTTP server providing Server-Sent Events (SSE) fallback for real-time updates.
Used when WebSocket connection is not available.
"""

import asyncio
import json
import logging
from datetime import datetime
from typing import Dict, Set
from aiohttp import web, WSMsgType
import aiohttp_cors

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class SSEFallbackServer:
    """SSE fallback server for real-time updates"""

    def __init__(self, host: str = "localhost", port: int = 8082):
        self.host = host
        self.port = port
        self.clients: Dict[str, web.StreamResponse] = {}
        self.job_subscriptions: Dict[str, Set[str]] = {}

    async def handle_sse(self, request: web.Request):
        """Handle SSE connection"""
        client_id = f"sse_{id(request)}"
        response = web.StreamResponse(
            status=200,
            reason='OK',
            headers={
                'Content-Type': 'text/event-stream',
                'Cache-Control': 'no-cache',
                'Connection': 'keep-alive',
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Headers': 'Cache-Control'
            }
        )

        await response.prepare(request)
        self.clients[client_id] = response

        # Send initial connection message
        await self.send_to_client(client_id, {
            "type": "connection_established",
            "client_id": client_id,
            "timestamp": datetime.now().isoformat()
        })

        logger.info(f"SSE client {client_id} connected")

        try:
            # Keep connection alive
            while True:
                # Check if client is still connected
                if request.transport is None or request.transport.is_closing():
                    break

                # Send heartbeat every 30 seconds
                await self.send_to_client(client_id, {
                    "type": "heartbeat",
                    "timestamp": datetime.now().isoformat()
                })

                await asyncio.sleep(30)

        except (ConnectionResetError, asyncio.CancelledError):
            pass
        finally:
            # Cleanup
            if client_id in self.clients:
                del self.clients[client_id]

            # Clean up subscriptions
            for job_id, clients in self.job_subscriptions.items():
                if client_id in clients:
                    clients.remove(client_id)

            logger.info(f"SSE client {client_id} disconnected")

    async def handle_subscribe(self, request: web.Request):
        """Handle job subscription via HTTP POST"""
        try:
            data = await request.json()
            client_id = data.get('client_id')
            job_id = data.get('job_id')

            if client_id and job_id:
                if job_id not in self.job_subscriptions:
                    self.job_subscriptions[job_id] = set()

                self.job_subscriptions[job_id].add(client_id)
                logger.info(f"Client {client_id} subscribed to job {job_id} via SSE")

                return web.json_response({"status": "subscribed"})

        except Exception as e:
            logger.error(f"Error in SSE subscribe: {e}")

        return web.json_response({"status": "error"}, status=400)

    async def send_to_client(self, client_id: str, data: dict):
        """Send message to SSE client"""
        if client_id in self.clients:
            try:
                message = f"data: {json.dumps(data)}\n\n"
                await self.clients[client_id].write(message.encode('utf-8'))
            except Exception as e:
                logger.error(f"Error sending to SSE client {client_id}: {e}")
                # Remove disconnected client
                if client_id in self.clients:
                    del self.clients[client_id]

    async def broadcast_job_update(self, job_id: str, data: dict):
        """Broadcast job update to subscribed SSE clients"""
        if job_id not in self.job_subscriptions:
            return

        message = {
            "type": "job_progress",
            "job_id": job_id,
            **data
        }

        for client_id in self.job_subscriptions[job_id]:
            await self.send_to_client(client_id, message)

    async def start_server(self):
        """Start the SSE server"""
        app = web.Application()

        # Configure CORS
        cors = aiohttp_cors.setup(app, defaults={
            "*": aiohttp_cors.ResourceOptions(
                allow_credentials=True,
                expose_headers="*",
                allow_headers="*",
                allow_methods="*"
            )
        })

        # Routes
        sse_route = app.router.add_get('/sse', self.handle_sse)
        subscribe_route = app.router.add_post('/sse/subscribe', self.handle_subscribe)

        # Add CORS to routes
        cors.add(sse_route)
        cors.add(subscribe_route)

        # Health check
        async def health_check(request):
            return web.json_response({"status": "healthy", "clients": len(self.clients)})

        app.router.add_get('/health', health_check)

        # Start server
        runner = web.AppRunner(app)
        await runner.setup()
        site = web.TCPSite(runner, self.host, self.port)

        logger.info(f"SSE fallback server started on {self.host}:{self.port}")
        await site.start()

# Global server instance
sse_server = None

async def get_sse_server() -> SSEFallbackServer:
    """Get or create the SSE server instance"""
    global sse_server
    if sse_server is None:
        sse_server = SSEFallbackServer()
    return sse_server

async def broadcast_to_sse_clients(job_id: str, data: dict):
    """Broadcast job update to SSE clients"""
    srv = await get_sse_server()
    await srv.broadcast_job_update(job_id, data)

if __name__ == "__main__":
    async def main():
        server = SSEFallbackServer()
        await server.start_server()

        try:
            # Keep running
            while True:
                await asyncio.sleep(3600)
        except KeyboardInterrupt:
            logger.info("Shutting down SSE server")

    asyncio.run(main())