#!/usr/bin/env python3
"""
translation_websocket_server.py

WebSocket server for real-time translation progress updates.
Handles multiple clients, job progress tracking, and connection management.
"""

import asyncio
import json
import logging
import time
from datetime import datetime
from typing import Dict, Set, Any, Optional
from dataclasses import dataclass, asdict
import websockets
from websockets.server import WebSocketServerProtocol

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@dataclass
class JobProgress:
    """Track progress of a translation job"""
    job_id: str
    status: str = "queued"  # queued, processing, completed, failed
    progress: float = 0.0  # 0.0 to 100.0
    stage: str = "initializing"  # initializing, extracting, translating, applying, finalizing
    tokens_processed: int = 0
    total_tokens: int = 0
    cost: float = 0.0
    estimated_cost: float = 0.0
    quality_score: Optional[float] = None
    error_message: Optional[str] = None
    file_name: str = ""
    file_size: int = 0
    start_time: Optional[float] = None
    eta_seconds: Optional[int] = None
    current_batch: int = 0
    total_batches: int = 0

class TranslationWebSocketServer:
    """WebSocket server for real-time translation updates"""

    def __init__(self, host: str = "localhost", port: int = 8081):
        self.host = host
        self.port = port
        self.clients: Dict[str, WebSocketServerProtocol] = {}  # client_id -> websocket
        self.job_subscriptions: Dict[str, Set[str]] = {}  # job_id -> set of client_ids
        self.jobs: Dict[str, JobProgress] = {}  # job_id -> JobProgress
        self.client_jobs: Dict[str, Set[str]] = {}  # client_id -> set of job_ids
        self._client_counter = 0

    async def register_client(self, websocket: WebSocketServerProtocol) -> str:
        """Register a new client connection"""
        self._client_counter += 1
        client_id = f"client_{self._client_counter}"
        self.clients[client_id] = websocket
        self.client_jobs[client_id] = set()

        # Send connection confirmation
        await self.send_to_client(client_id, {
            "type": "connection_established",
            "client_id": client_id,
            "timestamp": datetime.now().isoformat()
        })

        logger.info(f"Client {client_id} connected")
        return client_id

    async def unregister_client(self, client_id: str):
        """Unregister a client"""
        if client_id in self.clients:
            del self.clients[client_id]

            # Clean up job subscriptions
            for job_id in self.client_jobs.get(client_id, set()):
                if job_id in self.job_subscriptions:
                    self.job_subscriptions[job_id].discard(client_id)
                    if not self.job_subscriptions[job_id]:
                        del self.job_subscriptions[job_id]

            if client_id in self.client_jobs:
                del self.client_jobs[client_id]

            logger.info(f"Client {client_id} disconnected")

    async def subscribe_to_job(self, client_id: str, job_id: str):
        """Subscribe a client to job updates"""
        if client_id not in self.clients:
            return

        if job_id not in self.job_subscriptions:
            self.job_subscriptions[job_id] = set()

        self.job_subscriptions[job_id].add(client_id)
        self.client_jobs[client_id].add(job_id)

        # Send current job state if it exists
        if job_id in self.jobs:
            await self.send_to_client(client_id, {
                "type": "job_progress",
                "job_id": job_id,
                **asdict(self.jobs[job_id])
            })

        logger.info(f"Client {client_id} subscribed to job {job_id}")

    async def unsubscribe_from_job(self, client_id: str, job_id: str):
        """Unsubscribe a client from job updates"""
        if job_id in self.job_subscriptions:
            self.job_subscriptions[job_id].discard(client_id)

        if client_id in self.client_jobs:
            self.client_jobs[client_id].discard(job_id)

        logger.info(f"Client {client_id} unsubscribed from job {job_id}")

    async def create_job(self, job_id: str, file_name: str, file_size: int, estimated_tokens: int, estimated_cost: float):
        """Create a new job"""
        job = JobProgress(
            job_id=job_id,
            file_name=file_name,
            file_size=file_size,
            total_tokens=estimated_tokens,
            estimated_cost=estimated_cost,
            start_time=time.time()
        )
        self.jobs[job_id] = job
        await self.broadcast_job_update(job_id)
        logger.info(f"Created job {job_id} for {file_name}")

    async def update_job_progress(self, job_id: str, **kwargs):
        """Update job progress"""
        if job_id not in self.jobs:
            logger.warning(f"Attempted to update non-existent job {job_id}")
            return

        job = self.jobs[job_id]

        # Update fields
        for key, value in kwargs.items():
            if hasattr(job, key):
                setattr(job, key, value)

        # Calculate ETA if we have progress and start time
        if job.start_time and job.progress > 0:
            elapsed = time.time() - job.start_time
            if job.progress < 100:
                total_estimated = elapsed / (job.progress / 100)
                job.eta_seconds = int(total_estimated - elapsed)

        await self.broadcast_job_update(job_id)

    async def broadcast_job_update(self, job_id: str):
        """Broadcast job update to all subscribed clients"""
        if job_id not in self.jobs:
            return

        if job_id not in self.job_subscriptions:
            return

        message = {
            "type": "job_progress",
            "job_id": job_id,
            **asdict(self.jobs[job_id])
        }

        for client_id in self.job_subscriptions[job_id]:
            await self.send_to_client(client_id, message)

    async def send_to_client(self, client_id: str, message: Dict[str, Any]):
        """Send message to a specific client"""
        if client_id in self.clients:
            try:
                await self.clients[client_id].send(json.dumps(message))
            except websockets.exceptions.ConnectionClosed:
                await self.unregister_client(client_id)
            except Exception as e:
                logger.error(f"Error sending to client {client_id}: {e}")

    async def handle_client(self, websocket: WebSocketServerProtocol, path: str):
        """Handle a client connection"""
        client_id = await self.register_client(websocket)

        try:
            async for message in websocket:
                try:
                    data = json.loads(message)
                    await self.handle_message(client_id, data)
                except json.JSONDecodeError:
                    logger.error(f"Invalid JSON from client {client_id}")
                except Exception as e:
                    logger.error(f"Error handling message from client {client_id}: {e}")

        except websockets.exceptions.ConnectionClosed:
            pass
        finally:
            await self.unregister_client(client_id)

    async def handle_message(self, client_id: str, data: Dict[str, Any]):
        """Handle a message from a client"""
        message_type = data.get("type")

        if message_type == "subscribe":
            job_id = data.get("job_id")
            if job_id:
                await self.subscribe_to_job(client_id, job_id)

        elif message_type == "unsubscribe":
            job_id = data.get("job_id")
            if job_id:
                await self.unsubscribe_from_job(client_id, job_id)

        elif message_type == "ping":
            await self.send_to_client(client_id, {
                "type": "pong",
                "timestamp": datetime.now().isoformat()
            })

        elif message_type == "get_job_status":
            job_id = data.get("job_id")
            if job_id and job_id in self.jobs:
                await self.send_to_client(client_id, {
                    "type": "job_status",
                    "job_id": job_id,
                    **asdict(self.jobs[job_id])
                })

    async def send_heartbeat(self):
        """Send heartbeat to all connected clients"""
        while True:
            await asyncio.sleep(30)  # Send heartbeat every 30 seconds

            message = {
                "type": "heartbeat",
                "timestamp": datetime.now().isoformat(),
                "connected_clients": len(self.clients)
            }

            # Send to all clients
            for client_id in list(self.clients.keys()):
                await self.send_to_client(client_id, message)

    async def start_server(self):
        """Start the WebSocket server"""
        logger.info(f"Starting WebSocket server on {self.host}:{self.port}")

        # Start heartbeat task
        asyncio.create_task(self.send_heartbeat())

        # Start WebSocket server
        async with websockets.serve(self.handle_client, self.host, self.port):
            logger.info("WebSocket server started successfully")
            await asyncio.Future()  # Run forever

# Global server instance
server = None

async def get_server() -> TranslationWebSocketServer:
    """Get or create the server instance"""
    global server
    if server is None:
        server = TranslationWebSocketServer()
    return server

# Convenience functions for updating jobs
async def create_translation_job(job_id: str, file_name: str, file_size: int,
                               estimated_tokens: int, estimated_cost: float):
    """Create a new translation job"""
    srv = await get_server()
    await srv.create_job(job_id, file_name, file_size, estimated_tokens, estimated_cost)

async def update_job_progress(job_id: str, **kwargs):
    """Update job progress"""
    srv = await get_server()
    await srv.update_job_progress(job_id, **kwargs)

async def get_job_progress(job_id: str) -> Optional[JobProgress]:
    """Get current job progress"""
    srv = await get_server()
    return srv.jobs.get(job_id)

if __name__ == "__main__":
    # Run the server
    asyncio.run(TranslationWebSocketServer().start_server())