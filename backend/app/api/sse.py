from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.responses import StreamingResponse
import asyncio
import json
import logging
from datetime import datetime
from typing import Dict, Any

from ..services.auth_service import auth_service
from ..core.job_manager import job_manager

logger = logging.getLogger(__name__)
router = APIRouter()
security = HTTPBearer()

# Store active SSE connections
active_connections: Dict[str, asyncio.Queue] = {}

@router.get("/subscribe")
async def sse_subscribe(
    token: str,
    job_id: str = None,
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """SSE endpoint for real-time job updates"""
    try:
        user_id = auth_service.verify_token(credentials.credentials)
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid token"
        )

    # Verify job ownership if job_id is provided
    if job_id:
        job = await job_manager.get_job(job_id)
        if not job or job.user_id != user_id:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Job not found"
            )

    # Create a queue for this connection
    queue = asyncio.Queue()
    connection_id = f"{user_id}_{datetime.now().timestamp()}"
    active_connections[connection_id] = queue

    async def event_generator():
        try:
            # Send initial connection message
            yield f"data: {json.dumps({'type': 'connected', 'connection_id': connection_id})}\n\n"

            # Keep connection alive and send events
            while True:
                try:
                    # Wait for an event with timeout
                    event = await asyncio.wait_for(queue.get(), timeout=30.0)

                    # Send the event
                    yield f"data: {json.dumps(event)}\n\n"

                except asyncio.TimeoutError:
                    # Send heartbeat to keep connection alive
                    yield f"data: {json.dumps({'type': 'heartbeat', 'timestamp': datetime.now().isoformat()})}\n\n"

        except Exception as e:
            logger.error(f"SSE connection error: {e}")
        finally:
            # Clean up
            if connection_id in active_connections:
                del active_connections[connection_id]

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Headers": "*",
        }
    )

async def send_sse_update(user_id: str, event_type: str, data: Dict[str, Any]):
    """Send an update to all SSE connections for a user"""
    event = {
        "type": event_type,
        "timestamp": datetime.now().isoformat(),
        "data": data
    }

    # Find all connections for this user
    for conn_id, queue in active_connections.items():
        if conn_id.startswith(f"{user_id}_"):
            try:
                queue.put_nowait(event)
            except asyncio.QueueFull:
                logger.warning(f"SSE queue full for connection {conn_id}")

# Utility functions to send specific updates
async def send_job_update_sse(user_id: str, job_id: str, update_type: str, data: Dict[str, Any]):
    """Send job update via SSE"""
    await send_sse_update(user_id, "job_update", {
        "job_id": job_id,
        "update_type": update_type,
        **data
    })

async def send_notification_sse(user_id: str, notification_type: str, title: str, message: str, data: Dict[str, Any] = None):
    """Send notification via SSE"""
    await send_sse_update(user_id, "notification", {
        "notification_type": notification_type,
        "title": title,
        "message": message,
        "data": data or {}
    })