from fastapi import APIRouter, WebSocket, WebSocketDisconnect, Depends, HTTPException, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
import json
import logging

from .manager import manager
from ..services.auth_service import auth_service
from ..core.job_manager import job_manager

logger = logging.getLogger(__name__)
router = APIRouter()
security = HTTPBearer()

@router.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket, token: str):
    """WebSocket endpoint for real-time job updates"""
    try:
        # Verify the token
        user_id = auth_service.verify_token(token)
    except Exception as e:
        await websocket.close(code=status.WS_1008_POLICY_VIOLATION, reason="Invalid token")
        return

    # Accept the connection
    await manager.connect(websocket, user_id)

    try:
        while True:
            # Receive messages from the client
            data = await websocket.receive_text()
            try:
                message = json.loads(data)
                await handle_websocket_message(websocket, user_id, message)
            except json.JSONDecodeError:
                await websocket.send_text(json.dumps({
                    "type": "error",
                    "message": "Invalid JSON format"
                }))
            except Exception as e:
                logger.error(f"Error handling WebSocket message: {e}")
                await websocket.send_text(json.dumps({
                    "type": "error",
                    "message": "Failed to process message"
                }))

    except WebSocketDisconnect:
        manager.disconnect(websocket, user_id)
    except Exception as e:
        logger.error(f"WebSocket error: {e}")
        manager.disconnect(websocket, user_id)

async def handle_websocket_message(websocket: WebSocket, user_id: str, message: Dict):
    """Handle incoming WebSocket messages"""
    message_type = message.get("type")

    if message_type == "subscribe":
        # Subscribe to job updates
        job_id = message.get("job_id")
        if job_id:
            # Verify user owns the job
            job = await job_manager.get_job(job_id)
            if job and job.user_id == user_id:
                manager.subscribe_to_job(user_id, job_id)
                await websocket.send_text(json.dumps({
                    "type": "subscription_confirmed",
                    "job_id": job_id
                }))
            else:
                await websocket.send_text(json.dumps({
                    "type": "error",
                    "message": "Job not found or access denied"
                }))

    elif message_type == "unsubscribe":
        # Unsubscribe from job updates
        job_id = message.get("job_id")
        if job_id:
            manager.unsubscribe_from_job(user_id, job_id)
            await websocket.send_text(json.dumps({
                "type": "unsubscription_confirmed",
                "job_id": job_id
            }))

    elif message_type == "ping":
        # Respond to ping
        await websocket.send_text(json.dumps({
            "type": "pong",
            "timestamp": message.get("timestamp")
        }))

    elif message_type == "get_queue_status":
        # Send current queue status
        jobs, _ = await job_manager.list_jobs(user_id, limit=1000)
        status_counts = {}
        for job in jobs:
            status_counts[job.status] = status_counts.get(job.status, 0) + 1

        await websocket.send_text(json.dumps({
            "type": "queue_status",
            "data": {
                "status_counts": status_counts,
                "active_jobs": len([j for j in jobs if j.status in ["pending", "running"]])
            }
        }))

    else:
        await websocket.send_text(json.dumps({
            "type": "error",
            "message": f"Unknown message type: {message_type}"
        }))