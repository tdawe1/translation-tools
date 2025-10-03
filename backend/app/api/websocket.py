import json
from typing import Optional, Dict, Any
from datetime import datetime
from ..core.job_manager import job_manager

from fastapi import APIRouter, WebSocket, WebSocketDisconnect

router = APIRouter()

class ConnectionManager:
    def __init__(self):
        self.active_connections: list[WebSocket] = []
        self.batch_progress: Dict[str, Dict[str, Any]] = {}

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        self.active_connections.remove(websocket)

    async def broadcast(self, message: Dict[str, Any]):
        disconnected = []
        for connection in self.active_connections:
            try:
                await connection.send_text(json.dumps(message))
            except:
                disconnected.append(connection)
        for conn in disconnected:
            self.active_connections.remove(conn)

    async def broadcast_drive_update(self, job_id: str, status: str, manifest: Optional[Dict] = None):
        msg = {
            "type": "drive_job_update",
            "job_id": job_id,
            "status": status,
            "timestamp": datetime.utcnow().isoformat()
        }
        if manifest:
            msg["manifest"] = manifest
        await self.broadcast(msg)

    async def send_job_update(self, job_id: str, user_id: str, event: str, data: Dict[str, Any]) -> None:
        msg = {
            "type": "job_update",
            "job_id": job_id,
            "event": event,
            "data": data,
            "user_id": user_id,
            "timestamp": datetime.utcnow().isoformat()
        }
        await self.broadcast(msg)

        if event in ("completed", "failed"):
            job = await job_manager.get_job(job_id)
            if job and "batch_id" in job.metadata:
                batch_id = job.metadata["batch_id"]
                if batch_id in self.batch_progress:
                    self.batch_progress[batch_id]["completed"] += 1
                    total = self.batch_progress[batch_id]["total"]
                    progress = round((self.batch_progress[batch_id]["completed"] / total) * 100, 1)
                    self.batch_progress[batch_id]["progress"] = progress
                    status = "completed" if self.batch_progress[batch_id]["completed"] == total else "processing"
                    await self.broadcast_batch_update(batch_id, status, progress)

    async def broadcast_batch_update(self, batch_id: str, status: str, progress: Optional[float] = None):
        msg = {
            "type": "batch_update",
            "batch_id": batch_id,
            "status": status,
            "timestamp": datetime.utcnow().isoformat()
        }
        if progress is not None:
            msg["progress"] = progress
        await self.broadcast(msg)

    async def start_batch(self, batch_id: str, total: int):
        self.batch_progress[batch_id] = {
            "total": total,
            "completed": 0,
            "progress": 0.0,
            "job_ids": []
        }

    async def add_job_to_batch(self, batch_id: str, job_id: str):
        if batch_id in self.batch_progress:
            self.batch_progress[batch_id]["job_ids"].append(job_id)

    def get_batch_progress(self, batch_id: str) -> Optional[Dict[str, Any]]:
        return self.batch_progress.get(batch_id)

manager = ConnectionManager()

@router.websocket("/ws/drive")
async def drive_websocket(websocket: WebSocket):
    """WebSocket endpoint for drive job updates."""
    await manager.connect(websocket)
    try:
        while True:
            # Keep connection alive, handle pings if needed
            data = await websocket.receive_text()
            # Echo or handle
            await websocket.send_text(f"Message: {data}")
    except WebSocketDisconnect:
        manager.disconnect(websocket)