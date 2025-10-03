import asyncio
import json
import logging
from typing import Dict, Set, Any
from datetime import datetime

from fastapi import WebSocket, WebSocketDisconnect

logger = logging.getLogger(__name__)

class ConnectionManager:
    def __init__(self):
        # Store active connections by user_id
        self.active_connections: Dict[str, Set[WebSocket]] = {}
        # Store user subscriptions to job updates
        self.job_subscriptions: Dict[str, Set[str]] = {}  # user_id -> set of job_ids

    async def connect(self, websocket: WebSocket, user_id: str):
        """Accept a WebSocket connection"""
        await websocket.accept()
        if user_id not in self.active_connections:
            self.active_connections[user_id] = set()
        self.active_connections[user_id].add(websocket)
        logger.info(f"WebSocket connected for user {user_id}")

    def disconnect(self, websocket: WebSocket, user_id: str):
        """Remove a WebSocket connection"""
        if user_id in self.active_connections:
            self.active_connections[user_id].discard(websocket)
            if not self.active_connections[user_id]:
                del self.active_connections[user_id]
        logger.info(f"WebSocket disconnected for user {user_id}")

    async def send_personal_message(self, message: str, user_id: str):
        """Send a message to all connections of a specific user"""
        if user_id in self.active_connections:
            connections_to_remove = set()
            for connection in self.active_connections[user_id]:
                try:
                    await connection.send_text(message)
                except Exception as e:
                    logger.error(f"Failed to send message to WebSocket: {e}")
                    connections_to_remove.add(connection)

            # Clean up dead connections
            for conn in connections_to_remove:
                self.active_connections[user_id].discard(conn)

    async def broadcast_to_user(self, data: Dict[str, Any], user_id: str):
        """Broadcast data to all connections of a user"""
        message = json.dumps(data)
        await self.send_personal_message(message, user_id)

    async def send_job_update(self, job_id: str, user_id: str, update_type: str, data: Dict[str, Any]):
        """Send a job update to subscribed users"""
        if user_id in self.job_subscriptions and job_id in self.job_subscriptions[user_id]:
            message = {
                "type": "job_update",
                "job_id": job_id,
                "update_type": update_type,
                "timestamp": datetime.now().isoformat(),
                "data": data
            }
            await self.broadcast_to_user(message, user_id)

    def subscribe_to_job(self, user_id: str, job_id: str):
        """Subscribe a user to job updates"""
        if user_id not in self.job_subscriptions:
            self.job_subscriptions[user_id] = set()
        self.job_subscriptions[user_id].add(job_id)
        logger.info(f"User {user_id} subscribed to job {job_id}")

    def unsubscribe_from_job(self, user_id: str, job_id: str):
        """Unsubscribe a user from job updates"""
        if user_id in self.job_subscriptions:
            self.job_subscriptions[user_id].discard(job_id)
            if not self.job_subscriptions[user_id]:
                del self.job_subscriptions[user_id]
        logger.info(f"User {user_id} unsubscribed from job {job_id}")

    async def send_queue_update(self, user_id: str, queue_stats: Dict[str, Any]):
        """Send queue statistics update"""
        message = {
            "type": "queue_update",
            "timestamp": datetime.now().isoformat(),
            "data": queue_stats
        }
        await self.broadcast_to_user(message, user_id)

    async def send_notification(self, user_id: str, notification_type: str, title: str, message: str, data: Dict[str, Any] = None):
        """Send a notification to a user"""
        notification = {
            "type": "notification",
            "notification_type": notification_type,
            "title": title,
            "message": message,
            "timestamp": datetime.now().isoformat(),
            "data": data or {}
        }
        await self.broadcast_to_user(notification, user_id)

# Global connection manager instance
manager = ConnectionManager()