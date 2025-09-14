#!/usr/bin/env python3
"""
progress_tracker.py

Module for tracking and broadcasting translation progress via WebSocket.
Can be used by both PPTX and PDF translation scripts.
"""

import asyncio
import json
import logging
import time
from datetime import datetime
from typing import Optional, Dict, Any
from dataclasses import dataclass, asdict

logger = logging.getLogger(__name__)

@dataclass
class TranslationProgress:
    """Track progress of a translation job"""
    job_id: str
    status: str = "queued"  # queued, extracting, translating, applying, completed, failed
    progress: float = 0.0  # 0.0 to 100.0
    stage: str = "initializing"  # More detailed stage info
    tokens_processed: int = 0
    total_tokens: int = 0
    current_cost: float = 0.0
    estimated_cost: float = 0.0
    quality_score: Optional[float] = None
    error_message: Optional[str] = None
    file_name: str = ""
    file_size: int = 0
    start_time: Optional[float] = None
    eta_seconds: Optional[int] = None
    current_batch: int = 0
    total_batches: int = 0
    current_file_progress: float = 0.0  # For multi-file scenarios

class ProgressTracker:
    """Track and broadcast translation progress"""

    def __init__(self, websocket_url: str = "ws://localhost:8081"):
        self.websocket_url = websocket_url
        self.websocket = None
        self.progress = None
        self._connected = False
        self._reconnect_attempts = 0
        self._max_reconnect_attempts = 5
        self._broadcast_queue = asyncio.Queue()
        self._broadcast_task = None

    async def connect(self):
        """Connect to WebSocket server"""
        try:
            import websockets
            self.websocket = await websockets.connect(self.websocket_url)
            self._connected = True
            self._reconnect_attempts = 0
            logger.info("Connected to WebSocket server")

            # Start broadcast task
            if self._broadcast_task is None:
                self._broadcast_task = asyncio.create_task(self._broadcast_worker())

        except Exception as e:
            logger.error(f"Failed to connect to WebSocket: {e}")
            self._connected = False

    async def disconnect(self):
        """Disconnect from WebSocket server"""
        if self.websocket:
            await self.websocket.close()
            self.websocket = None
        self._connected = False

        # Cancel broadcast task
        if self._broadcast_task:
            self._broadcast_task.cancel()
            self._broadcast_task = None

    async def _broadcast_worker(self):
        """Worker to handle broadcasting progress updates"""
        while True:
            try:
                update = await self._broadcast_queue.get()
                await self._send_update(update)
            except asyncio.CancelledError:
                break
            except Exception as e:
                logger.error(f"Error in broadcast worker: {e}")

    async def _send_update(self, update: Dict[str, Any]):
        """Send update to WebSocket server"""
        if not self._connected:
            await self._ensure_connection()

        if self._connected and self.websocket:
            try:
                message = json.dumps(update)
                await self.websocket.send(message)
            except Exception as e:
                logger.error(f"Failed to send update: {e}")
                self._connected = False

    async def _ensure_connection(self):
        """Ensure connection to WebSocket server"""
        if not self._connected and self._reconnect_attempts < self._max_reconnect_attempts:
            self._reconnect_attempts += 1
            delay = min(2 ** self._reconnect_attempts, 30)  # Exponential backoff
            logger.info(f"Attempting to reconnect in {delay} seconds...")
            await asyncio.sleep(delay)
            await self.connect()

    async def start_job(self, job_id: str, file_name: str, file_size: int,
                       estimated_tokens: int, estimated_cost: float):
        """Start tracking a new job"""
        self.progress = TranslationProgress(
            job_id=job_id,
            file_name=file_name,
            file_size=file_size,
            total_tokens=estimated_tokens,
            estimated_cost=estimated_cost,
            start_time=time.time()
        )

        update = {
            "type": "job_started",
            "job_id": job_id,
            "file_name": file_name,
            "file_size": file_size,
            "estimated_tokens": estimated_tokens,
            "estimated_cost": estimated_cost
        }

        await self._broadcast_queue.put(update)
        await self.update_progress(status="extracting")

    async def update_progress(self, **kwargs):
        """Update job progress"""
        if not self.progress:
            return

        # Update fields
        for key, value in kwargs.items():
            if hasattr(self.progress, key):
                setattr(self.progress, key, value)

        # Calculate overall progress based on stage
        if self.progress.total_tokens > 0:
            # Weight progress by stage importance
            stage_weights = {
                "extracting": 0.2,
                "translating": 0.6,
                "applying": 0.15,
                "finalizing": 0.05
            }

            if self.progress.stage in stage_weights:
                stage_weight = stage_weights[self.progress.stage]
                stage_progress = self.progress.tokens_processed / max(self.progress.total_tokens, 1)

                # Calculate progress from previous stages
                prev_stages = list(stage_weights.keys())
                prev_index = prev_stages.index(self.progress.stage)
                prev_progress = sum(stage_weights[s] for s in prev_stages[:prev_index])

                self.progress.progress = prev_progress + (stage_weight * stage_progress)

        # Calculate ETA
        if self.progress.start_time and self.progress.progress > 0:
            elapsed = time.time() - self.progress.start_time
            if self.progress.progress < 100:
                total_estimated = elapsed / (self.progress.progress / 100)
                self.progress.eta_seconds = int(total_estimated - elapsed)

        # Create update message
        update = {
            "type": "job_progress",
            **asdict(self.progress)
        }

        await self._broadcast_queue.put(update)

    async def update_tokens(self, tokens_processed: int, cost_increment: float = 0.0):
        """Update token progress and cost"""
        if self.progress:
            self.progress.tokens_processed = tokens_processed
            self.progress.current_cost += cost_increment
            await self.update_progress()

    async def update_stage(self, stage: str, status: str = None):
        """Update processing stage"""
        updates = {"stage": stage}
        if status:
            updates["status"] = status
        await self.update_progress(**updates)

    async def update_batch_progress(self, current_batch: int, total_batches: int):
        """Update batch processing progress"""
        await self.update_progress(
            current_batch=current_batch,
            total_batches=total_batches
        )

    async def set_quality_score(self, score: float):
        """Set final quality score"""
        if self.progress:
            self.progress.quality_score = score
            await self.update_progress()

    async def complete_job(self, success: bool = True, error_message: str = None):
        """Mark job as completed"""
        if not self.progress:
            return

        self.progress.status = "completed" if success else "failed"
        self.progress.progress = 100.0
        self.progress.eta_seconds = 0

        if error_message:
            self.progress.error_message = error_message

        update = {
            "type": "job_completed" if success else "job_failed",
            "job_id": self.progress.job_id,
            "success": success,
            "error_message": error_message
        }

        await self._broadcast_queue.put(update)

    async def send_heartbeat(self):
        """Send heartbeat to keep connection alive"""
        if self._connected:
            try:
                await self.websocket.send(json.dumps({
                    "type": "ping",
                    "timestamp": datetime.now().isoformat()
                }))
            except Exception as e:
                logger.error(f"Heartbeat failed: {e}")
                self._connected = False

# Global tracker instance
_tracker = None

def get_tracker() -> ProgressTracker:
    """Get global progress tracker instance"""
    global _tracker
    if _tracker is None:
        _tracker = ProgressTracker()
    return _tracker

# Convenience functions for direct use
async def start_translation_job(job_id: str, file_name: str, file_size: int,
                              estimated_tokens: int, estimated_cost: float):
    """Start tracking a translation job"""
    tracker = get_tracker()
    await tracker.start_job(job_id, file_name, file_size, estimated_tokens, estimated_cost)

async def update_translation_progress(**kwargs):
    """Update translation progress"""
    tracker = get_tracker()
    await tracker.update_progress(**kwargs)

async def complete_translation_job(success: bool = True, error_message: str = None):
    """Complete translation job"""
    tracker = get_tracker()
    await tracker.complete_job(success, error_message)

# Synchronous wrapper for non-async contexts
def sync_update_progress(job_id: str, **kwargs):
    """Synchronously update progress (creates new event loop if needed)"""
    try:
        loop = asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

    if loop.is_running():
        # Create task if loop is running
        asyncio.create_task(update_translation_progress(job_id=job_id, **kwargs))
    else:
        # Run in new loop
        loop.run_until_complete(update_translation_progress(job_id=job_id, **kwargs))