import asyncio
import os
import uuid
import json
import logging
from datetime import datetime, timedelta
from typing import Dict, Optional, List, Tuple, Any
from pathlib import Path
import subprocess
import sys
import re
from collections import defaultdict
import sqlite3
import threading

from ..models.job import Job, JobStatus, TranslationRequest
from ..core.config import settings

logger = logging.getLogger(__name__)

# Import real-time update functions
try:
    from ..websocket.manager import manager
    from ..api.sse import send_job_update_sse, send_notification_sse
    HAS_REALTIME = True
except ImportError:
    HAS_REALTIME = False
    logger.warning("Real-time update modules not available")

class JobManager:
    def __init__(self):
        self.jobs: Dict[str, Job] = {}
        self.active_jobs: Dict[str, asyncio.Task] = {}
        self.job_semaphore = asyncio.Semaphore(settings.MAX_CONCURRENT_JOBS)
        self.job_logs: Dict[str, List[Dict[str, Any]]] = {}
        self.job_stats: Dict[str, Dict[str, Any]] = {}
        self.db_path = Path(settings.OUTPUT_DIR) / "jobs.db"
        self._init_db()

    def _init_db(self):
        """Initialize SQLite database for job persistence"""
        try:
            conn = sqlite3.connect(str(self.db_path))
            cursor = conn.cursor()

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS jobs (
                    id TEXT PRIMARY KEY,
                    user_id TEXT NOT NULL,
                    status TEXT NOT NULL,
                    input_file TEXT NOT NULL,
                    output_file TEXT,
                    request_json TEXT NOT NULL,
                    progress REAL DEFAULT 0.0,
                    message TEXT,
                    error TEXT,
                    created_at TEXT NOT NULL,
                    started_at TEXT,
                    completed_at TEXT,
                    metadata_json TEXT,
                    estimated_cost REAL,
                    actual_cost REAL
                )
            ''')

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS job_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_id TEXT NOT NULL,
                    timestamp TEXT NOT NULL,
                    level TEXT NOT NULL,
                    message TEXT NOT NULL,
                    data TEXT,
                    FOREIGN KEY (job_id) REFERENCES jobs (id)
                )
            ''')

            conn.commit()
            conn.close()
        except Exception as e:
            logger.error(f"Failed to initialize job database: {e}")

    async def initialize(self):
        """Initialize the job manager"""
        logger.info("Initializing JobManager")
        # Create necessary directories
        os.makedirs(settings.UPLOAD_DIR, exist_ok=True)
        os.makedirs(settings.OUTPUT_DIR, exist_ok=True)

    async def shutdown(self):
        """Shutdown the job manager and cancel all active jobs"""
        logger.info("Shutting down JobManager")
        for job_id, task in self.active_jobs.items():
            if not task.done():
                task.cancel()
                try:
                    await task
                except asyncio.CancelledError:
                    logger.info(f"Job {job_id} cancelled during shutdown")

    async def create_job(
        self,
        user_id: str,
        input_file: str,
        request: TranslationRequest
    ) -> Job:
        """Create a new translation job"""
        job_id = str(uuid.uuid4())

        job = Job(
            id=job_id,
            user_id=user_id,
            status=JobStatus.PENDING,
            input_file=input_file,
            request=request,
            created_at=datetime.now(),
            metadata={
                "file_type": request.file_type,
                "model": request.model,
                "file_size": os.path.getsize(input_file)
            }
        )

        self.jobs[job_id] = job

        # Save to database
        try:
            conn = sqlite3.connect(str(self.db_path))
            cursor = conn.cursor()

            cursor.execute("""
                INSERT INTO jobs (
                    id, user_id, status, input_file, request_json,
                    created_at, metadata_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """, [
                job_id,
                user_id,
                job.status.value,
                input_file,
                json.dumps(request.dict()),
                job.created_at.isoformat(),
                json.dumps(job.metadata)
            ])

            conn.commit()
            conn.close()
        except Exception as e:
            logger.error(f"Failed to save job to database: {e}")

        # Add initial log
        self._add_job_log(job_id, "INFO", f"Job created: {Path(input_file).name}")

        # Start the job in the background
        task = asyncio.create_task(self._run_job(job_id))
        self.active_jobs[job_id] = task

        return job

    async def get_job(self, job_id: str) -> Optional[Job]:
        """Get a job by ID"""
        return self.jobs.get(job_id)

    async def list_jobs(self, user_id: str, skip: int = 0, limit: int = 100,
                   filters: Optional[Dict[str, Any]] = None,
                   sort_by: str = "created_at",
                   sort_order: str = "desc") -> Tuple[List[Job], int]:
        """List jobs for a user with filtering and sorting"""
        try:
            conn = sqlite3.connect(str(self.db_path))
            cursor = conn.cursor()

            # Build WHERE clause
            where_clauses = ["user_id = ?"]
            params = [user_id]

            if filters:
                if filters.get("status"):
                    where_clauses.append("status = ?")
                    params.append(filters["status"])

                if filters.get("file_type"):
                    where_clauses.append("JSON_EXTRACT(request_json, '$.file_type') = ?")
                    params.append(filters["file_type"])

                if filters.get("date_from"):
                    where_clauses.append("created_at >= ?")
                    params.append(filters["date_from"])

                if filters.get("date_to"):
                    where_clauses.append("created_at <= ?")
                    params.append(filters["date_to"])

                if filters.get("search"):
                    search = f"%{filters['search']}%"
                    where_clauses.append("(input_file LIKE ? OR message LIKE ? OR error LIKE ?)")
                    params.extend([search, search, search])

            where_sql = " AND ".join(where_clauses)

            # Get total count
            count_sql = f"SELECT COUNT(*) FROM jobs WHERE {where_sql}"
            cursor.execute(count_sql, params)
            total = cursor.fetchone()[0]

            # Build ORDER BY
            valid_sort_fields = ["created_at", "started_at", "completed_at", "progress", "status"]
            if sort_by not in valid_sort_fields:
                sort_by = "created_at"

            order_sql = f"ORDER BY {sort_by} {'DESC' if sort_order.lower() == 'desc' else 'ASC'}"

            # Get paginated results
            sql = f"""
                SELECT * FROM jobs
                WHERE {where_sql}
                {order_sql}
                LIMIT ? OFFSET ?
            """
            params.extend([limit, skip])

            cursor.execute(sql, params)
            rows = cursor.fetchall()

            # Convert to Job objects
            jobs = []
            for row in rows:
                job_data = {
                    "id": row[0],
                    "user_id": row[1],
                    "status": row[2],
                    "input_file": row[3],
                    "output_file": row[4],
                    "request": json.loads(row[5]),
                    "progress": row[6],
                    "message": row[7],
                    "error": row[8],
                    "created_at": datetime.fromisoformat(row[9]),
                    "started_at": datetime.fromisoformat(row[10]) if row[10] else None,
                    "completed_at": datetime.fromisoformat(row[11]) if row[11] else None,
                    "metadata": json.loads(row[12]) if row[12] else {},
                    "estimated_cost": row[13],
                    "actual_cost": row[14]
                }
                jobs.append(Job(**job_data))

            conn.close()
            return jobs, total

        except Exception as e:
            logger.error(f"Failed to list jobs: {e}")
            # Fallback to in-memory jobs
            user_jobs = [
                job for job in self.jobs.values()
                if job.user_id == user_id
            ]
            user_jobs.sort(key=lambda j: getattr(j, sort_by, j.created_at),
                          reverse=(sort_order.lower() == "desc"))
            return user_jobs[skip:skip + limit], len(user_jobs)

    async def cancel_job(self, job_id: str, user_id: str = None) -> bool:
        """Cancel a running job"""
        job = self.jobs.get(job_id)
        if not job:
            return False

        if user_id and job.user_id != user_id:
            return False

        if job_id in self.active_jobs:
            task = self.active_jobs[job_id]
            if not task.done():
                task.cancel()
                try:
                    await task
                except asyncio.CancelledError:
                    pass

            del self.active_jobs[job_id]
            job.status = JobStatus.CANCELLED
            job.completed_at = datetime.now()
            return True

        return False

    async def cancel_jobs(self, user_id: str, job_ids: List[str]) -> Dict[str, bool]:
        """Cancel multiple jobs"""
        results = {}
        for job_id in job_ids:
            job = self.jobs.get(job_id)
            if job and job.user_id == user_id:
                results[job_id] = await self.cancel_job(job_id)
            else:
                results[job_id] = False
        return results

    async def retry_jobs(self, user_id: str, job_ids: List[str]) -> List[str]:
        """Retry multiple failed jobs"""
        retried = []
        for job_id in job_ids:
            job = self.jobs.get(job_id)
            if job and job.user_id == user_id and job.status == JobStatus.FAILED:
                # Create a new job with the same parameters
                new_job = await self.create_job(
                    user_id=user_id,
                    input_file=job.input_file,
                    request=job.request
                )
                retried.append(new_job.id)
        return retried

    async def get_job_statistics(self, user_id: str, days: int = 30) -> Dict[str, Any]:
        """Get job statistics for a user"""
        try:
            conn = sqlite3.connect(str(self.db_path))
            cursor = conn.cursor()

            # Date range
            date_from = (datetime.now() - timedelta(days=days)).isoformat()

            # Basic counts
            cursor.execute("""
                SELECT status, COUNT(*) FROM jobs
                WHERE user_id = ? AND created_at >= ?
                GROUP BY status
            """, [user_id, date_from])
            status_counts = dict(cursor.fetchall())

            # Total jobs
            cursor.execute("""
                SELECT COUNT(*) FROM jobs
                WHERE user_id = ? AND created_at >= ?
            """, [user_id, date_from])
            total_jobs = cursor.fetchone()[0]

            # Average duration (for completed jobs)
            cursor.execute("""
                SELECT AVG(
                    (julianday(completed_at) - julianday(started_at)) * 24 * 60
                ) FROM jobs
                WHERE user_id = ? AND status = 'completed'
                AND created_at >= ? AND started_at IS NOT NULL
            """, [user_id, date_from])
            avg_duration = cursor.fetchone()[0] or 0

            # Total cost
            cursor.execute("""
                SELECT SUM(actual_cost) FROM jobs
                WHERE user_id = ? AND actual_cost IS NOT NULL
                AND created_at >= ?
            """, [user_id, date_from])
            total_cost = cursor.fetchone()[0] or 0

            # Daily stats for charts
            cursor.execute("""
                SELECT
                    DATE(created_at) as date,
                    COUNT(*) as count,
                    SUM(CASE WHEN status = 'completed' THEN 1 ELSE 0 END) as completed,
                    SUM(CASE WHEN status = 'failed' THEN 1 ELSE 0 END) as failed
                FROM jobs
                WHERE user_id = ? AND created_at >= ?
                GROUP BY DATE(created_at)
                ORDER BY date DESC
            """, [user_id, date_from])
            daily_stats = [
                {
                    "date": row[0],
                    "total": row[1],
                    "completed": row[2],
                    "failed": row[3]
                }
                for row in cursor.fetchall()
            ]

            # File type distribution
            cursor.execute("""
                SELECT JSON_EXTRACT(request_json, '$.file_type'), COUNT(*)
                FROM jobs
                WHERE user_id = ? AND created_at >= ?
                GROUP BY JSON_EXTRACT(request_json, '$.file_type')
            """, [user_id, date_from])
            file_type_dist = dict(cursor.fetchall())

            conn.close()

            return {
                "total_jobs": total_jobs,
                "status_counts": status_counts,
                "average_duration_minutes": round(avg_duration, 2),
                "total_cost": round(total_cost, 4),
                "daily_stats": daily_stats,
                "file_type_distribution": file_type_dist,
                "period_days": days
            }

        except Exception as e:
            logger.error(f"Failed to get job statistics: {e}")
            return {
                "total_jobs": 0,
                "status_counts": {},
                "average_duration_minutes": 0,
                "total_cost": 0,
                "daily_stats": [],
                "file_type_distribution": {},
                "period_days": days
            }

    async def get_job_logs(self, job_id: str) -> List[Dict[str, Any]]:
        """Get logs for a specific job"""
        try:
            conn = sqlite3.connect(str(self.db_path))
            cursor = conn.cursor()

            cursor.execute("""
                SELECT timestamp, level, message, data
                FROM job_logs
                WHERE job_id = ?
                ORDER BY timestamp DESC
                LIMIT 1000
            """, [job_id])

            logs = [
                {
                    "timestamp": row[0],
                    "level": row[1],
                    "message": row[2],
                    "data": json.loads(row[3]) if row[3] else None
                }
                for row in cursor.fetchall()
            ]

            conn.close()
            return logs

        except Exception as e:
            logger.error(f"Failed to get job logs: {e}")
            return self.job_logs.get(job_id, [])

    def _add_job_log(self, job_id: str, level: str, message: str, data: Optional[Dict] = None):
        """Add a log entry for a job"""
        try:
            conn = sqlite3.connect(str(self.db_path))
            cursor = conn.cursor()

            cursor.execute("""
                INSERT INTO job_logs (job_id, timestamp, level, message, data)
                VALUES (?, ?, ?, ?, ?)
            """, [
                job_id,
                datetime.now().isoformat(),
                level,
                message,
                json.dumps(data) if data else None
            ])

            conn.commit()
            conn.close()
        except Exception as e:
            logger.error(f"Failed to add job log: {e}")

    async def export_job_report(self, user_id: str, format: str = "csv") -> str:
        """Export job data in specified format"""
        try:
            jobs, _ = await self.list_jobs(user_id, limit=10000)

            if format.lower() == "csv":
                import csv
                import io

                output = io.StringIO()
                writer = csv.writer(output)

                # Header
                writer.writerow([
                    "Job ID", "File Name", "File Type", "Status", "Progress",
                    "Created At", "Started At", "Completed At", "Duration (min)",
                    "Estimated Cost", "Actual Cost", "Error"
                ])

                # Data
                for job in jobs:
                    duration = 0
                    if job.started_at and job.completed_at:
                        duration = (job.completed_at - job.started_at).total_seconds() / 60

                    writer.writerow([
                        job.id,
                        Path(job.input_file).name,
                        job.request.file_type,
                        job.status,
                        job.progress,
                        job.created_at.isoformat(),
                        job.started_at.isoformat() if job.started_at else "",
                        job.completed_at.isoformat() if job.completed_at else "",
                        round(duration, 2),
                        job.metadata.get("estimated_cost", ""),
                        job.metadata.get("actual_cost", ""),
                        job.error or ""
                    ])

                return output.getvalue()

            elif format.lower() == "json":
                report = {
                    "exported_at": datetime.now().isoformat(),
                    "user_id": user_id,
                    "total_jobs": len(jobs),
                    "jobs": [job.dict() for job in jobs]
                }
                return json.dumps(report, indent=2, default=str)

            else:
                raise ValueError(f"Unsupported export format: {format}")

        except Exception as e:
            logger.error(f"Failed to export job report: {e}")
            raise

    async def _run_job(self, job_id: str):
        """Execute a translation job"""
        job = self.jobs.get(job_id)
        if not job:
            logger.error(f"Job {job_id} not found")
            return

        async with self.job_semaphore:
            try:
                logger.info(f"Starting job {job_id}")
                job.status = JobStatus.RUNNING
                job.started_at = datetime.now()
                job.progress = 5.0

                # Update database
                self._update_job_in_db(job)
                self._add_job_log(job_id, "INFO", "Job started")

                # Send real-time updates
                if HAS_REALTIME:
                    asyncio.create_task(manager.send_job_update(
                        job_id, job.user_id, "started",
                        {"progress": job.progress, "status": job.status.value}
                    ))
                    asyncio.create_task(send_job_update_sse(
                        job.user_id, job_id, "started",
                        {"progress": job.progress, "status": job.status.value}
                    ))

                # Generate output filename
                input_path = Path(job.input_file)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = f"{input_path.stem}_translated_{timestamp}{input_path.suffix}"
                output_path = Path(settings.OUTPUT_DIR) / output_filename
                job.output_file = str(output_path)

                # Build the translation command
                if job.request.file_type == "pptx":
                    script_path = Path(settings.SCRIPTS_DIR) / "translate_pptx_inplace.py"
                    cmd = [
                        sys.executable, str(script_path),
                        "--in", job.input_file,
                        "--out", str(output_path),
                        "--model", job.request.model,
                        "--temperature", str(job.request.temperature)
                    ]
                elif job.request.file_type == "pdf":
                    script_path = Path(settings.SCRIPTS_DIR) / "translate_pdf.py"
                    cmd = [
                        sys.executable, str(script_path),
                        "--in", job.input_file,
                        "--out", str(output_path),
                        "--model", job.request.model,
                        "--temperature", str(job.request.temperature)
                    ]
                    if job.request.pages:
                        cmd.extend(["--pages", job.request.pages])
                else:
                    raise ValueError(f"Unsupported file type: {job.request.file_type}")

                if job.request.offline:
                    cmd.append("--offline")

                # Run the translation script
                self._add_job_log(job_id, "INFO", f"Executing: {' '.join(cmd)}")
                process = await asyncio.create_subprocess_exec(
                    *cmd,
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE,
                    env={**os.environ, "OPENAI_API_KEY": settings.OPENAI_API_KEY}
                )

                # Monitor the process
                job.progress = 15.0  # Initial progress
                self._update_job_in_db(job)

                # Parse output for progress
                async def monitor_output():
                    last_progress = 15.0
                    while True:
                        line = await process.stdout.readline()
                        if not line:
                            break
                        line_str = line.decode().strip()
                        if line_str:
                            # Try to extract progress from output
                            progress_match = re.search(r'(\d+)%', line_str)
                            if progress_match:
                                progress = float(progress_match.group(1))
                                # Map to 15-95 range
                                job.progress = 15.0 + (progress * 0.8)
                                if job.progress > last_progress + 5:  # Update every 5%
                                    last_progress = job.progress
                                    self._update_job_in_db(job)
                                    self._add_job_log(job_id, "INFO", f"Progress: {progress}%")

                                    # Send real-time progress updates
                                    if HAS_REALTIME:
                                        asyncio.create_task(manager.send_job_update(
                                            job_id, job.user_id, "progress",
                                            {"progress": job.progress, "stage_progress": progress}
                                        ))
                                        asyncio.create_task(send_job_update_sse(
                                            job.user_id, job_id, "progress",
                                            {"progress": job.progress, "stage_progress": progress}
                                        ))

                            # Log important messages
                            if "error" in line_str.lower() or "failed" in line_str.lower():
                                self._add_job_log(job_id, "WARNING", line_str)

                # Start monitoring output
                monitor_task = asyncio.create_task(monitor_output())

                try:
                    stdout, stderr = await asyncio.wait_for(
                        process.communicate(),
                        timeout=settings.JOB_TIMEOUT
                    )

                    # Cancel monitor task
                    monitor_task.cancel()
                    try:
                        await monitor_task
                    except asyncio.CancelledError:
                        pass

                    if process.returncode == 0:
                        job.status = JobStatus.COMPLETED
                        job.progress = 100.0
                        job.completed_at = datetime.now()
                        job.metadata["output_file_size"] = os.path.getsize(str(output_path))

                        # Parse cost from output if available
                        stdout_str = stdout.decode()
                        cost_match = re.search(r'Total cost: \$([\d.]+)', stdout_str)
                        if cost_match:
                            job.actual_cost = float(cost_match.group(1))
                            job.metadata["actual_cost"] = job.actual_cost

                        self._update_job_in_db(job)
                        self._add_job_log(job_id, "INFO", f"Job completed successfully. Cost: ${job.actual_cost or 'unknown'}")

                        # Assess translation quality
                        try:
                            from ..services.quality_service import quality_service
                            job_logs = await job_manager.get_job_logs(job_id)
                            quality_metrics = await quality_service.assess_translation_quality(
                                job.input_file,
                                job.output_file,
                                job.request.file_type,
                                job_logs
                            )
                            job.quality_metrics = quality_metrics
                            job.metadata["quality_assessed"] = True
                            self._update_job_in_db(job)
                            self._add_job_log(job_id, "INFO", f"Quality assessment completed: {quality_metrics.get('quality_grade', 'unknown')}")
                        except Exception as e:
                            logger.error(f"Failed to assess translation quality: {e}")

                        # Send real-time completion updates
                        if HAS_REALTIME:
                            asyncio.create_task(manager.send_job_update(
                                job_id, job.user_id, "completed",
                                {
                                    "progress": 100.0,
                                    "status": job.status.value,
                                    "cost": job.actual_cost,
                                    "output_file": job.output_file,
                                    "output_size": job.metadata.get("output_file_size")
                                }
                            ))
                            asyncio.create_task(send_notification_sse(
                                job.user_id, "success", "Translation Complete",
                                f"Your {job.request.file_type.upper()} translation has completed successfully",
                                {"job_id": job_id, "cost": job.actual_cost}
                            ))
                            asyncio.create_task(send_job_update_sse(
                                job.user_id, job_id, "completed",
                                {
                                    "progress": 100.0,
                                    "status": job.status.value,
                                    "cost": job.actual_cost,
                                    "output_file": job.output_file,
                                    "output_size": job.metadata.get("output_file_size")
                                }
                            ))

                        logger.info(f"Job {job_id} completed successfully")
                    else:
                        job.status = JobStatus.FAILED
                        job.error = stderr.decode() if stderr else "Unknown error"
                        job.completed_at = datetime.now()
                        self._update_job_in_db(job)
                        self._add_job_log(job_id, "ERROR", f"Job failed: {job.error}")

                        # Send real-time failure updates
                        if HAS_REALTIME:
                            asyncio.create_task(manager.send_job_update(
                                job_id, job.user_id, "failed",
                                {
                                    "progress": job.progress,
                                    "status": job.status.value,
                                    "error": job.error
                                }
                            ))
                            asyncio.create_task(send_notification_sse(
                                job.user_id, "error", "Translation Failed",
                                f"Your {job.request.file_type.upper()} translation has failed",
                                {"job_id": job_id, "error": job.error}
                            ))
                            asyncio.create_task(send_job_update_sse(
                                job.user_id, job_id, "failed",
                                {
                                    "progress": job.progress,
                                    "status": job.status.value,
                                    "error": job.error
                                }
                            ))

                        logger.error(f"Job {job_id} failed: {job.error}")

                except asyncio.TimeoutError:
                    job.status = JobStatus.FAILED
                    job.error = "Job timed out"
                    job.completed_at = datetime.now()
                    self._update_job_in_db(job)
                    self._add_job_log(job_id, "ERROR", "Job timed out")
                    logger.error(f"Job {job_id} timed out")

            except Exception as e:
                job.status = JobStatus.FAILED
                job.error = str(e)
                job.completed_at = datetime.now()
                self._update_job_in_db(job)
                self._add_job_log(job_id, "ERROR", f"Job failed with exception: {str(e)}")
                logger.exception(f"Job {job_id} failed with exception")

            finally:
                # Clean up active job tracking
                if job_id in self.active_jobs:
                    del self.active_jobs[job_id]

                # Clean up input file
                try:
                    if os.path.exists(job.input_file):
                        os.remove(job.input_file)
                        logger.info(f"Cleaned up input file for job {job_id}")
                except Exception as e:
                    logger.warning(f"Failed to clean up input file for job {job_id}: {e}")

    def _update_job_in_db(self, job: Job):
        """Update job in database"""
        try:
            conn = sqlite3.connect(str(self.db_path))
            cursor = conn.cursor()

            cursor.execute("""
                UPDATE jobs SET
                    status = ?,
                    progress = ?,
                    message = ?,
                    error = ?,
                    started_at = ?,
                    completed_at = ?,
                    metadata_json = ?,
                    estimated_cost = ?,
                    actual_cost = ?
                WHERE id = ?
            """, [
                job.status.value,
                job.progress,
                job.message,
                job.error,
                job.started_at.isoformat() if job.started_at else None,
                job.completed_at.isoformat() if job.completed_at else None,
                json.dumps(job.metadata),
                job.metadata.get("estimated_cost"),
                job.metadata.get("actual_cost") or job.actual_cost,
                job.id
            ])

            conn.commit()
            conn.close()
        except Exception as e:
            logger.error(f"Failed to update job in database: {e}")


# Global instance
job_manager = JobManager()
