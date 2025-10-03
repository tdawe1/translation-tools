import re
import json
import logging
from typing import Dict, Any, Optional, List
from pathlib import Path
import subprocess
import sys
from datetime import datetime

logger = logging.getLogger(__name__)

class QualityService:
    """Service for assessing translation quality"""

    def __init__(self):
        self.quality_thresholds = {
            "excellent": 0.95,
            "good": 0.85,
            "fair": 0.70,
            "poor": 0.0
        }

    async def assess_translation_quality(
        self,
        input_file: str,
        output_file: str,
        file_type: str,
        job_logs: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Assess the quality of a translation job"""
        metrics = {}

        try:
            # Basic completion metrics
            metrics["completion_rate"] = self._calculate_completion_rate(job_logs)
            metrics["error_rate"] = self._calculate_error_rate(job_logs)
            metrics["processing_time"] = self._calculate_processing_time(job_logs)

            # File-based metrics
            if file_type == "pptx":
                metrics.update(await self._assess_pptx_quality(input_file, output_file))
            elif file_type == "pdf":
                metrics.update(await self._assess_pdf_quality(input_file, output_file))

            # Cost efficiency
            metrics["cost_efficiency"] = self._calculate_cost_efficiency(job_logs)

            # Overall quality score
            metrics["overall_score"] = self._calculate_overall_score(metrics)
            metrics["quality_grade"] = self._get_quality_grade(metrics["overall_score"])

            return metrics

        except Exception as e:
            logger.error(f"Failed to assess translation quality: {e}")
            return {
                "overall_score": 0.0,
                "quality_grade": "unknown",
                "error": str(e)
            }

    def _calculate_completion_rate(self, job_logs: List[Dict[str, Any]]) -> float:
        """Calculate the completion rate based on job logs"""
        if not job_logs:
            return 0.0

        total_steps = len([log for log in job_logs if log["level"] == "INFO"])
        completed_steps = len([log for log in job_logs if "completed" in log["message"].lower() or "progress: 100%" in log["message"]])

        return completed_steps / max(total_steps, 1)

    def _calculate_error_rate(self, job_logs: List[Dict[str, Any]]) -> float:
        """Calculate the error rate from job logs"""
        if not job_logs:
            return 0.0

        total_logs = len(job_logs)
        error_logs = len([log for log in job_logs if log["level"] == "ERROR" or "error" in log["message"].lower()])

        return error_logs / max(total_logs, 1)

    def _calculate_processing_time(self, job_logs: List[Dict[str, Any]]) -> float:
        """Calculate total processing time in minutes"""
        if not job_logs or len(job_logs) < 2:
            return 0.0

        try:
            start_time = None
            end_time = None

            for log in job_logs:
                if "started" in log["message"].lower() and not start_time:
                    start_time = datetime.fromisoformat(log["timestamp"])
                elif ("completed" in log["message"].lower() or "finished" in log["message"].lower()) and not end_time:
                    end_time = datetime.fromisoformat(log["timestamp"])

            if start_time and end_time:
                return (end_time - start_time).total_seconds() / 60
        except Exception as e:
            logger.error(f"Error calculating processing time: {e}")

        return 0.0

    async def _assess_pptx_quality(self, input_file: str, output_file: str) -> Dict[str, Any]:
        """Assess quality of PPTX translation"""
        metrics = {}

        try:
            # Check if output file exists and has content
            output_path = Path(output_file)
            if not output_path.exists():
                metrics["file_integrity"] = 0.0
                return metrics

            input_size = Path(input_file).stat().st_size
            output_size = output_path.stat().st_size

            # Size ratio (English text is typically shorter)
            metrics["size_ratio"] = output_size / max(input_size, 1)

            # Check for common issues
            issues = []
            if output_size == 0:
                issues.append("Empty output file")
                metrics["file_integrity"] = 0.0
            elif output_size < input_size * 0.3:
                issues.append("Output file significantly smaller")
                metrics["file_integrity"] = 0.5
            else:
                metrics["file_integrity"] = 1.0

            metrics["issues_found"] = issues

        except Exception as e:
            logger.error(f"Error assessing PPTX quality: {e}")
            metrics["file_integrity"] = 0.0
            metrics["assessment_error"] = str(e)

        return metrics

    async def _assess_pdf_quality(self, input_file: str, output_file: str) -> Dict[str, Any]:
        """Assess quality of PDF translation"""
        metrics = {}

        try:
            # Check if output file exists and has content
            output_path = Path(output_file)
            if not output_path.exists():
                metrics["file_integrity"] = 0.0
                return metrics

            input_size = Path(input_file).stat().st_size
            output_size = output_path.stat().st_size

            # Size ratio (can vary for PDFs)
            metrics["size_ratio"] = output_size / max(input_size, 1)

            # Basic integrity check
            if output_size == 0:
                metrics["file_integrity"] = 0.0
                metrics["issues_found"] = ["Empty output file"]
            else:
                metrics["file_integrity"] = 1.0
                metrics["issues_found"] = []

        except Exception as e:
            logger.error(f"Error assessing PDF quality: {e}")
            metrics["file_integrity"] = 0.0
            metrics["assessment_error"] = str(e)

        return metrics

    def _calculate_cost_efficiency(self, job_logs: List[Dict[str, Any]]) -> float:
        """Calculate cost efficiency based on processing time and errors"""
        # This is a simplified calculation
        # In a real implementation, you'd compare actual cost vs expected cost

        processing_time = self._calculate_processing_time(job_logs)
        error_rate = self._calculate_error_rate(job_logs)

        # Efficiency decreases with longer processing times and higher error rates
        time_factor = max(0, 1 - (processing_time / 60))  # Normalize to 60 minutes
        error_factor = max(0, 1 - error_rate)

        return (time_factor + error_factor) / 2

    def _calculate_overall_score(self, metrics: Dict[str, Any]) -> float:
        """Calculate overall quality score"""
        weights = {
            "completion_rate": 0.3,
            "error_rate": 0.25,
            "file_integrity": 0.25,
            "cost_efficiency": 0.2
        }

        score = 0.0
        total_weight = 0.0

        for metric, weight in weights.items():
            if metric in metrics:
                value = metrics[metric]
                if metric == "error_rate":
                    value = 1 - value  # Invert error rate
                score += value * weight
                total_weight += weight

        return score / max(total_weight, 1)

    def _get_quality_grade(self, score: float) -> str:
        """Get quality grade based on score"""
        if score >= self.quality_thresholds["excellent"]:
            return "excellent"
        elif score >= self.quality_thresholds["good"]:
            return "good"
        elif score >= self.quality_thresholds["fair"]:
            return "fair"
        else:
            return "poor"

    def get_quality_recommendations(self, metrics: Dict[str, Any]) -> List[str]:
        """Generate recommendations based on quality metrics"""
        recommendations = []

        if metrics.get("error_rate", 0) > 0.1:
            recommendations.append("High error rate detected. Consider checking input file format and content.")

        if metrics.get("completion_rate", 0) < 0.9:
            recommendations.append("Job completion rate is low. Review job logs for more details.")

        if metrics.get("file_integrity", 1.0) < 0.8:
            recommendations.append("Output file integrity issues detected. Verify the translated file.")

        if metrics.get("cost_efficiency", 0) < 0.5:
            recommendations.append("Cost efficiency is low. Consider optimizing translation settings.")

        return recommendations

# Global instance
quality_service = QualityService()