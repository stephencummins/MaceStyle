"""Structured monitoring for MaceStyle Validator — SOC 2 CC7.2

Provides:
  - Structured audit logging (JSON-format events to Application Insights)
  - Request tracking with correlation IDs
  - Operation metrics (duration, token usage, costs)
  - Health check endpoint data
  - Alert-worthy condition detection
"""

import os
import json
import time
import uuid
import logging
import datetime
import functools
from typing import Optional
from contextlib import contextmanager

logger = logging.getLogger("macestyle.monitoring")


class ValidationMetrics:
    """Tracks metrics for a single validation request."""

    def __init__(self, request_id: str, filename: str, caller: dict):
        self.request_id = request_id
        self.filename = filename
        self.caller = caller
        self.started_at = datetime.datetime.now(datetime.timezone.utc)
        self.ended_at: Optional[datetime.datetime] = None
        self.file_type: str = ""
        self.file_size_bytes: int = 0
        self.rules_loaded: int = 0
        self.ai_rules_count: int = 0
        self.claude_calls: int = 0
        self.claude_input_tokens: int = 0
        self.claude_output_tokens: int = 0
        self.issues_found: int = 0
        self.fixes_applied: int = 0
        self.status: str = "in_progress"
        self.error: Optional[str] = None
        self.sharepoint_calls: int = 0
        self.report_uploaded: bool = False
        self._timings: dict = {}
        self._current_phase: Optional[str] = None
        self._phase_start: Optional[float] = None

    def start_phase(self, phase: str):
        """Begin timing a named phase (e.g. 'claude_api', 'sharepoint_upload')."""
        self._current_phase = phase
        self._phase_start = time.monotonic()

    def end_phase(self):
        """End the current phase and record its duration."""
        if self._current_phase and self._phase_start:
            elapsed = time.monotonic() - self._phase_start
            self._timings[self._current_phase] = round(elapsed * 1000)  # ms
            self._current_phase = None
            self._phase_start = None

    def record_claude_usage(self, input_tokens: int, output_tokens: int):
        """Record token usage from a Claude API call."""
        self.claude_calls += 1
        self.claude_input_tokens += input_tokens
        self.claude_output_tokens += output_tokens

    def complete(self, status: str, issues: int, fixes: int):
        """Mark the validation as complete."""
        self.ended_at = datetime.datetime.now(datetime.timezone.utc)
        self.status = status
        self.issues_found = issues
        self.fixes_applied = fixes

    def fail(self, error: str):
        """Mark the validation as failed."""
        self.ended_at = datetime.datetime.now(datetime.timezone.utc)
        self.status = "error"
        self.error = error

    @property
    def duration_ms(self) -> int:
        end = self.ended_at or datetime.datetime.now(datetime.timezone.utc)
        return int((end - self.started_at).total_seconds() * 1000)

    @property
    def estimated_cost_usd(self) -> float:
        """Estimate Claude API cost (Haiku 4.5 pricing)."""
        input_cost = (self.claude_input_tokens / 1_000_000) * 0.80
        output_cost = (self.claude_output_tokens / 1_000_000) * 4.00
        return round(input_cost + output_cost, 6)

    def to_audit_entry(self) -> dict:
        """Produce a structured audit log entry."""
        return {
            "event_type": "validation_complete",
            "request_id": self.request_id,
            "timestamp": self.started_at.isoformat(),
            "completed_at": self.ended_at.isoformat() if self.ended_at else None,
            "duration_ms": self.duration_ms,
            "caller": self.caller,
            "document": {
                "filename": self.filename,
                "file_type": self.file_type,
                "file_size_bytes": self.file_size_bytes,
            },
            "validation": {
                "status": self.status,
                "rules_loaded": self.rules_loaded,
                "ai_rules_count": self.ai_rules_count,
                "issues_found": self.issues_found,
                "fixes_applied": self.fixes_applied,
                "report_uploaded": self.report_uploaded,
            },
            "ai_usage": {
                "claude_calls": self.claude_calls,
                "input_tokens": self.claude_input_tokens,
                "output_tokens": self.claude_output_tokens,
                "estimated_cost_usd": self.estimated_cost_usd,
            },
            "performance": {
                "total_ms": self.duration_ms,
                "phases_ms": self._timings,
                "sharepoint_calls": self.sharepoint_calls,
            },
            "error": self.error,
        }


def generate_request_id() -> str:
    """Generate a unique request correlation ID."""
    return f"msv-{uuid.uuid4().hex[:12]}"


def emit_audit_event(event: dict):
    """Emit a structured audit event.

    Events are logged as JSON to Application Insights via the standard
    Python logging integration. Azure Functions automatically ships
    these to Application Insights when the APPINSIGHTS_INSTRUMENTATIONKEY
    or APPLICATIONINSIGHTS_CONNECTION_STRING env var is set.
    """
    logger.info(f"AUDIT: {json.dumps(event, default=str)}")


def emit_alert(severity: str, message: str, context: dict):
    """Emit an alert-worthy event.

    Severity levels: INFO, WARNING, CRITICAL.
    """
    event = {
        "event_type": "alert",
        "severity": severity,
        "message": message,
        "context": context,
        "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    }
    if severity == "CRITICAL":
        logger.critical(f"ALERT: {json.dumps(event, default=str)}")
    elif severity == "WARNING":
        logger.warning(f"ALERT: {json.dumps(event, default=str)}")
    else:
        logger.info(f"ALERT: {json.dumps(event, default=str)}")


def get_health_status() -> dict:
    """Return system health status for the health check endpoint."""
    health = {
        "status": "healthy",
        "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "version": "5.1.0-governed",
        "checks": {},
    }

    # Check required env vars
    required_vars = [
        "SHAREPOINT_TENANT_ID",
        "SHAREPOINT_CLIENT_ID",
        "SHAREPOINT_CLIENT_SECRET",
        "SHAREPOINT_SITE_URL",
    ]
    missing = [v for v in required_vars if not os.environ.get(v)]
    health["checks"]["environment"] = {
        "status": "healthy" if not missing else "unhealthy",
        "missing_vars": missing,
    }

    # Check Claude API key
    has_claude = bool(os.environ.get("ANTHROPIC_API_KEY"))
    health["checks"]["claude_api"] = {
        "status": "healthy" if has_claude else "degraded",
        "detail": "API key configured" if has_claude else "ANTHROPIC_API_KEY not set — AI validation disabled",
    }

    # Check access control
    auth_mode = os.environ.get("MACESTYLE_AUTH_MODE", "api_key")
    has_api_key = bool(os.environ.get("MACESTYLE_API_KEY"))
    health["checks"]["access_control"] = {
        "status": "healthy" if (auth_mode != "api_key" or has_api_key) else "warning",
        "auth_mode": auth_mode,
        "detail": "API key configured" if has_api_key else "No API key set — access control is open" if auth_mode == "api_key" else f"Using {auth_mode} auth",
    }

    # Overall status
    statuses = [c["status"] for c in health["checks"].values()]
    if "unhealthy" in statuses:
        health["status"] = "unhealthy"
    elif "degraded" in statuses or "warning" in statuses:
        health["status"] = "degraded"

    return health


@contextmanager
def track_phase(metrics: ValidationMetrics, phase: str):
    """Context manager to time a named phase."""
    metrics.start_phase(phase)
    try:
        yield
    finally:
        metrics.end_phase()
