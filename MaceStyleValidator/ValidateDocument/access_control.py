"""Access control for MaceStyle Validator — SOC 2 CC6.1

Validates incoming requests against an API key or Azure AD bearer token.
Also provides a decorator for function-level access control.

Configuration (environment variables):
  MACESTYLE_API_KEY        — Static API key for service-to-service auth (e.g. Power Automate)
  MACESTYLE_AUTH_MODE      — "api_key" (default), "azure_ad", or "none" (dev only)
  MACESTYLE_ALLOWED_APPS   — Comma-separated Azure AD app IDs allowed to call this function
"""

import os
import json
import logging
import functools
import datetime
from typing import Optional

import azure.functions as func

logger = logging.getLogger("macestyle.access_control")

AUTH_MODE = os.environ.get("MACESTYLE_AUTH_MODE", "api_key")
API_KEY = os.environ.get("MACESTYLE_API_KEY", "")
ALLOWED_APPS = [
    app_id.strip()
    for app_id in os.environ.get("MACESTYLE_ALLOWED_APPS", "").split(",")
    if app_id.strip()
]


def validate_api_key(req: func.HttpRequest) -> Optional[str]:
    """Check X-Api-Key or Authorization: Bearer <key> header against MACESTYLE_API_KEY."""
    key = req.headers.get("x-api-key") or ""
    if not key:
        auth = req.headers.get("Authorization", "")
        if auth.startswith("Bearer "):
            key = auth[7:]

    if not API_KEY:
        logger.warning("MACESTYLE_API_KEY not set — access control is open")
        return None

    if key != API_KEY:
        return "Invalid or missing API key"

    return None


def validate_azure_ad(req: func.HttpRequest) -> Optional[str]:
    """Validate the Azure AD JWT from the X-MS-CLIENT-PRINCIPAL header.

    Azure Functions with built-in auth (EasyAuth) inject this header
    automatically. It contains a base64-encoded JSON with the caller's
    identity claims.
    """
    import base64

    principal_header = req.headers.get("X-MS-CLIENT-PRINCIPAL", "")
    if not principal_header:
        return "Missing Azure AD identity (X-MS-CLIENT-PRINCIPAL header)"

    try:
        claims = json.loads(base64.b64decode(principal_header))
    except Exception:
        return "Malformed Azure AD identity header"

    # Extract app ID (azp or appid claim)
    caller_app_id = None
    for claim in claims.get("claims", []):
        if claim.get("typ") in ("azp", "appid"):
            caller_app_id = claim.get("val")
            break

    if ALLOWED_APPS and caller_app_id not in ALLOWED_APPS:
        return f"App ID {caller_app_id} is not in the allowed list"

    return None


def check_access(req: func.HttpRequest) -> Optional[func.HttpResponse]:
    """Check access control based on configured auth mode.

    Returns None if access is granted, or an HttpResponse with 401/403 if denied.
    """
    if AUTH_MODE == "none":
        logger.debug("Auth mode is 'none' — skipping access control")
        return None

    if AUTH_MODE == "azure_ad":
        error = validate_azure_ad(req)
    else:
        error = validate_api_key(req)

    if error:
        logger.warning(f"Access denied: {error} | IP: {_get_client_ip(req)}")
        return func.HttpResponse(
            json.dumps({"error": "Unauthorised", "detail": error}),
            mimetype="application/json",
            status_code=401,
        )

    return None


def require_access(fn):
    """Decorator that enforces access control before the function runs."""

    @functools.wraps(fn)
    def wrapper(req: func.HttpRequest, *args, **kwargs) -> func.HttpResponse:
        denied = check_access(req)
        if denied:
            return denied
        return fn(req, *args, **kwargs)

    return wrapper


def get_caller_identity(req: func.HttpRequest) -> dict:
    """Extract caller identity from request headers for audit logging."""
    import base64

    identity = {
        "ip": _get_client_ip(req),
        "user_agent": req.headers.get("User-Agent", "unknown"),
        "auth_mode": AUTH_MODE,
    }

    # Try Azure AD principal
    principal = req.headers.get("X-MS-CLIENT-PRINCIPAL", "")
    if principal:
        try:
            claims = json.loads(base64.b64decode(principal))
            for claim in claims.get("claims", []):
                if claim.get("typ") == "name":
                    identity["name"] = claim.get("val")
                elif claim.get("typ") in ("azp", "appid"):
                    identity["app_id"] = claim.get("val")
                elif claim.get("typ") == "preferred_username":
                    identity["email"] = claim.get("val")
        except Exception:
            pass

    # Power Automate often sends a workflow ID header
    workflow_id = req.headers.get("X-MS-Workflow-Run-Id", "")
    if workflow_id:
        identity["power_automate_run_id"] = workflow_id

    return identity


def _get_client_ip(req: func.HttpRequest) -> str:
    """Extract client IP, respecting X-Forwarded-For."""
    forwarded = req.headers.get("X-Forwarded-For", "")
    if forwarded:
        return forwarded.split(",")[0].strip()
    return req.headers.get("X-Client-IP", "unknown")
