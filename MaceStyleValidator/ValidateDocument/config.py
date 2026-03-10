"""Configuration and authentication for MaceStyle Validator"""
import os
import msal

# Claude AI configuration
CLAUDE_MODEL = "claude-haiku-4-5-20251001"
CLAUDE_MAX_TOKENS = 8192
CLAUDE_TEMPERATURE = 0.3

# SharePoint list IDs - override via env vars for different tenants
DOC_LIBRARY_LIST_ID = os.environ.get(
    "SHAREPOINT_DOC_LIBRARY_ID", "800c67b1-816d-43f6-ac7d-d21bca8d140f"
)
VALIDATION_RESULTS_LIST_ID = os.environ.get(
    "SHAREPOINT_VALIDATION_RESULTS_ID", "d4f4cc72-7f68-4009-a1eb-e86d9e67a4dd"
)


def get_graph_token():
    """Get Microsoft Graph API access token using MSAL"""
    tenant_id = os.environ.get("SHAREPOINT_TENANT_ID")
    client_id = os.environ.get("SHAREPOINT_CLIENT_ID")
    client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET")

    if not all([tenant_id, client_id, client_secret]):
        raise ValueError(
            "Missing SharePoint credentials. Set SHAREPOINT_TENANT_ID, "
            "SHAREPOINT_CLIENT_ID, and SHAREPOINT_CLIENT_SECRET."
        )

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"Failed to acquire token: {result.get('error_description', result)}")


def get_site_info():
    """Get SharePoint site information from environment"""
    site_url = os.environ.get("SHAREPOINT_SITE_URL")
    if not site_url:
        raise ValueError("SHAREPOINT_SITE_URL environment variable not set")

    parts = site_url.replace("https://", "").split("/")
    hostname = parts[0]
    site_path = "/" + "/".join(parts[1:]) if len(parts) > 1 else ""

    return {"hostname": hostname, "site_path": site_path, "full_url": site_url}
