"""Configuration and authentication for MaceStyle Validator"""
import os
import msal
import requests

# Claude AI configuration
# Toggle AI validation via the ENABLE_CLAUDE_AI app setting ("true"/"false"). Default off.
ENABLE_CLAUDE_AI = os.environ.get("ENABLE_CLAUDE_AI", "false").lower() == "true"
# AI_PROVIDER selects which model backend serves the AI rules:
#   "anthropic"    - direct Anthropic API (ANTHROPIC_API_KEY) - current setup
#   "foundry"      - Claude in Microsoft Foundry (FOUNDRY_RESOURCE + FOUNDRY_API_KEY) - Mace target state.
#                    CLAUDE_MODEL must then match the Foundry deployment name (e.g. "claude-haiku-4-5").
#   "azure_openai" - Azure OpenAI GPT model on an Azure AI Foundry resource (AZURE_OPENAI_ENDPOINT +
#                    AZURE_OPENAI_API_KEY). Easiest for Mace to provision (Microsoft-native, no
#                    Anthropic Marketplace step). CLAUDE_MODEL must then be the GPT deployment name
#                    (e.g. "gpt-5-mini").
AI_PROVIDER = os.environ.get("AI_PROVIDER", "anthropic")
CLAUDE_MODEL = os.environ.get("CLAUDE_MODEL", "claude-haiku-4-5-20251001")

# Azure OpenAI settings (only used when AI_PROVIDER=azure_openai). The key defaults to the
# shared Foundry resource key since GPT and Claude can co-live on one AIServices resource.
AZURE_OPENAI_ENDPOINT = os.environ.get("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_API_KEY = os.environ.get("AZURE_OPENAI_API_KEY") or os.environ.get("FOUNDRY_API_KEY")
AZURE_OPENAI_API_VERSION = os.environ.get("AZURE_OPENAI_API_VERSION", "2024-12-01-preview")
# GPT-5 reasoning models spend tokens on hidden reasoning before output, so the completion cap
# needs headroom above the visible answer. "low" keeps latency/cost sane for an editing task.
AZURE_OPENAI_MAX_COMPLETION_TOKENS = int(os.environ.get("AZURE_OPENAI_MAX_COMPLETION_TOKENS", "16000"))
AZURE_OPENAI_REASONING_EFFORT = os.environ.get("AZURE_OPENAI_REASONING_EFFORT", "low")

# SharePoint write ownership.
# When False (default), the Power Automate flow owns ALL SharePoint writes
# (Validation Results item, report upload, ReportLink, doc properties/status)
# and the function only validates + returns data. Set True to let the function
# write directly (legacy behaviour) — but only if the flow's write actions are
# removed, otherwise you get duplicate list items / report files.
ENABLE_FUNCTION_SHAREPOINT_WRITES = False
CLAUDE_MAX_TOKENS = 8192
CLAUDE_TEMPERATURE = 0.3

# SharePoint list IDs - must be set via env vars
DOC_LIBRARY_LIST_ID = os.environ.get("SHAREPOINT_DOC_LIBRARY_ID")
VALIDATION_RESULTS_LIST_ID = os.environ.get("SHAREPOINT_VALIDATION_RESULTS_ID")


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


def get_style_rules_token():
    """Get Graph token for the style rules source (STYLE_RULES_* vars, falls back to SHAREPOINT_*)"""
    tenant_id = os.environ.get("STYLE_RULES_TENANT_ID") or os.environ.get("SHAREPOINT_TENANT_ID")
    client_id = os.environ.get("STYLE_RULES_CLIENT_ID") or os.environ.get("SHAREPOINT_CLIENT_ID")
    client_secret = os.environ.get("STYLE_RULES_CLIENT_SECRET") or os.environ.get("SHAREPOINT_CLIENT_SECRET")

    if not all([tenant_id, client_id, client_secret]):
        raise ValueError("Missing style rules credentials. Set STYLE_RULES_TENANT_ID/CLIENT_ID/CLIENT_SECRET.")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"Failed to acquire style rules token: {result.get('error_description', result)}")


def get_style_rules_site_info():
    """Get site info for the style rules source (STYLE_RULES_SITE_URL falls back to SHAREPOINT_SITE_URL)"""
    site_url = os.environ.get("STYLE_RULES_SITE_URL") or os.environ.get("SHAREPOINT_SITE_URL")
    if not site_url:
        raise ValueError("STYLE_RULES_SITE_URL environment variable not set")

    parts = site_url.replace("https://", "").split("/")
    hostname = parts[0]
    site_path = "/" + "/".join(parts[1:]) if len(parts) > 1 else ""

    return {"hostname": hostname, "site_path": site_path, "full_url": site_url}


def get_site_id(token=None):
    """Get SharePoint site ID from Graph API"""
    if token is None:
        token = get_graph_token()
    site = get_site_info()
    url = f"https://graph.microsoft.com/v1.0/sites/{site['hostname']}:{site['site_path']}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()["id"]
