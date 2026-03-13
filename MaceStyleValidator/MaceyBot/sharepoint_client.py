"""SharePoint integration — write site creation requests to the Site Creation list."""
import json
import logging
import os

import msal
import requests


def _get_graph_token() -> str:
    """Get Microsoft Graph API access token. Reuses MaceStyle's SHAREPOINT_* credentials."""
    tenant_id = os.environ.get("SP_TENANT_ID") or os.environ.get("SHAREPOINT_TENANT_ID")
    client_id = os.environ.get("SP_CLIENT_ID") or os.environ.get("SHAREPOINT_CLIENT_ID")
    client_secret = os.environ.get("SP_CLIENT_SECRET") or os.environ.get("SHAREPOINT_CLIENT_SECRET")

    if not all([tenant_id, client_id, client_secret]):
        raise ValueError(
            "Missing SharePoint credentials. "
            "Set SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, and SHAREPOINT_CLIENT_SECRET "
            "(or SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET)."
        )

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"Failed to acquire token: {result.get('error_description', result)}")


def _get_site_id(token: str) -> str:
    """Resolve the SharePoint site ID from SP_SITE_URL."""
    site_url = os.environ.get("SP_SITE_URL")
    if not site_url:
        raise ValueError("SP_SITE_URL environment variable not set")

    parts = site_url.replace("https://", "").split("/")
    hostname = parts[0]
    site_path = "/" + "/".join(parts[1:]) if len(parts) > 1 else ""

    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}",
        headers=headers,
    )
    resp.raise_for_status()
    return resp.json()["id"]


def _get_list_id(token: str, site_id: str) -> str:
    """Get the list ID for the Site Creation list."""
    list_name = os.environ.get("SP_LIST_NAME", "Site Creation")
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    # Try by display name first
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists",
        headers=headers,
        params={"$filter": f"displayName eq '{list_name}'"},
    )
    resp.raise_for_status()
    lists = resp.json().get("value", [])

    if lists:
        return lists[0]["id"]

    raise Exception(f"SharePoint list '{list_name}' not found on site {site_id}")


def list_columns() -> list[dict]:
    """Debug helper — fetch column names from the Site Creation list."""
    try:
        token = _get_graph_token()
        site_id = _get_site_id(token)
        list_id = _get_list_id(token, site_id)
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns",
            headers=headers,
        )
        resp.raise_for_status()
        cols = resp.json().get("value", [])
        result = [{"name": c["name"], "displayName": c.get("displayName", "")} for c in cols]
        logging.info(f"[MaceyBot SP] List columns: {json.dumps(result)}")
        return result
    except Exception as e:
        logging.error(f"[MaceyBot SP] Failed to list columns: {e}", exc_info=True)
        return []


def submit_to_sharepoint(params: dict) -> str:
    """Write a site creation request to the SharePoint list.

    Returns a JSON string describing success or failure (passed back to Claude as tool result).
    """
    try:
        token = _get_graph_token()
        site_id = _get_site_id(token)
        list_id = _get_list_id(token, site_id)

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        item_data = {
            "fields": {
                "Title": params["projectName"],
                "Sitedescription": params["projectDescription"],
                "Visibility": params["siteVisibility"],
                "SiteOwner": params["ownerEmail"],
                "Notes": params.get("additionalNotes", ""),
            }
        }

        logging.info(f"[MaceyBot SP] Creating list item: {json.dumps(item_data)}")

        resp = requests.post(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items",
            headers=headers,
            json=item_data,
        )
        if not resp.ok:
            logging.error(f"[MaceyBot SP] Graph API error {resp.status_code}: {resp.text}")
        resp.raise_for_status()

        item = resp.json()
        item_id = item.get("id")
        logging.info(f"[MaceyBot SP] Created item ID: {item_id}")

        return json.dumps({
            "status": "success",
            "message": f"Site creation request submitted successfully (ID: {item_id}). "
                       "The approval workflow has been triggered.",
            "item_id": item_id,
        })

    except Exception as e:
        logging.error(f"[MaceyBot SP] Failed to submit: {e}", exc_info=True)
        return json.dumps({
            "status": "error",
            "message": f"Failed to submit the request: {str(e)}. "
                       "Please try again or contact IT support.",
        })
