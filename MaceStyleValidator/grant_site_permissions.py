"""Grant per-site SharePoint permissions for the MaceStyle app using Sites.Selected.

This script grants the MaceStyle app registration 'write' access to a specific
SharePoint site via the Microsoft Graph API. This is a one-time setup step
required when using Sites.Selected instead of tenant-wide Sites.ReadWrite.All.

Prerequisites:
    - The app registration must have Sites.Selected admin-consented in Entra ID
    - An admin must run this script with credentials that have Sites.FullControl.All
      (or use the app's own credentials if it temporarily has FullControl for bootstrap)

Usage:
    # Set environment variables (same as used by the Function App)
    export SHAREPOINT_TENANT_ID="..."
    export SHAREPOINT_CLIENT_ID="..."
    export SHAREPOINT_CLIENT_SECRET="..."
    export SHAREPOINT_SITE_URL="https://tenant.sharepoint.com/sites/StyleValidation"

    # Optionally, use a separate admin credential for granting:
    export ADMIN_CLIENT_ID="..."
    export ADMIN_CLIENT_SECRET="..."

    python3 grant_site_permissions.py
"""

import os
import sys
import msal
import requests


def get_admin_token():
    """Get a Graph API token with permissions to grant site access.

    Uses ADMIN_CLIENT_ID/ADMIN_CLIENT_SECRET if set (recommended — an admin
    app with Sites.FullControl.All). Falls back to the MaceStyle app's own
    credentials.
    """
    tenant_id = os.environ.get("SHAREPOINT_TENANT_ID")
    client_id = os.environ.get("ADMIN_CLIENT_ID") or os.environ.get("SHAREPOINT_CLIENT_ID")
    client_secret = os.environ.get("ADMIN_CLIENT_SECRET") or os.environ.get("SHAREPOINT_CLIENT_SECRET")

    if not all([tenant_id, client_id, client_secret]):
        print("Error: Set SHAREPOINT_TENANT_ID and either ADMIN_CLIENT_ID/ADMIN_CLIENT_SECRET")
        print("       or SHAREPOINT_CLIENT_ID/SHAREPOINT_CLIENT_SECRET.")
        sys.exit(1)

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" in result:
        return result["access_token"]

    print(f"Error acquiring token: {result.get('error_description', result)}")
    sys.exit(1)


def get_site_id(token):
    """Resolve the SharePoint site URL to a Graph API site ID."""
    site_url = os.environ.get("SHAREPOINT_SITE_URL")
    if not site_url:
        print("Error: SHAREPOINT_SITE_URL not set.")
        sys.exit(1)

    parts = site_url.replace("https://", "").split("/")
    hostname = parts[0]
    site_path = "/" + "/".join(parts[1:]) if len(parts) > 1 else ""

    url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()["id"]


def grant_site_permission(token, site_id):
    """Grant the MaceStyle app 'write' access to the specified site."""
    app_client_id = os.environ.get("SHAREPOINT_CLIENT_ID")
    if not app_client_id:
        print("Error: SHAREPOINT_CLIENT_ID not set.")
        sys.exit(1)

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/permissions"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    body = {
        "roles": ["write"],
        "grantedToIdentities": [
            {
                "application": {
                    "id": app_client_id,
                    "displayName": "MaceStyleValidator-App",
                }
            }
        ],
    }

    resp = requests.post(url, headers=headers, json=body)

    if resp.status_code == 201:
        perm = resp.json()
        print(f"Permission granted successfully.")
        print(f"  Permission ID: {perm.get('id')}")
        print(f"  Roles: {perm.get('roles')}")
        print(f"  App: {app_client_id}")
        print(f"  Site: {site_id}")
    elif resp.status_code == 409:
        print("Permission already exists for this app on this site.")
    else:
        print(f"Error {resp.status_code}: {resp.text}")
        sys.exit(1)


def list_site_permissions(token, site_id):
    """List current permissions on the site."""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/permissions"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()

    perms = resp.json().get("value", [])
    print(f"\nCurrent site permissions ({len(perms)}):")
    for p in perms:
        identities = p.get("grantedToIdentitiesV2", p.get("grantedToIdentities", []))
        for identity in identities:
            app_info = identity.get("application", {})
            print(f"  - {app_info.get('displayName', 'Unknown')} ({app_info.get('id', '?')}): {p.get('roles')}")


if __name__ == "__main__":
    print("MaceStyle — Grant Per-Site SharePoint Permissions")
    print("=" * 50)

    token = get_admin_token()
    print("Authenticated successfully.")

    site_url = os.environ.get("SHAREPOINT_SITE_URL", "")
    print(f"Target site: {site_url}")

    site_id = get_site_id(token)
    print(f"Site ID: {site_id}")

    grant_site_permission(token, site_id)
    list_site_permissions(token, site_id)
