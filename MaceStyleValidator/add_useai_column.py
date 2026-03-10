"""
Add UseAI column to SharePoint Style Rules list
"""
import os
import msal
import requests
import json

# Configuration
TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID")
CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET")
SITE_URL = os.environ.get("SHAREPOINT_SITE_URL")

def get_token():
    """Get Microsoft Graph API access token"""
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    scope = ["https://graph.microsoft.com/.default"]

    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )

    result = app.acquire_token_for_client(scopes=scope)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Failed to acquire token: {result}")

def get_site_id(token):
    """Get SharePoint site ID"""
    parts = SITE_URL.replace("https://", "").split("/")
    hostname = parts[0]
    site_path = "/" + "/".join(parts[1:])

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
    response = requests.get(site_url, headers=headers)
    response.raise_for_status()
    return response.json()["id"]

def add_column(token, site_id):
    """Add UseAI boolean column to Style Rules list"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    columns_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/columns"

    column_data = {
        "name": "UseAI",
        "displayName": "UseAI",
        "description": "Whether to use AI (Claude) for this rule",
        "boolean": {}
    }

    response = requests.post(columns_url, headers=headers, json=column_data)
    response.raise_for_status()
    return response.json()

def main():
    print("Adding UseAI column to SharePoint Style Rules list\n")

    # Get access token
    print("Getting access token...")
    token = get_token()

    # Get site ID
    print("Getting site ID...")
    site_id = get_site_id(token)
    print(f"Site ID: {site_id}\n")

    # Add column
    print("Adding UseAI column...")
    result = add_column(token, site_id)
    print(f"✓ Success: Column added")
    print(f"Column ID: {result.get('id')}\n")

if __name__ == "__main__":
    main()
