"""
Inspect the Validation Results list schema
"""
import os
import msal
import requests
import json

# Configuration
TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID", "2ab0866e-23d6-4688-be97-ce9f447135d8")
CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID", "c7859dae-6997-448f-9530-7166fe857e75")
CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET", "DlD8Q~_NNgnpnVxKWsZTiz53DuNYrfrAjqkCDaP1")
SITE_URL = os.environ.get("SHAREPOINT_SITE_URL", "https://0rxf2.sharepoint.com/sites/StyleValidation")

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

def get_list_info(token, site_id, list_name):
    """Get list information and schema"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    # Get list info
    list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_name}"
    response = requests.get(list_url, headers=headers)
    response.raise_for_status()
    list_info = response.json()

    # Get columns
    columns_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_name}/columns"
    response = requests.get(columns_url, headers=headers)
    response.raise_for_status()
    columns = response.json()["value"]

    return list_info, columns

def main():
    print("🔍 Inspecting Validation Results list\n")

    # Get access token
    print("🔑 Getting access token...")
    token = get_token()

    # Get site ID
    print("🌐 Getting site ID...")
    site_id = get_site_id(token)
    print(f"   Site ID: {site_id}\n")

    # Get Validation Results list info
    print("📋 Getting Validation Results list schema...")
    try:
        list_info, columns = get_list_info(token, site_id, "Validation Results")

        print(f"\n✓ List Name: {list_info.get('displayName')}")
        print(f"✓ List ID: {list_info.get('id')}")
        print(f"✓ Description: {list_info.get('description', 'N/A')}")

        print(f"\n📊 Columns ({len(columns)} total):")
        print("=" * 80)

        for col in columns:
            col_name = col.get('displayName', col.get('name'))
            col_type = list(col.keys())[-1]  # Last key is usually the type
            required = col.get('required', False)

            # Skip system columns for cleaner output
            if col_name not in ['Content Type', 'Item Child Count', 'Folder Child Count',
                                'Compliance Asset Id', 'App Created By', 'App Modified By']:
                print(f"  • {col_name}")
                print(f"    Type: {col_type}")
                print(f"    Required: {required}")
                print(f"    Internal Name: {col.get('name')}")
                print()

        print("=" * 80)

    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 404:
            print("❌ 'Validation Results' list not found!")
            print("\nLet me check what lists exist...")

            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json"
            }
            lists_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
            response = requests.get(lists_url, headers=headers)
            lists = response.json()["value"]

            print(f"\n📚 Available lists ({len(lists)}):")
            for lst in lists:
                print(f"  • {lst.get('displayName')} (name: {lst.get('name')})")
        else:
            raise

if __name__ == "__main__":
    main()
