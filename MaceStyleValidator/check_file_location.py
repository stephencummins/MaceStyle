"""
Check actual file locations in SharePoint via Graph API
"""
import os
import msal
import requests
import json

# Configuration - set via environment variables
TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID")
CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET")
SITE_URL = os.environ.get("SHAREPOINT_SITE_URL")

def get_token():
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

def main():
    print("🔍 Checking SharePoint file locations\n")

    token = get_token()

    # Get site ID
    parts = SITE_URL.replace("https://", "").split("/")
    hostname = parts[0]
    site_path = "/" + "/".join(parts[1:])

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    graph_site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
    site_response = requests.get(graph_site_url, headers=headers)
    site_response.raise_for_status()
    site_id = site_response.json()["id"]

    print(f"Site ID: {site_id}\n")

    # List files in default drive (Shared Documents)
    files_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
    files_response = requests.get(files_url, headers=headers)
    files_response.raise_for_status()
    files_data = files_response.json()

    print("Files in default drive (Shared Documents):")
    print("=" * 60)

    for item in files_data.get("value", []):
        if item.get("file"):
            name = item.get("name")
            parent_path = item.get("parentReference", {}).get("path", "")
            web_url = item.get("webUrl")

            # Extract relative path from parent reference
            if "root:" in parent_path:
                relative_path = parent_path.split("root:")[1]
            else:
                relative_path = ""

            full_path = f"{relative_path}/{name}" if relative_path else f"/{name}"

            print(f"\nFile: {name}")
            print(f"  Graph API path: {full_path}")
            print(f"  Web URL: {web_url}")

    print("\n" + "=" * 60)
    print("\nNote: When using Graph API with /drive/root, the 'Shared Documents'")
    print("library is treated as the root, so files are accessed at '/' not '/Shared Documents/'")

if __name__ == "__main__":
    main()
