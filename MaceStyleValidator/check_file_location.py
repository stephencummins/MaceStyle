"""
Check actual file locations in SharePoint via Graph API
"""
import requests

from ValidateDocument.config import get_graph_token, get_site_info

def main():
    print("🔍 Checking SharePoint file locations\n")

    token = get_graph_token()

    # Get site ID
    site_info = get_site_info()
    hostname = site_info["hostname"]
    site_path = site_info["site_path"]

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
