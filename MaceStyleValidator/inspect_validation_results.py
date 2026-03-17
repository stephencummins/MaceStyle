"""
Inspect the Validation Results list schema
"""
import requests

from ValidateDocument.config import get_graph_token, get_site_id

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
    token = get_graph_token()

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
