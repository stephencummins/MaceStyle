"""
Add UseAI column to SharePoint Style Rules list
"""
import requests

from ValidateDocument.config import get_graph_token, get_site_id

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
    token = get_graph_token()

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
