"""
Setup script to populate SharePoint Style Rules list
Run this locally to add default validation rules to your SharePoint site
"""

import os
import msal
import requests
import json

# Configuration
TENANT_ID = "2ab0866e-23d6-4688-be97-ce9f447135d8"
CLIENT_ID = "c7859dae-6997-448f-9530-7166fe857e75"
CLIENT_SECRET = "DlD8Q~_NNgnpnVxKWsZTiz53DuNYrfrAjqkCDaP1"
SITE_URL = "https://0rxf2.sharepoint.com/sites/StyleValidation"

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
    site_path = "/" + "/".join(parts[1:]) if len(parts) > 1 else ""

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    graph_site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
    response = requests.get(graph_site_url, headers=headers)
    response.raise_for_status()

    return response.json()["id"]

def create_style_rules(token, site_id):
    """Create default validation rules in Style Rules list"""

    default_rules = [
        {
            "Title": "Heading 1 Font Check",
            "RuleType": "Font",
            "DocumentType": "Word",
            "CheckValue": "Heading1Font",
            "ExpectedValue": "Arial",
            "AutoFix": True,
            "Priority": 1
        },
        {
            "Title": "Heading 1 Color Check",
            "RuleType": "Color",
            "DocumentType": "Word",
            "CheckValue": "Heading1Color",
            "ExpectedValue": "0,51,153",
            "AutoFix": True,
            "Priority": 2
        },
        {
            "Title": "Body Font Check",
            "RuleType": "Font",
            "DocumentType": "Word",
            "CheckValue": "BodyFont",
            "ExpectedValue": "Calibri",
            "AutoFix": True,
            "Priority": 3
        }
    ]

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/items"

    created_rules = []
    for rule in default_rules:
        item_data = {"fields": rule}
        response = requests.post(list_url, headers=headers, json=item_data)

        if response.status_code == 201:
            created_rules.append(rule["Title"])
            print(f"✅ Created rule: {rule['Title']}")
        else:
            print(f"❌ Failed to create rule: {rule['Title']}")
            print(f"   Error: {response.text}")

    return created_rules

def main():
    print("🚀 MaceStyle Validator - SharePoint Setup")
    print("=" * 50)

    try:
        print("\n1. Getting access token...")
        token = get_token()
        print("   ✅ Token acquired")

        print("\n2. Getting site ID...")
        site_id = get_site_id(token)
        print(f"   ✅ Site ID: {site_id}")

        print("\n3. Creating validation rules...")
        created_rules = create_style_rules(token, site_id)

        print(f"\n✨ Setup complete! Created {len(created_rules)} validation rules:")
        for rule in created_rules:
            print(f"   - {rule}")

        print("\n🎯 Next steps:")
        print("   1. Check your Style Rules list in SharePoint")
        print("   2. Upload a Word document to test")
        print("   3. Set up Power Automate to trigger validation automatically")

    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
