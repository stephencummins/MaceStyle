"""
Update existing SharePoint Style Rules with UseAI field
Assumes the UseAI column already exists in the list
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

def get_all_items(token, site_id):
    """Get all items from Style Rules list"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/items?expand=fields"
    response = requests.get(list_url, headers=headers)
    response.raise_for_status()
    return response.json()["value"]

def update_item(token, site_id, item_id, useai_value):
    """Update a single item with UseAI value"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    item_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/items/{item_id}"

    update_data = {
        "fields": {
            "UseAI": useai_value
        }
    }

    response = requests.patch(item_url, headers=headers, json=update_data)
    response.raise_for_status()
    return response.json()

def create_style_rules():
    """
    Define all style rules with UseAI field
    Returns dict mapping Title to UseAI value
    """
    rules = {
        "All text must use Arial font": False,
        "Use British English spelling - 'colour' not 'color'": True,
        "Use British English spelling - 'aluminium' not 'aluminum'": True,
        "Use British English spelling - 'analyse' not 'analyze'": True,
        "Use British English spelling - 'centre' not 'center'": True,
        "Use British English spelling - 'licence' (noun) not 'license'": True,
        "Use British English spelling - 'organise' not 'organize'": True,
        "No contractions in formal text - use 'cannot' not 'can't'": True,
        "No contractions in formal text - use 'do not' not 'don't'": True,
        "No contractions in formal text - use 'is not' not 'isn't'": True,
        "No contractions in formal text - use 'will not' not 'won't'": True,
        "Date format in text: DD MONTH YEAR (e.g., 01 February 2015)": False,
        "Time format: 24-hour with colon (e.g., 09:00, 18:25)": False,
        "Numbers below 10 should be spelled out in text": False,
        "Use commas with numbers of 4+ digits (e.g., 1,000)": True,
        "Section titles should be capitalised": False,
        "Subsidiary headings: only first letter and proper nouns capitalised": False,
        "Job titles only capitalised when with person's name": False,
        "Do not capitalise for emphasis": False,
        "Use 'toward' not 'towards'": True,
        "Avoid 'etc.' - be specific instead": True,
        "Use 'will', 'must', 'shall' instead of 'should' or 'could'": False,
        "Use metric units where possible": False,
        "Use hyphens with suffix '-wide' (e.g., site-wide)": False,
        "Hyphenate compound modifiers (e.g., 15-page document)": False,
        "Use single quotes for special terms on first reference": False,
        "Use double quotes for direct speech": False,
        "Never use apostrophes for plurals": True,
        "Avoid ampersand (&) - use 'and' instead": True,
        "Spell out 'percent' in text (not %)": True,
        "Figures and tables must have captions": False,
    }
    return rules

def main():
    print("🔧 Updating SharePoint Style Rules with UseAI field\n")

    # Get access token
    print("🔑 Getting access token...")
    token = get_token()

    # Get site ID
    print("🌐 Getting site ID...")
    site_id = get_site_id(token)
    print(f"   Site ID: {site_id}\n")

    # Get all existing items
    print("📋 Getting existing items...")
    items = get_all_items(token, site_id)
    print(f"   Found {len(items)} existing items\n")

    # Get UseAI mappings
    useai_mappings = create_style_rules()

    # Update each item
    updated_count = 0
    failed_count = 0

    for i, item in enumerate(items, 1):
        title = item["fields"].get("Title", "")
        item_id = item["id"]

        if title in useai_mappings:
            useai_value = useai_mappings[title]
            try:
                print(f"[{i}/{len(items)}] Updating: {title}")
                print(f"   UseAI = {useai_value}")
                update_item(token, site_id, item_id, useai_value)
                updated_count += 1
                print(f"   ✓ Success")
            except Exception as e:
                failed_count += 1
                print(f"   ✗ Failed: {str(e)}")
        else:
            print(f"[{i}/{len(items)}] Skipping: {title} (not in mappings)")
        print()

    # Summary
    print("=" * 60)
    print(f"✅ Successfully updated: {updated_count}")
    print(f"❌ Failed: {failed_count}")
    print(f"📊 Total items: {len(items)}")
    print("=" * 60)

if __name__ == "__main__":
    main()
