"""
One-off script to update all SharePoint Style Rules with DocumentType='Word' to 'Both',
so they apply to both Word and Excel documents.
"""
import os
import msal
import requests

TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID", "2ab0866e-23d6-4688-be97-ce9f447135d8")
CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID", "c7859dae-6997-448f-9530-7166fe857e75")
CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET")
SITE_URL = os.environ.get("SHAREPOINT_SITE_URL", "https://0rxf2.sharepoint.com/sites/StyleValidation")


def get_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"Failed to acquire token: {result}")


def get_site_id(token):
    parts = SITE_URL.replace("https://", "").split("/")
    hostname = parts[0]
    site_path = "/" + "/".join(parts[1:])
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    resp = requests.get(f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}", headers=headers)
    resp.raise_for_status()
    return resp.json()["id"]


def main():
    print("Updating Style Rules: DocumentType 'Word'/'Both' -> 'All'\n")

    token = get_token()
    site_id = get_site_id(token)
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json", "Content-Type": "application/json"}

    # Fetch all list items
    list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/items?expand=fields&$top=200"
    resp = requests.get(list_url, headers=headers)
    resp.raise_for_status()
    items = resp.json().get("value", [])

    # First, show what DocumentType values exist
    doc_types = {}
    for item in items:
        fields = item.get("fields", {})
        doc_type = fields.get("DocumentType") or fields.get("field_2") or "<missing>"
        doc_types[doc_type] = doc_types.get(doc_type, 0) + 1

    print(f"Found {len(items)} items. DocumentType distribution:")
    for dt, count in sorted(doc_types.items()):
        print(f"  {dt}: {count}")

    # Show first item's fields for debugging
    if items:
        print(f"\nSample fields: {list(items[0].get('fields', {}).keys())}")
        f = items[0].get('fields', {})
        print(f"  Title: {f.get('Title')}")
        for i in range(1, 8):
            print(f"  field_{i}: {f.get(f'field_{i}')}")

    updated = 0
    skipped = 0

    for item in items:
        fields = item.get("fields", {})
        doc_type = fields.get("DocumentType") or fields.get("field_2") or ""
        title = fields.get("Title", "Unknown")
        item_id = item["id"]

        if doc_type in ("Word", "Both"):
            patch_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/items/{item_id}/fields"
            patch_resp = requests.patch(patch_url, headers=headers, json={"field_2": "All"})

            if patch_resp.ok:
                print(f"  Updated: {title}")
                updated += 1
            else:
                print(f"  FAILED: {title} - {patch_resp.status_code} {patch_resp.text}")
        else:
            skipped += 1

    print(f"\nDone. Updated: {updated}, Already non-Word: {skipped}")


if __name__ == "__main__":
    main()
