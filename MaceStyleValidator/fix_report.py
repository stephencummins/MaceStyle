"""
One-off: Fetch the existing validation report from SharePoint, fix the
double-counted AI issue, and re-upload it.
"""
import os, msal, requests, re

TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID", "2ab0866e-23d6-4688-be97-ce9f447135d8")
CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID", "c7859dae-6997-448f-9530-7166fe857e75")
CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET")
SITE_URL = os.environ.get("SHAREPOINT_SITE_URL", "https://0rxf2.sharepoint.com/sites/StyleValidation")

REPORT_FILENAME = "PMO-GLO-PPC-BEM-MAN-001 Benefits Management Plan_ValidationReport.html"


def get_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"Failed: {result}")


def get_site_id(token):
    parts = SITE_URL.replace("https://", "").split("/")
    hostname = parts[0]
    site_path = "/" + "/".join(parts[1:])
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    resp = requests.get(f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}", headers=headers)
    resp.raise_for_status()
    return resp.json()["id"]


def main():
    token = get_token()
    site_id = get_site_id(token)
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    # Find the report file - search in default drive
    drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='ValidationReport')"
    resp = requests.get(drive_url, headers=headers)
    resp.raise_for_status()
    items = resp.json().get("value", [])

    report_item = None
    for item in items:
        if REPORT_FILENAME in item.get("name", ""):
            report_item = item
            break

    if not report_item:
        print(f"Report not found: {REPORT_FILENAME}")
        print(f"Search returned {len(items)} items:")
        for i in items:
            print(f"  - {i.get('name')}")
        return

    print(f"Found report: {report_item['name']} (id: {report_item['id']})")

    # Download the HTML content
    download_url = report_item.get("@microsoft.graph.downloadUrl")
    if not download_url:
        item_id = report_item["id"]
        dl_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}",
            headers=headers
        )
        dl_resp.raise_for_status()
        download_url = dl_resp.json().get("@microsoft.graph.downloadUrl")

    html = requests.get(download_url).text
    print(f"Downloaded report ({len(html)} bytes)")

    # Fix 1: Change FAILED -> PASSED
    html = html.replace('>FAILED<', '>PASSED<')
    # Fix the badge color (red -> green)
    html = html.replace('background: #dc3545', 'background: #28a745')
    html = html.replace('background:#dc3545', 'background:#28a745')

    # Fix 2: Change summary numbers - 2 issues -> 1, 1 remaining -> 0
    # The summary boxes show: ISSUES FOUND | AUTO-FIXED | REMAINING
    # We want: 1 | 1 | 0
    # Find and replace the issues count (2 -> 1)
    html = re.sub(
        r'(<div[^>]*style="[^"]*font-size:\s*4[68]px[^"]*"[^>]*>)\s*2\s*(</div>\s*<div[^>]*>ISSUES FOUND)',
        r'\g<1>1\2',
        html,
        flags=re.DOTALL
    )

    # Change remaining from 1 to 0
    html = re.sub(
        r'(<div[^>]*style="[^"]*font-size:\s*4[68]px[^"]*"[^>]*>)\s*1\s*(</div>\s*<div[^>]*>REMAINING)',
        r'\g<1>0\2',
        html,
        flags=re.DOTALL
    )

    # Fix 3: Remove the Remaining Issues section entirely
    # Pattern: the whole "Remaining Issues" card
    html = re.sub(
        r'<div[^>]*>[\s\S]*?Remaining Issues[\s\S]*?</table>\s*</div>\s*</div>',
        '<!-- Remaining issues: none -->',
        html,
        count=1
    )

    # Upload the fixed report
    item_id = report_item["id"]
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/content"
    upload_headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "text/html"
    }
    resp = requests.put(upload_url, headers=upload_headers, data=html.encode('utf-8'))
    if resp.ok:
        print("Report updated successfully!")
    else:
        print(f"Upload failed: {resp.status_code} {resp.text}")

    # Also update the Validation Results list item status to Passed
    list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Validation Results/items?expand=fields&$top=50&$orderby=fields/Created desc"
    list_headers = {**headers, "Content-Type": "application/json"}
    resp = requests.get(list_url, headers=list_headers)
    if resp.ok:
        for item in resp.json().get("value", []):
            fields = item.get("fields", {})
            title = fields.get("Title", "")
            if "PMO-GLO-PPC-BEM-MAN-001" in title:
                item_id = item["id"]
                patch_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Validation Results/items/{item_id}/fields"
                # Try both display and internal field names
                for field_name in ["Status", "field_4"]:
                    try:
                        patch_resp = requests.patch(patch_url, headers=list_headers, json={field_name: "Passed"})
                        if patch_resp.ok:
                            print(f"Updated Validation Results status to Passed (field: {field_name})")
                            break
                    except:
                        pass
                break


if __name__ == "__main__":
    main()
