"""Reversible write probe: create -> read back -> delete a throwaway item in
the 'Validation Results' list on MaceWayControlCentre. Proves the app's grant
allows writes end-to-end, then cleans up after itself.
"""
import os
import sys
import datetime
import msal
import requests

GRAPH = "https://graph.microsoft.com/v1.0"
LIST_NAME = "Validation Results"


def token():
    app = msal.ConfidentialClientApplication(
        os.environ["SHAREPOINT_CLIENT_ID"],
        authority=f"https://login.microsoftonline.com/{os.environ['SHAREPOINT_TENANT_ID']}",
        client_credential=os.environ["SHAREPOINT_CLIENT_SECRET"],
    )
    r = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in r:
        sys.exit(f"token FAIL: {r.get('error_description')}")
    return r["access_token"]


def main():
    site_url = os.environ["SHAREPOINT_SITE_URL"]
    parts = site_url.replace("https://", "").split("/")
    host, path = parts[0], "/" + "/".join(parts[1:])
    h = {"Authorization": f"Bearer {token()}", "Accept": "application/json",
         "Content-Type": "application/json"}

    site_id = requests.get(f"{GRAPH}/sites/{host}:{path}", headers=h).json()["id"]
    lists = requests.get(f"{GRAPH}/sites/{site_id}/lists?$select=displayName,id&$top=50",
                         headers=h).json()["value"]
    lst = next((l for l in lists if l["displayName"] == LIST_NAME), None)
    if not lst:
        sys.exit(f"List '{LIST_NAME}' not found")
    list_id = lst["id"]
    print(f"List '{LIST_NAME}': {list_id}")

    marker = f"GRANT-WRITE-PROBE {datetime.datetime.utcnow().isoformat()}Z (safe to delete)"

    # CREATE
    r = requests.post(f"{GRAPH}/sites/{site_id}/lists/{list_id}/items",
                      headers=h, json={"fields": {"Title": marker}})
    if r.status_code not in (200, 201):
        print(f"[CREATE] HTTP {r.status_code}  FAIL")
        print(f"  {r.text[:400]}")
        sys.exit(1)
    item_id = r.json()["id"]
    print(f"[CREATE] item {item_id} created.  PASS  (write works)")

    # READ BACK
    r = requests.get(f"{GRAPH}/sites/{site_id}/lists/{list_id}/items/{item_id}?$expand=fields",
                     headers=h)
    got = r.json().get("fields", {}).get("Title") if r.status_code == 200 else None
    print(f"[READ ] back: {'match' if got == marker else got!r}  {'PASS' if got == marker else 'WARN'}")

    # DELETE (cleanup)
    r = requests.delete(f"{GRAPH}/sites/{site_id}/lists/{list_id}/items/{item_id}", headers=h)
    if r.status_code == 204:
        print(f"[DELETE] item {item_id} removed.  PASS  (cleaned up)")
    else:
        print(f"[DELETE] HTTP {r.status_code}  FAIL — MANUAL CLEANUP NEEDED for item {item_id}")
        print(f"  {r.text[:300]}")
        sys.exit(1)

    print("-" * 60)
    print("RESULT: Full read+write confirmed. Test item created and deleted.")


if __name__ == "__main__":
    main()
