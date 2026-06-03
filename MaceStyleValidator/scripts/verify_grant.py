"""Verify Sapna's Sites.Selected grant for MaceStyleValidator-App on MaceWayControlCentre.

Non-destructive checks (read-only):
  1. Acquire a Graph token as the app (client credentials flow).
  2. Resolve the site by URL  -> proves the Sites.Selected grant lets the app SEE the site.
  3. List site permissions     -> shows the role(s) granted to this app (expect 'write').
  4. Read the site's lists      -> proves read access actually works.

Reads creds from the same env vars the Function App uses:
  SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, SHAREPOINT_SITE_URL
"""
import os
import sys
import msal
import requests

GRAPH = "https://graph.microsoft.com/v1.0"


def token():
    tid = os.environ["SHAREPOINT_TENANT_ID"]
    cid = os.environ["SHAREPOINT_CLIENT_ID"]
    sec = os.environ["SHAREPOINT_CLIENT_SECRET"]
    app = msal.ConfidentialClientApplication(
        cid, authority=f"https://login.microsoftonline.com/{tid}", client_credential=sec
    )
    r = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in r:
        print(f"  FAIL token: {r.get('error')}: {r.get('error_description')}")
        sys.exit(1)
    return r["access_token"]


def main():
    site_url = os.environ["SHAREPOINT_SITE_URL"]
    cid = os.environ["SHAREPOINT_CLIENT_ID"]
    parts = site_url.replace("https://", "").split("/")
    host, path = parts[0], "/" + "/".join(parts[1:])

    print(f"App:  {cid}")
    print(f"Site: {site_url}")
    print("-" * 60)

    t = token()
    print("[1] Token acquired (client credentials).  PASS")

    h = {"Authorization": f"Bearer {t}", "Accept": "application/json"}

    r = requests.get(f"{GRAPH}/sites/{host}:{path}", headers=h)
    if r.status_code != 200:
        print(f"[2] Resolve site: HTTP {r.status_code}  FAIL")
        print(f"    {r.text[:300]}")
        print("    -> Grant NOT effective: app cannot see the site.")
        sys.exit(1)
    site_id = r.json()["id"]
    print(f"[2] Site resolved.  PASS  (Sites.Selected grant is active)")
    print(f"    site id: {site_id}")

    r = requests.get(f"{GRAPH}/sites/{site_id}/permissions", headers=h)
    if r.status_code == 200:
        ours = []
        for p in r.json().get("value", []):
            ids = p.get("grantedToIdentitiesV2", p.get("grantedToIdentities", []))
            for i in ids:
                appinfo = i.get("application", {})
                if appinfo.get("id") == cid:
                    ours.append(p.get("roles"))
        print(f"[3] Site permissions readable.  roles for this app: {ours or 'NONE VISIBLE'}")
    else:
        print(f"[3] Site permissions: HTTP {r.status_code} (app may lack permission to read perms; not fatal)")

    r = requests.get(f"{GRAPH}/sites/{site_id}/lists?$select=displayName,id&$top=50", headers=h)
    if r.status_code != 200:
        print(f"[4] Read lists: HTTP {r.status_code}  FAIL")
        print(f"    {r.text[:300]}")
        sys.exit(1)
    lists = r.json().get("value", [])
    print(f"[4] Lists readable: {len(lists)} found.  PASS  (read access confirmed)")
    for lst in lists[:15]:
        print(f"    - {lst['displayName']}")

    print("-" * 60)
    print("RESULT: Grant is working for READ. Run a write probe separately to confirm 'write'.")


if __name__ == "__main__":
    main()
