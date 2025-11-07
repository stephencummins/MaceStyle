import azure.functions as func
import datetime
import json
import logging
import os
import msal
import requests

app = func.FunctionApp()

@app.route(route="ValidateDocument", methods=["GET", "POST"])
def ValidateDocument(req: func.HttpRequest) -> func.HttpResponse:
    """HTTP trigger function to validate documents"""
    from ValidateDocument import main
    return main(req)

@app.route(route="TestSharePoint", methods=["GET"])
def TestSharePoint(req: func.HttpRequest) -> func.HttpResponse:
    """Test SharePoint connectivity using Graph API"""
    try:
        # Get credentials
        tenant_id = os.environ.get("SHAREPOINT_TENANT_ID")
        client_id = os.environ.get("SHAREPOINT_CLIENT_ID")
        client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET")
        site_url = os.environ.get("SHAREPOINT_SITE_URL")

        # Get access token using MSAL
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        scope = ["https://graph.microsoft.com/.default"]

        app = msal.ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret
        )

        result = app.acquire_token_for_client(scopes=scope)

        if "access_token" not in result:
            return func.HttpResponse(
                json.dumps({"error": "Failed to acquire token", "details": result}),
                mimetype="application/json",
                status_code=500
            )

        token = result["access_token"]

        # Extract hostname and site path
        parts = site_url.replace("https://", "").split("/")
        hostname = parts[0]
        site_path = "/" + "/".join(parts[1:]) if len(parts) > 1 else ""

        # Get site information using Graph API
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        }

        graph_site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
        site_response = requests.get(graph_site_url, headers=headers)
        site_response.raise_for_status()
        site_data = site_response.json()

        return func.HttpResponse(
            json.dumps({
                "status": "success",
                "siteTitle": site_data.get("displayName", "N/A"),
                "siteId": site_data.get("id"),
                "siteUrl": site_url,
                "webUrl": site_data.get("webUrl")
            }),
            mimetype="application/json"
        )
    except Exception as e:
        import traceback
        return func.HttpResponse(
            json.dumps({
                "error": str(e),
                "traceback": traceback.format_exc()
            }),
            mimetype="application/json",
            status_code=500
        )

@app.route(route="ListDocuments", methods=["GET"])
def ListDocuments(req: func.HttpRequest) -> func.HttpResponse:
    """List documents in SharePoint library"""
    try:
        # Get credentials
        tenant_id = os.environ.get("SHAREPOINT_TENANT_ID")
        client_id = os.environ.get("SHAREPOINT_CLIENT_ID")
        client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET")
        site_url = os.environ.get("SHAREPOINT_SITE_URL")

        # Get access token
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        scope = ["https://graph.microsoft.com/.default"]

        app_client = msal.ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret
        )

        result = app_client.acquire_token_for_client(scopes=scope)

        if "access_token" not in result:
            return func.HttpResponse(
                json.dumps({"error": "Failed to acquire token", "details": result}),
                mimetype="application/json",
                status_code=500
            )

        token = result["access_token"]

        # Extract hostname and site path
        parts = site_url.replace("https://", "").split("/")
        hostname = parts[0]
        site_path = "/" + "/".join(parts[1:]) if len(parts) > 1 else ""

        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        }

        # Get site ID
        graph_site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
        site_response = requests.get(graph_site_url, headers=headers)
        site_response.raise_for_status()
        site_id = site_response.json()["id"]

        # List files in default document library
        files_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
        files_response = requests.get(files_url, headers=headers)
        files_response.raise_for_status()
        files_data = files_response.json()

        documents = []
        for item in files_data.get("value", []):
            if item.get("file"):  # Only files, not folders
                documents.append({
                    "id": item.get("id"),
                    "name": item.get("name"),
                    "size": item.get("size"),
                    "webUrl": item.get("webUrl"),
                    "lastModified": item.get("lastModifiedDateTime")
                })

        return func.HttpResponse(
            json.dumps({
                "status": "success",
                "documentsCount": len(documents),
                "documents": documents
            }, indent=2),
            mimetype="application/json"
        )
    except Exception as e:
        import traceback
        return func.HttpResponse(
            json.dumps({
                "error": str(e),
                "traceback": traceback.format_exc()
            }),
            mimetype="application/json",
            status_code=500
        )