"""
Test saving to Validation Results list
"""
import os
import sys
sys.path.insert(0, 'ValidateDocument')

from sharepoint_results import save_validation_result
import msal
import requests

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

def main():
    print("🧪 Testing Validation Results save functionality\n")

    try:
        # Get access token
        print("🔑 Getting access token...")
        token = get_token()
        print("   ✓ Token acquired\n")

        # Get site ID
        print("🌐 Getting site ID...")
        site_id = get_site_id(token)
        print(f"   ✓ Site ID: {site_id}\n")

        # Test save
        print("💾 Attempting to save validation result...")
        html_report = """
        <html>
        <head><title>Test Report</title></head>
        <body>
            <h1>Test Validation Report</h1>
            <p>This is a test report.</p>
        </body>
        </html>
        """

        result = save_validation_result(
            token=token,
            site_id=site_id,
            filename="test_document.docx",
            issues_count=5,
            fixes_count=3,
            status="Passed",
            html_report=html_report
        )

        print("   ✓ Validation result saved successfully!\n")
        print(f"📋 Result Details:")
        print(f"   Item ID: {result['item_id']}")
        print(f"   Report URL: {result['report_url']}")
        print(f"   List Item URL: {result['list_item_url']}")
        print("\n✅ Test PASSED - Validation Results save is working!")

    except Exception as e:
        print(f"\n❌ Test FAILED")
        print(f"   Error: {str(e)}")
        print(f"   Error type: {type(e).__name__}")

        import traceback
        print(f"\n📋 Full traceback:")
        print(traceback.format_exc())
        return 1

    return 0

if __name__ == "__main__":
    exit(main())
