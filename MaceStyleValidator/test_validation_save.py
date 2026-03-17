"""
Test saving to Validation Results list
"""
import sys
sys.path.insert(0, 'ValidateDocument')

from sharepoint_results import save_validation_result
from ValidateDocument.config import get_graph_token, get_site_id

def main():
    print("🧪 Testing Validation Results save functionality\n")

    try:
        # Get access token
        print("🔑 Getting access token...")
        token = get_graph_token()
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
