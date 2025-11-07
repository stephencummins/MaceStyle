"""
Test script to validate a document from SharePoint without list item tracking
This directly uses the validation modules to test the full workflow
"""

import sys
sys.path.append('ValidateDocument')

from ValidateDocument import (
    get_graph_token,
    download_file,
    fetch_validation_rules,
    validate_word_document,
    generate_report,
    upload_file
)
from io import BytesIO

def main():
    print("🧪 Testing SharePoint Document Validation")
    print("=" * 60)

    # Configuration
    file_path = "/test.docx"  # File path relative to drive root
    file_name = "test.docx"

    try:
        # Step 1: Get access token
        print("\n1. Acquiring Microsoft Graph API token...")
        token = get_graph_token()
        print("   ✅ Token acquired")

        # Step 2: Fetch validation rules from Style Rules list
        print("\n2. Fetching validation rules from SharePoint...")
        rules = fetch_validation_rules(token)
        print(f"   ✅ Loaded {len(rules)} validation rules:")
        for rule in rules:
            print(f"      - {rule['title']} ({rule['rule_type']})")

        # Step 3: Download document
        print(f"\n3. Downloading document: {file_path}...")
        file_stream = download_file(token, file_path)
        print("   ✅ Document downloaded")

        # Step 4: Validate document
        print("\n4. Validating document against rules...")
        result = validate_word_document(file_stream, rules)
        print(f"   ✅ Validation complete")
        print(f"      - Issues found: {len(result['issues'])}")
        print(f"      - Fixes applied: {len(result['fixes_applied'])}")

        # Step 5: Display issues
        if result['issues']:
            print("\n   📋 Issues detected:")
            for issue in result['issues']:
                print(f"      - {issue}")
        else:
            print("\n   ✅ No issues detected!")

        # Step 6: Display fixes
        if result['fixes_applied']:
            print("\n   🔧 Auto-fixes applied:")
            for fix in result['fixes_applied']:
                print(f"      - {fix}")

        # Step 7: Generate report
        print("\n5. Generating validation report...")
        report_html = generate_report(file_name, result['issues'], result['fixes_applied'])
        print("   ✅ Report generated")

        # Step 8: Save fixed document locally for inspection
        print("\n6. Saving fixed document locally...")
        result['document'].save('test_output_fixed.docx')
        print("   ✅ Saved to: test_output_fixed.docx")

        # Step 9: Save report locally
        with open('test_output_report.html', 'w') as f:
            f.write(report_html)
        print("   ✅ Saved report to: test_output_report.html")

        # Step 10: Optionally upload fixed document back to SharePoint
        if result['fixes_applied']:
            print("\n7. Uploading fixed document back to SharePoint...")
            fixed_stream = BytesIO()
            result['document'].save(fixed_stream)
            fixed_stream.seek(0)

            upload_path = "/test_FIXED.docx"
            web_url = upload_file(token, fixed_stream, upload_path)
            print(f"   ✅ Fixed document uploaded")
            print(f"      URL: {web_url}")

        print("\n" + "=" * 60)
        print("✨ Validation test completed successfully!")
        print("\nSummary:")
        print(f"  - Validation rules checked: {len(rules)}")
        print(f"  - Issues found: {len(result['issues'])}")
        print(f"  - Auto-fixes applied: {len(result['fixes_applied'])}")
        print(f"  - Status: {'⚠️  Failed' if result['issues'] else '✅ Passed'}")

    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1

    return 0

if __name__ == "__main__":
    sys.exit(main())
