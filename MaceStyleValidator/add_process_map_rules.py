"""
Add Process Map Validation Rules to SharePoint Style Rules List

Rules based on Process Map Template requirements:
- Top/bottom line requirements
- Stage Gates (RIBA Plan of Work)
- Box formatting and colors
- Swim lane requirements
- Activity types and decision points
- Document references
"""
import os
import msal
import requests
import json

# Configuration
TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID", "2ab0866e-23d6-4688-be97-ce9f447135d8")
CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID", "c7859dae-6997-448f-9530-7166fe857e75")
CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET")
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

def add_rule(token, site_id, rule):
    """Add a single rule to the Style Rules list"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/items"

    item_data = {
        "fields": rule
    }

    response = requests.post(list_url, headers=headers, json=item_data)
    response.raise_for_status()
    return response.json()

def create_process_map_rules():
    """
    Define process map validation rules based on template requirements

    These rules apply to Visio/PowerPoint process maps for the New Hospitals Programme
    """

    rules = [
        # ============================================
        # SLIDE/PAGE STRUCTURE RULES
        # ============================================
        {
            "Title": "Process Map - Slide Size 10x5.62 inches",
            "RuleType": "PageDimensions",
            "DocumentType": "Both",
            "CheckValue": "PageSize",
            "ExpectedValue": "10.0x5.62",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 50
        },

        # ============================================
        # REQUIRED CONTENT RULES (AI-POWERED)
        # ============================================
        {
            "Title": "Process Map - Top Line Workstream Name Required",
            "RuleType": "Layout",
            "DocumentType": "Both",
            "CheckValue": "TopLineContent",
            "ExpectedValue": "Workstream name must be present at top",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 110
        },
        {
            "Title": "Process Map - Bottom Line Process Name Required",
            "RuleType": "Layout",
            "DocumentType": "Both",
            "CheckValue": "BottomLineContent",
            "ExpectedValue": "Process name must be present at bottom",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 111
        },
        {
            "Title": "Process Map - Document Reference Required",
            "RuleType": "Layout",
            "DocumentType": "Both",
            "CheckValue": "DocumentReference",
            "ExpectedValue": "Doc Ref: [value] must be present",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 112
        },
        {
            "Title": "Process Map - Set Up or Start Activity Required",
            "RuleType": "Layout",
            "DocumentType": "Both",
            "CheckValue": "StartActivity",
            "ExpectedValue": "Must contain 'Set Up' or 'Start' activity",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 113
        },
        {
            "Title": "Process Map - Feedback Activity Required",
            "RuleType": "Layout",
            "DocumentType": "Both",
            "CheckValue": "FeedbackActivity",
            "ExpectedValue": "Must contain feedback activity at end",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 114
        },

        # ============================================
        # SWIM LANE RULES
        # ============================================
        {
            "Title": "Process Map - All 5 Swim Lanes Required",
            "RuleType": "Layout",
            "DocumentType": "Both",
            "CheckValue": "SwimLanes",
            "ExpectedValue": "All five swim lanes must be present: New Hospitals Programme, NHS, Healthy Delivery Partnership, Delivery Team, Contractor/Supply Chain",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 115
        },

        # ============================================
        # STAGE GATES RULES
        # ============================================
        {
            "Title": "Process Map - RIBA Stage Gates Format",
            "RuleType": "Language",
            "DocumentType": "Both",
            "CheckValue": "StageGates",
            "ExpectedValue": "Stage Gates must follow RIBA Plan of Work format",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 116
        },

        # ============================================
        # BOX FORMATTING RULES
        # ============================================
        {
            "Title": "Process Map - Box Headers One Word or Short Sentence",
            "RuleType": "Grammar",
            "DocumentType": "Both",
            "CheckValue": "BoxHeaders",
            "ExpectedValue": "Each box header shall be one word or short sentence",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 117
        },
        {
            "Title": "Process Map - Box Body Concise Description",
            "RuleType": "Grammar",
            "DocumentType": "Both",
            "CheckValue": "BoxBodyContent",
            "ExpectedValue": "Box body must be concise description without role ownership",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 118
        },
        {
            "Title": "Process Map - Multi-Lane Activities Grey Color",
            "RuleType": "Color",
            "DocumentType": "Both",
            "CheckValue": "MultiLaneActivityColor",
            "ExpectedValue": "#808080",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 119
        },
        {
            "Title": "Process Map - Interface Boxes Dark Grey",
            "RuleType": "Color",
            "DocumentType": "Both",
            "CheckValue": "InterfaceBoxColor",
            "ExpectedValue": "#404040",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 120
        },

        # ============================================
        # SWIM LANE COLOR RULES
        # ============================================
        {
            "Title": "Process Map - New Hospitals Programme Lane Color",
            "RuleType": "Color",
            "DocumentType": "Both",
            "CheckValue": "SwimLane1Color",
            "ExpectedValue": "Boxes in New Hospitals Programme lane use lane color in header",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 121
        },
        {
            "Title": "Process Map - NHS Lane Color",
            "RuleType": "Color",
            "DocumentType": "Both",
            "CheckValue": "SwimLane2Color",
            "ExpectedValue": "Boxes in NHS lane use lane color in header",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 122
        },
        {
            "Title": "Process Map - Healthy Delivery Partnership Lane Color",
            "RuleType": "Color",
            "DocumentType": "Both",
            "CheckValue": "SwimLane3Color",
            "ExpectedValue": "Boxes in Healthy Delivery Partnership lane use lane color in header",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 123
        },
        {
            "Title": "Process Map - Delivery Team Lane Color",
            "RuleType": "Color",
            "DocumentType": "Both",
            "CheckValue": "SwimLane4Color",
            "ExpectedValue": "Boxes in Delivery Team lane use lane color in header",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 124
        },
        {
            "Title": "Process Map - Contractor Lane Color",
            "RuleType": "Color",
            "DocumentType": "Both",
            "CheckValue": "SwimLane5Color",
            "ExpectedValue": "Boxes in Contractor/Supply Chain lane use lane color in header",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 125
        },

        # ============================================
        # DECISION POINT RULES
        # ============================================
        {
            "Title": "Process Map - Decision Points Use Diamond Shapes",
            "RuleType": "Layout",
            "DocumentType": "Visio",
            "CheckValue": "DecisionShapes",
            "ExpectedValue": "Decision points must use diamond shapes",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 126
        },
        {
            "Title": "Process Map - Decision Arrows Have YES/NO Labels",
            "RuleType": "Language",
            "DocumentType": "Both",
            "CheckValue": "DecisionLabels",
            "ExpectedValue": "Arrows from decision points must contain YES or NO",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 127
        },

        # ============================================
        # ACTIVITY GUIDE RULES
        # ============================================
        {
            "Title": "Process Map - Underlined Headers Need Activity Guide",
            "RuleType": "Layout",
            "DocumentType": "Both",
            "CheckValue": "UnderlinedHeaders",
            "ExpectedValue": "Underlined box headers must have corresponding Activity Guide",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 128
        },
        {
            "Title": "Process Map - Activity Guide References Included",
            "RuleType": "Language",
            "DocumentType": "Both",
            "CheckValue": "ActivityGuideReferences",
            "ExpectedValue": "Activity Guide references shall be included where applicable",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 129
        },

        # ============================================
        # ACTIVITY REFERENCE FORMAT RULES
        # ============================================
        {
            "Title": "Process Map - Activity Reference Format XX-000001",
            "RuleType": "Language",
            "DocumentType": "Both",
            "CheckValue": "ActivityReferenceFormat",
            "ExpectedValue": "XX-000001",
            "AutoFix": False,
            "UseAI": True,
            "Priority": 130
        },
    ]

    return rules

def main():
    print("=" * 70)
    print("ADDING PROCESS MAP VALIDATION RULES TO SHAREPOINT")
    print("=" * 70)
    print()

    try:
        # Step 1: Get authentication token
        print("Step 1: Getting authentication token...")
        token = get_token()
        print("✓ Token acquired successfully")
        print()

        # Step 2: Get site ID
        print("Step 2: Getting SharePoint site ID...")
        site_id = get_site_id(token)
        print(f"✓ Site ID: {site_id}")
        print()

        # Step 3: Get rules
        print("Step 3: Preparing process map validation rules...")
        rules = create_process_map_rules()
        print(f"✓ Prepared {len(rules)} process map validation rules")
        print()

        # Step 4: Add rules
        print("Step 4: Adding rules to SharePoint...")
        print()

        success_count = 0
        error_count = 0

        for i, rule in enumerate(rules, 1):
            try:
                result = add_rule(token, site_id, rule)
                print(f"  ✓ [{i:2d}/{len(rules)}] Added: {rule['Title']}")
                success_count += 1
            except requests.exceptions.HTTPError as e:
                error_count += 1
                print(f"  ✗ [{i:2d}/{len(rules)}] Failed: {rule['Title']}")
                error_detail = e.response.text if hasattr(e.response, 'text') else str(e)
                print(f"      Error: {e.response.status_code} - {error_detail[:200]}")
            except Exception as e:
                error_count += 1
                print(f"  ✗ [{i:2d}/{len(rules)}] Failed: {rule['Title']}")
                print(f"      Error: {str(e)}")

        print()
        print("=" * 70)
        print("SUMMARY")
        print("=" * 70)
        print(f"Total rules:      {len(rules)}")
        print(f"Added:            {success_count}")
        print(f"Failed:           {error_count}")
        print()

        if success_count > 0:
            print("✓ Process map validation rules added successfully!")
            print()
            print("Next steps:")
            print("1. Go to SharePoint → Style Rules list to verify")
            print("2. Upload a process map to test validation")
            print("3. Check the validation report for results")
            print()
            print("Rule Categories Added:")
            print("  - Page Structure: 1 rule")
            print("  - Required Content: 5 rules")
            print("  - Swim Lanes: 6 rules (including color rules)")
            print("  - Box Formatting: 4 rules")
            print("  - Decision Points: 2 rules")
            print("  - Activity Guides: 2 rules")
            print("  - Stage Gates: 1 rule")
            print("  - Activity References: 1 rule")
            print()
            print("IMPORTANT NOTES:")
            print("  - Most rules use AI validation (UseAI: True)")
            print("  - Some rules cannot auto-fix structural requirements")
            print("  - Color rules for grey/dark grey boxes can auto-fix")
            print("  - Page size can auto-fix to 10.0x5.62")
            print()

        if error_count > 0:
            print("⚠ Some rules failed to add. Check errors above.")
            print()

    except Exception as e:
        print()
        print("=" * 70)
        print("ERROR")
        print("=" * 70)
        print(f"Failed to add rules: {str(e)}")
        print()
        import traceback
        traceback.print_exc()
        return False

    return success_count > 0

if __name__ == "__main__":
    import sys
    success = main()
    sys.exit(0 if success else 1)
