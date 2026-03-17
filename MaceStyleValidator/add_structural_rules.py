"""
Add Structural Validation Rules to SharePoint Style Rules List

This script adds Visio structural validation rules including:
- Page dimension standardization
- Shape size validation
- Position validation (margins and exact placement)

Run this script to populate your SharePoint list with structural rules.
"""
import requests

from ValidateDocument.config import get_graph_token, get_site_id

def check_tolerance_column(token, site_id):
    """Check if Tolerance column exists in Style Rules list"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    columns_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/columns"
    response = requests.get(columns_url, headers=headers)
    response.raise_for_status()

    columns = response.json().get("value", [])
    tolerance_exists = any(col.get("name") == "Tolerance" for col in columns)

    return tolerance_exists

def add_tolerance_column(token, site_id):
    """Add Tolerance column to Style Rules list"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    columns_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/columns"

    # Simpler column definition that should work with SharePoint
    column_data = {
        "name": "Tolerance",
        "displayName": "Tolerance",
        "description": "Acceptable variance for dimension/position validation (in inches)",
        "number": {}
    }

    try:
        response = requests.post(columns_url, headers=headers, json=column_data)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.HTTPError as e:
        # Try with even simpler definition
        column_data = {
            "name": "Tolerance",
            "text": {}
        }
        response = requests.post(columns_url, headers=headers, json=column_data)
        response.raise_for_status()
        return response.json()

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

def create_structural_rules():
    """
    Define structural validation rules for Visio diagrams

    These rules enforce:
    - Page dimensions (Letter, A4, etc.)
    - Shape sizes (title boxes, icons, process boxes)
    - Positions (margins, exact placement)
    """

    rules = [
        # ============================================
        # PAGE DIMENSION RULES
        # ============================================
        {
            "Title": "Visio - Page Size Letter Landscape",
            "RuleType": "PageDimensions",
            "DocumentType": "Visio",
            "CheckValue": "PageSize",
            "ExpectedValue": "11.0x8.5",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 80
        },

        # ============================================
        # SHAPE SIZE RULES
        # ============================================
        {
            "Title": "Visio - Title Box 3x1 inches",
            "RuleType": "Size",
            "DocumentType": "Visio",
            "CheckValue": "TitleBoxSize",
            "ExpectedValue": "3.0x1.0",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 90,
            "Tolerance": 0.1
        },
        {
            "Title": "Visio - Icons 0.5 inch Square",
            "RuleType": "Size",
            "DocumentType": "Visio",
            "CheckValue": "IconSize",
            "ExpectedValue": "0.5x0.5",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 92,
            "Tolerance": 0.05
        },
        {
            "Title": "Visio - Process Box 2x1.5 inches",
            "RuleType": "Size",
            "DocumentType": "Visio",
            "CheckValue": "ProcessBoxSize",
            "ExpectedValue": "2.0x1.5",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 91,
            "Tolerance": 0.1
        },

        # ============================================
        # POSITION RULES - MARGINS
        # ============================================
        {
            "Title": "Visio - Top Margin 2 inches",
            "RuleType": "Position",
            "DocumentType": "Visio",
            "CheckValue": "TopMargin",
            "ExpectedValue": "2.0",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 85,
            "Tolerance": 0.1
        },
        {
            "Title": "Visio - Left Margin 1 inch",
            "RuleType": "Position",
            "DocumentType": "Visio",
            "CheckValue": "LeftMargin",
            "ExpectedValue": "1.0",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 86,
            "Tolerance": 0.1
        },
        {
            "Title": "Visio - Right Margin at 10 inches",
            "RuleType": "Position",
            "DocumentType": "Visio",
            "CheckValue": "RightMargin",
            "ExpectedValue": "10.0",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 87,
            "Tolerance": 0.1
        },
        {
            "Title": "Visio - Bottom Margin 1 inch",
            "RuleType": "Position",
            "DocumentType": "Visio",
            "CheckValue": "BottomMargin",
            "ExpectedValue": "1.0",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 88,
            "Tolerance": 0.1
        },

        # ============================================
        # POSITION RULES - EXACT PLACEMENT
        # ============================================
        {
            "Title": "Visio - Logo at Top-Left (0.5, 7.5)",
            "RuleType": "Position",
            "DocumentType": "Visio",
            "CheckValue": "ExactPosition",
            "ExpectedValue": "0.5,7.5",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 75,
            "Tolerance": 0.05
        },
        {
            "Title": "Visio - Footer Bottom-Right (9.5, 0.5)",
            "RuleType": "Position",
            "DocumentType": "Visio",
            "CheckValue": "ExactPosition",
            "ExpectedValue": "9.5,0.5",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 76,
            "Tolerance": 0.05
        },

        # ============================================
        # FONT RULES FOR VISIO
        # ============================================
        {
            "Title": "Visio - All Fonts Must Be Arial",
            "RuleType": "Font",
            "DocumentType": "Visio",
            "CheckValue": "AllTextFont",
            "ExpectedValue": "Arial",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 100
        },

        # ============================================
        # COLOR RULES FOR VISIO
        # ============================================
        {
            "Title": "Visio - Shape Fill Color Brand Blue",
            "RuleType": "Color",
            "DocumentType": "Visio",
            "CheckValue": "ShapeFillColor",
            "ExpectedValue": "#003399",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 101
        },
        {
            "Title": "Visio - Text Color Black",
            "RuleType": "Color",
            "DocumentType": "Visio",
            "CheckValue": "ShapeTextColor",
            "ExpectedValue": "#000000",
            "AutoFix": True,
            "UseAI": False,
            "Priority": 102
        },
    ]

    return rules

def main():
    print("=" * 70)
    print("ADDING STRUCTURAL VALIDATION RULES TO SHAREPOINT")
    print("=" * 70)
    print()

    try:
        # Step 1: Get authentication token
        print("Step 1: Getting authentication token...")
        token = get_graph_token()
        print("✓ Token acquired successfully")
        print()

        # Step 2: Get site ID
        print("Step 2: Getting SharePoint site ID...")
        site_id = get_site_id(token)
        print(f"✓ Site ID: {site_id}")
        print()

        # Step 3: Check if Tolerance column exists
        print("Step 3: Checking if Tolerance column exists...")
        tolerance_exists = check_tolerance_column(token, site_id)

        if tolerance_exists:
            print("✓ Tolerance column already exists")
        else:
            print("! Tolerance column not found")
            print()
            print("  MANUAL STEP REQUIRED:")
            print("  Please add a 'Tolerance' column to your Style Rules list:")
            print("  1. Go to SharePoint → Style Rules list")
            print("  2. Click '+ Add column' → Number")
            print("  3. Name: Tolerance")
            print("  4. Description: Acceptable variance for dimension/position validation (in inches)")
            print("  5. Click Save")
            print()
            print("  Continuing without Tolerance column...")
            print("  (Rules with Tolerance values will omit that field)")
            print()
        print()

        # Step 4: Get rules
        print("Step 4: Preparing structural validation rules...")
        rules = create_structural_rules()
        print(f"✓ Prepared {len(rules)} structural validation rules")
        print()

        # Step 5: Add rules
        print("Step 5: Adding rules to SharePoint...")
        print()

        success_count = 0
        error_count = 0

        for i, rule in enumerate(rules, 1):
            try:
                # If Tolerance column doesn't exist, remove Tolerance field from rule
                if not tolerance_exists and 'Tolerance' in rule:
                    print(f"  ! [{i:2d}/{len(rules)}] Removing Tolerance field from: {rule['Title']}")
                    rule_copy = rule.copy()
                    del rule_copy['Tolerance']
                    result = add_rule(token, site_id, rule_copy)
                else:
                    result = add_rule(token, site_id, rule)

                print(f"  ✓ [{i:2d}/{len(rules)}] Added: {rule['Title']}")
                success_count += 1
            except requests.exceptions.HTTPError as e:
                error_count += 1
                print(f"  ✗ [{i:2d}/{len(rules)}] Failed: {rule['Title']}")
                print(f"      Error: {e.response.status_code} - {e.response.text}")
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
            print("✓ Structural validation rules added successfully!")
            print()

            if not tolerance_exists:
                print("⚠ IMPORTANT: Tolerance column not found")
                print("  To enable tolerance-based validation:")
                print("  1. Go to SharePoint → Style Rules list")
                print("  2. Click '+ Add column' → Number")
                print("  3. Name: Tolerance")
                print("  4. For each rule with tolerance, edit and add the value manually")
                print()
                print("  Recommended Tolerance values:")
                print("    - Shape sizes: 0.1 (standard)")
                print("    - Icons: 0.05 (precise)")
                print("    - Margins: 0.1 (standard)")
                print("    - Exact positions: 0.05 (precise)")
                print()

            print("Next steps:")
            print("1. Go to SharePoint → Style Rules list to verify")
            print("2. Upload a Visio file to test validation")
            print("3. Check the validation report for results")
            print()
            print("Rule Categories Added:")
            print("  - Page Dimensions: 1 rule")
            print("  - Shape Sizes: 3 rules")
            print("  - Position Margins: 4 rules")
            print("  - Exact Positions: 2 rules")
            print("  - Font: 1 rule")
            print("  - Color: 2 rules")
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
