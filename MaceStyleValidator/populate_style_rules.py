"""
Populate SharePoint Style Rules list from Control Centre Writing Style Guide
"""
import os
import msal
import requests
import json

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

def create_style_rules():
    """
    Define all style rules extracted from the Writing Style Guide PDF

    Rule structure:
    - Title: Rule name/description
    - RuleType: Font, Grammar, Punctuation, Capitalisation, Language, Layout
    - DocumentType: Word, Visio, or Both
    - CheckValue: What to check (e.g., 'AllTextFont', 'Contractions', etc.)
    - ExpectedValue: The expected/correct value
    - AutoFix: True/False - whether this can be auto-fixed
    - Priority: Numeric priority (lower = higher priority)
    """

    rules = [
        # ============================================
        # FONT RULES
        # ============================================
        {
            "Title": "All text must use Arial font",
            "RuleType": "Font",
            "DocumentType": "Word",
            "CheckValue": "AllTextFont",
            "ExpectedValue": "Arial",
            "AutoFix": True,
            "Priority": 1
        },

        # ============================================
        # LANGUAGE RULES - British English
        # ============================================
        {
            "Title": "Use British English spelling - 'colour' not 'color'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_colour",
            "ExpectedValue": "colour",
            "AutoFix": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'aluminium' not 'aluminum'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_aluminium",
            "ExpectedValue": "aluminium",
            "AutoFix": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'analyse' not 'analyze'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_analyse",
            "ExpectedValue": "analyse",
            "AutoFix": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'centre' not 'center'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_centre",
            "ExpectedValue": "centre",
            "AutoFix": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'licence' (noun) not 'license'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_licence",
            "ExpectedValue": "licence",
            "AutoFix": False,  # Context dependent (noun vs verb)
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'organise' not 'organize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_organise",
            "ExpectedValue": "organise",
            "AutoFix": True,
            "Priority": 10
        },

        # ============================================
        # GRAMMAR RULES - Contractions
        # ============================================
        {
            "Title": "No contractions in formal text - use 'cannot' not 'can't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_cant",
            "ExpectedValue": "cannot",
            "AutoFix": True,
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'do not' not 'don't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_dont",
            "ExpectedValue": "do not",
            "AutoFix": True,
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'is not' not 'isn't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_isnt",
            "ExpectedValue": "is not",
            "AutoFix": True,
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'will not' not 'won't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_wont",
            "ExpectedValue": "will not",
            "AutoFix": True,
            "Priority": 15
        },

        # ============================================
        # PUNCTUATION RULES - Date/Time
        # ============================================
        {
            "Title": "Date format in text: DD MONTH YEAR (e.g., 01 February 2015)",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "DateFormat_Text",
            "ExpectedValue": "DD MONTH YYYY",
            "AutoFix": False,
            "Priority": 20
        },
        {
            "Title": "Time format: 24-hour with colon (e.g., 09:00, 18:25)",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "TimeFormat",
            "ExpectedValue": "HH:MM",
            "AutoFix": False,
            "Priority": 20
        },

        # ============================================
        # PUNCTUATION RULES - Numbers
        # ============================================
        {
            "Title": "Numbers below 10 should be spelled out in text",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "NumbersBelowTen",
            "ExpectedValue": "Spelled",
            "AutoFix": False,  # Context dependent
            "Priority": 25
        },
        {
            "Title": "Use commas with numbers of 4+ digits (e.g., 1,000)",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "NumberCommas",
            "ExpectedValue": "WithCommas",
            "AutoFix": True,
            "Priority": 25
        },

        # ============================================
        # CAPITALISATION RULES
        # ============================================
        {
            "Title": "Section titles should be capitalised",
            "RuleType": "Capitalisation",
            "DocumentType": "Word",
            "CheckValue": "SectionTitles",
            "ExpectedValue": "Capitalised",
            "AutoFix": True,
            "Priority": 30
        },
        {
            "Title": "Subsidiary headings: only first letter and proper nouns capitalised",
            "RuleType": "Capitalisation",
            "DocumentType": "Word",
            "CheckValue": "SubsidiaryHeadings",
            "ExpectedValue": "SentenceCase",
            "AutoFix": False,
            "Priority": 30
        },
        {
            "Title": "Job titles only capitalised when with person's name",
            "RuleType": "Capitalisation",
            "DocumentType": "Word",
            "CheckValue": "JobTitles",
            "ExpectedValue": "ContextDependent",
            "AutoFix": False,
            "Priority": 30
        },
        {
            "Title": "Do not capitalise for emphasis",
            "RuleType": "Capitalisation",
            "DocumentType": "Word",
            "CheckValue": "NoEmphasisCaps",
            "ExpectedValue": "NoCapsForEmphasis",
            "AutoFix": False,
            "Priority": 30
        },

        # ============================================
        # LANGUAGE RULES - Word Choice
        # ============================================
        {
            "Title": "Use 'toward' not 'towards'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "Word_toward",
            "ExpectedValue": "toward",
            "AutoFix": True,
            "Priority": 35
        },
        {
            "Title": "Avoid 'etc.' - be specific instead",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "AvoidEtc",
            "ExpectedValue": "NoEtc",
            "AutoFix": False,
            "Priority": 35
        },
        {
            "Title": "Use 'will', 'must', 'shall' instead of 'should' or 'could'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "AvoidShould",
            "ExpectedValue": "will/must/shall",
            "AutoFix": False,
            "Priority": 35
        },
        {
            "Title": "Use metric units where possible",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "PreferMetric",
            "ExpectedValue": "Metric",
            "AutoFix": False,
            "Priority": 40
        },

        # ============================================
        # HYPHENATION RULES
        # ============================================
        {
            "Title": "Use hyphens with suffix '-wide' (e.g., site-wide)",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "Hyphen_wide",
            "ExpectedValue": "Hyphenated",
            "AutoFix": True,
            "Priority": 45
        },
        {
            "Title": "Hyphenate compound modifiers (e.g., 15-page document)",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "CompoundModifiers",
            "ExpectedValue": "Hyphenated",
            "AutoFix": False,
            "Priority": 45
        },

        # ============================================
        # QUOTATION RULES
        # ============================================
        {
            "Title": "Use single quotes for special terms on first reference",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "SpecialTerms",
            "ExpectedValue": "SingleQuotes",
            "AutoFix": False,
            "Priority": 50
        },
        {
            "Title": "Use double quotes for direct speech",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "DirectSpeech",
            "ExpectedValue": "DoubleQuotes",
            "AutoFix": False,
            "Priority": 50
        },

        # ============================================
        # APOSTROPHE RULES
        # ============================================
        {
            "Title": "Never use apostrophes for plurals",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "NoApostrophePlurals",
            "ExpectedValue": "NoApostrophe",
            "AutoFix": True,
            "Priority": 55
        },

        # ============================================
        # SYMBOLS
        # ============================================
        {
            "Title": "Avoid ampersand (&) - use 'and' instead",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "NoAmpersand",
            "ExpectedValue": "and",
            "AutoFix": True,
            "Priority": 60
        },
        {
            "Title": "Spell out 'percent' in text (not %)",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "PercentSymbol",
            "ExpectedValue": "percent",
            "AutoFix": True,
            "Priority": 60
        },

        # ============================================
        # LAYOUT RULES
        # ============================================
        {
            "Title": "Figures and tables must have captions",
            "RuleType": "Layout",
            "DocumentType": "Word",
            "CheckValue": "FigureTableCaptions",
            "ExpectedValue": "Required",
            "AutoFix": False,
            "Priority": 70
        },
    ]

    return rules

def main():
    print("🔧 Populating SharePoint Style Rules list from Writing Style Guide PDF\n")

    # Get access token
    print("🔑 Getting access token...")
    token = get_token()

    # Get site ID
    print("🌐 Getting site ID...")
    site_id = get_site_id(token)
    print(f"   Site ID: {site_id}\n")

    # Get all rules
    rules = create_style_rules()
    print(f"📋 Found {len(rules)} style rules to add\n")

    # Add each rule
    added_count = 0
    failed_count = 0

    for i, rule in enumerate(rules, 1):
        try:
            print(f"[{i}/{len(rules)}] Adding: {rule['Title']}")
            add_rule(token, site_id, rule)
            added_count += 1
            print(f"   ✓ Success")
        except Exception as e:
            failed_count += 1
            print(f"   ✗ Failed: {str(e)}")
        print()

    # Summary
    print("=" * 60)
    print(f"✅ Successfully added: {added_count}")
    print(f"❌ Failed: {failed_count}")
    print(f"📊 Total rules: {len(rules)}")
    print("=" * 60)

if __name__ == "__main__":
    main()
