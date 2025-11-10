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
    - UseAI: True/False - whether to use AI (Claude) for this rule
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
            "UseAI": False,  # Use hard-coded validation for fonts
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
            "UseAI": True,  # Use Claude for comprehensive British spelling
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'aluminium' not 'aluminum'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_aluminium",
            "ExpectedValue": "aluminium",
            "AutoFix": True,
            "UseAI": True,  # Use Claude for comprehensive British spelling
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'analyse' not 'analyze'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_analyse",
            "ExpectedValue": "analyse",
            "AutoFix": True,
            "UseAI": True,  # Use Claude for comprehensive British spelling
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'centre' not 'center'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_centre",
            "ExpectedValue": "centre",
            "AutoFix": True,
            "UseAI": True,  # Use Claude for comprehensive British spelling
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'licence' (noun) not 'license'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_licence",
            "ExpectedValue": "licence",
            "AutoFix": False,  # Context dependent (noun vs verb)
            "UseAI": True,  # Use Claude for comprehensive British spelling
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'organise' not 'organize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_organise",
            "ExpectedValue": "organise",
            "AutoFix": True,
            "UseAI": True,  # Use Claude for comprehensive British spelling
            "Priority": 10
        },
        # Additional British English spellings from PDF page 21
        {
            "Title": "Use British English spelling - 'analogue' not 'analog'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_analogue",
            "ExpectedValue": "analogue",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'authorise' not 'authorize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_authorise",
            "ExpectedValue": "authorise",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'calibre' not 'caliber'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_calibre",
            "ExpectedValue": "calibre",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'catalogue' not 'catalog'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_catalogue",
            "ExpectedValue": "catalogue",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'characterise' not 'characterize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_characterise",
            "ExpectedValue": "characterise",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'defence' not 'defense'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_defence",
            "ExpectedValue": "defence",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'finalised' not 'finalized'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_finalised",
            "ExpectedValue": "finalised",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'dialogue' not 'dialog'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_dialogue",
            "ExpectedValue": "dialogue",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'fibre' not 'fiber'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_fibre",
            "ExpectedValue": "fibre",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'grey' not 'gray'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_grey",
            "ExpectedValue": "grey",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'harbour' not 'harbor'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_harbour",
            "ExpectedValue": "harbour",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'labour' not 'labor'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_labour",
            "ExpectedValue": "labour",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'learnt' not 'learned' (past tense)",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_learnt",
            "ExpectedValue": "learnt",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'litre' not 'liter'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_litre",
            "ExpectedValue": "litre",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'manoeuvre' not 'maneuver'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_manoeuvre",
            "ExpectedValue": "manoeuvre",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'maximise' not 'maximize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_maximise",
            "ExpectedValue": "maximise",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'metre' not 'meter'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_metre",
            "ExpectedValue": "metre",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'minimise' not 'minimize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_minimise",
            "ExpectedValue": "minimise",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'mobilise' not 'mobilize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_mobilise",
            "ExpectedValue": "mobilise",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'modelling' not 'modeling'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_modelling",
            "ExpectedValue": "modelling",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'neighbour' not 'neighbor'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_neighbour",
            "ExpectedValue": "neighbour",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'neutralise' not 'neutralize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_neutralise",
            "ExpectedValue": "neutralise",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'normalise' not 'normalize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_normalise",
            "ExpectedValue": "normalise",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'optimise' not 'optimize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_optimise",
            "ExpectedValue": "optimise",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'programme' not 'program'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_programme",
            "ExpectedValue": "programme",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'realise' not 'realize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_realise",
            "ExpectedValue": "realise",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'skilful' not 'skillful'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_skilful",
            "ExpectedValue": "skilful",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'spelt' not 'spelled' (past tense)",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_spelt",
            "ExpectedValue": "spelt",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'stabilise' not 'stabilize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_stabilise",
            "ExpectedValue": "stabilise",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'summarise' not 'summarize'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_summarise",
            "ExpectedValue": "summarise",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 10
        },
        {
            "Title": "Use British English spelling - 'tunnelling' not 'tunneling'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "BritishSpelling_tunnelling",
            "ExpectedValue": "tunnelling",
            "AutoFix": True,
            "UseAI": True,
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
            "UseAI": True,  # Use Claude to catch all contractions
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'do not' not 'don't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_dont",
            "ExpectedValue": "do not",
            "AutoFix": True,
            "UseAI": True,  # Use Claude to catch all contractions
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'is not' not 'isn't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_isnt",
            "ExpectedValue": "is not",
            "AutoFix": True,
            "UseAI": True,  # Use Claude to catch all contractions
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'will not' not 'won't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_wont",
            "ExpectedValue": "will not",
            "AutoFix": True,
            "UseAI": True,  # Use Claude to catch all contractions
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'could not' not 'couldn't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_couldnt",
            "ExpectedValue": "could not",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'did not' not 'didn't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_didnt",
            "ExpectedValue": "did not",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'does not' not 'doesn't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_doesnt",
            "ExpectedValue": "does not",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'has not' not 'hasn't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_hasnt",
            "ExpectedValue": "has not",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'have not' not 'haven't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_havent",
            "ExpectedValue": "have not",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'should not' not 'shouldn't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_shouldnt",
            "ExpectedValue": "should not",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 15
        },
        {
            "Title": "No contractions in formal text - use 'would not' not 'wouldn't'",
            "RuleType": "Grammar",
            "DocumentType": "Word",
            "CheckValue": "NoContraction_wouldnt",
            "ExpectedValue": "would not",
            "AutoFix": True,
            "UseAI": True,
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
            "UseAI": False,  # Not implemented yet
            "Priority": 20
        },
        {
            "Title": "Time format: 24-hour with colon (e.g., 09:00, 18:25)",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "TimeFormat",
            "ExpectedValue": "HH:MM",
            "AutoFix": False,
            "UseAI": False,  # Not implemented yet
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
            "UseAI": False,  # Not implemented yet
            "Priority": 25
        },
        {
            "Title": "Use commas with numbers of 4+ digits (e.g., 1,000)",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "NumberCommas",
            "ExpectedValue": "WithCommas",
            "AutoFix": True,
            "UseAI": True,  # Use Claude to catch all number formatting
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
            "UseAI": False,  # Not implemented yet
            "Priority": 30
        },
        {
            "Title": "Subsidiary headings: only first letter and proper nouns capitalised",
            "RuleType": "Capitalisation",
            "DocumentType": "Word",
            "CheckValue": "SubsidiaryHeadings",
            "ExpectedValue": "SentenceCase",
            "AutoFix": False,
            "UseAI": False,  # Not implemented yet
            "Priority": 30
        },
        {
            "Title": "Job titles only capitalised when with person's name",
            "RuleType": "Capitalisation",
            "DocumentType": "Word",
            "CheckValue": "JobTitles",
            "ExpectedValue": "ContextDependent",
            "AutoFix": False,
            "UseAI": False,  # Not implemented yet
            "Priority": 30
        },
        {
            "Title": "Do not capitalise for emphasis",
            "RuleType": "Capitalisation",
            "DocumentType": "Word",
            "CheckValue": "NoEmphasisCaps",
            "ExpectedValue": "NoCapsForEmphasis",
            "AutoFix": False,
            "UseAI": False,  # Not implemented yet
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
            "UseAI": True,  # Use Claude for word choice
            "Priority": 35
        },
        {
            "Title": "Avoid 'etc.' - be specific instead",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "AvoidEtc",
            "ExpectedValue": "NoEtc",
            "AutoFix": False,
            "UseAI": True,  # Use Claude for word choice
            "Priority": 35
        },
        {
            "Title": "Use 'will', 'must', 'shall' instead of 'should' or 'could'",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "AvoidShould",
            "ExpectedValue": "will/must/shall",
            "AutoFix": False,
            "UseAI": False,  # Too context dependent, not implemented
            "Priority": 35
        },
        {
            "Title": "Use metric units where possible",
            "RuleType": "Language",
            "DocumentType": "Word",
            "CheckValue": "PreferMetric",
            "ExpectedValue": "Metric",
            "AutoFix": False,
            "UseAI": False,  # Not implemented yet
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
            "UseAI": False,  # Not implemented yet
            "Priority": 45
        },
        {
            "Title": "Hyphenate compound modifiers (e.g., 15-page document)",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "CompoundModifiers",
            "ExpectedValue": "Hyphenated",
            "AutoFix": False,
            "UseAI": False,  # Not implemented yet
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
            "UseAI": False,  # Not implemented yet
            "Priority": 50
        },
        {
            "Title": "Use double quotes for direct speech",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "DirectSpeech",
            "ExpectedValue": "DoubleQuotes",
            "AutoFix": False,
            "UseAI": False,  # Not implemented yet
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
            "UseAI": True,  # Use Claude to catch apostrophe errors
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
            "UseAI": True,  # Use Claude to replace ampersands
            "Priority": 60
        },
        {
            "Title": "Spell out 'percent' in text (not %)",
            "RuleType": "Punctuation",
            "DocumentType": "Word",
            "CheckValue": "PercentSymbol",
            "ExpectedValue": "percent",
            "AutoFix": True,
            "UseAI": True,  # Use Claude to replace percent symbols
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
            "UseAI": False,  # Not implemented yet
            "Priority": 70
        },

        # ============================================
        # VISIO-SPECIFIC RULES (from PPM-GLO-DAE-DMT-PRO-001 analysis)
        # ============================================
        {
            "Title": "Reference codes must be uppercase (e.g., DMT-ACT-001)",
            "RuleType": "Capitalisation",
            "DocumentType": "Visio",
            "CheckValue": "ReferenceCodeCase",
            "ExpectedValue": "UPPERCASE",
            "AutoFix": True,
            "UseAI": True,  # Use AI to detect and fix reference codes
            "Priority": 75
        },
        {
            "Title": "Avoid ampersand (&) in Visio diagrams - use 'and' instead",
            "RuleType": "Punctuation",
            "DocumentType": "Visio",
            "CheckValue": "NoAmpersand",
            "ExpectedValue": "and",
            "AutoFix": True,
            "UseAI": True,  # Use Claude to replace ampersands
            "Priority": 76
        },
        {
            "Title": "British spelling: 'Programme' not 'Program' (in UK/Mace context)",
            "RuleType": "Language",
            "DocumentType": "Both",
            "CheckValue": "BritishSpelling_programme",
            "ExpectedValue": "Programme",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 11
        },
        {
            "Title": "Use 'Constructability' not 'Constructibility'",
            "RuleType": "Language",
            "DocumentType": "Both",
            "CheckValue": "Constructability",
            "ExpectedValue": "Constructability",
            "AutoFix": True,
            "UseAI": True,
            "Priority": 77
        },
        {
            "Title": "Project phases: Use G0-G1, G2-G4, G5 notation consistently",
            "RuleType": "Language",
            "DocumentType": "Visio",
            "CheckValue": "PhaseNotation",
            "ExpectedValue": "G0-G1",
            "AutoFix": False,
            "UseAI": False,  # Formatting rule
            "Priority": 79
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
