# CLAUDE.md

## Project Overview

MaceStyle Validator is an automated document validation system that enforces the Mace Control Centre Writing Style Guide. It uses Azure Functions, Claude AI (Anthropic), and SharePoint integration to validate Word (.docx) and Visio (.vsdx) documents for British English compliance, grammar, punctuation, font consistency, and style adherence.

## Tech Stack

- **Runtime:** Python 3.11, Azure Functions v4 (serverless)
- **AI:** Claude 3 Haiku via `anthropic` SDK
- **Document Processing:** `python-docx` (Word), `vsdx` (Visio)
- **Integration:** Microsoft Graph API, MSAL authentication, SharePoint Online
- **Workflow:** Power Automate triggers the Azure Function

## Repository Structure

```
MaceStyle/
├── MaceStyleValidator/                # Main application
│   ├── function_app.py               # Azure Functions entry point
│   ├── host.json                     # Azure Functions config
│   ├── requirements.txt              # Python dependencies
│   ├── ValidateDocument/             # Core validation module
│   │   ├── __init__.py               # Main validation logic & orchestration
│   │   ├── claude_validator.py       # AI-powered validation via Claude
│   │   ├── enhanced_validators.py    # Regex-based validators
│   │   ├── sharepoint_results.py     # SharePoint results integration
│   │   └── test_helpers.py           # Mock objects for local testing
│   ├── populate_style_rules.py       # Populate SharePoint rules list
│   ├── setup_sharepoint.py           # SharePoint site initialization
│   ├── add_process_map_rules.py      # Visio process map rules
│   ├── add_structural_rules.py       # Visio structural rules
│   ├── test_local.py                 # Local validation testing
│   └── create_test_document.py       # Generate test documents
├── docs/                             # Documentation
├── test_files/                       # Test documents
└── [root scripts]                    # PDF/Visio test utilities
```

## Key Files

- `MaceStyleValidator/ValidateDocument/__init__.py` — Main validation pipeline (~1400 lines). Handles document loading, rule fetching, validation orchestration, and HTML report generation.
- `MaceStyleValidator/ValidateDocument/claude_validator.py` — AI validation. Builds prompts from SharePoint rules, calls Claude API, parses JSON corrections.
- `MaceStyleValidator/ValidateDocument/enhanced_validators.py` — Regex-based validators for patterns that don't need AI.
- `MaceStyleValidator/ValidateDocument/sharepoint_results.py` — Posts validation results back to SharePoint.
- `MaceStyleValidator/function_app.py` — Azure Function HTTP trigger entry point.

## Development Setup

```bash
cd MaceStyleValidator
pip install -r requirements.txt
```

### Required Environment Variables

- `SHAREPOINT_TENANT_ID` — Azure AD tenant ID
- `SHAREPOINT_CLIENT_ID` — App registration client ID
- `SHAREPOINT_CLIENT_SECRET` — App registration client secret
- `SHAREPOINT_SITE_URL` — SharePoint site URL
- `ANTHROPIC_API_KEY` — Anthropic API key for Claude

### Local Testing

```bash
cd MaceStyleValidator
python test_local.py
```

This runs validation without SharePoint connectivity using mock rules from `test_helpers.py`.

### Deploy to Azure

```bash
cd MaceStyleValidator
func azure functionapp publish <function-app-name>
```

## Testing

There is no formal test framework (no pytest/unittest). Testing is done through:

- `test_local.py` — Validates documents locally with mock rules
- `create_test_document.py` — Generates Word documents with intentional violations
- Various `test_*.py` and `create_test_*.py` scripts at root level for Visio testing
- Integration testing against live SharePoint environments

## Code Conventions

### Style

- Standard Python 3.11 style
- No linter/formatter configured (no flake8, black, or isort)
- Use `logging` module throughout — `logging.info()`, `logging.error()`, etc.
- Descriptive docstrings on functions
- Imports grouped at top of files

### Error Handling

- Try-catch with logging at all API boundaries
- Graceful degradation: continue processing if one validation step fails
- Detailed error messages with context

### Data Patterns

Validation rules are dictionaries with this shape:
```python
{
    'title': str,
    'rule_type': str,    # Font, Language, Grammar, Punctuation, etc.
    'doc_type': str,     # Word, Visio, Both
    'check_value': str,
    'expected_value': str,
    'auto_fix': bool,
    'use_ai': bool,
    'priority': int
}
```

Validation results return:
```python
{
    'document': Document,
    'issues': list[str],
    'fixes_applied': list[str]
}
```

### Validation Pipeline Flow

1. Receive document via HTTP trigger (Power Automate)
2. Fetch validation rules from SharePoint
3. Load document (Word or Visio)
4. Run regex-based validators (enhanced_validators.py)
5. Run AI validation via Claude (claude_validator.py)
6. Generate HTML report
7. Upload corrected document and report to SharePoint
8. Post results to SharePoint Validation Results list

## Important Notes

- **British English** is enforced throughout (e.g., "finalised" not "finalized")
- Visio font detection uses numeric font IDs (0 = Arial)
- Documents up to ~100 pages are supported; larger files may timeout
- No CI/CD pipeline exists — deployment is manual via Azure Functions Core Tools
- Credentials must never be committed; use environment variables or Azure Key Vault
