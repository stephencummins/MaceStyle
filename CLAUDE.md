# MaceStyle - Mace Style Validator

Automated document validation system enforcing the Mace Control Centre Writing Style Guide using Azure Functions, Claude AI, and SharePoint integration.

## Architecture

- **Azure Function** (Python 3.11) triggered via Power Automate on SharePoint document upload
- **SharePoint Online** stores documents, style rules (`Style Rules` list), and validation results (`Validation Results` list)
- **Claude AI** (Haiku) for intelligent language corrections (rules with `UseAI: Yes`)
- **Microsoft Graph API** for SharePoint integration, authenticated via Azure AD (MSAL)

## Project Structure

```
MaceStyleValidator/
  function_app.py              # Azure Function entry point (routes: ValidateDocument, TestSharePoint)
  ValidateDocument/
    __init__.py                # Main validation logic (large - ~67KB)
    claude_validator.py        # Claude AI validation integration
    enhanced_validators.py     # Additional validation rules
    sharepoint_results.py      # SharePoint results upload
    test_helpers.py            # Test utilities
  populate_style_rules.py      # Seed style rules to SharePoint
  add_process_map_rules.py     # Add Visio process map rules
  add_structural_rules.py      # Add Visio structural rules
  create_test_document.py      # Generate test .docx with violations
  requirements.txt             # Python deps
docs/                          # Architecture, user guide, config/setup docs
test_files/                    # Test documents and guides
```

## Key Files

- `ValidateDocument/__init__.py` - Core validation engine. Handles Word (.docx) and Visio (.vsdx) files. Fetches rules from SharePoint, applies fixes, generates HTML reports, uploads corrected files.
- `ValidateDocument/claude_validator.py` - Sends text to Claude API for style corrections.
- `populate_style_rules.py` - Populates the SharePoint `Style Rules` list with all validation rules.

## Validation Rules

Rules cover: British English spelling, contraction expansion, symbol replacement (& -> "and", % -> "percent"), number formatting (comma-separated thousands), font standardisation (Arial).

## Environment Variables

Required in Azure Function App Settings or `local.settings.json`:
- `SHAREPOINT_TENANT_ID`, `SHAREPOINT_CLIENT_ID`, `SHAREPOINT_CLIENT_SECRET`
- `SHAREPOINT_SITE_URL` (e.g. `https://tenant.sharepoint.com/sites/StyleValidation`)
- `ANTHROPIC_API_KEY`

## Development

```bash
cd MaceStyleValidator
pip install -r requirements.txt
func start                    # Run locally with Azure Functions Core Tools v4
python3 create_test_document.py  # Generate test document
python3 test_local.py         # Run local tests
```

## Notes

- `local.settings.json` and `.env` are gitignored - never commit secrets
- Visio font detection uses numeric IDs (font 0 = Arial)
- Large files (>100 pages) may timeout on Azure Functions consumption plan
- `ValidateDocument/__init__.py` is the largest file and contains most business logic - consider splitting if it grows further
