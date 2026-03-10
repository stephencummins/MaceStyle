# MaceStyle - Mace Style Validator

Automated document validation system enforcing the Mace Control Centre Writing Style Guide using Azure Functions, Claude AI, and SharePoint integration.

## Architecture

- **Azure Function** (Python 3.11) triggered via Power Automate on SharePoint document upload
- **SharePoint Online** stores documents, style rules (`Style Rules` list), and validation results (`Validation Results` list)
- **Claude AI** (Haiku 4.5) for intelligent language corrections (rules with `UseAI: Yes`)
- **Microsoft Graph API** for SharePoint integration, authenticated via Azure AD (MSAL)

## Project Structure

```
MaceStyleValidator/
  function_app.py              # Azure Function entry point (routes: ValidateDocument, TestSharePoint, ListDocuments)
  ValidateDocument/
    __init__.py                # Main routing and orchestration
    config.py                  # Auth, constants, Claude model config
    ai_client.py               # Centralised Claude API client
    sharepoint_client.py       # SharePoint/Graph API operations
    report.py                  # HTML validation report generation
    word_validator.py          # Word (.docx) validation
    visio_validator.py         # Visio (.vsdx) validation
    excel_validator.py         # Excel (.xlsx) validation
    powerpoint_validator.py    # PowerPoint (.pptx) validation
    enhanced_validators.py     # Hard-coded text rules (spelling, contractions, symbols)
    sharepoint_results.py      # Validation results list operations
    test_helpers.py            # Test utilities
  populate_style_rules.py      # Seed style rules to SharePoint
  add_process_map_rules.py     # Add Visio process map rules
  add_structural_rules.py      # Add Visio structural rules
  create_test_document.py      # Generate test .docx with violations
  requirements.txt             # Python deps
docs/                          # Architecture, user guide, config/setup docs
test_files/                    # Test documents and guides
```

## Supported Formats

- **Word** (.docx, .doc) - Full validation + auto-fix + corrected file upload
- **Visio** (.vsdx, .vsd) - Full validation + auto-fix (text, fonts, colours, sizes, positions, page dims)
- **Excel** (.xlsx, .xls) - Full validation + auto-fix (text corrections written back to cells)
- **PowerPoint** (.pptx, .ppt) - Full validation + auto-fix (text, fonts)

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
- All SharePoint credentials must be set via environment variables (no hardcoded defaults)
- Claude model and settings configured in `config.py` (currently Haiku 4.5, 8192 max tokens)
- Visio font detection uses numeric IDs (font 0 = Arial)
- Large files (>100 pages) may timeout on Azure Functions consumption plan
