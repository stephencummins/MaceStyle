# MaceStyle - Mace Style Validator

Automated document validation system enforcing the Mace Control Centre Writing Style Guide using Azure Functions, Claude AI, and SharePoint integration.

## Architecture

- **Azure Function** (Python 3.11) triggered via Power Automate on SharePoint document upload
- **SharePoint Online** stores documents, style rules (`Style Rules` list), and validation results (`Validation Results` list)
- **Claude AI** (Haiku 4.5) for intelligent language corrections (rules with `UseAI: Yes`)
- **Microsoft Graph API** for SharePoint integration, authenticated via Azure AD (MSAL)

## Security

- **Site-scoped permissions**: Uses `Sites.Selected` (not tenant-wide `Sites.ReadWrite.All`). Per-site `write` access granted via Graph API. See `grant_site_permissions.py`.
- **Endpoint auth**: All HTTP routes require a function key (`auth_level=FUNCTION`), except MaceyBot (ANONYMOUS, required for Teams webhook). Keys managed in Azure Portal > App keys.
- **Credentials**: All Azure AD identifiers and secrets are env-var only. No hardcoded defaults anywhere.
- **Data classification**: Documents >50K chars trigger a warning log before being sent to the external AI service. Ensure document classification permits external processing.
- **Dependencies**: Critical packages pinned to exact versions (`anthropic==0.84.0`, `openpyxl==3.1.5`, `python-pptx==1.0.2`).

## Project Structure

```
MaceStyleValidator/
  function_app.py              # Azure Function entry point (5 routes: ValidateDocument, TestSharePoint, ListDocuments, HealthCheck [FUNCTION auth], MaceyBot [ANONYMOUS])
  ValidateDocument/
    __init__.py                # Main routing and orchestration (v5.1.0-governed)
    config.py                  # Shared auth (get_graph_token, get_site_id, get_site_info), constants, Claude model config
    ai_client.py               # Centralised Claude API client (with data classification warning and token tracking)
    access_control.py          # SOC 2 CC6.1 — API key / Azure AD access control
    monitoring.py              # SOC 2 CC7.2 — Structured audit logging, metrics, health checks
    sharepoint_client.py       # SharePoint/Graph API operations
    report.py                  # HTML validation report generation
    word_validator.py          # Word (.docx) validation
    visio_validator.py         # Visio (.vsdx) validation
    excel_validator.py         # Excel (.xlsx) validation
    powerpoint_validator.py    # PowerPoint (.pptx) validation
    enhanced_validators.py     # Hard-coded text rules (spelling, contractions, symbols)
    sharepoint_results.py      # Validation results list operations
    test_helpers.py            # Test utilities
  MaceyBot/                    # Teams bot (Claude-powered site creation assistant)
  governance_check.py          # AGT v2.1.0 compliance verification script
  populate_style_rules.py      # Seed style rules to SharePoint
  add_process_map_rules.py     # Add Visio process map rules
  add_structural_rules.py      # Add Visio structural rules
  create_test_document.py      # Generate test .docx with violations
  requirements.txt             # Python deps (pinned)
docs/                          # Architecture, user guide, config/setup docs
test_files/                    # Test documents and guides
```

All utility scripts import auth from `ValidateDocument.config` (no duplicated MSAL boilerplate).

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
- `SHAREPOINT_DOC_LIBRARY_ID` - SharePoint document library list GUID
- `SHAREPOINT_VALIDATION_RESULTS_ID` - SharePoint validation results list GUID
- `ANTHROPIC_API_KEY`

Optional (MaceyBot):
- `MACEY_MODEL` - Override Claude model for MaceyBot (default: `claude-sonnet-4-20250514`)
- `BOT_TENANT_ID` - Azure AD tenant for Teams bot channel auth
- `SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET` - Alternative credential names for MaceyBot

Optional (governance):
- `MACESTYLE_API_KEY` — Static API key for access control (required if auth mode is `api_key`)
- `MACESTYLE_AUTH_MODE` — `api_key` (default), `azure_ad`, or `none` (dev only)
- `MACESTYLE_ALLOWED_APPS` — Comma-separated Azure AD app IDs allowed to call this function

## Development

```bash
cd MaceStyleValidator
pip install -r requirements.txt
func start                    # Run locally with Azure Functions Core Tools v4
python3 create_test_document.py  # Generate test document
python3 test_local.py         # Run local tests
```

## Governance (Microsoft Agent Governance Toolkit v2.1.0)

Checked and verified on 22 March 2026. All frameworks COMPLIANT.

- **SOC 2 CC6.1** — `access_control.py`: API key or Azure AD bearer token validation on every request. Caller identity extraction for audit. Configurable via `MACESTYLE_AUTH_MODE`.
- **SOC 2 CC7.2** — `monitoring.py`: Structured JSON audit events emitted to Application Insights. Per-request metrics (duration, Claude tokens, cost estimate, SharePoint calls). `/api/HealthCheck` endpoint. Alert emission for failures.
- **GDPR** — No PII processed. Data minimisation confirmed.
- **ISO 27001** — Compliant.

Run the check: `python3 governance_check.py` (requires `.venv` with `agent-governance-toolkit`).

## Notes

- `local.settings.json` and `.env` are gitignored - never commit secrets
- All SharePoint credentials and list IDs must be set via environment variables (no hardcoded defaults)
- Claude model and settings configured in `config.py` (currently Haiku 4.5, 8192 max tokens)
- Visio font detection uses numeric IDs (font 0 = Arial)
- Large files (>100 pages) may timeout on Azure Functions consumption plan
