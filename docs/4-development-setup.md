# Mace Style Validator - Development Setup

## IDE

**Recommended:** Visual Studio Code

### Required Extensions

- **Azure Functions** (`ms-azuretools.vscode-azurefunctions`) — run/debug/deploy Azure Functions locally
- **Python** (`ms-python.python`) — IntelliSense, linting, debugging
- **Pylance** (`ms-python.vscode-pylance`) — type checking and auto-imports

### Optional Extensions

- **Azure Account** (`ms-vscode.azure-account`) — sign in to Azure from VS Code
- **REST Client** (`humao.rest-client`) — test API endpoints without Postman
- **GitLens** (`eamodio.gitlens`) — inline blame, history

## Local Environment

### Prerequisites

- Python 3.11+
- [Azure Functions Core Tools v4](https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local)
- Git

### Setup

```bash
cd MaceStyleValidator
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Configuration

Create `local.settings.json` (gitignored):

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "python",
    "SHAREPOINT_TENANT_ID": "<your-tenant-id>",
    "SHAREPOINT_CLIENT_ID": "<your-client-id>",
    "SHAREPOINT_CLIENT_SECRET": "<your-client-secret>",
    "SHAREPOINT_SITE_URL": "https://<tenant>.sharepoint.com/sites/StyleValidation",
    "SHAREPOINT_DOC_LIBRARY_ID": "<library-guid>",
    "SHAREPOINT_VALIDATION_RESULTS_ID": "<results-list-guid>",
    "ANTHROPIC_API_KEY": "<your-key> (optional — AI currently disabled)"
  }
}
```

### Running Locally

```bash
func start
```

Endpoints:
- `POST http://localhost:7071/api/ValidateDocument`
- `GET  http://localhost:7071/api/TestSharePoint`
- `GET  http://localhost:7071/api/ListDocuments`

### Testing

```bash
python3 create_test_document.py   # Generate a .docx with style violations
python3 test_local.py             # Run local validation tests
```

## VS Code Debug Configuration

Add to `.vscode/launch.json`:

```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "name": "Attach to Python Functions",
      "type": "python",
      "request": "attach",
      "port": 9091,
      "preLaunchTask": "func: host start"
    }
  ]
}
```

## Deployment

```bash
az login
az account set --subscription "<subscription-name>"
func azure functionapp publish func-mace-validator-prod
```

See [3-configuration-setup.md](3-configuration-setup.md) for full deployment steps.

---

**Last Updated:** March 2026
