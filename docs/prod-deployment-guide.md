# MaceStyle Validator — Production Deployment Guide

What the Azure Admin needs to set up to transfer the MaceStyle Validator to the production tenant.

---

## 1. Azure Resources

- **Resource Group** for the solution
- **Azure Function App** (Python 3.11, Consumption or App Service plan, UK South region)
- **Storage Account** (created automatically with the Function App)

## 2. Azure AD App Registration

- **New App Registration** in the prod tenant for SharePoint access
- **API Permissions** (Application, not Delegated):
  - `Sites.Selected` (site-scoped access — no tenant-wide permissions)
- **Admin consent** granted for this permission
- **Per-site access** granted via Graph API (see [azure-admin-setup.md](azure-admin-setup.md#grant-per-site-access))
- **Client Secret** generated (note the expiry)
- Provide Stephen with the **Tenant ID**, **Client ID**, and **Client Secret**

## 3. SharePoint Site

A **SharePoint site** (e.g. `/sites/StyleValidation`) with:

- **Document Library** for uploading documents to be validated
- **Style Rules** list with columns:
  - Title (text)
  - RuleType (text)
  - DocumentType (text)
  - CheckValue (text)
  - ExpectedValue (text)
  - AutoFix (yes/no)
  - Priority (number)
  - UseAI (yes/no)
- **Validation Results** list with columns:
  - Title (text)
  - Status (text)
  - IssuesFound (number)
  - IssuesFixed (number)
  - ReportHTML (multiple lines of text)
  - ReportUrl (hyperlink)
- **Custom columns on the Document Library**:
  - ValidationStatus (Choice: Not Validated, Validate Now, Validating..., Passed, Review Required, Failed)
  - LastValidated (Date and Time)
  - ValidationReport (Hyperlink)
  - ValidationResultLink (Hyperlink)

## 4. Function App Settings (Environment Variables)

Add these in the Function App Configuration > Application Settings:

| Setting | Example Value | Required |
|---------|--------------|----------|
| `SHAREPOINT_TENANT_ID` | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` | Yes |
| `SHAREPOINT_CLIENT_ID` | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` | Yes |
| `SHAREPOINT_CLIENT_SECRET` | *(from App Registration)* | Yes |
| `SHAREPOINT_SITE_URL` | `https://macegroup.sharepoint.com/sites/StyleValidation` | Yes |
| `ANTHROPIC_API_KEY` | *(Claude API key)* | Yes |
| `SHAREPOINT_DOC_LIBRARY_ID` | *(GUID of the Document Library list)* | Optional* |
| `SHAREPOINT_VALIDATION_RESULTS_ID` | *(GUID of the Validation Results list)* | Optional* |

*List IDs default to the dev tenant values. For production, get the list GUIDs from SharePoint (Site Settings > Site Contents > list settings URL contains the GUID).

## 5. Logic App (replaces Power Automate)

An ARM template is provided at `infra/logic-app.json` that deploys a **Consumption Logic App** replicating the Power Automate flow. This is preferable for production because it lives in the Azure subscription alongside the Function App and can be deployed via ARM/DevOps.

The Logic App:

1. **Triggers** when a document is created or modified in the SharePoint Document Library
2. **Filters** to supported file types (.docx, .xlsx, .pptx, .vsdx, etc.)
3. **Sets** ValidationStatus to "Validating..."
4. **Gets file content** and base64-encodes it
5. **Calls the Function App** (`POST /api/ValidateDocument`) with JSON body:
   ```json
   {
     "itemId": "<SharePoint item ID>",
     "fileName": "<file name with extension>",
     "fileContent": "<base64 encoded file content>",
     "fileUrl": "<server-relative path>"
   }
   ```
6. **Parses the JSON response** (fields: `status`, `description`, `issuesFound`, `issuesFixed`, `reportUrl`, `fixedFileContent`)
7. If fixes were applied, **uploads the corrected file** back to SharePoint
8. **Updates** ValidationStatus, Description, ValidationReport, ValidationResultLink, and LastValidated columns

### Deploying the Logic App

```bash
az deployment group create \
  --resource-group rg-mace-style-validator \
  --template-file infra/logic-app.json \
  --parameters infra/logic-app.parameters.json \
  --parameters functionAppKey="<function-host-key>"
```

Fill in `functionAppKey` and `sharepointDocLibraryId` in the parameters file (or pass on the command line). After deployment, open the Logic App in the Azure Portal and **authorise the SharePoint connection** (API Connections > sharepointonline > Edit > Authorize).

## 6. Azure DevOps CI/CD Pipeline

An `azure-pipelines.yml` is provided in the repo root. It automates:

- **Build**: Install deps, run tests, create deployment zip
- **Deploy to Dev**: Deploys Function App to `func-mace-validator-dev`
- **Deploy to Prod**: Deploys Function App to `func-mace-validator-prod` (requires approval)
- **Deploy Logic App**: Deploys the ARM template to the resource group

### Setup in Azure DevOps

1. Create a **Service Connection** (type: Azure Resource Manager) named `MaceStyle-ServiceConnection` with access to the target subscription
2. Create **Environments** named `dev` and `production` in Azure DevOps Pipelines > Environments
3. Add an **Approval gate** on the `production` environment
4. Import the pipeline from `azure-pipelines.yml`
5. Update the variables at the top of the YAML if resource names differ

Pushes to `main` trigger the full pipeline: build → dev → prod (with approval).

## 7. Network / Firewall

The Function App needs **outbound HTTPS access** to:

| Endpoint | Purpose |
|----------|---------|
| `graph.microsoft.com` | SharePoint / Microsoft Graph API |
| `api.anthropic.com` | Claude AI for style corrections |
| `login.microsoftonline.com` | Azure AD authentication |

## 8. Supported File Types

The validator handles:

- **Word**: .docx, .doc, .docm, .dotx, .dotm
- **Excel**: .xlsx, .xls, .xlsm
- **PowerPoint**: .pptx, .ppt, .pptm, .potx, .potm
- **Visio**: .vsdx, .vsd

## What Stephen Will Provide

- The **code** (GitHub repo — includes Function App, Logic App ARM template, and DevOps pipeline)
- The `populate_style_rules.py` script to seed the Style Rules list with all validation rules
- ARM template + parameters for the Logic App (`infra/logic-app.json`)
- Azure DevOps pipeline definition (`azure-pipelines.yml`)

---

*Contact: stephen.cummins@macegroup.com*
