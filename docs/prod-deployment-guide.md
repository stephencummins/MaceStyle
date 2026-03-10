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
  - `Sites.ReadWrite.All` (read/write SharePoint sites)
  - `Files.ReadWrite.All` (read/write files)
- **Admin consent** granted for those permissions
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

## 5. Power Automate Flow

A **Power Automate flow** that:

1. **Triggers** when a document is uploaded or modified in the SharePoint Document Library
2. **Gets the file content** (base64 encoded)
3. **Calls the Function App** HTTP endpoint (`POST /api/ValidateDocument`) with JSON body:
   ```json
   {
     "itemId": "<SharePoint item ID>",
     "fileName": "<file name with extension>",
     "fileContent": "<base64 encoded file content>"
   }
   ```
4. **Parses the JSON response** (fields: `status`, `description`, `issuesFound`, `issuesFixed`, `reportUrl`)
5. **Updates the document's ValidationStatus** column using the `status` field from the response

Stephen can export the existing dev flow as a .zip for reference/import.

## 6. Network / Firewall

The Function App needs **outbound HTTPS access** to:

| Endpoint | Purpose |
|----------|---------|
| `graph.microsoft.com` | SharePoint / Microsoft Graph API |
| `api.anthropic.com` | Claude AI for style corrections |
| `login.microsoftonline.com` | Azure AD authentication |

## 7. Supported File Types

The validator handles:

- **Word**: .docx, .doc, .docm, .dotx, .dotm
- **Excel**: .xlsx, .xls, .xlsm
- **PowerPoint**: .pptx, .ppt, .pptm, .potx, .potm
- **Visio**: .vsdx, .vsd

## What Stephen Will Provide

- The **code** (GitHub repo or zip deployment)
- The `populate_style_rules.py` script to seed the Style Rules list with all validation rules
- The Power Automate flow export (.zip)
- Deployment via VS Code Azure Functions extension or `func azure functionapp publish`

---

*Contact: stephen.cummins@macegroup.com*
