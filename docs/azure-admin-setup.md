# MaceStyle Validator - Azure Admin Setup Guide

**Prepared for:** Azure Global Admin
**Date:** March 2026
**Contact:** Stephen Cummins (stephencummins@gmail.com)

---

## What This App Does

MaceStyle is an automated document style checker for the Mace Control Centre. When a user uploads a Word (.docx) or Visio (.vsdx) file to a designated SharePoint document library, the system automatically validates it against the Mace Writing Style Guide — checking British English spelling, grammar, font consistency, and formatting — then uploads a corrected file and an HTML report back to SharePoint.

The validation takes 5-15 seconds per document and runs serverlessly with no ongoing infrastructure management.

---

## What We Need Configured

### 1. App Registration (Entra ID)

| Setting | Value |
|---------|-------|
| Name | `MaceStyleValidator-App` |
| Account type | Single tenant |
| Redirect URI | None |

**Client secret** required (24-month expiry recommended).

**API Permissions** (Application, not Delegated):

| Permission | Type | Description |
|------------|------|-------------|
| `Sites.ReadWrite.All` | Application | Read/write SharePoint site content |
| `Files.ReadWrite.All` | Application | Read/write files in SharePoint libraries |

Both require **admin consent granted**.

**Values we need back:**
- Tenant ID
- Client ID (Application ID)
- Client Secret value

### 2. Azure Function App

| Setting | Value |
|---------|-------|
| Resource Group | `rg-macestyle` (new) |
| Function App name | `func-mace-validator-prod` |
| Runtime | Python 3.11 |
| OS | Linux |
| Plan | Consumption (Serverless) |
| Region | UK South (closest to SharePoint tenant) |

**Application Settings** (we will configure these once the App Registration is created):

| Name | Value |
|------|-------|
| `SHAREPOINT_TENANT_ID` | From App Registration |
| `SHAREPOINT_CLIENT_ID` | From App Registration |
| `SHAREPOINT_CLIENT_SECRET` | From App Registration |
| `SHAREPOINT_SITE_URL` | `https://[tenant].sharepoint.com/sites/StyleValidation` |
| `ANTHROPIC_API_KEY` | Provided separately (third-party AI API key) |

### 3. SharePoint Site

| Setting | Value |
|---------|-------|
| Site type | Team site |
| Site name | `Style Validation` |
| URL path | `/sites/StyleValidation` |

We will handle list/library creation and rule population once the site exists.

---

## Security Summary

- **No user impersonation** - the app authenticates as a service principal only
- **No data persistence** - documents are processed in-memory and not stored outside SharePoint
- **Secrets** are stored in Azure Function App Settings (encrypted at rest), never in code
- **Audit trail** maintained in a SharePoint list (Validation Results)
- **External API call** - text content is sent to Anthropic's Claude AI API (hosted in the US) for language correction. No files are stored by Anthropic. If this raises data residency concerns, we can disable AI rules and rely on rule-based validation only.

---

## Cost Estimate

| Service | Monthly Cost (est. 100 docs/month) |
|---------|-------------------------------------|
| Azure Function (Consumption) | < $5 |
| Application Insights | < $5 |
| Claude AI (Anthropic) | ~$1 |
| SharePoint | Included in M365 |
| **Total** | **~$5-10/month** |

---

## Deployment

Once the above resources are provisioned, we deploy the function code via Azure Functions Core Tools:

```bash
func azure functionapp publish func-mace-validator-prod
```

A Power Automate flow then connects the SharePoint document library to the Azure Function (configured by us, not by the Azure admin).

---

## Summary Checklist

For the Azure Admin:

- [ ] Create App Registration (`MaceStyleValidator-App`) in Entra ID
- [ ] Create client secret and share Tenant ID, Client ID, and Secret securely
- [ ] Grant admin consent for `Sites.ReadWrite.All` and `Files.ReadWrite.All` (Application)
- [ ] Create Azure Function App (`func-mace-validator-prod`, Python 3.11, Linux, Consumption)
- [ ] Create Resource Group `rg-macestyle`
- [ ] Create SharePoint site `/sites/StyleValidation` (or confirm an existing site to use)
- [ ] Grant the App Registration access to the SharePoint site (if site-scoped permissions are preferred over tenant-wide)

Everything else (function deployment, SharePoint list setup, Power Automate flow, rule configuration) will be handled by the development team.
