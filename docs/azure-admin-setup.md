# MaceStyle Validator - Azure Admin Setup Guide

**Prepared for:** Azure Global Admin
**Date:** March 2026
**Contact:** Stephen Cummins (stephen.cummins@macegroup.com)

---

## What This App Does

MaceStyle is an automated document style checker for the Mace Control Centre. When a user uploads a document to a designated SharePoint document library, the system automatically validates it against the Mace Writing Style Guide — checking British English spelling, grammar, font consistency, and formatting — then uploads a corrected file and an HTML report back to SharePoint.

**Supported formats:** Word (.docx, .doc, .docm), Excel (.xlsx, .xls, .xlsm), PowerPoint (.pptx, .ppt, .pptm), Visio (.vsdx, .vsd).

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
| `Sites.Selected` | Application | Access only explicitly granted SharePoint sites (none by default) |

This requires **admin consent granted**.

> **Why `Sites.Selected` instead of `Sites.ReadWrite.All`?**
> `Sites.ReadWrite.All` grants access to **every** SharePoint site in the tenant. `Sites.Selected` grants access to **no sites** by default — access is then explicitly granted per-site via the Graph API. This follows the principle of least privilege and ensures the app can only read/write the designated Style Validation site.

**After the App Registration is created**, per-site access must be granted using the Graph API (see [Grant Per-Site Access](#grant-per-site-access) below).

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

## Grant Per-Site Access

Once the App Registration exists and `Sites.Selected` has been admin-consented, you must grant the app access to the specific SharePoint site. This requires a one-time Graph API call made by an admin with `Sites.FullControl.All` (or a Global Admin using PowerShell).

**Option A: Use the included helper script**

```bash
cd MaceStyleValidator
python3 grant_site_permissions.py
```

This script uses the app's own credentials (from environment variables) plus an admin token to grant `write` access on the target site.

**Option B: PowerShell (run by a Global Admin)**

```powershell
# Install module if needed
Install-Module Microsoft.Graph -Scope CurrentUser

# Connect as admin
Connect-MgGraph -Scopes "Sites.FullControl.All"

# Get the site ID
$site = Get-MgSite -Search "StyleValidation"

# Grant the app write access to this site only
New-MgSitePermission -SiteId $site.Id -BodyParameter @{
    roles = @("write")
    grantedToIdentities = @(@{
        application = @{
            id = "<MaceStyleValidator-App Client ID>"
            displayName = "MaceStyleValidator-App"
        }
    })
}
```

**Option C: Graph API (via curl or Postman)**

```http
POST https://graph.microsoft.com/v1.0/sites/{siteId}/permissions
Authorization: Bearer <admin-token-with-Sites.FullControl.All>
Content-Type: application/json

{
    "roles": ["write"],
    "grantedToIdentities": [{
        "application": {
            "id": "<MaceStyleValidator-App Client ID>",
            "displayName": "MaceStyleValidator-App"
        }
    }]
}
```

Available roles: `read`, `write`, `owner`, `fullcontrol`. The `write` role is sufficient for MaceStyle (read files, write corrected files, update list items).

> **Note:** The admin who makes this grant needs `Sites.FullControl.All` permission, but MaceStyle itself only ever receives `Sites.Selected` with `write` on the single site. The admin permission is only needed for this one-time setup step.

---

## Security Summary

- **Site-scoped access** — uses `Sites.Selected` permission, granting access to only the designated SharePoint site (not the entire tenant)
- **No user impersonation** — the app authenticates as a service principal only
- **No data persistence** — documents are processed in-memory and not stored outside SharePoint
- **Secrets** are stored in Azure Function App Settings (encrypted at rest), never in code
- **Audit trail** maintained in a SharePoint list (Validation Results)
- **External API call** — for Word documents only, text content is sent to Anthropic's Claude AI API (hosted in the US) for language correction. No files are stored by Anthropic. Excel, PowerPoint, and Visio use rule-based validation only (no AI). If data residency is a concern, AI can be disabled entirely.

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

A Logic App (deployed via ARM template) connects the SharePoint document library to the Azure Function. The ARM template is included in the repo at `infra/logic-app.json`. An Azure DevOps CI/CD pipeline (`azure-pipelines.yml`) automates future deployments.

---

## Summary Checklist

For the Azure Admin:

- [ ] Create App Registration (`MaceStyleValidator-App`) in Entra ID
- [ ] Create client secret and share Tenant ID, Client ID, and Secret securely
- [ ] Grant admin consent for `Sites.Selected` (Application permission)
- [ ] Create Azure Function App (`func-mace-validator-prod`, Python 3.11, Linux, Consumption)
- [ ] Create Resource Group `rg-macestyle`
- [ ] Create SharePoint site `/sites/StyleValidation` (or confirm an existing site to use)
- [ ] Grant the App Registration `write` access to the SharePoint site (see [Grant Per-Site Access](#grant-per-site-access))

Everything else (function deployment, Logic App deployment, SharePoint list setup, rule configuration) will be handled by me (Stephen Cummins).
