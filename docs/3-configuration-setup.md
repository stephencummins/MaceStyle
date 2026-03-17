# Mace Style Validator - Configuration & Setup Guide

## Prerequisites

Before beginning setup, ensure you have:

- [ ] **Azure Subscription** with Owner or Contributor permissions
- [ ] **Microsoft 365 Tenant** with SharePoint Online
- [ ] **Azure subscription** access (for Logic App deployment)
- [ ] **Azure AD** admin access for app registration
- [ ] **Anthropic API Key** (Claude AI) — *optional, AI validation currently disabled*
- [ ] **Development tools**:
  - Python 3.11+
  - Azure Functions Core Tools v4
  - Git
  - VS Code (recommended)

---

## Part 1: Azure Setup

### 1.1 Create Azure App Registration

This app will authenticate the Azure Function to access SharePoint.

1. **Navigate to Azure Portal**
   - Go to: https://portal.azure.com
   - Sign in with admin credentials

2. **Create App Registration**
   - Search for **"Azure Active Directory"** or **"Microsoft Entra ID"**
   - Click **App registrations** → **New registration**

   **Settings:**
   - **Name:** `MaceStyleValidator-App`
   - **Supported account types:** Single tenant
   - **Redirect URI:** Leave blank
   - Click **Register**

3. **Note the Application Details**
   ```
   Tenant ID: [Copy from Overview page]
   Client ID (Application ID): [Copy from Overview page]
   ```

   **Save these values - you'll need them later!**

4. **Create Client Secret**
   - Click **Certificates & secrets** → **New client secret**
   - **Description:** `MaceStyleValidator-Secret`
   - **Expires:** 24 months (or as per your policy)
   - Click **Add**
   - **Copy the secret value immediately** (won't be shown again)

   ```
   Client Secret: [Copy the Value, not the Secret ID]
   ```

5. **Configure API Permissions**
   - Click **API permissions** → **Add a permission**
   - Select **Microsoft Graph** → **Application permissions**

   **Add these permissions:**
   - `Sites.ReadWrite.All` - Read and write items in all site collections
   - `Files.ReadWrite.All` - Read and write files in all site collections

   - Click **Add permissions**
   - Click **Grant admin consent for [Tenant Name]**
   - Confirm by clicking **Yes**

   ✅ **Verify:** Permissions show "Granted for [Tenant]"

---

### 1.2 Create Azure Function App

1. **Create Function App Resource**
   - In Azure Portal, click **Create a resource**
   - Search for **"Function App"**
   - Click **Create**

   **Basic Settings:**
   ```
   Resource Group: Create new → "rg-macestyle"
   Function App name: "func-mace-validator-prod"
   Publish: Code
   Runtime stack: Python
   Version: 3.11
   Region: [Choose closest to your SharePoint tenant]
   Operating System: Linux
   Plan type: Consumption (Serverless)
   ```

   - Click **Review + create** → **Create**
   - Wait for deployment (2-3 minutes)

2. **Configure Application Settings**
   - Go to Function App → **Configuration** → **Application settings**
   - Click **+ New application setting** for each:

   ```
   Name: SHAREPOINT_TENANT_ID
   Value: [Your Tenant ID from step 1.1]

   Name: SHAREPOINT_CLIENT_ID
   Value: [Your Client ID from step 1.1]

   Name: SHAREPOINT_CLIENT_SECRET
   Value: [Your Client Secret from step 1.1]

   Name: SHAREPOINT_SITE_URL
   Value: https://[yourtenant].sharepoint.com/sites/StyleValidation

   Name: ANTHROPIC_API_KEY
   Value: [Your Claude API key - get from console.anthropic.com] (optional — AI currently disabled)
   ```

   - Click **Save** → **Continue**

3. **Configure CORS** (if calling from web app)
   - Go to **CORS** under API section
   - Add allowed origins if needed
   - Click **Save**

---

### 1.3 Get Anthropic API Key (Optional)

> **Note:** Claude AI validation is currently disabled via the `ENABLE_CLAUDE_AI` flag. This step is only needed if you plan to re-enable AI validation.

1. **Sign up for Anthropic**
   - Go to: https://console.anthropic.com
   - Create account or sign in

2. **Create API Key**
   - Navigate to **API Keys**
   - Click **Create Key**
   - **Name:** `MaceStyleValidator`
   - Click **Create**
   - **Copy the key** (starts with `sk-ant-...`)

3. **Fund your account**
   - Add billing information
   - Add credits (recommended: $10 minimum)
   - Claude Haiku 4.5 costs ~$0.01 per document (Word only)

---

## Part 2: SharePoint Setup

### 2.1 Create SharePoint Site

1. **Create Site Collection** (if not exists)
   - Go to SharePoint Admin Center
   - **Sites** → **Active sites** → **Create**
   - **Type:** Team site
   - **Site name:** `Style Validation`
   - **Site address:** `/sites/StyleValidation`
   - **Primary administrator:** [Your account]
   - Click **Finish**

2. **Verify Site URL**
   ```
   Expected: https://[tenant].sharepoint.com/sites/StyleValidation
   ```

---

### 2.2 Create Style Rules List

1. **Navigate to site**
   - Go to: `https://[tenant].sharepoint.com/sites/StyleValidation`

2. **Create List**
   - Click **New** → **List**
   - **Name:** `Style Rules`
   - **Description:** `Validation rules for style checking`
   - Click **Create**

3. **Add Columns**

   **Column 1: RuleType**
   - Type: Choice
   - Choices:
     ```
     Font
     Language
     Grammar
     Punctuation
     Capitalisation
     Layout
     ```
   - Default: (none)

   **Column 2: DocumentType**
   - Type: Choice
   - Choices:
     ```
     Word
     Visio
     Excel
     PowerPoint
     Both
     All
     ```
   - Default: Word

   **Column 3: CheckValue**
   - Type: Single line of text

   **Column 4: ExpectedValue**
   - Type: Single line of text

   **Column 5: AutoFix**
   - Type: Yes/No
   - Default: Yes

   **Column 6: UseAI**
   - Type: Yes/No
   - Default: No

   **Column 7: Priority**
   - Type: Number
   - Min: 1
   - Max: 999
   - Default: 100

4. **Populate Rules**
   - Option A: Use the `populate_style_rules.py` script
   - Option B: Manually add rules from the style guide

   **To use the script:**
   ```bash
   cd MaceStyleValidator

   # Set environment variables
   export SHAREPOINT_TENANT_ID="[your-tenant-id]"
   export SHAREPOINT_CLIENT_ID="[your-client-id]"
   export SHAREPOINT_CLIENT_SECRET="[your-client-secret]"
   export SHAREPOINT_SITE_URL="https://[tenant].sharepoint.com/sites/StyleValidation"

   # Run script
   python3 populate_style_rules.py
   ```

   This will add ~70 style rules automatically.

---

### 2.3 Create Validation Results List

1. **Create List**
   - **Name:** `Validation Results`
   - **Description:** `History of document validations`
   - Click **Create**

2. **Add Columns**

   **Column 1: FileName**
   - Type: Single line of text

   **Column 2: ValidationDate**
   - Type: Date and Time
   - Format: Date & Time

   **Column 3: Status**
   - Type: Choice
   - Choices:
     ```
     Passed
     Review Required
     Failed
     ```
   - Default: (none)

   **Column 4: IssuesFound**
   - Type: Single line of text
   - (Stored as text for simplicity)

   **Column 5: IssuesFixed**
   - Type: Single line of text

   **Column 6: ReportLink**
   - Type: Hyperlink or Picture

3. **Get List ID** (needed for code)
   - Go to List Settings → copy URL
   - Extract ID from URL: `List={GUID}`
   - Or run: `python3 inspect_validation_results.py`

   **Note:** List IDs are configurable via environment variables:
   - `SHAREPOINT_VALIDATION_RESULTS_ID` — GUID of the Validation Results list
   - `SHAREPOINT_DOC_LIBRARY_ID` — GUID of the Document Library

   Set these in Azure Function App Settings for production.

---

### 2.4 Create/Configure Document Library

1. **Use Default or Create New**
   - **Option A:** Use existing "Documents" library
   - **Option B:** Create new library named "Validated Documents"

2. **Add Custom Columns**

   **Column 1: ValidationStatus**
   - Type: Choice
   - Choices:
     ```
     Not Validated
     Validate Now
     Validating...
     Passed
     Review Required
     Failed
     ```
   - Default: Not Validated
   - Colour coding (optional):
     - Validating... = Yellow
     - Passed = Green
     - Review Required = Amber
     - Failed = Red

   **Column 2: ValidationResultLink**
   - Type: Hyperlink or Picture

   **Column 3: LastValidated**
   - Type: Date and Time
   - Format: Date & Time

3. **Enable Versioning** (Important!)
   - Library Settings → **Versioning settings**
   - **Document Version History:** Yes
   - **Create major versions:** Yes
   - **Keep versions:** At least 50 (or more)
   - Click **OK**

4. **Set Permissions** (if needed)
   - Ensure users have:
     - Read: To view documents
     - Edit: To upload and modify
     - Contribute: To trigger validation

---

## Part 3: Azure Function Deployment

### 3.1 Clone Repository

```bash
# Clone the repository
git clone https://github.com/stephencummins/MaceStyle.git
cd MaceStyle/MaceStyleValidator
```

### 3.2 Local Development Setup

1. **Create virtual environment**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Create local.settings.json**
   ```bash
   cat > local.settings.json << 'EOF'
   {
     "IsEncrypted": false,
     "Values": {
       "AzureWebJobsStorage": "UseDevelopmentStorage=true",
       "FUNCTIONS_WORKER_RUNTIME": "python",
       "SHAREPOINT_TENANT_ID": "[your-tenant-id]",
       "SHAREPOINT_CLIENT_ID": "[your-client-id]",
       "SHAREPOINT_CLIENT_SECRET": "[your-client-secret]",
       "SHAREPOINT_SITE_URL": "https://[tenant].sharepoint.com/sites/StyleValidation",
       "ANTHROPIC_API_KEY": "[your-anthropic-api-key]"
     }
   }
   EOF
   ```

   **Replace all `[...]` placeholders with your actual values**

4. **Test locally**
   ```bash
   func start
   ```

   You should see:
   ```
   Functions:
       ListDocuments: [GET,POST] http://localhost:7071/api/listdocuments
       TestSharePoint: [GET,POST] http://localhost:7071/api/testsharepoint
       ValidateDocument: [POST] http://localhost:7071/api/validatedocument
   ```

### 3.3 Deploy to Azure

1. **Login to Azure**
   ```bash
   az login
   ```

2. **Set subscription** (if you have multiple)
   ```bash
   az account set --subscription "[Your Subscription Name or ID]"
   ```

3. **Deploy Function**
   ```bash
   func azure functionapp publish func-mace-validator-prod
   ```

   Wait for deployment to complete (2-3 minutes).

4. **Verify deployment**
   - Go to Azure Portal → Function App
   - Check **Functions** tab - should show 3 functions
   - Check **App Keys** → copy **Host key (default)**

5. **Get Function URL**
   ```
   https://func-mace-validator-prod.azurewebsites.net/api/validatedocument
   ```

   **Save this URL - needed for Power Automate!**

---

## Part 4: Logic App Setup (Replaces Power Automate)

An ARM template is provided at `infra/logic-app.json` that deploys a Consumption Logic App. This is preferred over Power Automate for production as it lives in the Azure subscription and can be deployed via ARM/DevOps.

### 4.1 Deploy Logic App

```bash
az deployment group create \
  --resource-group rg-macestyle \
  --template-file infra/logic-app.json \
  --parameters infra/logic-app.parameters.json \
  --parameters functionAppKey="<function-host-key>" \
  --parameters sharepointDocLibraryId="<library-guid>"
```

### 4.2 Authorise SharePoint Connection

After deployment:
1. Go to Azure Portal → Resource Group → API Connections → `sharepointonline`
2. Click **Edit API connection**
3. Click **Authorize** and sign in with a SharePoint account
4. Click **Save**

### 4.3 Logic App Flow

The Logic App:
1. Triggers on file create/modify in the SharePoint Document Library
2. Filters to supported file types (.docx, .xlsx, .pptx, .vsdx, etc.)
3. Sets ValidationStatus to "Validating..."
4. Gets file content and base64-encodes it
5. Calls the Function App (`POST /api/ValidateDocument`) with 10-minute timeout
6. Parses the JSON response
7. If fixes were applied, uploads the corrected file back
8. Updates ValidationStatus, Description, ValidationReport, ValidationResultLink, and LastValidated

### 4.4 Azure DevOps CI/CD (Optional)

An `azure-pipelines.yml` is provided in the repo root for automated deployment:
- Build + test on push to `main`
- Deploy to dev (automatic)
- Deploy to prod (with approval gate)
- Deploy Logic App ARM template

See `docs/prod-deployment-guide.md` for DevOps setup instructions.

### 4.5 Test

1. Upload a test document to the SharePoint library
2. Check Logic App run history in Azure Portal
3. Verify document status updated, report generated, and results saved

---

## Part 5: Testing & Validation

### 5.1 Create Test Document

Create `test_validation.docx` with:

```
Content to test:

1. American spelling: finalized, color, center, analyze
2. Contractions: can't, don't, won't
3. Symbols: M&S, 50%
4. Wrong font: (Set some text to Calibri or Times New Roman)
5. Numbers: 1000, 2000, 3000
```

### 5.2 Test Validation

1. **Upload to SharePoint**
   - Upload `test_validation.docx`
   - Wait for "Validating..." status

2. **Check Results** (should complete in 10-20 seconds)
   - Status: Should be "Passed"
   - ValidationResultLink: Should have link

3. **Open HTML Report**
   - Click ValidationResultLink
   - Should show:
     - Fixes Applied: Multiple corrections
     - Issues: All issues listed

4. **Download Fixed Document**
   - Download the document
   - Open in Word
   - Verify:
     - ✅ finalized → finalised
     - ✅ color → colour
     - ✅ can't → cannot
     - ✅ M&S → M and S
     - ✅ 50% → 50 percent
     - ✅ All text is Arial
     - ✅ Numbers: 1,000, 2,000, 3,000

### 5.3 Test Error Scenarios

**Test 1: Invalid file type**
- Upload `test.txt`
- Expected: Error or status "Failed"

**Test 2: Large file**
- Upload 100+ page document
- Expected: Completes (may take 30+ seconds)

**Test 3: Corrupted file**
- Upload corrupted .docx
- Expected: Status "Failed" with error

---

## Part 6: Monitoring & Maintenance

### 6.1 Azure Function Monitoring

**Application Insights:**
```
Azure Portal → Function App → Application Insights

Key Metrics:
- Request rate
- Average duration
- Failure rate
- Dependency calls (Graph API, Claude API)
```

**Live Metrics:**
```
Application Insights → Live Metrics

Monitor:
- Real-time requests
- Failures as they happen
- Performance issues
```

**Log Stream:**
```
Function App → Log stream

View:
- Real-time execution logs
- Debugging information
- Error details
```

### 6.2 Cost Monitoring

**Azure Costs:**
```
Azure Portal → Cost Management + Billing → Cost Analysis

Filter by:
- Resource Group: rg-macestyle
- Service: Azure Functions, Application Insights
```

**Anthropic API Costs:**
```
console.anthropic.com → Usage

Track:
- Total API calls
- Tokens used
- Cost per month
```

**Expected Monthly Costs:**
```
Azure Functions: $0-10 (Consumption plan)
Application Insights: $0-5 (Basic tier)
Claude API: ~$0.01 per document
SharePoint: Included in M365

Example:
- 1,000 documents/month = ~$10-20 total
- 100 documents/month = ~$1-5 total
```

### 6.3 Performance Tuning

**Optimize validation speed:**

1. **Reduce AI rules** (if too slow)
   - Disable UseAI for simple rules
   - Keep AI for complex language checks

2. **Batch operations** (for large documents)
   - Split very large files
   - Process in chunks

3. **Cache Graph API tokens**
   - Already implemented (60-minute cache)

4. **Use faster Claude model** (if needed)
   - Current: Haiku (fast, cheap)
   - Alternative: Sonnet (slower, more accurate)

### 6.4 Regular Maintenance

**Weekly:**
- [ ] Check error logs for patterns
- [ ] Review failed validations
- [ ] Verify costs are within budget

**Monthly:**
- [ ] Review and update style rules
- [ ] Analyze validation metrics
- [ ] Check for API updates (Graph, Anthropic)

**Quarterly:**
- [ ] Review permissions and security
- [ ] Update dependencies
- [ ] Rotate secrets if required

---

## Part 7: Security Best Practices

### 7.1 Secret Management

**Do:**
✅ Store secrets in Azure Key Vault
✅ Use managed identities when possible
✅ Rotate secrets every 12 months
✅ Use different secrets for dev/prod

**Don't:**
❌ Commit secrets to Git
❌ Share secrets via email/chat
❌ Use same secrets across environments
❌ Store secrets in code

### 7.2 Access Control

**Principle of Least Privilege:**

1. **Azure Function**
   - Only needs Sites.ReadWrite.All
   - No user impersonation

2. **SharePoint Users**
   - Read: View documents and reports
   - Edit: Upload and modify documents
   - No Contribute to Style Rules list

3. **Administrators**
   - Manage Style Rules
   - View logs
   - Modify configuration

### 7.3 Compliance

**Data Privacy:**
- Documents processed in-memory only
- No persistent storage of content
- Audit trail in Validation Results

**GDPR Considerations:**
- Personal data in documents (if any)
- Retention policy for Validation Results
- Right to deletion (version history cleanup)

---

## Part 8: Troubleshooting

### Common Issues

#### Issue 1: "401 Unauthorised" errors

**Cause:** App registration permissions not granted

**Solution:**
1. Azure Portal → App Registration
2. API Permissions → Grant admin consent
3. Verify green checkmarks

---

#### Issue 2: Validation not triggering

**Cause:** Logic App disabled or broken

**Solution:**
1. Azure Portal → Logic App → Check run history
2. Review last run error
3. Check SharePoint API connection authorisation

---

#### Issue 3: "List not found" errors

**Cause:** List ID mismatch

**Solution:**
1. Run `inspect_validation_results.py`
2. Get actual list ID
3. Update in `sharepoint_results.py` line 35

---

#### Issue 4: Claude API errors

> **Note:** Claude AI is currently disabled. This section only applies if `ENABLE_CLAUDE_AI = True`.

**Cause:** Invalid API key or insufficient credits

**Solution:**
1. Verify API key in Function App settings
2. Check Anthropic console for credits
3. Add credits if balance is low

---

#### Issue 5: HTML report not uploading

**Cause:** Missing fileUrl parameter

**Solution:**
1. Check Logic App request body includes fileUrl
2. Review Logic App run history for the specific run

---

## Part 9: Backup & Recovery

### 9.1 Backup Configuration

**SharePoint Lists:**
- Export Style Rules to Excel monthly
- Export Validation Results quarterly
- Use SharePoint Online backup (included in M365)

**Azure Function Code:**
- Git repository (primary backup)
- Azure DevOps / GitHub (secondary)
- Local developer machines

**Secrets:**
- Document in secure password manager
- Store offline backup in safe
- Keep recovery contacts updated

### 9.2 Disaster Recovery

**Scenario: SharePoint site deleted**
1. Restore from SharePoint Recycle Bin (93 days)
2. Or restore from M365 admin center backup
3. Re-run populate_style_rules.py to restore rules

**Scenario: Azure Function deleted**
1. Redeploy from Git
2. Reconfigure App Settings
3. Test with sample document

**Scenario: App registration deleted**
1. Create new app registration
2. Update secrets in Function App
3. Re-grant API permissions

---

## Appendix A: Complete Environment Variables

```bash
# Azure Function App Settings
SHAREPOINT_TENANT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
SHAREPOINT_CLIENT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
SHAREPOINT_CLIENT_SECRET="your-client-secret-value"
SHAREPOINT_SITE_URL="https://tenant.sharepoint.com/sites/StyleValidation"
ANTHROPIC_API_KEY="sk-ant-xxxxxxxxxxxxxxxxxxxxx"
```

---

## Appendix B: PowerShell Helper Scripts

### Get SharePoint Site ID
```powershell
Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/StyleValidation" -Interactive
Get-PnPSite | Select Id
```

### List All SharePoint Lists
```powershell
Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/StyleValidation" -Interactive
Get-PnPList | Select Title, Id
```

### Export Style Rules
```powershell
Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/StyleValidation" -Interactive
Get-PnPListItem -List "Style Rules" | Export-Csv "StyleRules_Backup.csv"
```

---

## Appendix C: Testing Checklist

Before going live, verify:

- [ ] App registration created with correct permissions
- [ ] Azure Function deployed and responding
- [ ] SharePoint lists created with all columns
- [ ] Style Rules populated
- [ ] Document library configured with custom columns
- [ ] Logic App deployed and SharePoint connection authorised
- [ ] Test document validates successfully
- [ ] HTML report generates and uploads
- [ ] Validation Results list populates
- [ ] Document metadata updates with result link
- [ ] All secrets stored securely
- [ ] Monitoring and logging configured
- [ ] User guide distributed to team
- [ ] Admin contacts documented

---

## Support & Resources

### Official Documentation
- [Azure Functions Python](https://docs.microsoft.com/azure/azure-functions/functions-reference-python)
- [Microsoft Graph API](https://docs.microsoft.com/graph/overview)
- [Azure Logic Apps](https://docs.microsoft.com/azure/logic-apps/)
- [Claude API](https://docs.anthropic.com/claude/reference/getting-started-with-the-api)

### Community
- Azure Functions GitHub: https://github.com/Azure/azure-functions-python-worker
- Power Automate Community: https://powerusers.microsoft.com/

### Version Information
- **Last Updated:** March 2026
- **Version:** 5.0
- **Author:** [Your Name/Team]
- **License:** [Your License]
