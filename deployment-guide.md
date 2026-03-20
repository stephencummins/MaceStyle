# MaceyBot Deployment Guide

Deploy the MaceyBot Teams bot powered by Claude API to replace Copilot Studio.

## Prerequisites

- Azure subscription with access to create resources
- Azure Functions Core Tools v4 (or VS Code Azure Functions extension)
- An Anthropic API key
- SharePoint site with a "Site Creation" list
- Azure AD app registration with `Sites.Selected` (application permission) and per-site access granted

## Step 1: Azure Bot Service

1. Go to [Azure Portal](https://portal.azure.com) → Create a resource → "Azure Bot"
2. Configure:
   - **Bot handle**: `MaceyBot`
   - **Pricing tier**: Standard (S1) — free for Teams-only
   - **Type of app**: Multi Tenant
   - **Creation type**: Create new Microsoft App ID
3. Click **Create**
4. Once created, go to the resource → **Configuration**:
   - Note the **Microsoft App ID** (this is `BOT_APP_ID`)
   - Click **Manage Password** → **New client secret** → copy the value (this is `BOT_APP_PASSWORD`)
5. Set the **Messaging endpoint** to:
   ```
   https://<your-function-app>.azurewebsites.net/api/MaceyBot
   ```
   (For `func-mace-validator-dev`, this would be:
   `https://func-mace-validator-dev.azurewebsites.net/api/MaceyBot`)

## Step 2: Enable Teams Channel

1. In the Azure Bot resource → **Channels** → **Microsoft Teams**
2. Click **Apply** (accept the terms)
3. The Teams channel is now active

## Step 3: SharePoint App Registration

If you already have an app registration for MaceStyle with `Sites.Selected`, you can reuse it. Otherwise:

1. Azure Portal → Azure Active Directory → App registrations → New registration
2. Name: `MaceyBot-SP` (or reuse existing)
3. API permissions → Add → Microsoft Graph → Application → `Sites.Selected`
4. Grant admin consent
5. Certificates & secrets → New client secret → copy value
6. Note the **Application (client) ID** and **Directory (tenant) ID**

## Step 4: Environment Variables

Add these to your Azure Function App Settings (Configuration → Application settings):

| Variable | Value | Notes |
|----------|-------|-------|
| `BOT_APP_ID` | Microsoft App ID from Step 1 | Azure Bot registration |
| `BOT_APP_PASSWORD` | Client secret from Step 1 | Azure Bot password |
| `ANTHROPIC_API_KEY` | Your Anthropic API key | Claude API access |
| `SP_TENANT_ID` | Azure AD tenant ID | SharePoint auth |
| `SP_CLIENT_ID` | App registration client ID | SharePoint auth |
| `SP_CLIENT_SECRET` | App registration client secret | SharePoint auth |
| `SP_SITE_URL` | `https://0rxf2.sharepoint.com/sites/pdms` | Target SharePoint site |
| `SP_LIST_NAME` | `Site Creation` | SharePoint list name |
| `MACEY_MODEL` | `claude-sonnet-4-20250514` | Optional: override Claude model |

## Step 5: Deploy the Function

Using VS Code:
1. Open `MaceStyleValidator/` in VS Code
2. Cmd+Shift+P → "Azure Functions: Deploy to Function App"
3. Select `func-mace-validator-dev`

Using CLI:
```bash
cd MaceStyleValidator
func azure functionapp publish func-mace-validator-dev
```

## Step 6: Teams App Manifest

1. Edit `teams-manifest/manifest.json`:
   - Replace `{{BOT_APP_ID}}` with your actual Bot App ID (from Step 1)
2. Replace the placeholder icons (`color.png`, `outline.png`) with branded versions if desired:
   - `color.png`: 192×192 px, full colour
   - `outline.png`: 32×32 px, transparent background with white outline
3. Zip the manifest:
   ```bash
   cd teams-manifest
   zip MaceyBot.zip manifest.json color.png outline.png
   ```
4. In Teams Admin Center (or Teams → Apps → Manage your apps → Upload):
   - Upload `MaceyBot.zip`
   - Or use "Upload a custom app" in Teams

## Step 7: Test

1. In Teams, find "Macey" in your apps
2. Start a chat — you should see the welcome message
3. Walk through the conversation flow:
   - Provide a project name
   - Provide a description
   - Choose visibility
   - Provide an owner email
   - Add optional notes
   - Confirm submission
4. Check the SharePoint "Site Creation" list for the new item

## Local Testing

Test the conversation flow without deploying:

```bash
cd MaceStyleValidator
export ANTHROPIC_API_KEY=your_key_here
python3 test-maceybot.py
```

This simulates the chat locally, calls Claude API, and mocks the SharePoint submission.

## Troubleshooting

### Bot doesn't respond in Teams
- Check the messaging endpoint URL in Azure Bot Configuration
- Verify `BOT_APP_ID` and `BOT_APP_PASSWORD` are correct
- Check Azure Function logs: Portal → Function App → Monitor

### Claude API errors
- Verify `ANTHROPIC_API_KEY` is set correctly
- Check the Function App logs for error details
- Try running `test-maceybot.py` locally to isolate the issue

### SharePoint write fails
- Verify the app registration has `Sites.Selected` with admin consent and per-site access granted
- Check `SP_SITE_URL` points to the correct site
- Verify the list name matches `SP_LIST_NAME`
- Check that list columns match: Title, SiteDescription, Visibility, SiteOwnerClaims, Notes

### "Unauthorized" errors
- For Bot Framework: check `BOT_APP_ID`/`BOT_APP_PASSWORD`
- For SharePoint: check `SP_TENANT_ID`/`SP_CLIENT_ID`/`SP_CLIENT_SECRET`
- Ensure the app registration is in the correct tenant

## Architecture

```
User in Teams
    ↓ (chat message)
Azure Bot Service (Teams channel)
    ↓
Python Azure Function (/api/MaceyBot)
    ↓ (sends conversation to Claude API)
Anthropic Claude API (claude-sonnet-4-20250514)
    ↓ (returns response + tool calls)
Python Azure Function
    ├── Replies to user in Teams
    └── When confirmed → writes to SharePoint "Site Creation" list
                              ↓ (triggers existing flow)
                         Power Automate "PDMS - Create Site and Group"
```

## Cost Estimate

- **Azure Bot Service**: Free for Teams-only channel
- **Azure Functions**: Consumption plan (existing — negligible additional cost)
- **Claude API**: ~$0.003–0.01 per conversation (Sonnet, ~5–8 turns, ~2K tokens/turn)
- **Estimated monthly**: <$5 for moderate usage (100 site requests/month)
