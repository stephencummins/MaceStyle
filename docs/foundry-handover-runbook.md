# MaceStyle — Microsoft Foundry Handover Runbook

Move MaceStyle's AI inference from a personal Anthropic API key onto **Claude in Microsoft Foundry**, billed and governed inside a **Mace Azure subscription**.

This is the technical companion to the [Mace Funding Brief](./mace-funding-brief.md). The whole path below has been rehearsed end-to-end in a test resource — it is execution, not discovery.

## Outcome

The `func-mace-validator-dev` Function calls Claude via a Foundry resource in Mace's tenant. Claude usage bills to Mace's Azure invoice. **No application code changes** — the switch is app settings only, because `ai_client.py` is provider-switchable via `AI_PROVIDER`.

## Prerequisites

- A **Mace Azure subscription** with rights to create Cognitive Services / AI resources (Contributor on the target resource group).
- Rights to **accept an Azure Marketplace offer** in that subscription (required once — see Step 2). This is the one step that needs the portal and cannot be scripted.
- **Azure CLI** ≥ 2.79 (`az login` against the Mace tenant), or use the Foundry portal (ai.azure.com).
- Access to the `func-mace-validator-dev` Function App settings (or wherever the Function ends up living — see Step 5).
- Data-governance sign-off to process Control Centre document text via a US-hosted model (Sapna).

## Key facts from the rehearsal

| Item | Value used in rehearsal |
|---|---|
| Foundry resource | `foundry-mace-validator` (rename for Mace, e.g. `foundry-macestyle-prod`) |
| Region | `eastus2` — Claude is **not** offered in `uksouth`; US data zone |
| Model / version | `claude-haiku-4-5`, version `2`, `GlobalStandard` |
| Rate limit (capacity 10) | ~80,000 tokens/min, 80 requests/min |
| Inference endpoint | `https://<resource>.services.ai.azure.com/anthropic` |
| SDK client | `AnthropicFoundry(resource=..., api_key=...)`, needs `anthropic>=0.103` |

---

## Step 1: Create the Foundry resource (Mace tenant)

```bash
az login                                   # sign in to the Mace tenant
az account set --subscription "<Mace subscription id>"

az cognitiveservices account create \
  --name foundry-macestyle-prod \
  --resource-group <mace-rg> \
  --kind AIServices \
  --sku S0 \
  --location eastus2 \
  --custom-domain foundry-macestyle-prod \
  --yes
```

`--custom-domain` is required — the Anthropic endpoint is served from `https://<name>.services.ai.azure.com`.

## Step 2: Deploy the Claude model (⚠️ one-time Marketplace step)

**This is the step that tripped up the rehearsal.** Deploying an Anthropic model requires accepting the Azure Marketplace terms *once per subscription*, which creates a Marketplace SaaS subscription linking Azure to Anthropic's offer. **A pure CLI/ARM deployment fails silently** (`provisioningState: Failed`, no error message) until those terms have been accepted.

**Do this the first time, in the portal:**

1. Go to **https://ai.azure.com** → open the `foundry-macestyle-prod` resource.
2. **Model catalog** → search `claude-haiku-4-5` → **Deploy**.
3. Fill in the provider details: Organisation `Mace Group`, Country `United Kingdom`, Industry `Construction`.
4. **Accept the Marketplace/publisher terms** when prompted (this is the linkage that unblocks everything).
5. Keep the deployment name **`claude-haiku-4-5`** and confirm.

After the terms are accepted once, subsequent deployments in the same subscription *can* be scripted:

```bash
SUB=$(az account show --query id -o tsv)
az rest --method PUT \
  --url "https://management.azure.com/subscriptions/$SUB/resourceGroups/<mace-rg>/providers/Microsoft.CognitiveServices/accounts/foundry-macestyle-prod/deployments/claude-haiku-4-5?api-version=2025-10-01-preview" \
  --body '{
    "sku": {"name": "GlobalStandard", "capacity": 10},
    "properties": {
      "model": {"format": "Anthropic", "name": "claude-haiku-4-5", "version": "2"},
      "modelProviderData": {"organizationName": "Mace Group", "countryCode": "GB", "industry": "construction"},
      "versionUpgradeOption": "OnceNewDefaultVersionAvailable",
      "raiPolicyName": "Microsoft.DefaultV2"
    }
  }'
```

> **Gotchas:** `modelProviderData` is mandatory (`organizationName`, `countryCode`, `industry` in lowercase) **and** the API version must be `2025-10-01-preview` — older versions strip the field and return `InvalidModelProviderData`.

Confirm it reaches `Succeeded`:

```bash
az cognitiveservices account deployment show \
  -n foundry-macestyle-prod -g <mace-rg> \
  --deployment-name claude-haiku-4-5 \
  --query "properties.provisioningState" -o tsv
```

## Step 3: Retrieve the endpoint and key

```bash
# Key
az cognitiveservices account keys list \
  -n foundry-macestyle-prod -g <mace-rg> --query key1 -o tsv

# Endpoint is: https://foundry-macestyle-prod.services.ai.azure.com/anthropic
```

**Smoke-test before touching the Function:**

```bash
FKEY=$(az cognitiveservices account keys list -n foundry-macestyle-prod -g <mace-rg> --query key1 -o tsv)
curl -s -X POST "https://foundry-macestyle-prod.services.ai.azure.com/anthropic/v1/messages" \
  -H "x-api-key: $FKEY" -H "anthropic-version: 2023-06-01" -H "content-type: application/json" \
  -d '{"model":"claude-haiku-4-5","max_tokens":30,"messages":[{"role":"user","content":"Reply with exactly: Foundry OK"}]}'
```

## Step 4: Decide where the Function lives

Two options:

- **A. Keep the Function where it is, point it at Mace's Foundry.** Fastest. Only the AI cost moves to Mace; the Function's (trivial) hosting cost stays put. Fine as an interim.
- **B. Move the Function into the Mace subscription too.** Cleanest end state — Mace owns the whole service. Redeploy the Function App into `<mace-rg>` via VS Code (the existing deploy path — `func publish` is known to fail for this project), then re-apply all app settings.

Either way, the AI provider switch in Step 5 is identical.

## Step 5: Point the Function at Foundry (app settings)

Set these on the target Function App (`func-mace-validator-dev`, or its Mace-tenant replacement):

```bash
az functionapp config appsettings set -n <function-app> -g <function-rg> --settings \
  ENABLE_CLAUDE_AI=true \
  AI_PROVIDER=foundry \
  CLAUDE_MODEL=claude-haiku-4-5 \
  FOUNDRY_RESOURCE=foundry-macestyle-prod \
  FOUNDRY_API_KEY="<key from Step 3>"
```

> **Critical:** `CLAUDE_MODEL` must be the **Foundry deployment name** (`claude-haiku-4-5`), **not** the dated Anthropic string (`claude-haiku-4-5-20251001`). The dated string 404s against Foundry.

Then remove the personal-account setting so nothing falls back to it:

```bash
az functionapp config appsettings delete -n <function-app> -g <function-rg> \
  --setting-names ANTHROPIC_API_KEY
```

## Step 6: Verify

1. `GET https://<function-app>.azurewebsites.net/api/HealthCheck` (with the function key) → the `claude_api` check should read `API key configured (foundry)`.
2. Trigger a real validation from SharePoint (set a document's `ValidationStatus` to **Validate Now**) and confirm the Validation Results entry is produced.
3. Check the Foundry resource's **Metrics** (or Monitoring) in the portal — you should see the inference request land there, confirming it billed to Mace.

## Rollback

Instant, no redeploy — flip the provider back to the personal key (only valid while `ANTHROPIC_API_KEY` is still set):

```bash
az functionapp config appsettings set -n <function-app> -g <function-rg> --settings \
  AI_PROVIDER=anthropic CLAUDE_MODEL=claude-haiku-4-5-20251001
```

Or disable AI validation entirely (hard rules still run): `ENABLE_CLAUDE_AI=false`.

## Decommissioning the personal setup

Once the Mace path is verified and stable:

- Delete the personal Foundry resource `foundry-mace-validator` in `rg-mace-validator` (eastus2) if it was ever created for rehearsal.
- Rotate/revoke the personal `ANTHROPIC_API_KEY` that was used during the pilot.

## Data governance notes (for sign-off)

- **What leaves Azure:** the text of documents flagged with `UseAI: Yes` rules is sent to the Claude model for correction. Documents over 50,000 characters emit a warning log before transmission (`ai_client.py`).
- **Where it goes:** a US data zone (Foundry does not currently offer Claude in UK regions). Anthropic operates the model under Microsoft's Foundry terms; usage is governed by the Azure agreement, not a separate Anthropic contract.
- **Controls available:** Entra ID auth, RBAC on the resource, per-deployment rate limits/quotas, and Azure spend alerts. Configure these on the Mace resource per Mace policy.
- **No training on inputs:** Claude in Foundry does not use customer inputs to train models (per Microsoft/Anthropic Foundry terms — confirm the current terms at acceptance).
