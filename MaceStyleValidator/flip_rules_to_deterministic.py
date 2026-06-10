"""
Flip the deterministic style rules from UseAI=True back to UseAI=False.

Background: contraction, British-spelling, symbol and number rules were marked
UseAI=True in the Style Rules list, but ENABLE_CLAUDE_AI is False, so they were
routed to the (disabled) Claude path and never ran. The hard-coded validators in
enhanced_validators.py already implement these checks deterministically, so we
flip them back to UseAI=False to make them run again — no external AI needed.

Only rules whose CheckValue maps to an IMPLEMENTED hard-coded checker are
touched. Anything else (e.g. genuinely judgement-based rules, or the dispatched-
but-unimplemented 'AvoidShould') is left exactly as-is.

DRY RUN by default — shows what would change. Pass --apply to write.

    cd MaceStyleValidator
    python3 flip_rules_to_deterministic.py            # preview
    python3 flip_rules_to_deterministic.py --apply    # write changes

Requires SHAREPOINT_TENANT_ID / CLIENT_ID / CLIENT_SECRET / SITE_URL in the
environment (or local.settings.json values exported). Pull them with:
    az functionapp config appsettings list -g rg-mace-validator -n func-mace-validator-dev
"""
import sys
import requests

from ValidateDocument.config import get_graph_token, get_site_id

# CheckValue values (or prefixes) with a real implementation in
# enhanced_validators.py. Keep this in lockstep with the dispatchers there.
IMPLEMENTED_PREFIXES = ("BritishSpelling_", "NoContraction_")
IMPLEMENTED_EXACT = {
    "Word_toward",
    "AvoidEtc",
    "NoAmpersand",
    "PercentSymbol",
    "NoApostrophePlurals",
    "NumberCommas",
}


def is_deterministic(check_value):
    if not check_value:
        return False
    if check_value in IMPLEMENTED_EXACT:
        return True
    return any(check_value.startswith(p) for p in IMPLEMENTED_PREFIXES)


def truthy(val):
    """Normalise SharePoint's UseAI (bool, 'Yes'/'No', 1/0) to a bool."""
    if isinstance(val, bool):
        return val
    if isinstance(val, str):
        return val.strip().lower() in ("yes", "true", "1")
    return bool(val)


def main(apply_changes):
    token = get_graph_token()
    site_id = get_site_id(token)
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }

    list_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/"
        f"items?expand=fields&$top=500"
    )
    items = []
    url = list_url
    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        body = resp.json()
        items.extend(body.get("value", []))
        url = body.get("@odata.nextLink")  # follow pagination if present

    print(f"Fetched {len(items)} Style Rules.\n")

    to_flip = []
    for item in items:
        fields = item.get("fields", {})
        check_value = fields.get("CheckValue") or fields.get("field_3")
        use_ai = fields.get("UseAI")
        if use_ai is None:
            use_ai = fields.get("field_7")
        if is_deterministic(check_value) and truthy(use_ai):
            to_flip.append((item["id"], fields.get("Title", "?"), check_value))

    if not to_flip:
        print("Nothing to flip — all deterministic rules already have UseAI=False.")
        return

    print(f"{len(to_flip)} deterministic rule(s) currently UseAI=True:\n")
    for _id, title, cv in to_flip:
        print(f"  [{cv}] {title}")

    if not apply_changes:
        print("\nDRY RUN — no changes written. Re-run with --apply to flip these to UseAI=False.")
        return

    print("\nApplying UseAI=False ...")
    for _id, title, cv in to_flip:
        item_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/items/{_id}"
        r = requests.patch(item_url, headers=headers, json={"fields": {"UseAI": False}})
        r.raise_for_status()
        print(f"  flipped: [{cv}] {title}")
    print(f"\nDone — {len(to_flip)} rule(s) now deterministic.")


if __name__ == "__main__":
    main(apply_changes="--apply" in sys.argv)
