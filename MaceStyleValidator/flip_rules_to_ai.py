"""
Flip a curated set of *editorial* Style Rules to UseAI=True — the rules Claude
handles better than deterministic checks (wordy-phrase simplification, tone,
homophones, redundancy, contractions, text symbols, and an iconic British-
spelling set).

Deliberately LEAVES deterministic: fonts, colours, all Visio/Process-Map rules,
structural/format rules (captions, TOC, hyperlinks, line-breaks, hyphenation,
quotes, date/time/number formats) and the long tail of British-spelling checks,
which the hard-coded validators already do perfectly and for free.

The showcase is fully reversible via --revert (same curated selection):
    python3 flip_rules_to_ai.py                    # preview (dry run)
    python3 flip_rules_to_ai.py --apply            # write UseAI=True
    python3 flip_rules_to_ai.py --revert --apply   # set the same 69 back to UseAI=False

NOTE: the UseAI column's INTERNAL name is auto-generated ('field_7'), not 'UseAI'.
Writes that PATCH {"fields": {"UseAI": ...}} fail with "Field not recognized" —
this script resolves the internal name at runtime.

Requires SHAREPOINT_* creds in the environment (or local.settings.json exported).
"""
import os
import re
import sys
import json
import requests

# Load gitignored local.settings.json into the environment if present.
_ls = os.path.join(os.path.dirname(__file__), "local.settings.json")
if os.path.exists(_ls):
    for k, v in json.load(open(_ls)).get("Values", {}).items():
        os.environ.setdefault(k, str(v))

from ValidateDocument.config import get_graph_token, get_site_id

# British-spelling target words (the first quoted token in the rule title) that
# we route to AI as a representative showcase set. The rest stay deterministic.
ICONIC_BRITISH = {
    "colour", "centre", "organise", "analyse", "licence", "programme",
    "aluminium", "authorise", "realise", "metre", "defence", "grey",
}

# Substring keyword matches (lower-cased) for the editorial rules we flip.
JUDGEMENT_KEYWORDS = (
    "professional tone", "parallel structure", "consistent terminology",
    "homophone", "close proximity", "do not use 'feel'", "avoid 'etc.'",
    "'etc.' alongside", "do not use 'etc.'", "and/or", "toward' not 'towards",
    "'should' or 'could'", "constructability", "forecast' as past tense",
)
SYMBOL_KEYWORDS = (
    "apostrophes for plurals", "forward slash",
)
CAP_KEYWORDS = (
    "capitalise for emphasis", "brand names: capitalise",
    "job titles only capitalised", "proper noun derivations",
    "fields of study lowercase", "govt bodies",
)


def should_flip(title, doc_type):
    t = (title or "").strip()
    tl = t.lower()
    # --- exclusions: keep deterministic ---
    if (doc_type or "").strip().lower() == "visio":
        return False
    if t.startswith("Visio") or "Process Map" in t or t.startswith("[DISABLED"):
        return False
    # --- inclusions: editorial rules AI does best ---
    if t.startswith("Replace '"):                       # wordy-phrase simplification
        return True
    if "no contractions" in tl:                          # contractions
        return True
    if "ampersand" in tl and "visio" not in tl:          # & -> and (text)
        return True
    if "percent" in tl and "spell out" in tl:            # % -> percent
        return True
    if any(k in tl for k in SYMBOL_KEYWORDS):
        return True
    if any(k in tl for k in JUDGEMENT_KEYWORDS):
        return True
    if any(k in tl for k in CAP_KEYWORDS):
        return True
    if t.startswith("Use British English spelling"):
        m = re.search(r"'([^']+)'", t)
        return bool(m and m.group(1).lower() in ICONIC_BRITISH)
    return False


def truthy(val):
    if isinstance(val, bool):
        return val
    if isinstance(val, str):
        return val.strip().lower() in ("yes", "true", "1")
    return bool(val)


def resolve_useai_internal_name(site_id, headers):
    """The 'UseAI' column's INTERNAL name is auto-generated (e.g. field_7) and
    differs from its display name; PATCH must use the internal name."""
    cols = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/columns",
        headers=headers).json().get("value", [])
    for c in cols:
        if c.get("displayName") == "UseAI" or c.get("name") == "UseAI":
            return c["name"]
    return "field_7"  # confirmed fallback for this list


def main(apply_changes, revert=False):
    target = not revert  # revert => set the same curated set back to False
    token = get_graph_token()
    site_id = get_site_id(token)
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json",
               "Content-Type": "application/json"}
    useai_field = resolve_useai_internal_name(site_id, headers)

    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/"
           f"items?expand=fields&$top=500")
    items = []
    while url:
        b = requests.get(url, headers=headers).json()
        items.extend(b.get("value", []))
        url = b.get("@odata.nextLink")
    print(f"Fetched {len(items)} Style Rules.\n")

    to_flip = []
    for it in items:
        f = it["fields"]
        title = f.get("Title", "")
        doc_type = f.get("DocType") or f.get("RuleType") or f.get("field_2") or ""
        use_ai = truthy(f.get(useai_field, f.get("UseAI")))
        if should_flip(title, doc_type) and use_ai != target:
            to_flip.append((it["id"], title))

    verb = "flip to UseAI=True" if target else "revert to UseAI=False"
    print(f"{len(to_flip)} editorial rule(s) to {verb}:\n")
    for _id, title in sorted(to_flip, key=lambda x: x[1]):
        print(f"  • {title}")

    if not apply_changes:
        print(f"\nDRY RUN — no changes written (UseAI column = '{useai_field}'). "
              f"Re-run with --apply to {verb}.")
        return

    print(f"\nApplying UseAI={target} (column '{useai_field}') ...")
    ok = 0
    for _id, title in to_flip:
        u = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/items/{_id}"
        r = requests.patch(u, headers=headers, json={"fields": {useai_field: target}})
        r.raise_for_status()
        ok += 1
    print(f"\nDone — {ok} rule(s) now UseAI={target}.")


if __name__ == "__main__":
    main(apply_changes="--apply" in sys.argv, revert="--revert" in sys.argv)
