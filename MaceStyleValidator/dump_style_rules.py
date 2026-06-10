"""Export the live SharePoint 'Style Rules' list to rules_snapshot.json.

Run on a machine where the SharePoint credentials are set in the environment
(the same SHAREPOINT_* / STYLE_RULES_* vars the function uses — e.g. exported
from local.settings.json). Then analyse the snapshot offline:

    python3 dump_style_rules.py
    python3 rule_doctor.py rules_snapshot.json

This is the bridge between "the engine works" (test_rule_coverage.py) and
"the real rules are shaped to use it" (rule_doctor.py).
"""
import json
import os
import sys

from ValidateDocument.sharepoint_client import fetch_validation_rules


def main():
    try:
        # fetch_validation_rules re-derives its own token from the
        # STYLE_RULES_* / SHAREPOINT_* env vars; the argument is unused for the
        # rules fetch, so an empty string is fine.
        rules = fetch_validation_rules("")
    except Exception as e:
        print(f"Failed to fetch rules: {e}", file=sys.stderr)
        print("Ensure SHAREPOINT_TENANT_ID / CLIENT_ID / CLIENT_SECRET / SITE_URL "
              "(or the STYLE_RULES_* equivalents) are set in the environment.",
              file=sys.stderr)
        return 1

    out = os.environ.get("RULES_SNAPSHOT", "rules_snapshot.json")
    with open(out, "w") as f:
        json.dump(rules, f, indent=2, default=str)
    print(f"Wrote {len(rules)} rules to {out}")
    print(f"Next: python3 rule_doctor.py {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
