"""Authoritative coverage test against the testers' real rule set.

Loads style_rules_fixture.json (derived from the testers' 'Broken Rules
Tracker' — refresh with the snippet at the bottom) and asserts that EVERY rule
is handled: either implemented by a validator (deterministic) or routed to the
AI path (UseAI). Any rule that is neither — a silent no-op — fails the test,
UNLESS it is on KNOWN_GAPS: rules that genuinely need linguistic judgement,
page rendering, or Visio process-map logic and are tracked as future work.

This is the regression gate for "does the tool actually cover the testers'
rules". Run:  python3 test_tracker_rules.py
"""
import json
import os
import sys

from rule_registry import classify, handled_by

FIXTURE = os.path.join(os.path.dirname(__file__), "style_rules_fixture.json")

# check_values that are genuinely NOT cleanly deterministic. Disposition:
#   ai     — flip to UseAI=Yes in SharePoint; Claude handles these well
#   visio  — needs Visio process-map logic (shape colours), not text analysis
#   render — needs page rendering / structure the validator can't see offline
KNOWN_GAPS = {
    "ConsistentFonts": "ai", "JobTitles": "ai", "SectionTitles": "ai",
    "SubsidiaryHeadings": "ai", "DirectSpeech": "ai", "SpecialTerms": "ai",
    "FigureTableCaptions": "ai", "CaptionPosition": "ai", "PhaseNotation": "ai",
    "NoBreakNumberUnit": "ai", "MaxConsecutiveHyphens": "ai",
    "InterfaceBoxColor": "visio", "MultiLaneActivityColor": "visio",
    "HyperlinksWorking": "render", "NoOutstandingComments": "render",
    "NoWidowsOrphans": "render", "TOCComplete": "render",
}


def load_rules():
    with open(FIXTURE) as f:
        return json.load(f)


def run():
    rules = load_rules()
    det = [r for r in rules if classify(r)[0] == "deterministic"]
    ai = [r for r in rules if classify(r)[0] == "ai"]
    gaps = [r for r in rules if classify(r)[0] == "gap"]

    word_det = sum(1 for r in det if "word" in handled_by(r))
    print(f"\nTracker coverage — {len(rules)} rules from the testers' set\n")
    print(f"  deterministic : {len(det):3}  ({word_det} apply to Word)")
    print(f"  AI-routed     : {len(ai):3}")
    print(f"  gaps          : {len(gaps):3}")

    # 1. Data hygiene: booleans must be real booleans.
    bad_bool = [r.get("title") for r in rules
                for f in ("auto_fix", "use_ai")
                if r.get(f) is not None and not isinstance(r.get(f), bool)]
    # 2. No gap outside the documented allow-list.
    unexpected = [r for r in gaps if (r.get("check_value") not in KNOWN_GAPS)]

    ok = True
    if bad_bool:
        ok = False
        print(f"\n  ✗ {len(bad_bool)} rule(s) store auto_fix/use_ai as a string, not a boolean:")
        for t in bad_bool[:10]:
            print(f"      - {t}")
    if unexpected:
        ok = False
        print(f"\n  ✗ {len(unexpected)} rule(s) silently do nothing and are NOT in KNOWN_GAPS:")
        for r in unexpected:
            print(f"      - {r.get('title')} [{r.get('rule_type')}/{r.get('check_value')}]")
        print("    Implement a validator branch, set UseAI=Yes, or add to KNOWN_GAPS.")

    # Actionable worklist for the documented gaps (not a failure).
    if gaps:
        by_disp = {}
        for r in gaps:
            d = KNOWN_GAPS.get(r.get("check_value"), "?")
            by_disp.setdefault(d, []).append(r.get("check_value"))
        print("\n  Known gaps (tracked, not failing) — recommended disposition:")
        labels = {"ai": "flip to UseAI=Yes (Claude handles these)",
                  "visio": "needs Visio process-map logic",
                  "render": "needs page rendering / document structure"}
        for disp, cvs in sorted(by_disp.items()):
            print(f"    {labels.get(disp, disp)}:")
            for cv in sorted(set(cvs)):
                print(f"      - {cv}")

    print()
    if ok:
        print("  ✓ Every rule is handled deterministically, by AI, or is a documented gap.")
        return 0
    print("  ✗ Unhandled rules above — see messages.")
    return 1


if __name__ == "__main__":
    sys.exit(run())

# To refresh style_rules_fixture.json from a new tracker export:
#   import openpyxl, json
#   def b(v): return isinstance(v,bool) and v or str(v).strip().lower() in ("true","yes","1")
#   ws = openpyxl.load_workbook("Broken Rules Tracker.xlsx", data_only=True).active
#   rows = list(ws.iter_rows(values_only=True)); h = {k:i for i,k in enumerate(rows[0])}
#   out = [{"title":r[h["Title"]],"rule_type":r[h["RuleType"]],"doc_type":r[h["DocumentType"]],
#           "check_value":r[h["CheckValue"]],"expected_value":r[h["ExpectedValue"]],
#           "auto_fix":b(r[h["AutoFix"]]),"use_ai":b(r[h["UseAI"]]),"priority":r[h["Priority"]]}
#          for r in rows[1:] if r[h["Title"]]]
#   json.dump(out, open("style_rules_fixture.json","w"), indent=2, default=str)
