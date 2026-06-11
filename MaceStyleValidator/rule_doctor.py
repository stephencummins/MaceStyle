"""Rule doctor — will my style rules actually fire?

Static analysis of the 'Style Rules' data against what the validators really
implement (see rule_registry.py for the single source of truth). Reports, per
rule, whether it is:

  * deterministic — a validator for its doc_type implements its check_value
  * ai            — UseAI is set, so Claude applies it from the rule text
  * GAP           — neither: the rule silently does nothing

and flags the data-shape problems that cause silent no-ops:

  * auto_fix / use_ai stored as a STRING ("Yes"/"No") instead of a boolean
    (any non-empty string is truthy in Python, so it routes/fixes wrongly)
  * doc_type that matches no validator
  * rule_type no validator dispatches on (e.g. 'Layout')
  * check_value with no validator branch

Usage:
  python3 dump_style_rules.py                 # on a box with SharePoint creds
  python3 rule_doctor.py rules_snapshot.json  # analyse real rules offline

With no argument it analyses a built-in reference set.
"""
import json
import sys

from rule_registry import classify, handled_by


def diagnose(rule):
    """Return (problems, warnings) for one rule dict."""
    problems, warnings = [], []

    # Boolean fields stored as strings — the most insidious failure.
    for field in ("auto_fix", "use_ai"):
        val = rule.get(field)
        if val is not None and not isinstance(val, bool):
            problems.append(
                f"{field}={val!r} is {type(val).__name__}, not a boolean — "
                f"any non-empty string is truthy, so this routes/fixes wrongly"
            )

    kind, detail = classify(rule)
    if kind == "gap":
        problems.append(detail)

    if not rule.get("title"):
        warnings.append("no Title — reports will show 'Unknown'")
    return problems, warnings


def _reference_rules():
    return [
        {"title": "British spelling", "rule_type": "Language", "doc_type": "Word",
         "check_value": "BritishSpelling_color", "auto_fix": True, "use_ai": False},
        {"title": "Visio margins", "rule_type": "Position", "doc_type": "Visio",
         "check_value": "LeftMargin", "auto_fix": True, "use_ai": False},
        {"title": "AI-routed tone rule", "rule_type": "Language", "doc_type": "All",
         "check_value": "ProfessionalTone", "auto_fix": False, "use_ai": True},
        # deliberately broken examples:
        {"title": "Broken bool", "rule_type": "Punctuation", "doc_type": "Word",
         "check_value": "PercentSymbol", "auto_fix": "No", "use_ai": "No"},
        {"title": "Layout rule, no validator, no AI", "rule_type": "Layout", "doc_type": "All",
         "check_value": "HyperlinksWorking", "auto_fix": False, "use_ai": False},
    ]


def run(rules, source):
    print(f"\nRule doctor — analysing {len(rules)} rules from {source}\n")
    det = ai = 0
    flagged = []
    word_applicable = 0
    for rule in rules:
        problems, warnings = diagnose(rule)
        if problems:
            flagged.append((rule.get("title") or "(untitled)", problems, warnings))
            continue
        kind, _ = classify(rule)
        if kind == "deterministic":
            det += 1
            if "word" in handled_by(rule):
                word_applicable += 1
        elif kind == "ai":
            ai += 1

    if flagged:
        print("  Rules that will NOT work as intended:\n")
        for title, problems, warnings in flagged:
            print(f"  ✗ {title}")
            for p in problems:
                print(f"      - {p}")
            for w in warnings:
                print(f"      ~ {w}")
            print()

    print(f"  {det} rule(s) handled deterministically ({word_applicable} apply to Word documents).")
    print(f"  {ai} rule(s) routed to the AI path (UseAI).")
    if flagged:
        print(f"  {len(flagged)} rule(s) flagged above — silently do nothing until fixed.")
        return 1
    print("  No gaps: every rule is handled deterministically or by the AI path.")
    return 0


def main():
    if len(sys.argv) > 1:
        try:
            with open(sys.argv[1]) as f:
                rules = json.load(f)
        except Exception as e:
            print(f"Could not read {sys.argv[1]}: {e}", file=sys.stderr)
            return 2
        source = sys.argv[1]
    else:
        rules, source = _reference_rules(), "built-in reference set (no file given)"
    return run(rules, source)


if __name__ == "__main__":
    sys.exit(main())
