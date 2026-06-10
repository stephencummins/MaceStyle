"""Rule doctor — will my style rules actually fire?

Static analysis of the 'Style Rules' data against what the validators really
implement. Catches the silent failure modes that make a rule a no-op even
though it looks correct in SharePoint:

  * auto_fix / use_ai stored as a STRING ("Yes"/"No") instead of a boolean.
    Python treats any non-empty string as True, so a "No" auto_fix still
    triggers a fix, and a "No" UseAI still routes the rule to the AI path
    instead of the deterministic one.
  * doc_type that no validator matches. Word documents only see rules whose
    doc_type is Word / Both / All — anything else is silently dropped.
  * rule_type the dispatcher does not recognise — silently dropped.
  * check_value with no matching validator branch — silently skipped (the code
    logs "Unknown ... check" and moves on).

Usage:
  python3 dump_style_rules.py                 # on a box with SharePoint creds
  python3 rule_doctor.py rules_snapshot.json  # analyse the real rules offline

With no argument it analyses a built-in reference set so you can see what a
healthy result looks like.
"""
import json
import sys

# --- what the Word validator actually dispatches on (see word_validator.py
#     and enhanced_validators.py). Keep in sync if new checks are added. ---
WORD_RULE_TYPES = {"Font", "Color", "Language", "Grammar", "Punctuation", "Capitalisation"}
KNOWN_RULE_TYPES = WORD_RULE_TYPES | {"Structure", "Layout", "Page", "Font/Size"}
WORD_DOC_TYPES = {"Word", "Both", "All"}
KNOWN_DOC_TYPES = {"Word", "Visio", "Excel", "PowerPoint", "Both", "All"}

EXACT_CHECKS = {
    "Font": {"AllTextFont", "Heading1Font"},
    "Color": {"Heading1Color"},
    "Language": {"Word_toward", "AvoidEtc", "AvoidShould", "ProximityRedundant",
                 "NoMinMaxApprox", "ForecastPastTense", "Constructability"},
    "Punctuation": {"NoAmpersand", "PercentSymbol", "NoApostrophePlurals", "NumberCommas",
                    "NoDoubleSpaces", "NoHyphenInSitu", "NoHyphenOffOn", "AvoidAndOr"},
    "Capitalisation": {"ReferenceCodeCase"},
    "Grammar": set(),
}
PREFIX_CHECKS = {
    "Language": ("BritishSpelling_", "NoContraction_", "PhraseReplace_"),
    "Grammar": ("NoContraction_",),
}


def _check_value_implemented(rule_type, check_value):
    if rule_type not in EXACT_CHECKS:
        return False
    if check_value in EXACT_CHECKS[rule_type]:
        return True
    return any(check_value.startswith(p) for p in PREFIX_CHECKS.get(rule_type, ()))


def diagnose(rule):
    """Return (problems, warnings) for one rule dict."""
    problems, warnings = [], []
    rule_type = rule.get("rule_type")
    doc_type = rule.get("doc_type")
    check_value = rule.get("check_value") or ""

    # 1. Boolean fields stored as strings (the most insidious failure).
    for field in ("auto_fix", "use_ai"):
        val = rule.get(field)
        if val is not None and not isinstance(val, bool):
            problems.append(
                f"{field}={val!r} is {type(val).__name__}, not a boolean — "
                f"any non-empty string is truthy, so this routes/fixes wrongly"
            )

    # 2. doc_type must match a validator.
    if doc_type not in KNOWN_DOC_TYPES:
        problems.append(f"doc_type={doc_type!r} matches no validator — rule never runs")

    # 3. rule_type must be recognised.
    if rule_type not in KNOWN_RULE_TYPES:
        problems.append(f"rule_type={rule_type!r} not recognised by the dispatcher")

    # 4. check_value must have a validator branch — but only for the hard-coded
    #    path. use_ai rules are applied by Claude from the rule text and never
    #    touch the check_value dispatch, so an "unimplemented" check_value is
    #    irrelevant for them. (Mirror the engine's truthiness on use_ai.)
    is_ai = bool(rule.get("use_ai"))
    if doc_type in WORD_DOC_TYPES and rule_type in WORD_RULE_TYPES and not is_ai:
        if not _check_value_implemented(rule_type, check_value):
            problems.append(
                f"check_value={check_value!r} has no validator branch for "
                f"rule_type {rule_type!r} — silently skipped"
            )
    elif rule_type in WORD_RULE_TYPES and doc_type not in WORD_DOC_TYPES:
        warnings.append(f"Word-style rule_type {rule_type!r} but doc_type {doc_type!r} "
                        f"— will not apply to Word documents")

    if not rule.get("title"):
        warnings.append("no Title — reports will show 'Unknown'")

    return problems, warnings


def _reference_rules():
    return [
        {"title": "British spelling", "rule_type": "Language", "doc_type": "Word",
         "check_value": "BritishSpelling_color", "expected_value": "colour",
         "auto_fix": True, "use_ai": False, "priority": 10},
        {"title": "Ampersand", "rule_type": "Punctuation", "doc_type": "Both",
         "check_value": "NoAmpersand", "expected_value": "and",
         "auto_fix": True, "use_ai": False, "priority": 20},
        # Two deliberately-broken examples so the demo shows what failures look like:
        {"title": "Broken bool", "rule_type": "Punctuation", "doc_type": "Word",
         "check_value": "PercentSymbol", "expected_value": "percent",
         "auto_fix": "No", "use_ai": "No", "priority": 30},
        {"title": "Wrong doc_type", "rule_type": "Language", "doc_type": "Document",
         "check_value": "Word_toward", "expected_value": "toward",
         "auto_fix": True, "use_ai": False, "priority": 40},
    ]


def run(rules, source):
    print(f"\nRule doctor — analysing {len(rules)} rules from {source}\n")
    ok = 0
    word_applicable = 0
    flagged = []
    for rule in rules:
        problems, warnings = diagnose(rule)
        title = rule.get("title") or "(untitled)"
        if problems:
            flagged.append((title, problems, warnings))
        else:
            ok += 1
            if rule.get("doc_type") in WORD_DOC_TYPES and rule.get("rule_type") in WORD_RULE_TYPES:
                word_applicable += 1

    if flagged:
        print("  Rules that will NOT work as intended:\n")
        for title, problems, warnings in flagged:
            print(f"  ✗ {title}")
            for p in problems:
                print(f"      - {p}")
            for w in warnings:
                print(f"      ~ {w}")
            print()

    print(f"  {ok}/{len(rules)} rules are correctly shaped.")
    print(f"  {word_applicable} of those will actually apply to Word documents.")
    if flagged:
        print(f"  {len(flagged)} rule(s) flagged above — fix the data in the "
              f"SharePoint 'Style Rules' list.")
        return 1
    print("  No problems found.")
    return 0


def main():
    if len(sys.argv) > 1:
        path = sys.argv[1]
        try:
            with open(path) as f:
                rules = json.load(f)
        except Exception as e:
            print(f"Could not read {path}: {e}", file=sys.stderr)
            return 2
        source = path
    else:
        rules = _reference_rules()
        source = "built-in reference set (no file given)"
    return run(rules, source)


if __name__ == "__main__":
    sys.exit(main())
