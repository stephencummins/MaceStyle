"""Offline rule-coverage regression test.

Proves the validation ENGINE actually fires for each implemented Word
check_value, by building a .docx that contains a known violation for every
rule and asserting the validator reports it. No SharePoint / Azure / network.

Why this exists: the previous local test (test_local.py) only exercised 2 of
the 6 rule types with idealised data, so it passed while real documents failed.
This covers every implemented check, and FAILS LOUDLY (exit 1) if any rule
stops catching its violation — so it works as a regression gate.

Run:  python3 test_rule_coverage.py
Pair with rule_doctor.py, which checks your REAL SharePoint rules are shaped
to actually reach this engine.
"""
import sys
from io import BytesIO

from docx import Document
from docx.shared import RGBColor

from ValidateDocument.word_validator import validate_word_document


# Each case: a synthetic rule (in the exact shape fetch_validation_rules emits)
# plus document text that genuinely violates it. auto_fix=False so every
# violation surfaces as an ISSUE we can assert on by title.
CASES = [
    {
        "title": "British spelling (color->colour)",
        "rule": {"title": "British spelling (color->colour)", "rule_type": "Language",
                 "doc_type": "Word", "check_value": "BritishSpelling_color",
                 "expected_value": "colour", "auto_fix": False, "use_ai": False, "priority": 10},
        "text": "The color of the cladding was approved.",
    },
    {
        "title": "No contraction (don't)",
        "rule": {"title": "No contraction (don't)", "rule_type": "Language",
                 "doc_type": "Word", "check_value": "NoContraction_dont",
                 "expected_value": "do not", "auto_fix": False, "use_ai": False, "priority": 10},
        "text": "We don't accept late submissions.",
    },
    {
        "title": "Word choice (towards->toward)",
        "rule": {"title": "Word choice (towards->toward)", "rule_type": "Language",
                 "doc_type": "Word", "check_value": "Word_toward",
                 "expected_value": "toward", "auto_fix": False, "use_ai": False, "priority": 10},
        "text": "Progress was made towards completion.",
    },
    {
        "title": "Avoid etc.",
        "rule": {"title": "Avoid etc.", "rule_type": "Language",
                 "doc_type": "Word", "check_value": "AvoidEtc",
                 "expected_value": "", "auto_fix": False, "use_ai": False, "priority": 10},
        "text": "Bring drawings, specifications, etc. to the meeting.",
    },
    {
        "title": "No ampersand (&)",
        "rule": {"title": "No ampersand (&)", "rule_type": "Punctuation",
                 "doc_type": "Word", "check_value": "NoAmpersand",
                 "expected_value": "and", "auto_fix": False, "use_ai": False, "priority": 10},
        "text": "Design & build was selected.",
    },
    {
        "title": "Percent symbol (%)",
        "rule": {"title": "Percent symbol (%)", "rule_type": "Punctuation",
                 "doc_type": "Word", "check_value": "PercentSymbol",
                 "expected_value": "percent", "auto_fix": False, "use_ai": False, "priority": 10},
        "text": "The works are 85% complete.",
    },
    {
        "title": "No apostrophe plurals (CD's)",
        "rule": {"title": "No apostrophe plurals (CD's)", "rule_type": "Punctuation",
                 "doc_type": "Word", "check_value": "NoApostrophePlurals",
                 "expected_value": "", "auto_fix": False, "use_ai": False, "priority": 10},
        "text": "Several CD's were issued to the team.",
    },
    {
        "title": "Number commas (1000000)",
        "rule": {"title": "Number commas (1000000)", "rule_type": "Punctuation",
                 "doc_type": "Word", "check_value": "NumberCommas",
                 "expected_value": "", "auto_fix": False, "use_ai": False, "priority": 10},
        "text": "The budget is 1000000 pounds for this phase.",
    },
    {
        "title": "All text font (Arial)",
        "rule": {"title": "All text font (Arial)", "rule_type": "Font",
                 "doc_type": "Word", "check_value": "AllTextFont",
                 "expected_value": "Arial", "auto_fix": False, "use_ai": False, "priority": 10},
        "text": "__FONT_TIMES__This sentence is in Times New Roman.",
    },
]


def _build_document():
    """One .docx containing every violation above, plus a wrong-font/colour H1."""
    doc = Document()

    heading = doc.add_heading("Project Report", level=1)
    if heading.runs:
        heading.runs[0].font.name = "Times New Roman"        # Heading1Font violation
        heading.runs[0].font.color.rgb = RGBColor(255, 0, 0)  # Heading1Color violation

    for case in CASES:
        text = case["text"]
        if text.startswith("__FONT_TIMES__"):
            p = doc.add_paragraph()
            run = p.add_run(text.replace("__FONT_TIMES__", ""))
            run.font.name = "Times New Roman"
        else:
            doc.add_paragraph(text)

    stream = BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream


def _heading_cases():
    """Heading font + colour rules, asserted against the H1 built above."""
    return [
        {"title": "Heading 1 font (Arial)",
         "rule": {"title": "Heading 1 font (Arial)", "rule_type": "Font", "doc_type": "Word",
                  "check_value": "Heading1Font", "expected_value": "Arial",
                  "auto_fix": False, "use_ai": False, "priority": 5}},
        {"title": "Heading 1 colour (Mace blue)",
         "rule": {"title": "Heading 1 colour (Mace blue)", "rule_type": "Color", "doc_type": "Word",
                  "check_value": "Heading1Color", "expected_value": "0,51,153",
                  "auto_fix": False, "use_ai": False, "priority": 6}},
    ]


def run():
    all_cases = CASES + _heading_cases()
    rules = [c["rule"] for c in all_cases]

    result = validate_word_document(_build_document(), rules)

    # A rule "fired" if it produced an issue or a fix referencing its title.
    fired = {i.get("rule_name") for i in result["issues"]}
    fired |= {f.get("rule_name") for f in result["fixes_applied"]}

    print("\nRule coverage — does each implemented check actually fire?\n")
    print(f"  {'RESULT':6}  {'CHECK_VALUE':24}  RULE")
    print(f"  {'-'*6}  {'-'*24}  {'-'*30}")
    missed = []
    for case in all_cases:
        title = case["title"]
        cv = case["rule"]["check_value"]
        ok = title in fired
        print(f"  {'PASS' if ok else 'FAIL':6}  {cv:24}  {title}")
        if not ok:
            missed.append(title)

    print(f"\n  {len(all_cases) - len(missed)}/{len(all_cases)} rules fired.")
    if missed:
        print("\n  ✗ These rules did NOT catch their violation:")
        for m in missed:
            print(f"      - {m}")
        print("\n  The engine is not applying these checks. Investigate the matching"
              "\n  validator in enhanced_validators.py / word_validator.py.")
        return 1
    print("\n  ✓ Every implemented check caught its violation. The engine works;"
          "\n    if a real rule isn't firing, run rule_doctor.py on the real rules.")
    return 0


if __name__ == "__main__":
    sys.exit(run())
