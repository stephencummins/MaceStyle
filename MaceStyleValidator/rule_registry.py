"""Canonical map of which validator implements which (rule_type, check_value).

Single source of truth shared by rule_doctor.py and test_tracker_rules.py, so
"is this rule actually handled?" is answered the same way everywhere. Keep in
sync with the validators when checks are added.

A rule is *deterministically handled* if a validator that runs for its
doc_type implements its check_value. Otherwise it is only handled if UseAI is
set (the AI path). Otherwise it is a GAP — silently does nothing.
"""

# Word — enhanced_validators.py / word_validator.py
_WORD = {
    "Font": {"AllTextFont", "Heading1Font"},
    "Color": {"Heading1Color"},
    "Language": {"Word_toward", "AvoidEtc", "AvoidShould", "ProximityRedundant", "NoMinMaxApprox",
                 "ForecastPastTense", "Constructability", "NoFeelTechnical", "NoAboveBelow", "PreferMetric"},
    "Punctuation": {"NoAmpersand", "PercentSymbol", "NoApostrophePlurals", "NumberCommas", "NoDoubleSpaces",
                    "NoHyphenInSitu", "NoHyphenOffOn", "AvoidAndOr", "TimeFormat", "DateFormat_Text",
                    "DateFormat_Table", "YearIntervalFormat", "NoSpacesAroundSlash", "AvoidForwardSlash",
                    "HyphenInWords", "HyphenSuffixes", "HyphenAlwaysPrefix", "Hyphen_wide",
                    "PunctuationBeforeEgIe", "OxfordComma", "NumbersBelowTen", "CaptionNoPeriod",
                    "CompoundModifiers"},
    "Grammar": {"NoSentenceStartEgIe", "NoEtcWithEgIe", "ClientNameNotTheClient", "OrgSingular"},
    "Capitalisation": {"ReferenceCodeCase", "ProperNounDerivations", "NoEmphasisCaps"},
}
_WORD_PREFIX = {
    "Language": ("BritishSpelling_", "NoContraction_", "PhraseReplace_"),
    "Grammar": ("NoContraction_",),
}

# Visio — visio_validator.py. Size/PageDimensions dispatch on rule_type and
# accept any check_value (parsed as WxH), so they are marked "*".
_VISIO = {
    "Color": {"ShapeFillColor", "ShapeTextColor"},
    "Font": {"AllTextFont"},
    "Size": {"*"},
    "Position": {"TopMargin", "LeftMargin", "RightMargin", "BottomMargin", "ExactPosition"},
    "PageDimensions": {"*"},
}

# Excel + PowerPoint share the same subset.
_SHEET = {
    "Font": {"AllTextFont"},
    "Language": {"Word_toward", "AvoidEtc"},
    "Punctuation": {"NoAmpersand", "NoApostrophePlurals", "NumberCommas", "PercentSymbol"},
}

VALIDATORS = {
    "word": (_WORD, _WORD_PREFIX),
    "visio": (_VISIO, {}),
    "excel": (_SHEET, {}),
    "powerpoint": (_SHEET, {}),
}

DOC_TYPE_VALIDATORS = {
    "All": ("word", "visio", "excel", "powerpoint"),
    "Both": ("word",),
    "Word": ("word",),
    "Visio": ("visio",),
    "Excel": ("excel",),
    "PowerPoint": ("powerpoint",),
}


def _implements(validator, rule_type, check_value):
    exact, prefix = VALIDATORS[validator]
    cvs = exact.get(rule_type)
    if not cvs:
        return False
    if "*" in cvs or check_value in cvs:
        return True
    return any((check_value or "").startswith(p) for p in prefix.get(rule_type, ()))


def applicable_validators(doc_type):
    return DOC_TYPE_VALIDATORS.get(doc_type, ())


def handled_by(rule):
    """Validators that deterministically implement this rule (may be empty)."""
    dt, rt, cv = rule.get("doc_type"), rule.get("rule_type"), rule.get("check_value") or ""
    return [v for v in applicable_validators(dt) if _implements(v, rt, cv)]


def classify(rule):
    """('deterministic'|'ai'|'gap', detail) — how this rule is (or isn't) handled."""
    hb = handled_by(rule)
    if hb:
        return ("deterministic", "+".join(hb))
    if bool(rule.get("use_ai")):
        return ("ai", "UseAI")
    dt, rt, cv = rule.get("doc_type"), rule.get("rule_type"), rule.get("check_value")
    if dt not in DOC_TYPE_VALIDATORS:
        return ("gap", f"doc_type {dt!r} matches no validator")
    owners = [v for v in applicable_validators(dt) if rt in VALIDATORS[v][0]]
    if not owners:
        return ("gap", f"rule_type {rt!r} handled by no validator")
    return ("gap", f"check_value {cv!r} not implemented for rule_type {rt!r}")
