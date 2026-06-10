"""
Enhanced validation functions for all style rules
"""
import re
import logging


def iter_all_paragraphs(container):
    """Return every paragraph in the container, descending into table cells
    (and nested tables).

    python-docx's ``iter_all_paragraphs(doc)`` only yields top-level body paragraphs and
    silently skips anything inside tables. Many Mace documents (e.g. activity
    guides) hold all their content in tables, so checkers that walked only
    ``iter_all_paragraphs(doc)`` never saw that text. Horizontally/vertically merged cells
    expose the same underlying cell more than once, so we de-duplicate by the
    cell's XML element to avoid double-counting.
    """
    paras = list(getattr(container, 'paragraphs', []))
    seen = set()
    for table in getattr(container, 'tables', []):
        for row in table.rows:
            for cell in row.cells:
                tc_id = id(cell._tc)
                if tc_id in seen:
                    continue
                seen.add(tc_id)
                paras.extend(iter_all_paragraphs(cell))
    return paras

# ============================================
# LANGUAGE VALIDATORS
# ============================================

# British vs American spelling mappings
BRITISH_SPELLINGS = {
    # Common American -> British replacements
    'color': 'colour',
    'colors': 'colours',
    'colored': 'coloured',
    'coloring': 'colouring',
    'center': 'centre',
    'centers': 'centres',
    'centered': 'centred',
    'analyze': 'analyse',
    'analyzes': 'analyses',
    'analyzed': 'analysed',
    'analyzing': 'analysing',
    'organization': 'organisation',
    'organizations': 'organisations',
    'organize': 'organise',
    'organizes': 'organises',
    'organized': 'organised',
    'organizing': 'organising',
    'aluminum': 'aluminium',
    'license': 'licence',  # Note: verb 'license' stays the same
    'harbor': 'harbour',
    'harbors': 'harbours',
    'finalize': 'finalise',
    'finalizes': 'finalises',
    'finalized': 'finalised',
    'finalizing': 'finalising',
    'labor': 'labour',
    'catalog': 'catalogue',
    'catalogs': 'catalogues',
    'defense': 'defence',
    'minimize': 'minimise',
    'minimizes': 'minimises',
    'minimized': 'minimised',
    'minimizing': 'minimising',
    'utilize': 'utilise',
    'utilizes': 'utilises',
    'utilized': 'utilised',
    'utilizing': 'utilising',
    'fiber': 'fibre',
    'fibers': 'fibres',
    'theater': 'theatre',
    'theaters': 'theatres',
    'authorize': 'authorise',
    'authorizes': 'authorises',
    'authorized': 'authorised',
    'authorizing': 'authorising',
    'summarize': 'summarise',
    'summarizes': 'summarises',
    'summarized': 'summarised',
    'summarizing': 'summarising',
    'recognize': 'recognise',
    'recognizes': 'recognises',
    'recognized': 'recognised',
    'realize': 'realise',
    'realizes': 'realises',
    'realized': 'realised',
    'prioritize': 'prioritise',
    'prioritized': 'prioritised',
    'standardize': 'standardise',
    'standardized': 'standardised',
    'optimize': 'optimise',
    'optimized': 'optimised',
    'specialize': 'specialise',
    'specialized': 'specialised',
    'maximize': 'maximise',
    'minimise': 'minimise',
    'customize': 'customise',
    'customized': 'customised',
    'program': 'programme',  # Mace house style prefers 'programme'
    'programs': 'programmes',
    'meter': 'metre',
    'meters': 'metres',
    'liter': 'litre',
    'liters': 'litres',
    'modeling': 'modelling',
    'modeled': 'modelled',
    'traveling': 'travelling',
    'traveled': 'travelled',
    'enrollment': 'enrolment',
}

# Reverse map: British word -> list of American spellings that should be
# corrected to it. Used when a rule stores the British (target) word in its
# CheckValue/ExpectedValue rather than the American one.
_AMERICAN_FOR_BRITISH = {}
for _am, _br in BRITISH_SPELLINGS.items():
    _AMERICAN_FOR_BRITISH.setdefault(_br.lower(), []).append(_am)

# Contractions to expand
CONTRACTIONS = {
    "can't": "cannot",
    "couldn't": "could not",
    "didn't": "did not",
    "don't": "do not",
    "doesn't": "does not",
    "hasn't": "has not",
    "haven't": "have not",
    "isn't": "is not",
    "shouldn't": "should not",
    "won't": "will not",
    "wouldn't": "would not",
    "aren't": "are not",
    "wasn't": "was not",
    "weren't": "were not",
    "we're": "we are",
    "they're": "they are",
    "you're": "you are",
    "it's": "it is",
    "that's": "that is",
    "there's": "there is",
    "could've": "could have",
    "should've": "should have",
    "would've": "would have",
}

def _resolve_spelling_rule(rule):
    """Work out the American form(s) to search for and the British replacement.

    Rules in the list may store EITHER the American word (e.g.
    'BritishSpelling_color') OR — as in the live Mace list — the British target
    word ('BritishSpelling_programme'). Returns (american_forms, british_word)
    or (None, None) when no mapping is known, so we never "correct" a word to
    itself (which produced false positives counting correct British words as
    American)."""
    suffix = rule['check_value'].replace('BritishSpelling_', '')
    expected = rule.get('expected_value') or ''
    key = suffix.lower()

    if key in BRITISH_SPELLINGS:
        # Suffix is the American word.
        return [suffix], (expected or BRITISH_SPELLINGS[key])
    if key in _AMERICAN_FOR_BRITISH:
        # Suffix is the British target word — search for its American form(s).
        return _AMERICAN_FOR_BRITISH[key], (expected or suffix)
    ekey = expected.lower()
    if ekey in _AMERICAN_FOR_BRITISH:
        return _AMERICAN_FOR_BRITISH[ekey], expected
    return None, None


def check_british_spelling(doc, rule):
    """Check and fix American spellings"""
    issues = []
    fixes = []
    changes = []

    american_forms, british_word = _resolve_spelling_rule(rule)
    if not american_forms:
        logging.info(f"No American-spelling mapping for rule '{rule.get('check_value')}' — skipping")
        return {'issues': issues, 'fixes': fixes, 'changes': changes}

    def replace_preserve_case(match):
        word = match.group(0)
        if word.isupper():
            return british_word.upper()
        elif word[0].isupper():
            return british_word.capitalize()
        return british_word

    issue_count = 0
    fix_count = 0

    for para_idx, paragraph in enumerate(iter_all_paragraphs(doc)):
        for run in paragraph.runs:
            if not run.text:
                continue
            for american_word in american_forms:
                pattern = r'\b' + re.escape(american_word) + r'\b'
                matches = re.findall(pattern, run.text, re.IGNORECASE)
                if not matches:
                    continue
                issue_count += len(matches)
                if rule['auto_fix']:
                    before = run.text
                    run.text = re.sub(pattern, replace_preserve_case, run.text, flags=re.IGNORECASE)
                    changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                    fix_count += len(matches)

    if issue_count > 0:
        issues.append(f"Found {issue_count} instances of American spelling (use '{british_word}')")
    if fix_count > 0:
        fixes.append(f"Fixed {fix_count} instances to British spelling '{british_word}'")

    return {'issues': issues, 'fixes': fixes, 'changes': changes}

# CheckValue stores the contraction with the apostrophe stripped
# (e.g. 'NoContraction_shouldnt'), so map the stripped form back to the
# canonical apostrophe form used in real text.
_APOSTROPHELESS_CONTRACTIONS = {k.replace("'", ""): k for k in CONTRACTIONS}

# Apostrophe variants that appear in Word documents: straight (') and the
# typographic/curly right single quote (’) Word auto-substitutes.
_APOSTROPHES = "'’"


def _contraction_pattern(canonical):
    """Word-bounded, case-insensitive regex matching a contraction with either
    a straight or a curly apostrophe."""
    base = re.escape(canonical).replace("'", f"[{_APOSTROPHES}]")
    return re.compile(r'\b' + base + r'\b', re.IGNORECASE)


def check_contractions(doc, rule):
    """Check and fix contractions"""
    issues = []
    fixes = []
    changes = []

    stripped = rule['check_value'].replace('NoContraction_', '')
    canonical = _APOSTROPHELESS_CONTRACTIONS.get(stripped.lower(), stripped)
    expanded = rule['expected_value']
    pattern = _contraction_pattern(canonical)

    issue_count = 0
    fix_count = 0

    def _expand(match):
        # Preserve a leading capital (e.g. start of sentence)
        return expanded.capitalize() if match.group(0)[0].isupper() else expanded

    # Check all paragraphs (including those inside tables)
    for para_idx, paragraph in enumerate(iter_all_paragraphs(doc)):
        for run in paragraph.runs:
            if not run.text:
                continue
            matches = pattern.findall(run.text)
            if matches:
                issue_count += len(matches)
                if rule['auto_fix']:
                    before = run.text
                    run.text = pattern.sub(_expand, run.text)
                    changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                    fix_count += len(matches)

    if issue_count > 0:
        issues.append(f"Found {issue_count} instances of contraction '{canonical}'")
    if fix_count > 0:
        fixes.append(f"Fixed {fix_count} contractions to '{expanded}'")

    return {'issues': issues, 'fixes': fixes, 'changes': changes}

def check_word_choice(doc, rule):
    """Check word choice violations"""
    issues = []
    fixes = []
    changes = []

    check_value = rule['check_value']

    # Handle specific word choice rules
    if check_value == 'Word_toward':
        # Replace 'towards' with 'toward'
        issue_count = 0
        fix_count = 0

        for para_idx, paragraph in enumerate(iter_all_paragraphs(doc)):
            for run in paragraph.runs:
                if run.text and 'towards' in run.text.lower():
                    matches = len(re.findall(r'\btowards\b', run.text, re.IGNORECASE))
                    issue_count += matches

                    if rule['auto_fix']:
                        before = run.text
                        run.text = re.sub(r'\btowards\b', 'toward', run.text, flags=re.IGNORECASE)
                        changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                        fix_count += matches

        if issue_count > 0:
            issues.append(f"Found {issue_count} instances of 'towards'")
        if fix_count > 0:
            fixes.append(f"Fixed {fix_count} instances to 'toward'")

    elif check_value == 'AvoidEtc':
        # Flag usage of 'etc.'
        issue_count = 0

        for paragraph in iter_all_paragraphs(doc):
            for run in paragraph.runs:
                if run.text and 'etc.' in run.text.lower():
                    matches = len(re.findall(r'\betc\.?\b', run.text, re.IGNORECASE))
                    issue_count += matches

        if issue_count > 0:
            issues.append(f"Found {issue_count} instances of 'etc.' - be specific instead")

    return {'issues': issues, 'fixes': fixes, 'changes': changes}

# ============================================
# PUNCTUATION VALIDATORS
# ============================================

def check_symbols(doc, rule):
    """Check and fix symbol usage"""
    issues = []
    fixes = []
    changes = []

    check_value = rule['check_value']

    if check_value == 'NoAmpersand':
        # Replace & with 'and'
        issue_count = 0
        fix_count = 0

        for para_idx, paragraph in enumerate(iter_all_paragraphs(doc)):
            for run in paragraph.runs:
                if run.text and '&' in run.text:
                    # Count ampersands (exclude &nbsp; and other HTML entities)
                    matches = len([c for c in run.text if c == '&'])
                    issue_count += matches

                    if rule['auto_fix']:
                        before = run.text
                        run.text = run.text.replace('&', 'and')
                        changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                        fix_count += matches

        if issue_count > 0:
            issues.append(f"Found {issue_count} ampersands (&)")
        if fix_count > 0:
            fixes.append(f"Fixed {fix_count} ampersands to 'and'")

    elif check_value == 'PercentSymbol':
        # Replace % with 'percent' in text (not in numbers like "85%")
        issue_count = 0
        fix_count = 0

        for para_idx, paragraph in enumerate(iter_all_paragraphs(doc)):
            for run in paragraph.runs:
                if run.text and '%' in run.text:
                    # Find number% patterns
                    matches = re.findall(r'\d+%', run.text)
                    issue_count += len(matches)

                    if rule['auto_fix']:
                        before = run.text
                        # Replace number% with number percent
                        run.text = re.sub(r'(\d+)%', r'\1 percent', run.text)
                        changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                        fix_count += len(matches)

        if issue_count > 0:
            issues.append(f"Found {issue_count} percent symbols (%)")
        if fix_count > 0:
            fixes.append(f"Fixed {fix_count} percent symbols to 'percent'")

    elif check_value == 'NoApostrophePlurals':
        # Detect incorrect apostrophes in plurals (e.g., CD's, SME's)
        issue_count = 0

        for paragraph in iter_all_paragraphs(doc):
            for run in paragraph.runs:
                if run.text:
                    # Pattern: word ending with 's followed by 's or other letters
                    # This is a simplified check
                    matches = re.findall(r"\b[A-Z]{2,}'s\b", run.text)  # e.g., CD's, SME's
                    issue_count += len(matches)

        if issue_count > 0:
            issues.append(f"Found {issue_count} incorrect apostrophes in plurals (e.g., CD's should be CDs)")

    return {'issues': issues, 'fixes': fixes, 'changes': changes}

def check_numbers(doc, rule):
    """Check number formatting"""
    issues = []
    fixes = []
    changes = []

    check_value = rule['check_value']

    if check_value == 'NumberCommas':
        # Check for numbers 1000+ without commas
        issue_count = 0
        fix_count = 0

        for para_idx, paragraph in enumerate(iter_all_paragraphs(doc)):
            for run in paragraph.runs:
                if run.text:
                    # Find numbers with 4+ digits without commas
                    matches = re.findall(r'\b\d{4,}\b', run.text)
                    # Filter out years (1900-2099)
                    matches = [m for m in matches if not (1900 <= int(m) <= 2099)]
                    issue_count += len(matches)

                    if rule['auto_fix'] and matches:
                        before = run.text
                        # Add commas to numbers
                        for match in matches:
                            formatted = '{:,}'.format(int(match))
                            run.text = run.text.replace(match, formatted)
                            fix_count += 1
                        changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})

        if issue_count > 0:
            issues.append(f"Found {issue_count} numbers missing commas")
        if fix_count > 0:
            fixes.append(f"Added commas to {fix_count} numbers")

    return {'issues': issues, 'fixes': fixes, 'changes': changes}

# ============================================
# MAIN DISPATCHER
# ============================================

def validate_language_rules(doc, rule):
    """Dispatch language rule validation"""
    check_value = rule['check_value']

    if check_value.startswith('BritishSpelling_'):
        return check_british_spelling(doc, rule)
    elif check_value.startswith('NoContraction_'):
        return check_contractions(doc, rule)
    elif check_value in ['Word_toward', 'AvoidEtc', 'AvoidShould']:
        return check_word_choice(doc, rule)
    else:
        logging.warning(f"Unknown language check: {check_value}")
        return {'issues': [], 'fixes': []}

def validate_punctuation_rules(doc, rule):
    """Dispatch punctuation rule validation"""
    check_value = rule['check_value']

    if check_value in ['NoAmpersand', 'PercentSymbol', 'NoApostrophePlurals']:
        return check_symbols(doc, rule)
    elif check_value == 'NumberCommas':
        return check_numbers(doc, rule)
    else:
        logging.info(f"Punctuation check '{check_value}' not yet implemented")
        return {'issues': [], 'fixes': []}

def validate_grammar_rules(doc, rule):
    """Dispatch grammar rule validation"""
    check_value = rule['check_value']

    if check_value.startswith('NoContraction_'):
        return check_contractions(doc, rule)
    else:
        logging.info(f"Grammar check '{check_value}' not yet implemented")
        return {'issues': [], 'fixes': []}
