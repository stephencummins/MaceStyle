"""
Enhanced validation functions for all style rules
"""
import re
import logging

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
}

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

def check_british_spelling(doc, rule):
    """Check and fix American spellings"""
    issues = []
    fixes = []

    american_word = rule['check_value'].replace('BritishSpelling_', '')
    british_word = rule['expected_value']

    issue_count = 0
    fix_count = 0

    # Check all paragraphs
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run.text:
                # Use word boundaries to avoid partial matches
                pattern = r'\b' + re.escape(american_word) + r'\b'
                matches = re.findall(pattern, run.text, re.IGNORECASE)

                if matches:
                    issue_count += len(matches)

                    if rule['auto_fix']:
                        # Replace while preserving case
                        def replace_preserve_case(match):
                            word = match.group(0)
                            if word.isupper():
                                return british_word.upper()
                            elif word[0].isupper():
                                return british_word.capitalize()
                            else:
                                return british_word

                        run.text = re.sub(pattern, replace_preserve_case, run.text, flags=re.IGNORECASE)
                        fix_count += len(matches)

    if issue_count > 0:
        issues.append(f"Found {issue_count} instances of American spelling '{american_word}'")
    if fix_count > 0:
        fixes.append(f"Fixed {fix_count} instances to British spelling '{british_word}'")

    return {'issues': issues, 'fixes': fixes}

def check_contractions(doc, rule):
    """Check and fix contractions"""
    issues = []
    fixes = []

    contraction = rule['check_value'].replace('NoContraction_', '')
    expanded = rule['expected_value']

    issue_count = 0
    fix_count = 0

    # Check all paragraphs
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run.text and contraction in run.text:
                # Count occurrences
                count = run.text.count(contraction)
                issue_count += count

                if rule['auto_fix']:
                    run.text = run.text.replace(contraction, expanded)
                    fix_count += count

    if issue_count > 0:
        issues.append(f"Found {issue_count} instances of contraction '{contraction}'")
    if fix_count > 0:
        fixes.append(f"Fixed {fix_count} contractions to '{expanded}'")

    return {'issues': issues, 'fixes': fixes}

def check_word_choice(doc, rule):
    """Check word choice violations"""
    issues = []
    fixes = []

    check_value = rule['check_value']

    # Handle specific word choice rules
    if check_value == 'Word_toward':
        # Replace 'towards' with 'toward'
        issue_count = 0
        fix_count = 0

        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if run.text and 'towards' in run.text.lower():
                    matches = len(re.findall(r'\btowards\b', run.text, re.IGNORECASE))
                    issue_count += matches

                    if rule['auto_fix']:
                        run.text = re.sub(r'\btowards\b', 'toward', run.text, flags=re.IGNORECASE)
                        fix_count += matches

        if issue_count > 0:
            issues.append(f"Found {issue_count} instances of 'towards'")
        if fix_count > 0:
            fixes.append(f"Fixed {fix_count} instances to 'toward'")

    elif check_value == 'AvoidEtc':
        # Flag usage of 'etc.'
        issue_count = 0

        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if run.text and 'etc.' in run.text.lower():
                    matches = len(re.findall(r'\betc\.?\b', run.text, re.IGNORECASE))
                    issue_count += matches

        if issue_count > 0:
            issues.append(f"Found {issue_count} instances of 'etc.' - be specific instead")

    return {'issues': issues, 'fixes': fixes}

# ============================================
# PUNCTUATION VALIDATORS
# ============================================

def check_symbols(doc, rule):
    """Check and fix symbol usage"""
    issues = []
    fixes = []

    check_value = rule['check_value']

    if check_value == 'NoAmpersand':
        # Replace & with 'and'
        issue_count = 0
        fix_count = 0

        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if run.text and '&' in run.text:
                    # Count ampersands (exclude &nbsp; and other HTML entities)
                    matches = len([c for c in run.text if c == '&'])
                    issue_count += matches

                    if rule['auto_fix']:
                        run.text = run.text.replace('&', 'and')
                        fix_count += matches

        if issue_count > 0:
            issues.append(f"Found {issue_count} ampersands (&)")
        if fix_count > 0:
            fixes.append(f"Fixed {fix_count} ampersands to 'and'")

    elif check_value == 'PercentSymbol':
        # Replace % with 'percent' in text (not in numbers like "85%")
        issue_count = 0
        fix_count = 0

        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if run.text and '%' in run.text:
                    # Find number% patterns
                    matches = re.findall(r'\d+%', run.text)
                    issue_count += len(matches)

                    if rule['auto_fix']:
                        # Replace number% with number percent
                        run.text = re.sub(r'(\d+)%', r'\1 percent', run.text)
                        fix_count += len(matches)

        if issue_count > 0:
            issues.append(f"Found {issue_count} percent symbols (%)")
        if fix_count > 0:
            fixes.append(f"Fixed {fix_count} percent symbols to 'percent'")

    elif check_value == 'NoApostrophePlurals':
        # Detect incorrect apostrophes in plurals (e.g., CD's, SME's)
        issue_count = 0

        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if run.text:
                    # Pattern: word ending with 's followed by 's or other letters
                    # This is a simplified check
                    matches = re.findall(r"\b[A-Z]{2,}'s\b", run.text)  # e.g., CD's, SME's
                    issue_count += len(matches)

        if issue_count > 0:
            issues.append(f"Found {issue_count} incorrect apostrophes in plurals (e.g., CD's should be CDs)")

    return {'issues': issues, 'fixes': fixes}

def check_numbers(doc, rule):
    """Check number formatting"""
    issues = []
    fixes = []

    check_value = rule['check_value']

    if check_value == 'NumberCommas':
        # Check for numbers 1000+ without commas
        issue_count = 0
        fix_count = 0

        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if run.text:
                    # Find numbers with 4+ digits without commas
                    matches = re.findall(r'\b\d{4,}\b', run.text)
                    # Filter out years (1900-2099)
                    matches = [m for m in matches if not (1900 <= int(m) <= 2099)]
                    issue_count += len(matches)

                    if rule['auto_fix'] and matches:
                        # Add commas to numbers
                        for match in matches:
                            formatted = '{:,}'.format(int(match))
                            run.text = run.text.replace(match, formatted)
                            fix_count += 1

        if issue_count > 0:
            issues.append(f"Found {issue_count} numbers missing commas")
        if fix_count > 0:
            fixes.append(f"Added commas to {fix_count} numbers")

    return {'issues': issues, 'fixes': fixes}

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
