"""
Enhanced validation functions for all style rules
"""
import re
import logging
import datetime
from copy import deepcopy

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ============================================
# TRACKED-CHANGE HELPERS (suggested changes)
# ============================================
# For rules with auto_fix=False we apply the suggested change as a native Word
# tracked revision (a w:del of the old text + a w:ins of the new). The reviewer
# opens the document in Word and Accepts or Rejects each one — a single click,
# leaving no residue to clean up (unlike a highlight). Confident auto-fixes are
# still applied silently; only suggestions become tracked changes.

_TRACK_AUTHOR = 'Mace Style Validator'
_rev_counter = [1000]


def _rev_id():
    _rev_counter[0] += 1
    return str(_rev_counter[0])


def _mark_revision(el):
    el.set(qn('w:id'), _rev_id())
    el.set(qn('w:author'), _TRACK_AUTHOR)
    el.set(qn('w:date'), datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'))


def _run_shell(src_r):
    """Clone a <w:r>, preserving its formatting (rPr) but stripping any text/
    break children so the caller can attach fresh text."""
    new_r = deepcopy(src_r)
    for child in list(new_r):
        if child.tag in (qn('w:t'), qn('w:delText'), qn('w:tab'), qn('w:br'), qn('w:cr')):
            new_r.remove(child)
    return new_r


def _plain_run(src_r, text):
    new_r = _run_shell(src_r)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    new_r.append(t)
    return new_r


def _tracked_replace_in_run(run, compiled, repl):
    """Apply each genuine change of `compiled` in `run` as a Word tracked
    revision (delete old + insert new), splitting the run so surrounding text
    and formatting are preserved. `repl` is a string or callable(match)->str.
    Unchanged matches are left as normal text. Returns the number of changes.
    """
    text = run.text
    if not text:
        return 0
    parts = []  # ('t', text) | ('c', old, new)
    last = 0
    n = 0
    for m in compiled.finditer(text):
        original = m.group(0)
        replacement = repl(m) if callable(repl) else repl
        if replacement == original:
            continue  # no real change — leave as normal text
        if m.start() > last:
            parts.append(('t', text[last:m.start()]))
        parts.append(('c', original, replacement))
        last = m.end()
        n += 1
    if n == 0:
        return 0
    if last < len(text):
        parts.append(('t', text[last:]))

    r = run._r
    parent = r.getparent()
    idx = list(parent).index(r)
    for part in parts:
        if part[0] == 't':
            if part[1]:
                parent.insert(idx, _plain_run(r, part[1]))
                idx += 1
            continue
        _, old, new = part
        # Tracked deletion of the old text
        del_el = OxmlElement('w:del')
        _mark_revision(del_el)
        del_r = _run_shell(r)
        dtext = OxmlElement('w:delText')
        dtext.set(qn('xml:space'), 'preserve')
        dtext.text = old
        del_r.append(dtext)
        del_el.append(del_r)
        parent.insert(idx, del_el)
        idx += 1
        # Tracked insertion of the new text
        ins_el = OxmlElement('w:ins')
        _mark_revision(ins_el)
        ins_el.append(_plain_run(r, new))
        parent.insert(idx, ins_el)
        idx += 1
    parent.remove(r)
    return n


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
    # --- Additional mappings to cover the full live Mace rule set (2026-06-22).
    # Previously 27 of 51 BritishSpelling_* rules silently no-op'd because their
    # British target had no American->British entry here (Natasha/Jade test doc).
    'analog': 'analogue',
    'analogs': 'analogues',
    'caliber': 'calibre',
    'calibers': 'calibres',
    'characterize': 'characterise',
    'characterizes': 'characterises',
    'characterized': 'characterised',
    'characterizing': 'characterising',
    'dialog': 'dialogue',
    'dialogs': 'dialogues',
    'dike': 'dyke',
    'dikes': 'dykes',
    'enroll': 'enrol',
    'enrolls': 'enrols',
    'equalize': 'equalise',
    'equalizes': 'equalises',
    'equalized': 'equalised',
    'equalizing': 'equalising',
    'fiberglass': 'fibreglass',
    'fulfill': 'fulfil',
    'fulfills': 'fulfils',
    'gray': 'grey',
    'homogenize': 'homogenise',
    'homogenizes': 'homogenises',
    'homogenized': 'homogenised',
    'homogenizing': 'homogenising',
    'instill': 'instil',
    'instills': 'instils',
    'maneuver': 'manoeuvre',
    'maneuvers': 'manoeuvres',
    'maneuvered': 'manoeuvred',
    'maneuvering': 'manoeuvring',
    'mobilize': 'mobilise',
    'mobilizes': 'mobilises',
    'mobilized': 'mobilised',
    'mobilizing': 'mobilising',
    'neighbor': 'neighbour',
    'neighbors': 'neighbours',
    'neighboring': 'neighbouring',
    'neighborhood': 'neighbourhood',
    'neutralize': 'neutralise',
    'neutralizes': 'neutralises',
    'neutralized': 'neutralised',
    'neutralizing': 'neutralising',
    'normalize': 'normalise',
    'normalizes': 'normalises',
    'normalized': 'normalised',
    'normalizing': 'normalising',
    'odor': 'odour',
    'odors': 'odours',
    'recognizing': 'recognising',
    'skillful': 'skilful',
    'skillfully': 'skilfully',
    'spelled': 'spelt',
    'stabilize': 'stabilise',
    'stabilizes': 'stabilises',
    'stabilized': 'stabilised',
    'stabilizing': 'stabilising',
    'sulfide': 'sulphide',
    'sulfides': 'sulphides',
    'tunneling': 'tunnelling',
    'tunneled': 'tunnelled',
    'tire': 'tyre',
    'tires': 'tyres',
    'learned': 'learnt',
    'percent': 'per cent',
    # Deliberately NOT mapped: curb->kerb. "curb" is also the standard verb
    # (curb costs/risk) in both UK and US English, so auto-fixing every "curb"
    # to "kerb" would corrupt correct text. The kerb rule stays a no-op rather
    # than risk false positives.
}

# Reverse map: British word -> list of American spellings that should be
# corrected to it. Used when a rule stores the British (target) word in its
# CheckValue/ExpectedValue rather than the American one.
_AMERICAN_FOR_BRITISH = {}
for _am, _br in BRITISH_SPELLINGS.items():
    # Skip any accidental self-maps (american == british): searching for the
    # British word and "fixing" it to itself is a no-op that still inflates the
    # issue/fix counts.
    if _am.lower() == _br.lower():
        continue
    _AMERICAN_FOR_BRITISH.setdefault(_br.lower(), []).append(_am)


def _derive_american_forms(british):
    """Best-effort algorithmic derivation of the American spelling(s) of a
    British word using regular orthographic patterns, so a NEW BritishSpelling_*
    rule whose target isn't in the explicit map above still does something
    instead of silently no-op'ing. Conservative: only fires on well-known
    regular patterns and never returns the British word itself."""
    b = british.lower()
    out = set()
    if b.endswith('ise'):  # organise -> organize (+ inflections)
        stem = b[:-3]
        out.update({stem + 'ize', stem + 'izes', stem + 'ized', stem + 'izing'})
    if b.endswith('isation'):  # organisation -> organization
        out.add(b[:-7] + 'ization')
    if b.endswith('our'):  # colour -> color
        stem = b[:-3]
        out.update({stem + 'or', stem + 'ors'})
    if b.endswith('re') and len(b) > 4:  # centre -> center, fibre -> fiber
        out.update({b[:-2] + 'er', b[:-2] + 'ers'})
    if b.endswith('ogue'):  # catalogue -> catalog, dialogue -> dialog
        out.update({b[:-4] + 'og', b[:-4] + 'ogs'})
    out.discard(b)
    return sorted(out)

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
    # Fallback: derive the American form from the (British) target word so a new
    # rule that isn't in the explicit map still works for the regular patterns.
    derived = _derive_american_forms(suffix) or _derive_american_forms(expected)
    if derived:
        return derived, (expected or suffix)
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

    combined = re.compile(
        r'\b(?:' + '|'.join(re.escape(w) for w in american_forms) + r')\b', re.IGNORECASE)
    auto = rule['auto_fix']
    fix_count = 0
    suggest_count = 0

    for para_idx, run in _iter_runs(doc):
        matches = combined.findall(run.text)
        if not matches:
            continue
        if auto:
            before = run.text
            run.text = combined.sub(replace_preserve_case, run.text)
            if run.text != before:
                changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                fix_count += len(matches)
        else:
            suggest_count += _tracked_replace_in_run(run, combined, replace_preserve_case)

    if fix_count > 0:
        fixes.append(f"Fixed {fix_count} instances to British spelling '{british_word}'")
    if suggest_count > 0:
        fixes.append(f"Suggested {suggest_count} change(s) to British spelling '{british_word}' — proposed as tracked changes to accept or reject")

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

    auto = rule['auto_fix']
    fix_count = 0
    suggest_count = 0

    def _expand(match):
        # Preserve a leading capital (e.g. start of sentence)
        return expanded.capitalize() if match.group(0)[0].isupper() else expanded

    # Check all paragraphs (including those inside tables)
    for para_idx, run in _iter_runs(doc):
        matches = pattern.findall(run.text)
        if not matches:
            continue
        if auto:
            before = run.text
            run.text = pattern.sub(_expand, run.text)
            changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
            fix_count += len(matches)
        else:
            suggest_count += _tracked_replace_in_run(run, pattern, _expand)

    if fix_count > 0:
        fixes.append(f"Fixed {fix_count} contractions to '{expanded}'")
    if suggest_count > 0:
        fixes.append(f"Suggested {suggest_count} change(s) expanding '{canonical}' to '{expanded}' — proposed as tracked changes to accept or reject")

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
        toward_pat = re.compile(r'\btowards\b', re.IGNORECASE)

        def _toward(m):
            return 'Toward' if m.group(0)[0].isupper() else 'toward'

        auto = rule['auto_fix']
        fix_count = 0
        suggest_count = 0

        for para_idx, run in _iter_runs(doc):
            if not toward_pat.search(run.text):
                continue
            if auto:
                before = run.text
                run.text = toward_pat.sub(_toward, run.text)
                changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                fix_count += len(toward_pat.findall(before))
            else:
                suggest_count += _tracked_replace_in_run(run, toward_pat, _toward)

        if fix_count > 0:
            fixes.append(f"Fixed {fix_count} instances to 'toward'")
        if suggest_count > 0:
            fixes.append(f"Suggested {suggest_count} change(s) 'towards' to 'toward' — proposed as tracked changes to accept or reject")

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
        amp_pat = re.compile(r'&')
        auto = rule['auto_fix']
        fix_count = 0
        suggest_count = 0

        for para_idx, run in _iter_runs(doc):
            if '&' not in run.text:
                continue
            if auto:
                before = run.text
                run.text = run.text.replace('&', 'and')
                changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                fix_count += before.count('&')
            else:
                suggest_count += _tracked_replace_in_run(run, amp_pat, 'and')

        if fix_count > 0:
            fixes.append(f"Fixed {fix_count} ampersands to 'and'")
        if suggest_count > 0:
            fixes.append(f"Suggested {suggest_count} change(s) '&' to 'and' — proposed as tracked changes to accept or reject")

    elif check_value == 'PercentSymbol':
        # Replace number% with 'number percent' (e.g. "85%" -> "85 percent")
        pct_pat = re.compile(r'(\d+)%')
        auto = rule['auto_fix']
        fix_count = 0
        suggest_count = 0

        for para_idx, run in _iter_runs(doc):
            matches = pct_pat.findall(run.text)
            if not matches:
                continue
            if auto:
                before = run.text
                run.text = pct_pat.sub(r'\1 percent', run.text)
                changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                fix_count += len(matches)
            else:
                suggest_count += _tracked_replace_in_run(run, pct_pat, lambda m: f"{m.group(1)} percent")

        if fix_count > 0:
            fixes.append(f"Fixed {fix_count} percent symbols to 'percent'")
        if suggest_count > 0:
            fixes.append(f"Suggested {suggest_count} change(s) '%' to 'percent' — proposed as tracked changes to accept or reject")

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
        # Numbers with 4+ digits and no commas (excluding years 1900-2099)
        num_pat = re.compile(r'\b\d{4,}\b')

        def _is_target(token):
            return not (1900 <= int(token) <= 2099)

        def _comma(m):
            tok = m.group(0)
            return '{:,}'.format(int(tok)) if _is_target(tok) else tok

        auto = rule['auto_fix']
        fix_count = 0
        suggest_count = 0

        for para_idx, run in _iter_runs(doc):
            targets = [m for m in num_pat.findall(run.text) if _is_target(m)]
            if not targets:
                continue
            if auto:
                before = run.text
                run.text = num_pat.sub(_comma, run.text)
                changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                fix_count += len(targets)
            else:
                # highlight only the comma-needing numbers (years pass through untouched)
                tgt_pat = re.compile(
                    r'\b(?:' + '|'.join(re.escape(t) for t in dict.fromkeys(targets)) + r')\b')
                suggest_count += _tracked_replace_in_run(run, tgt_pat, _comma)

        if fix_count > 0:
            fixes.append(f"Added commas to {fix_count} numbers")
        if suggest_count > 0:
            fixes.append(f"Suggested comma formatting on {suggest_count} number(s) — proposed as tracked changes to accept or reject")

    return {'issues': issues, 'fixes': fixes, 'changes': changes}

# ============================================
# GENERIC REGEX HELPERS
# ============================================

def _iter_runs(doc):
    """Yield (paragraph_index, run) for every non-empty run, tables included."""
    for para_idx, paragraph in enumerate(iter_all_paragraphs(doc)):
        for run in paragraph.runs:
            if run.text:
                yield para_idx, run


def _flag_regex(doc, pattern, label):
    """Detection-only: count regex matches and report them as an issue."""
    count = sum(len(pattern.findall(run.text)) for _idx, run in _iter_runs(doc))
    issues = [f"Found {count} instance(s): {label}"] if count else []
    return {'issues': issues, 'fixes': [], 'changes': []}


def _replace_regex(doc, rule, pattern, repl, label):
    """Apply a regex replacement. If the rule auto-fixes, replace silently; if
    not, apply the suggested replacement and highlight it yellow for review.
    `repl` may be a string or a callable (re.sub semantics)."""
    fix_count = 0
    suggest_count = 0
    changes = []
    auto = rule.get('auto_fix')
    for para_idx, run in _iter_runs(doc):
        matches = pattern.findall(run.text)
        if not matches:
            continue
        if auto:
            before = run.text
            run.text = pattern.sub(repl, run.text)
            if run.text != before:
                changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                fix_count += len(matches)
        else:
            suggest_count += _tracked_replace_in_run(run, pattern, repl)
    fixes = []
    if fix_count:
        fixes.append(f"Fixed {fix_count} instance(s): {label}")
    if suggest_count:
        fixes.append(f"Suggested {suggest_count} change(s) — proposed as tracked changes to accept or reject: {label}")
    return {'issues': [], 'fixes': fixes, 'changes': changes}


# ============================================
# WORDY-PHRASE REPLACEMENTS (PhraseReplace_*)
# ============================================

# CheckValue -> the wordy phrase to detect. The concise replacement comes from
# the rule's ExpectedValue. These rules are detection-only (auto_fix=False) in
# the live list, but _replace_regex honours auto_fix if that changes.
PHRASE_REPLACE_PHRASES = {
    'PhraseReplace_atpresenttime': 'at the present time',
    'PhraseReplace_conductinvestigation': 'conduct an investigation of',
    'PhraseReplace_commence': 'commence',
    'PhraseReplace_despitethefact': 'despite the fact that',
    'PhraseReplace_duringthetime': 'during the time that',
    'PhraseReplace_duetothefact': 'due to the fact that',
    'PhraseReplace_forthepurpose': 'for the purpose of',
    'PhraseReplace_inorderto': 'in order to',
    'PhraseReplace_inthecourseof': 'in the course of',
    'PhraseReplace_intheeventof': 'in the event of',
    'PhraseReplace_vicinity': 'in the vicinity of',
    'PhraseReplace_initiateprep': 'initiate the preparation of',
    'PhraseReplace_isapplicable': 'is applicable',
    'PhraseReplace_makeadecision': 'make a decision',
    'PhraseReplace_mostofthetime': 'most of the time',
    'PhraseReplace_voluntarybasis': 'on a voluntary basis',
    'PhraseReplace_annualbasis': 'on an annual basis',
    'PhraseReplace_priorto': 'prior to',
    'PhraseReplace_providedescription': 'provide a description of',
    'PhraseReplace_subsequentto': 'subsequent to',
    'PhraseReplace_takeintoconsideration': 'take into consideration',
    'PhraseReplace_majority': 'the majority',
}


def _phrase_pattern(phrase):
    """Word-bounded, case-insensitive, whitespace-flexible pattern for a phrase."""
    return re.compile(r'\b' + r'\s+'.join(re.escape(w) for w in phrase.split()) + r'\b', re.IGNORECASE)


def check_phrase_replace(doc, rule):
    phrase = PHRASE_REPLACE_PHRASES.get(rule['check_value'])
    if not phrase:
        return {'issues': [], 'fixes': [], 'changes': []}
    suggestion = (rule.get('expected_value') or '').strip()
    label = f"wordy phrase '{phrase}'" + (f" — use '{suggestion}'" if suggestion else '')
    pattern = _phrase_pattern(phrase)
    # A concrete one-to-one suggestion (no "X/Y" ambiguity) can be applied and,
    # when the rule is suggest-only, highlighted yellow. _replace_regex handles
    # both the auto-fix and the suggest+highlight paths.
    if suggestion and '/' not in suggestion:
        return _replace_regex(doc, rule, pattern, suggestion, label)
    return _flag_regex(doc, pattern, label)


# ============================================
# OTHER SIMPLE LANGUAGE / PUNCTUATION CHECKS
# ============================================

_OFFON_MAP = {'off-site': 'offsite', 'on-line': 'online', 'on-site': 'onsite', 'off-line': 'offline'}


def check_simple_language(doc, rule):
    cv = rule['check_value']
    if cv == 'ProximityRedundant':
        return _flag_regex(doc, re.compile(r'\bclose\s+proximity\b', re.I),
                           "'close proximity' is redundant — use 'proximity'")
    if cv == 'NoMinMaxApprox':
        return _flag_regex(doc, re.compile(r'\b(?:min|max|approx)\.', re.I),
                           "abbreviated minimum/maximum/approximately — spell out in full")
    if cv == 'ForecastPastTense':
        return _flag_regex(doc, re.compile(r'\bforecasted\b', re.I),
                           "'forecasted' — use 'forecast' as the past tense")
    if cv == 'Constructability':
        return _replace_regex(doc, rule, re.compile(r'\bConstructibility\b', re.I),
                              'Constructability', "'Constructibility' — use 'Constructability'")
    return {'issues': [], 'fixes': [], 'changes': []}


def check_simple_punctuation(doc, rule):
    cv = rule['check_value']
    if cv == 'NoDoubleSpaces':
        return _replace_regex(doc, rule, re.compile(r'  +'), ' ',
                              "double spaces — use a single space")
    if cv == 'NoHyphenInSitu':
        return _replace_regex(doc, rule, re.compile(r'\b(in|ex)-situ\b', re.I),
                              lambda m: f"{m.group(1)} situ", "hyphenated 'in/ex situ' — remove hyphen")
    if cv == 'NoHyphenOffOn':
        return _replace_regex(doc, rule, re.compile(r'\b(off-site|on-line|on-site|off-line)\b', re.I),
                              lambda m: _OFFON_MAP[m.group(1).lower()],
                              "hyphenated offsite/online/onsite/offline — remove hyphen")
    if cv == 'AvoidAndOr':
        return _flag_regex(doc, re.compile(r'\band\s*/\s*or\b', re.I),
                           "'and/or' — use 'X or Y or both'")
    return {'issues': [], 'fixes': [], 'changes': []}


# ============================================
# CAPITALISATION CHECKS
# ============================================

_REF_CODE_RE = re.compile(r'\b[A-Za-z0-9]{2,}(?:-[A-Za-z0-9]{2,}){2,}\b')


def _looks_like_ref_code(token):
    """A hyphenated token is treated as a reference code only if at least one
    segment is all-uppercase or all-digits (so 'state-of-the-art' is ignored)."""
    segs = token.split('-')
    return any(s.isupper() or s.isdigit() for s in segs)


def check_reference_code_case(doc, rule):
    issues = []
    fixes = []
    changes = []
    issue_count = 0
    fix_count = 0
    auto = rule.get('auto_fix')

    for para_idx, paragraph in enumerate(iter_all_paragraphs(doc)):
        for run in paragraph.runs:
            if not run.text:
                continue
            mixed = [m.group(0) for m in _REF_CODE_RE.finditer(run.text)
                     if _looks_like_ref_code(m.group(0)) and any(c.islower() for c in m.group(0))]
            if not mixed:
                continue
            issue_count += len(mixed)
            if auto:
                before = run.text

                def _upper(m):
                    tok = m.group(0)
                    return tok.upper() if (_looks_like_ref_code(tok) and any(c.islower() for c in tok)) else tok

                run.text = _REF_CODE_RE.sub(_upper, run.text)
                if run.text != before:
                    changes.append({'before': before, 'after': run.text, 'location': f'Paragraph {para_idx + 1}'})
                    fix_count += len(mixed)

    if issue_count and not auto:
        issues.append(f"Found {issue_count} reference code(s) not fully uppercase")
    if fix_count:
        fixes.append(f"Uppercased {fix_count} reference code(s)")
    return {'issues': issues, 'fixes': fixes, 'changes': changes}


# ============================================
# DETERMINISTIC CHECKS (batch 2)
# ============================================
# Each maps to a check_value in the live 'Style Rules' list that previously had
# no validator branch (and so was silently skipped). Detection-only unless the
# rule's auto_fix is set and the correction is unambiguous. Deliberately
# precision-first: rules needing linguistic judgement (homophones, tone, title
# case, terminology consistency) are left for the AI path, not approximated here.

_NONE = {'issues': [], 'fixes': [], 'changes': []}

# -- Punctuation --
_TIME_12H = re.compile(r'\b\d{1,2}(?::\d{2})?\s*[ap]\.?m\.?\b', re.I)
_NUMERIC_DATE = re.compile(r'\b(?:\d{1,2}[/.]\d{1,2}[/.]\d{2,4}|\d{4}-\d{1,2}-\d{1,2})\b')
_YEAR_RANGE = re.compile(r'\b(?:19|20)\d{2}\s*(?:[-–—]\s*(?:19|20)\d{2}|to\s+(?:19|20)\d{2})\b', re.I)
_SLASH_SPACED = re.compile(r'\s*/\s*')
_SLASH_WORDS = re.compile(r'\b[A-Za-z]{2,}/[A-Za-z]{2,}\b')
_HYPHEN_IN = re.compile(r'\bin\s+(?:depth|house|line|place|service|text)\b', re.I)
_HYPHEN_SUFFIX = re.compile(r'\b\w+\s+(?:related|type)\b', re.I)
_HYPHEN_PREFIX = re.compile(r'\b(?:self|quasi)\s+[a-z]+\b', re.I)
_WIDE = re.compile(r'\b(site|company|organisation|organization|nation|country|world|industry|estate|network|project|programme|system|region)\s+wide\b', re.I)
_EGIE_COMMA_AFTER = re.compile(r'\b[ei]\.[ge]\.,')
_EGIE_NO_PUNCT_BEFORE = re.compile(r'(?<=\w)\s+(?:e\.g\.|i\.e\.)', re.I)
# "A, B and C" (one comma, then 'word and/or word') — a 3-item list missing the
# serial comma. "A, B, and C" (already has it) does not match, since the word
# before and/or is followed by a comma, not the conjunction.
_OXFORD = re.compile(r'\b\w+,\s+\w+\s+(?:and|or)\s+\w+\b')
_NUM_BELOW_TEN = re.compile(r'(?<![\w./:-])[1-9](?![\w./:%-])')
# number + SINGULAR unit + a following word = a compound modifier that should be
# hyphenated ('15 page document' -> '15-page document'). Singular unit only, so
# '5 years later' (not a modifier) is not flagged.
_COMPOUND_MOD = re.compile(
    r'\b\d+\s+(?:page|year|month|week|day|storey|metre|mile|hour|minute|foot|bed|bedroom|'
    r'lane|stage|phase|step|tonne|litre|person)\s+[a-z]+', re.I)
_NUM_EXCL_PREFIX = re.compile(
    r'\b(?:figure|fig|table|section|level|phase|stage|step|chapter|part|no|item|day|week|year|'
    r'option|appendix|volume|grade|band|tier|class|type|page|version|rev|para|paragraph|clause|'
    r'note|row|column|col|point|task|unit|zone|lane|gate)\.?\s*$', re.I)
_UNIT_AFTER = re.compile(r'^\s*(?:%|mm|cm|km|kg|ml|pp|st|nd|rd|th|am|pm|m\b|g\b|l\b|t\b|x\b|:|/)', re.I)

# -- Grammar --
_SENT_EGIE = re.compile(r'(?:^|[.!?]\s+)(?:E\.g\.|I\.e\.)')
_THE_CLIENT = re.compile(r'\bthe client\b', re.I)
_ORG_SINGULAR = re.compile(
    r"\bthe\s+(?:team|company|organisation|organization|department|committee|board|government|"
    r"council|group|authority|contractor|client|firm|business)\s+"
    r"(?:are|were|have|do|aren't|weren't|don't|haven't)\b", re.I)
_EGIE = re.compile(r'\b(?:e\.g\.|i\.e\.)', re.I)
_ETC = re.compile(r'\betc\b', re.I)

# -- Language --
_FEEL = re.compile(r'\bfeel(?:s|ing|t)?\b', re.I)
_ABOVE_BELOW = re.compile(
    r'\b(?:see|shown|noted|listed|described|mentioned|detailed|figure|table|stated|outlined|the)\s+'
    r'(?:above|below)\b', re.I)
_IMPERIAL = re.compile(
    r'\b\d+(?:\.\d+)?\s*(?:feet|foot|inch(?:es)?|miles?|yards?|pounds?|lbs?|ounces?|gallons?|'
    r'°?\s?fahrenheit|°F)\b', re.I)

# -- Capitalisation (case-sensitive: match the lowercase form only) --
_NATIONALITY = re.compile(
    r'\b(?:welsh|english|scottish|irish|british|european|japanese|chinese|american|french|german|'
    r'spanish|italian|russian|indian|australian|canadian|portuguese|dutch|greek|roman|latin|'
    r'arabic|hebrew|nordic|asian|african)\b')
_EMPHASIS_CAPS = re.compile(r'\b[A-Z]{2,}(?:\s+[A-Z]{2,})+\b')


def _check_punct_egie(doc, rule):
    n = sum(len(_EGIE_COMMA_AFTER.findall(r.text)) + len(_EGIE_NO_PUNCT_BEFORE.findall(r.text))
            for _i, r in _iter_runs(doc))
    return {'issues': [f"Found {n} instance(s): e.g./i.e. punctuation — comma/colon/hyphen before, "
                       f"no comma after"] if n else [], 'fixes': [], 'changes': []}


def _check_numbers_below_ten(doc, rule):
    count = 0
    for _idx, run in _iter_runs(doc):
        text = run.text
        for m in _NUM_BELOW_TEN.finditer(text):
            if _NUM_EXCL_PREFIX.search(text[:m.start()]):
                continue
            if _UNIT_AFTER.match(text[m.end():]):
                continue
            count += 1
    return {'issues': [f"Found {count} digit(s) below ten in running text — spell out (one to nine)"]
            if count else [], 'fixes': [], 'changes': []}


def _check_caption_no_period(doc, rule):
    count = 0
    for paragraph in iter_all_paragraphs(doc):
        style = (paragraph.style.name or '') if paragraph.style else ''
        if 'caption' in style.lower():
            t = paragraph.text.rstrip()
            if t.endswith('.') and not t.endswith('...'):
                count += 1
    return {'issues': [f"Found {count} caption(s) ending with a full stop — remove it"]
            if count else [], 'fixes': [], 'changes': []}


def _check_no_etc_with_egie(doc, rule):
    count = sum(1 for p in iter_all_paragraphs(doc)
                if _EGIE.search(p.text) and _ETC.search(p.text))
    return {'issues': [f"Found {count} paragraph(s) using 'etc.' alongside e.g./i.e. — drop 'etc.'"]
            if count else [], 'fixes': [], 'changes': []}


def _check_proper_noun_derivations(doc, rule):
    count = sum(len(_NATIONALITY.findall(r.text)) for _i, r in _iter_runs(doc))
    return {'issues': [f"Found {count} lowercase proper-noun derivation(s) — capitalise "
                       f"(e.g. 'welsh' to 'Welsh')"] if count else [], 'fixes': [], 'changes': []}


# check_value -> handler. Detection-only handlers use _flag_regex; a couple
# auto-fix where the live rule sets AutoFix and the correction is unambiguous.
_LANGUAGE_CHECKS = {
    'NoFeelTechnical': lambda d, r: _flag_regex(d, _FEEL, "'feel' in technical writing — use 'think'/'believe'/'consider'"),
    'NoAboveBelow': lambda d, r: _flag_regex(d, _ABOVE_BELOW, "'above'/'below' cross-reference — cite the figure/table/section number"),
    'PreferMetric': lambda d, r: _flag_regex(d, _IMPERIAL, "imperial unit — use metric where possible"),
}
_PUNCTUATION_CHECKS = {
    'TimeFormat': lambda d, r: _flag_regex(d, _TIME_12H, "12-hour clock time — use 24-hour HH:MM (e.g. 09:00, 18:25)"),
    'DateFormat_Text': lambda d, r: _flag_regex(d, _NUMERIC_DATE, "numeric date — use DD MONTH YYYY (e.g. 01 February 2015)"),
    'DateFormat_Table': lambda d, r: _flag_regex(d, _NUMERIC_DATE, "numeric date — use DD-MMM-YYYY in tables (e.g. 28-Feb-2020)"),
    'YearIntervalFormat': lambda d, r: _flag_regex(d, _YEAR_RANGE, "year range — use YYYY-YY (e.g. 2019-20)"),
    'NoSpacesAroundSlash': lambda d, r: _replace_regex(d, r, _SLASH_SPACED, '/', "spaces around '/' — close up (e.g. km/s)"),
    'AvoidForwardSlash': lambda d, r: _flag_regex(d, _SLASH_WORDS, "forward slash between words — use words to avoid ambiguity"),
    'HyphenInWords': lambda d, r: _flag_regex(d, _HYPHEN_IN, "missing hyphen — e.g. 'in-depth', 'in-house', 'in-text'"),
    'HyphenSuffixes': lambda d, r: _flag_regex(d, _HYPHEN_SUFFIX, "missing hyphen before -related/-type (e.g. 'quality-related')"),
    'HyphenAlwaysPrefix': lambda d, r: _flag_regex(d, _HYPHEN_PREFIX, "missing hyphen after self-/quasi- (e.g. 'self-made')"),
    'Hyphen_wide': lambda d, r: _replace_regex(d, r, _WIDE, lambda m: f"{m.group(1)}-wide", "missing hyphen before '-wide' (e.g. 'site-wide')"),
    'PunctuationBeforeEgIe': _check_punct_egie,
    'OxfordComma': lambda d, r: _flag_regex(d, _OXFORD, "list of 3+ items may be missing an Oxford comma before 'and'/'or'"),
    'NumbersBelowTen': _check_numbers_below_ten,
    'CaptionNoPeriod': _check_caption_no_period,
    'CompoundModifiers': lambda d, r: _flag_regex(d, _COMPOUND_MOD, "number+unit used as a modifier — hyphenate (e.g. '15-page document')"),
}
_GRAMMAR_CHECKS = {
    'NoSentenceStartEgIe': lambda d, r: _flag_regex(d, _SENT_EGIE, "sentence starts with e.g./i.e. — rephrase (e.g. 'for example')"),
    'NoEtcWithEgIe': _check_no_etc_with_egie,
    'ClientNameNotTheClient': lambda d, r: _flag_regex(d, _THE_CLIENT, "'the client' — use the client's actual name"),
    'OrgSingular': lambda d, r: _flag_regex(d, _ORG_SINGULAR, "organisation with plural verb — use the singular ('the team is')"),
}
_CAPITALISATION_CHECKS = {
    'ProperNounDerivations': _check_proper_noun_derivations,
    'NoEmphasisCaps': lambda d, r: _flag_regex(d, _EMPHASIS_CAPS, "consecutive ALL-CAPS words used for emphasis — use normal case"),
}


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
    elif check_value.startswith('PhraseReplace_'):
        return check_phrase_replace(doc, rule)
    elif check_value in ['Word_toward', 'AvoidEtc', 'AvoidShould']:
        return check_word_choice(doc, rule)
    elif check_value in ['ProximityRedundant', 'NoMinMaxApprox', 'ForecastPastTense', 'Constructability']:
        return check_simple_language(doc, rule)
    elif check_value in _LANGUAGE_CHECKS:
        return _LANGUAGE_CHECKS[check_value](doc, rule)
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
    elif check_value in ['NoDoubleSpaces', 'NoHyphenInSitu', 'NoHyphenOffOn', 'AvoidAndOr']:
        return check_simple_punctuation(doc, rule)
    elif check_value in _PUNCTUATION_CHECKS:
        return _PUNCTUATION_CHECKS[check_value](doc, rule)
    else:
        logging.info(f"Punctuation check '{check_value}' not yet implemented")
        return {'issues': [], 'fixes': []}


def validate_capitalisation_rules(doc, rule):
    """Dispatch capitalisation rule validation.

    Only mechanical capitalisation checks are implemented. Context-dependent
    ones (job titles, govt bodies, fields of study, document/section titles)
    need judgement and are left for the AI path rather than risk false
    positives."""
    check_value = rule['check_value']

    if check_value == 'ReferenceCodeCase':
        return check_reference_code_case(doc, rule)
    elif check_value in _CAPITALISATION_CHECKS:
        return _CAPITALISATION_CHECKS[check_value](doc, rule)
    else:
        logging.info(f"Capitalisation check '{check_value}' not yet implemented")
        return {'issues': [], 'fixes': []}

def validate_grammar_rules(doc, rule):
    """Dispatch grammar rule validation"""
    check_value = rule['check_value']

    if check_value.startswith('NoContraction_'):
        return check_contractions(doc, rule)
    elif check_value in _GRAMMAR_CHECKS:
        return _GRAMMAR_CHECKS[check_value](doc, rule)
    else:
        logging.info(f"Grammar check '{check_value}' not yet implemented")
        return {'issues': [], 'fixes': []}
