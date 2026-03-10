"""Visio document (.vsdx) validation"""
import os
import logging
import tempfile
from vsdx import VisioFile
from .ai_client import call_claude


def validate_visio_document(file_stream, rules):
    """Validate Visio document against rules"""
    logging.info("Loading Visio document...")

    # VisioFile expects a filename, not a stream
    with tempfile.NamedTemporaryFile(suffix='.vsdx', delete=False) as tmp:
        tmp.write(file_stream.read())
        tmp_path = tmp.name
    file_stream.seek(0)
    visio = VisioFile(tmp_path)

    page_count = len(visio.pages)
    logging.info(f"Visio document loaded. Pages: {page_count}")

    issues = []
    fixes_applied = []

    visio_rules = [r for r in rules if r['doc_type'] in ['Visio', 'Both', 'All']]
    ai_rules = [r for r in visio_rules if r.get('use_ai', False)]
    hard_coded_rules = [r for r in visio_rules if not r.get('use_ai', False)]

    logging.info(f"AI rules: {len(ai_rules)}, Hard-coded rules: {len(hard_coded_rules)}")

    # Extract all text from Visio shapes
    shape_texts = []
    for page in visio.pages:
        shape_texts.extend(_extract_shape_texts(page, page.child_shapes))

    logging.info(f"Extracted text from {len(shape_texts)} shapes")

    # AI-powered corrections
    if ai_rules and shape_texts:
        try:
            combined_text = "\n\n".join([st['text'] for st in shape_texts if st['text'].strip()])
            if combined_text.strip():
                result = call_claude(ai_rules, combined_text)
                if result and result['changes_made'] > 0:
                    # Report AI issues but don't apply text changes to Visio
                    # (splitting corrected text back to shapes is unreliable
                    #  and can corrupt the document)
                    issues.append({
                        'rule_name': 'AI Style Corrections',
                        'rule_type': 'AI',
                        'description': f"Found {result['changes_made']} style violations requiring manual review",
                        'location': 'Document-wide',
                        'priority': 3
                    })
                    logging.info(f"Claude found {result['changes_made']} Visio style issues (report only, no auto-fix)")
        except Exception as e:
            logging.error(f"Claude validation failed for Visio: {e}")
            issues.append({
                'rule_name': 'AI Validation',
                'rule_type': 'AI',
                'description': f"AI validation failed: {e}",
                'location': 'N/A',
                'priority': 1
            })

    # Hard-coded rules
    for rule in hard_coded_rules:
        result = None
        if rule['rule_type'] == 'Color':
            result = _check_colors(visio, rule)
        elif rule['rule_type'] == 'Font':
            result = _check_fonts(visio, rule)
        elif rule['rule_type'] == 'Size':
            result = _check_shape_size(visio, rule)
        elif rule['rule_type'] == 'Position':
            result = _check_position(visio, rule)
        elif rule['rule_type'] == 'PageDimensions':
            result = _check_page_dimensions(visio, rule)

        if result:
            issues.extend(result['issues'])
            fixes_applied.extend(result['fixes'])

    logging.info(f"Visio validation complete. Issues: {len(issues)}, Fixes: {len(fixes_applied)}")

    # Clean up temp file
    try:
        os.unlink(tmp_path)
    except Exception:
        pass

    return {'document': visio, 'issues': issues, 'fixes_applied': fixes_applied}


def _extract_shape_texts(page, shapes, shape_list=None):
    """Recursively extract text from all shapes"""
    if shape_list is None:
        shape_list = []

    for shape in shapes:
        if hasattr(shape, 'text') and shape.text:
            text = str(shape.text).strip()
            if text:
                shape_list.append({'shape': shape, 'text': text, 'page': page.name})

        if hasattr(shape, 'child_shapes') and shape.child_shapes:
            _extract_shape_texts(page, shape.child_shapes, shape_list)

    return shape_list


def _process_shapes(visio, check_fn):
    """Process all shapes across all pages with a check function"""
    for page in visio.pages:
        if hasattr(page, 'child_shapes'):
            _process_shapes_recursive(page.child_shapes, check_fn)


def _process_shapes_recursive(shapes, check_fn):
    """Recursively process shapes"""
    for shape in shapes:
        if hasattr(shape, 'text') and shape.text and str(shape.text).strip():
            check_fn(shape)
        if hasattr(shape, 'child_shapes') and shape.child_shapes:
            _process_shapes_recursive(shape.child_shapes, check_fn)


def _check_colors(visio, rule):
    """Check and fix colors in Visio diagrams"""
    issues = []
    fixes = []
    expected_color = rule['expected_value']
    counts = {'issues': 0, 'fixes': 0}

    def check(shape):
        try:
            if rule['check_value'] == 'ShapeFillColor':
                current = getattr(shape, 'fill_color', None)
                if current and current != expected_color:
                    counts['issues'] += 1
                    if rule['auto_fix']:
                        shape.fill_color = expected_color
                        counts['fixes'] += 1

            elif rule['check_value'] == 'ShapeTextColor':
                current = getattr(shape, 'text_color', None)
                if current and current != expected_color:
                    counts['issues'] += 1
                    if rule['auto_fix']:
                        shape.text_color = expected_color
                        counts['fixes'] += 1
        except Exception as e:
            logging.warning(f"Could not check/set color for shape: {e}")

    _process_shapes(visio, check)

    if counts['issues'] > 0:
        issues.append(f"Found {counts['issues']} shapes with incorrect {rule['check_value']}")
    if counts['fixes'] > 0:
        fixes.append(f"Fixed {counts['fixes']} shapes to {expected_color}")

    return {'issues': issues, 'fixes': fixes}


def _check_fonts(visio, rule):
    """Check and fix fonts in Visio diagrams"""
    issues = []
    fixes = []
    expected_font = rule['expected_value']
    counts = {'issues': 0, 'fixes': 0}

    def check(shape):
        try:
            if rule['check_value'] == 'AllTextFont':
                try:
                    current_font = shape.cells.get('Char.Font', None)
                    if current_font is not None:
                        font_value = current_font.value if hasattr(current_font, 'value') else str(current_font)
                        if font_value != '0':
                            counts['issues'] += 1
                            if rule['auto_fix']:
                                shape.set_cell_value('Char.Font', '0')
                                counts['fixes'] += 1
                except (AttributeError, KeyError):
                    if rule['auto_fix']:
                        try:
                            shape.set_cell_value('Char.Font', '0')
                            counts['fixes'] += 1
                            counts['issues'] += 1
                        except Exception:
                            pass
        except Exception as e:
            logging.warning(f"Could not check/set font for shape: {e}")

    _process_shapes(visio, check)

    if counts['issues'] > 0:
        issues.append(f"Found {counts['issues']} shapes with incorrect font")
    if counts['fixes'] > 0:
        fixes.append(f"Fixed {counts['fixes']} shapes to {expected_font}")

    return {'issues': issues, 'fixes': fixes}


def _check_shape_size(visio, rule):
    """Check and fix shape dimensions"""
    issues = []
    fixes = []

    expected_value = rule['expected_value']
    tolerance = float(rule.get('tolerance', 0.1))

    try:
        expected_width, expected_height = map(float, expected_value.lower().split('x'))
    except ValueError:
        logging.warning(f"Could not parse size value: {expected_value}")
        return {'issues': issues, 'fixes': fixes}

    counts = {'issues': 0, 'fixes': 0}

    def check(shape):
        try:
            w = getattr(shape, 'width', None)
            h = getattr(shape, 'height', None)
            if w is not None and h is not None:
                if abs(w - expected_width) > tolerance or abs(h - expected_height) > tolerance:
                    counts['issues'] += 1
                    if rule['auto_fix']:
                        shape.width = expected_width
                        shape.height = expected_height
                        counts['fixes'] += 1
        except Exception as e:
            logging.warning(f"Could not check/set size: {e}")

    _process_shapes(visio, check)

    if counts['issues'] > 0:
        issues.append(f"Found {counts['issues']} shapes with incorrect dimensions (expected {expected_width}x{expected_height})")
    if counts['fixes'] > 0:
        fixes.append(f"Resized {counts['fixes']} shapes to {expected_width}x{expected_height}")

    return {'issues': issues, 'fixes': fixes}


def _check_position(visio, rule):
    """Check and fix shape positions"""
    issues = []
    fixes = []
    check_type = rule['check_value']
    expected_value = rule['expected_value']
    tolerance = float(rule.get('tolerance', 0.1))
    counts = {'issues': 0, 'fixes': 0}

    def check(shape):
        try:
            x = getattr(shape, 'x', None)
            y = getattr(shape, 'y', None)
            if x is None or y is None:
                return

            if check_type == 'TopMargin':
                max_y = float(expected_value)
                if y > max_y + tolerance:
                    counts['issues'] += 1
                    if rule['auto_fix']:
                        shape.y = max_y
                        counts['fixes'] += 1

            elif check_type == 'LeftMargin':
                min_x = float(expected_value)
                if x < min_x - tolerance:
                    counts['issues'] += 1
                    if rule['auto_fix']:
                        shape.x = min_x
                        counts['fixes'] += 1

            elif check_type == 'RightMargin':
                max_x = float(expected_value)
                if x > max_x + tolerance:
                    counts['issues'] += 1
                    if rule['auto_fix']:
                        shape.x = max_x
                        counts['fixes'] += 1

            elif check_type == 'BottomMargin':
                min_y = float(expected_value)
                if y < min_y - tolerance:
                    counts['issues'] += 1
                    if rule['auto_fix']:
                        shape.y = min_y
                        counts['fixes'] += 1

            elif check_type == 'ExactPosition':
                expected_x, expected_y = map(float, expected_value.split(','))
                if abs(x - expected_x) > tolerance or abs(y - expected_y) > tolerance:
                    counts['issues'] += 1
                    if rule['auto_fix']:
                        shape.x = expected_x
                        shape.y = expected_y
                        counts['fixes'] += 1

        except Exception as e:
            logging.warning(f"Could not check/set position: {e}")

    _process_shapes(visio, check)

    if counts['issues'] > 0:
        issues.append(f"Found {counts['issues']} shapes with incorrect position ({check_type})")
    if counts['fixes'] > 0:
        fixes.append(f"Repositioned {counts['fixes']} shapes for {check_type}")

    return {'issues': issues, 'fixes': fixes}


def _check_page_dimensions(visio, rule):
    """Check and fix page dimensions"""
    issues = []
    fixes = []

    try:
        expected_width, expected_height = map(float, rule['expected_value'].lower().split('x'))
    except ValueError:
        logging.warning(f"Could not parse page size: {rule['expected_value']}")
        return {'issues': issues, 'fixes': fixes}

    issue_count = 0
    fix_count = 0

    for page in visio.pages:
        try:
            w = getattr(page, 'width', None)
            h = getattr(page, 'height', None)
            if w is not None and h is not None:
                if w != expected_width or h != expected_height:
                    issue_count += 1
                    if rule['auto_fix']:
                        page.width = expected_width
                        page.height = expected_height
                        fix_count += 1
        except Exception as e:
            logging.warning(f"Could not check/set page dimensions for '{page.name}': {e}")

    if issue_count > 0:
        issues.append(f"Found {issue_count} pages with incorrect dimensions (expected {expected_width}x{expected_height})")
    if fix_count > 0:
        fixes.append(f"Resized {fix_count} pages to {expected_width}x{expected_height}")

    return {'issues': issues, 'fixes': fixes}
