"""HTML validation report generation with Mace branding"""
from datetime import datetime, timezone


def _escape_html(text):
    """Escape HTML special characters"""
    if text is None:
        return ''
    return str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def generate_report(file_name, issues, fixes_applied):
    """Generate validation report as HTML with Mace branding.

    Args:
        file_name: Name of the validated document.
        issues: List of dicts with keys: rule_name, rule_type, description, location, priority.
        fixes_applied: List of dicts with keys: rule_name, rule_type, found_value, fixed_value, location.
    """
    remaining_issues = [i for i in issues if isinstance(i, dict)]
    remaining_count = len(remaining_issues)
    fixes_count = len(fixes_applied)
    total_issues_found = remaining_count + fixes_count

    if remaining_count == 0:
        status = "Passed"
        status_color = "#28a745"
    elif fixes_count > 0:
        status = "Review Required"
        status_color = "#f0ad4e"
    else:
        status = "Failed"
        status_color = "#dc3545"

    validation_time = datetime.now(timezone.utc).strftime('%d %B %Y at %H:%M:%S UTC')

    # Build description text
    if fixes_count > 0 and remaining_count == 0:
        description = f"{fixes_count} issue{'s' if fixes_count != 1 else ''} found and auto-fixed. Document is compliant."
    elif fixes_count > 0:
        description = f"{fixes_count} issue{'s' if fixes_count != 1 else ''} auto-fixed, {remaining_count} remaining for manual review."
    elif total_issues_found > 0:
        description = f"{total_issues_found} issue{'s' if total_issues_found != 1 else ''} found. Manual correction required."
    else:
        description = "No issues found. Document fully complies with the Mace Writing Style Guide."

    # Build fixes table rows
    fixes_rows = ''
    for fix in fixes_applied:
        if isinstance(fix, dict):
            fixes_rows += f"""<tr>
                <td>{_escape_html(fix.get('rule_name', ''))}</td>
                <td><span class="rule-type-badge">{_escape_html(fix.get('rule_type', ''))}</span></td>
                <td>{_escape_html(fix.get('found_value', ''))}</td>
                <td>{_escape_html(fix.get('fixed_value', ''))}</td>
                <td>{_escape_html(fix.get('location', ''))}</td>
            </tr>"""
        else:
            fixes_rows += f"""<tr>
                <td colspan="5">{_escape_html(str(fix))}</td>
            </tr>"""

    # Build remaining issues table rows
    issues_rows = ''
    for issue in remaining_issues:
        priority = issue.get('priority', 999)
        priority_label = 'High' if priority <= 3 else ('Medium' if priority <= 6 else 'Low')
        priority_color = '#dc3545' if priority <= 3 else ('#f0ad4e' if priority <= 6 else '#6c757d')
        issues_rows += f"""<tr>
            <td>{_escape_html(issue.get('rule_name', ''))}</td>
            <td><span class="rule-type-badge">{_escape_html(issue.get('rule_type', ''))}</span></td>
            <td>{_escape_html(issue.get('description', ''))}</td>
            <td>{_escape_html(issue.get('location', ''))}</td>
            <td><span style="color:{priority_color};font-weight:bold;">{priority_label}</span></td>
        </tr>"""

    if fixes_applied:
        fixes_section = f"""<div class="section">
            <h2>Changes Made ({len(fixes_applied)})</h2>
            <table>
                <thead><tr>
                    <th>Rule Name</th><th>Rule Type</th><th>What Was Found</th><th>What It Was Changed To</th><th>Location</th>
                </tr></thead>
                <tbody>{fixes_rows}</tbody>
            </table>
        </div>"""
    else:
        fixes_section = ''

    issues_class = "section review-section" if status == "Review Required" else "section"
    if remaining_issues:
        issues_section = f"""<div class="{issues_class}">
            <h2>Remaining Issues ({remaining_count})</h2>
            <table>
                <thead><tr>
                    <th>Rule Name</th><th>Rule Type</th><th>Issue Description</th><th>Location</th><th>Priority</th>
                </tr></thead>
                <tbody>{issues_rows}</tbody>
            </table>
        </div>"""
    else:
        issues_section = ''

    # Build detailed changes section (collapsible diffs)
    detailed_changes_html = ''
    details_items = ''
    for fix in fixes_applied:
        if not isinstance(fix, dict) or not fix.get('changes'):
            continue
        rule_name = _escape_html(fix.get('rule_name', 'Unknown'))
        fix_changes = fix['changes']
        total = len(fix_changes)
        capped = fix_changes[:50]
        details_items += f'<details><summary>{rule_name} ({total} change{"s" if total != 1 else ""})</summary>'
        details_items += '<table><thead><tr><th>Location</th><th>Before</th><th>After</th></tr></thead><tbody>'
        for change in capped:
            details_items += f"""<tr>
                <td>{_escape_html(change.get('location', ''))}</td>
                <td><span class="diff-before">{_escape_html(change.get('before', ''))}</span></td>
                <td><span class="diff-after">{_escape_html(change.get('after', ''))}</span></td>
            </tr>"""
        if total > 50:
            details_items += f'<tr><td colspan="3" style="text-align:center;color:#6c757d;font-style:italic;">and {total - 50} more...</td></tr>'
        details_items += '</tbody></table></details>'

    if details_items:
        detailed_changes_html = f"""<div class="section">
            <h2>Detailed Changes</h2>
            {details_items}
        </div>"""

    if total_issues_found == 0:
        no_issues_section = """<div class="section passed-section">
            <h2>All Clear</h2>
            <p>This document fully complies with the Mace Control Centre Writing Style Guide. No issues were found.</p>
        </div>"""
    else:
        no_issues_section = ''

    report_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Validation Report - {_escape_html(file_name)}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: Arial, Helvetica, sans-serif;
            max-width: 1100px;
            margin: 0 auto;
            padding: 24px;
            background: #f4f6f9;
            color: #333;
            line-height: 1.5;
        }}
        .header {{
            background: linear-gradient(135deg, #1F4E79 0%, #1671CB 100%);
            color: #fff;
            padding: 32px;
            border-radius: 8px;
            margin-bottom: 24px;
        }}
        .header h1 {{
            font-size: 22px;
            margin-bottom: 16px;
            font-weight: 700;
        }}
        .status-badge {{
            display: inline-block;
            padding: 6px 18px;
            background: {status_color};
            color: #fff;
            border-radius: 16px;
            font-weight: 700;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .meta-info {{
            margin-top: 14px;
            font-size: 13px;
            opacity: 0.92;
        }}
        .meta-info div {{ margin-bottom: 3px; }}
        .summary {{
            background: #fff;
            padding: 24px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }}
        .summary h2 {{
            font-size: 16px;
            color: #1F4E79;
            margin-bottom: 16px;
            padding-bottom: 8px;
            border-bottom: 2px solid #1F4E79;
        }}
        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 16px;
        }}
        .summary-card {{
            background: #f4f6f9;
            padding: 18px;
            border-radius: 6px;
            text-align: center;
            border-left: 4px solid #1F4E79;
        }}
        .summary-card .number {{
            font-size: 32px;
            font-weight: 700;
            color: #1F4E79;
            margin-bottom: 4px;
        }}
        .summary-card .label {{
            font-size: 12px;
            color: #6c757d;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .section {{
            background: #fff;
            padding: 24px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }}
        .section h2 {{
            font-size: 16px;
            color: #1F4E79;
            margin-top: 0;
            margin-bottom: 16px;
            padding-bottom: 8px;
            border-bottom: 2px solid #1F4E79;
        }}
        .passed-section {{
            border-left: 4px solid #28a745;
        }}
        .passed-section h2 {{
            color: #28a745;
            border-bottom-color: #28a745;
        }}
        .passed-section p {{
            color: #555;
            padding: 20px 0;
            text-align: center;
            font-size: 15px;
        }}
        .review-section {{
            border-left: 4px solid #f0ad4e;
        }}
        .review-section h2 {{
            color: #e69500;
            border-bottom-color: #f0ad4e;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }}
        thead th {{
            background: #1F4E79;
            color: #fff;
            padding: 10px 12px;
            text-align: left;
            font-weight: 600;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }}
        thead th:first-child {{ border-radius: 4px 0 0 0; }}
        thead th:last-child {{ border-radius: 0 4px 0 0; }}
        tbody td {{
            padding: 10px 12px;
            border-bottom: 1px solid #e9ecef;
            vertical-align: top;
        }}
        tbody tr:hover {{ background: #f8f9fb; }}
        tbody tr:last-child td {{ border-bottom: none; }}
        .rule-type-badge {{
            display: inline-block;
            padding: 2px 8px;
            background: #e8f0fe;
            color: #1671CB;
            border-radius: 10px;
            font-size: 11px;
            font-weight: 600;
        }}
        .diff-before {{ background: #ffeef0; text-decoration: line-through; color: #b31d28; padding: 2px 4px; }}
        .diff-after {{ background: #e6ffec; color: #22863a; padding: 2px 4px; }}
        details {{ margin-bottom: 8px; }}
        summary {{ cursor: pointer; font-weight: 600; color: #1F4E79; padding: 6px 0; }}
        .footer {{
            text-align: center;
            margin-top: 28px;
            padding: 16px;
            color: #999;
            font-size: 11px;
        }}
        .footer span {{ color: #1F4E79; font-weight: 600; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>Mace Style Validation Report</h1>
        <span class="status-badge">{status}</span>
        <div class="meta-info">
            <div style="margin-top:8px;font-size:14px;opacity:1;">{description}</div>
            <div style="margin-top:10px;"><strong>Document:</strong> {_escape_html(file_name)}</div>
            <div><strong>Validated:</strong> {validation_time}</div>
        </div>
    </div>

    <div class="summary">
        <h2>Summary</h2>
        <div class="summary-grid">
            <div class="summary-card">
                <div class="number">{total_issues_found}</div>
                <div class="label">Issues Found</div>
            </div>
            <div class="summary-card">
                <div class="number">{len(fixes_applied)}</div>
                <div class="label">Auto-Fixed</div>
            </div>
            <div class="summary-card" style="border-left-color: {'#f0ad4e' if remaining_count > 0 and fixes_count > 0 else '#dc3545' if remaining_count > 0 else '#28a745'};">
                <div class="number" style="color: {'#f0ad4e' if remaining_count > 0 and fixes_count > 0 else '#dc3545' if remaining_count > 0 else '#28a745'};">{remaining_count}</div>
                <div class="label">Remaining</div>
            </div>
        </div>
    </div>

    {no_issues_section}
    {fixes_section}
    {detailed_changes_html}
    {issues_section}

    <div class="footer">
        <span>Mace Style Validator</span> &middot; Control Centre Writing Style Guide<br>
        Powered by Azure Functions &amp; Claude AI
    </div>
</body>
</html>"""
    return report_html
