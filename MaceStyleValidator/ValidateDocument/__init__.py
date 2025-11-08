import azure.functions as func
import logging
import json
import os
from io import BytesIO
from datetime import datetime, timezone
from docx import Document
from docx.shared import Pt, RGBColor
from vsdx import VisioFile
import msal
import requests
from .enhanced_validators import validate_language_rules, validate_punctuation_rules, validate_grammar_rules

# ============================================
# CONFIGURATION
# ============================================
def get_graph_token():
    """Get Microsoft Graph API access token using MSAL"""
    tenant_id = os.environ.get("SHAREPOINT_TENANT_ID", "2ab0866e-23d6-4688-be97-ce9f447135d8")
    client_id = os.environ.get("SHAREPOINT_CLIENT_ID", "c7859dae-6997-448f-9530-7166fe857e75")
    client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET", "DlD8Q~_NNgnpnVxKWsZTiz53DuNYrfrAjqkCDaP1")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scope = ["https://graph.microsoft.com/.default"]

    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )

    result = app.acquire_token_for_client(scopes=scope)

    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Failed to acquire token: {result.get('error_description', result)}")

def get_site_info():
    """Get SharePoint site information"""
    site_url = os.environ.get("SHAREPOINT_SITE_URL", "https://0rxf2.sharepoint.com/sites/StyleValidation")

    # Extract hostname and site path
    # Format: https://tenant.sharepoint.com/sites/sitename
    parts = site_url.replace("https://", "").split("/")
    hostname = parts[0]
    site_path = "/" + "/".join(parts[1:]) if len(parts) > 1 else ""

    return {
        "hostname": hostname,
        "site_path": site_path,
        "full_url": site_url
    }

# ============================================
# SHAREPOINT OPERATIONS
# ============================================
def fetch_validation_rules(token):
    """Fetch rules from SharePoint 'Style Rules' list using Graph API"""
    site_info = get_site_info()
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    # Get site ID first
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_info['hostname']}:{site_info['site_path']}"
    site_response = requests.get(site_url, headers=headers)
    site_response.raise_for_status()
    site_id = site_response.json()["id"]

    # Get list items
    list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/items?expand=fields"
    list_response = requests.get(list_url, headers=headers)
    list_response.raise_for_status()

    rules = []
    for item in list_response.json().get("value", []):
        fields = item.get("fields", {})
        rules.append({
            'title': fields.get('Title'),
            'rule_type': fields.get('RuleType'),
            'doc_type': fields.get('DocumentType'),
            'check_value': fields.get('CheckValue'),
            'expected_value': fields.get('ExpectedValue'),
            'auto_fix': fields.get('AutoFix'),
            'priority': fields.get('Priority', 999)
        })

    rules.sort(key=lambda x: x['priority'])
    return rules

def download_file(token, file_path):
    """Download file from SharePoint using Graph API"""
    if not file_path:
        raise ValueError("file_path cannot be None or empty")

    logging.info(f"Downloading file: {file_path}")

    site_info = get_site_info()
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    # Get site ID
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_info['hostname']}:{site_info['site_path']}"
    site_response = requests.get(site_url, headers=headers)
    site_response.raise_for_status()
    site_id = site_response.json()["id"]

    # Convert server-relative URL to drive-relative path
    # Remove site path and "Shared Documents" from the path
    # E.g., "/sites/StyleValidation/Shared Documents/test.docx" -> "/test.docx"
    drive_relative_path = file_path
    if "Shared Documents/" in file_path:
        drive_relative_path = "/" + file_path.split("Shared Documents/", 1)[1]
    elif "Shared Documents" in file_path and file_path.endswith("Shared Documents"):
        # Handle case where path ends with "Shared Documents" (folder itself)
        drive_relative_path = "/"

    logging.info(f"Using drive-relative path: {drive_relative_path}")

    # Download file
    file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{drive_relative_path}:/content"
    file_response = requests.get(file_url, headers=headers)
    file_response.raise_for_status()

    logging.info(f"File downloaded successfully, size: {len(file_response.content)} bytes")
    return BytesIO(file_response.content)

def upload_file(token, file_stream, target_path):
    """Upload file to SharePoint using Graph API"""
    if not target_path:
        raise ValueError("target_path cannot be None or empty")

    logging.info(f"Uploading file to: {target_path}")

    site_info = get_site_info()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream"
    }

    # Get site ID
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_info['hostname']}:{site_info['site_path']}"
    site_response = requests.get(site_url, headers=headers.copy())
    site_response.raise_for_status()
    site_id = site_response.json()["id"]

    # Convert server-relative URL to drive-relative path
    # Remove site path and "Shared Documents" from the path
    # E.g., "/sites/StyleValidation/Shared Documents/test.docx" -> "/test.docx"
    drive_relative_path = target_path
    if "Shared Documents/" in target_path:
        drive_relative_path = "/" + target_path.split("Shared Documents/", 1)[1]
    elif "Shared Documents" in target_path and target_path.endswith("Shared Documents"):
        # Handle case where path ends with "Shared Documents" (folder itself)
        drive_relative_path = "/"

    logging.info(f"Using drive-relative path for upload: {drive_relative_path}")

    # Upload file
    file_stream.seek(0)
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{drive_relative_path}:/content"
    upload_response = requests.put(upload_url, headers=headers, data=file_stream.read())
    upload_response.raise_for_status()

    web_url = upload_response.json().get("webUrl")
    logging.info(f"File uploaded successfully: {web_url}")
    return web_url

def update_validation_status(token, item_id, status, report_url, list_name="Documents"):
    """Update ValidationStatus column in document library using Graph API"""
    site_info = get_site_info()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    # Get site ID
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_info['hostname']}:{site_info['site_path']}"
    site_response = requests.get(site_url, headers=headers)
    site_response.raise_for_status()
    site_id = site_response.json()["id"]

    # Update list item
    update_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_name}/items/{item_id}/fields"
    update_data = {
        "ValidationStatus": status,
        "LastValidated": datetime.now(timezone.utc).isoformat()
    }
    if report_url:
        update_data["ValidationReport"] = report_url

    update_response = requests.patch(update_url, headers=headers, json=update_data)
    update_response.raise_for_status()

# ============================================
# VALIDATION LOGIC - WORD
# ============================================
def validate_word_document(file_stream, rules):
    """Validate Word document against rules"""
    logging.info("Loading Word document...")
    doc = Document(file_stream)
    logging.info(f"Document loaded. Paragraphs: {len(doc.paragraphs)}, Tables: {len(doc.tables)}")

    issues = []
    fixes_applied = []

    # Filter rules for Word documents
    word_rules = [r for r in rules if r['doc_type'] == 'Word']
    logging.info(f"Applying {len(word_rules)} Word validation rules")

    for idx, rule in enumerate(word_rules):
        logging.info(f"Processing rule {idx+1}/{len(word_rules)}: {rule.get('title', 'Unknown')}")

        if rule['rule_type'] == 'Font':
            result = check_word_fonts(doc, rule)
            issues.extend(result['issues'])
            fixes_applied.extend(result['fixes'])

        elif rule['rule_type'] == 'Color':
            result = check_word_colors(doc, rule)
            issues.extend(result['issues'])
            fixes_applied.extend(result['fixes'])

        elif rule['rule_type'] == 'Language':
            result = validate_language_rules(doc, rule)
            issues.extend(result['issues'])
            fixes_applied.extend(result['fixes'])

        elif rule['rule_type'] == 'Grammar':
            result = validate_grammar_rules(doc, rule)
            issues.extend(result['issues'])
            fixes_applied.extend(result['fixes'])

        elif rule['rule_type'] == 'Punctuation':
            result = validate_punctuation_rules(doc, rule)
            issues.extend(result['issues'])
            fixes_applied.extend(result['fixes'])

        else:
            logging.info(f"Rule type '{rule['rule_type']}' not yet implemented")

    logging.info(f"Validation complete. Issues: {len(issues)}, Fixes: {len(fixes_applied)}")

    return {
        'document': doc,
        'issues': issues,
        'fixes_applied': fixes_applied
    }

def check_word_fonts(doc, rule):
    """Check and fix font issues in Word doc"""
    issues = []
    fixes = []
    expected_font = rule['expected_value']

    logging.info(f"Checking fonts: {rule['check_value']} -> {expected_font}")

    # Check all text fonts
    if rule['check_value'] == 'AllTextFont':
        issue_count = 0
        fix_count = 0

        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if run.text.strip():  # Only check runs with actual text
                    current_font = run.font.name

                    # Handle None font names or mismatches
                    if current_font is None or current_font != expected_font:
                        issue_count += 1

                        if rule['auto_fix']:
                            run.font.name = expected_font
                            fix_count += 1

        if issue_count > 0:
            issues.append(f"Found {issue_count} text runs with incorrect font")
        if fix_count > 0:
            fixes.append(f"Fixed {fix_count} text runs to {expected_font}")

        logging.info(f"Font check complete: {issue_count} issues, {fix_count} fixes")

    # Check Heading 1 font
    elif rule['check_value'] == 'Heading1Font':
        for paragraph in doc.paragraphs:
            if paragraph.style.name == 'Heading 1':
                current_font = paragraph.runs[0].font.name if paragraph.runs else None

                if current_font is None or current_font != expected_font:
                    issues.append(f"Heading 1 has incorrect font: {current_font}")

                    if rule['auto_fix']:
                        for run in paragraph.runs:
                            run.font.name = expected_font
                        fixes.append(f"Fixed Heading 1 font to {expected_font}")

    return {'issues': issues, 'fixes': fixes}

def check_word_colors(doc, rule):
    """Check and fix color issues in Word doc"""
    issues = []
    fixes = []
    
    # Example: Check heading color
    if rule['check_value'] == 'Heading1Color':
        # Parse expected RGB from rule (e.g., "0,51,153")
        expected_rgb = tuple(map(int, rule['expected_value'].split(',')))
        
        for paragraph in doc.paragraphs:
            if paragraph.style.name == 'Heading 1':
                for run in paragraph.runs:
                    if run.font.color.rgb:
                        current_rgb = run.font.color.rgb
                        
                        if current_rgb != expected_rgb:
                            issues.append(f"Heading 1 color incorrect: {current_rgb}")
                            
                            if rule['auto_fix']:
                                run.font.color.rgb = RGBColor(*expected_rgb)
                                fixes.append(f"Fixed Heading 1 color to {expected_rgb}")
    
    return {'issues': issues, 'fixes': fixes}

# ============================================
# VALIDATION LOGIC - VISIO
# ============================================
def validate_visio_document(file_stream, rules):
    """Validate Visio document against rules"""
    visio = VisioFile(file_stream)
    issues = []
    fixes_applied = []
    
    # Filter rules for Visio documents
    visio_rules = [r for r in rules if r['doc_type'] == 'Visio']
    
    for rule in visio_rules:
        if rule['rule_type'] == 'Color':
            result = check_visio_colors(visio, rule)
            issues.extend(result['issues'])
            fixes_applied.extend(result['fixes'])
            
        elif rule['rule_type'] == 'Font':
            result = check_visio_fonts(visio, rule)
            issues.extend(result['issues'])
            fixes_applied.extend(result['fixes'])
    
    return {
        'document': visio,
        'issues': issues,
        'fixes_applied': fixes_applied
    }

def check_visio_colors(visio, rule):
    """Check and fix colors in Visio diagrams"""
    issues = []
    fixes = []
    
    # Example: Check shape fill colors
    if rule['check_value'] == 'ShapeFillColor':
        expected_color = rule['expected_value']  # e.g., "#003399"
        
        for page in visio.pages:
            for shape in page.shapes:
                # TODO: Implement color checking logic
                # vsdx library shape color extraction
                pass
    
    return {'issues': issues, 'fixes': fixes}

def check_visio_fonts(visio, rule):
    """Check and fix fonts in Visio diagrams"""
    issues = []
    fixes = []
    
    # TODO: Implement Visio font checking
    # Note: vsdx library has limited font manipulation capabilities
    
    return {'issues': issues, 'fixes': fixes}

# ============================================
# REPORT GENERATION
# ============================================
def generate_report(file_name, issues, fixes_applied):
    """Generate validation report as HTML"""
    report_html = f"""
    <html>
    <head><title>Validation Report - {file_name}</title></head>
    <body>
        <h1>Style Validation Report</h1>
        <p><strong>File:</strong> {file_name}</p>
        <p><strong>Date:</strong> {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')} UTC</p>
        
        <h2>Summary</h2>
        <p>Issues Found: {len(issues)}</p>
        <p>Issues Fixed: {len(fixes_applied)}</p>
        
        <h2>Issues Detected</h2>
        <ul>
            {''.join(f'<li>{issue}</li>' for issue in issues)}
        </ul>
        
        <h2>Fixes Applied</h2>
        <ul>
            {''.join(f'<li>{fix}</li>' for fix in fixes_applied)}
        </ul>
    </body>
    </html>
    """
    return report_html

# ============================================
# MAIN FUNCTION
# ============================================
def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('=== STYLE VALIDATION FUNCTION TRIGGERED ===')

    try:
        # 1. Parse request
        logging.info('Step 1: Parsing request...')
        req_body = req.get_json()
        logging.info(f"Request keys: {list(req_body.keys())}")

        item_id = req_body.get('itemId') or req_body.get('ID')

        # Try multiple possible parameter names for file name
        file_name = (req_body.get('fileName') or
                    req_body.get('FileLeafRef') or
                    req_body.get('Name'))

        file_extension = os.path.splitext(file_name)[1].lower() if file_name else ''

        # Check if file content is provided directly (base64 encoded)
        file_content_base64 = req_body.get('fileContent')

        # Try multiple possible parameter names for file URL (for legacy support)
        file_url = (req_body.get('fileUrl') or
                   req_body.get('FileRef') or
                   req_body.get('ServerRelativeUrl') or
                   req_body.get('fileRef'))

        logging.info(f"Request params - File: {file_name}, ID: {item_id}, Has content: {bool(file_content_base64)}, URL: {file_url}")

        # 2. Get Microsoft Graph API token
        logging.info('Step 2: Acquiring Graph API token...')
        token = get_graph_token()
        logging.info('Token acquired successfully')

        # 3. Update status to "Validating..."
        logging.info('Step 3: Updating validation status...')
        update_validation_status(token, item_id, "Validating...", None)

        # 4. Fetch validation rules
        logging.info('Step 4: Fetching validation rules...')
        rules = fetch_validation_rules(token)
        logging.info(f"Loaded {len(rules)} validation rules")

        # 5. Get file content
        if file_content_base64:
            # File content provided directly as base64
            logging.info('Step 5: Decoding file content from base64...')
            import base64
            file_bytes = base64.b64decode(file_content_base64)
            file_stream = BytesIO(file_bytes)
            logging.info(f"File decoded successfully, size: {len(file_bytes)} bytes")
        elif file_url:
            # Download file from SharePoint using Graph API
            logging.info('Step 5: Downloading file from SharePoint...')
            file_stream = download_file(token, file_url)
        else:
            raise ValueError("Either fileContent or fileUrl must be provided")
        
        # 6. Validate based on file type
        logging.info(f'Step 6: Validating document type {file_extension}...')
        if file_extension in ['.docx', '.doc']:
            result = validate_word_document(file_stream, rules)

            # Save fixed document
            logging.info('Saving fixed document to stream...')
            fixed_stream = BytesIO()
            result['document'].save(fixed_stream)
            fixed_stream.seek(0)

        elif file_extension in ['.vsdx', '.vsd']:
            result = validate_visio_document(file_stream, rules)

            # Save fixed document
            fixed_stream = BytesIO()
            result['document'].save_vsdx(fixed_stream)
            fixed_stream.seek(0)

        else:
            logging.error(f"Unsupported file type: {file_extension}")
            return func.HttpResponse(
                json.dumps({"error": f"Unsupported file type: {file_extension}"}),
                status_code=400
            )

        # 7. Upload fixed file (overwrite original) - only if file_url is provided
        if file_url and result['fixes_applied']:
            logging.info(f'Step 7: Uploading fixed document ({len(result["fixes_applied"])} fixes)...')
            upload_file(token, fixed_stream, file_url)
        elif not file_url:
            logging.info('Step 7: Skipping upload (no file URL provided, file content was sent directly)')
        else:
            logging.info('Step 7: No fixes to upload')

        # 8. Generate report
        logging.info('Step 8: Generating report...')
        report_html = generate_report(file_name, result['issues'], result['fixes_applied'])
        report_url = None

        # Upload report only if we have a file URL
        if file_url:
            logging.info('Uploading report to SharePoint...')
            report_stream = BytesIO(report_html.encode('utf-8'))
            report_filename = f"{os.path.splitext(file_name)[0]}_ValidationReport.html"
            report_folder = os.path.dirname(file_url)
            report_path = f"{report_folder}/{report_filename}" if report_folder else f"/{report_filename}"
            logging.info(f"Report will be uploaded to: {report_path}")
            report_url = upload_file(token, report_stream, report_path)
        else:
            logging.info('Skipping report upload (no file URL provided)')

        # 9. Update validation status
        logging.info('Step 9: Updating final validation status...')
        # Pass if no issues, or if all issues were auto-fixed
        unfixed_issues = len(result['issues']) - len(result['fixes_applied'])
        final_status = "Passed" if unfixed_issues == 0 else "Failed"
        logging.info(f"Issues: {len(result['issues'])}, Fixes: {len(result['fixes_applied'])}, Unfixed: {unfixed_issues}")
        update_validation_status(token, item_id, final_status, report_url)

        # 10. Return response with fixed file content
        logging.info(f'=== VALIDATION COMPLETE: {final_status} ===')

        response_data = {
            "status": final_status,
            "issuesFound": len(result['issues']),
            "issuesFixed": len(result['fixes_applied']),
            "reportUrl": report_url
        }

        # Include fixed file content if fixes were applied
        if result['fixes_applied']:
            import base64
            fixed_stream.seek(0)
            fixed_content_base64 = base64.b64encode(fixed_stream.read()).decode('utf-8')
            response_data["fixedFileContent"] = fixed_content_base64
            logging.info(f"Returning fixed file content ({len(fixed_content_base64)} chars)")

        return func.HttpResponse(
            json.dumps(response_data),
            mimetype="application/json",
            status_code=200
        )

    except Exception as e:
        logging.error(f"=== VALIDATION FAILED ===")
        logging.error(f"Error: {str(e)}")
        logging.error(f"Error type: {type(e).__name__}")
        import traceback
        logging.error(f"Traceback: {traceback.format_exc()}")
        return func.HttpResponse(
            json.dumps({
                "error": str(e),
                "error_type": type(e).__name__
            }),
            mimetype="application/json",
            status_code=500
        )