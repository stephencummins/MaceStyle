"""MaceStyle Document Validator - Azure Function entry point"""
import azure.functions as func
import logging
import json
import os
import base64
from io import BytesIO

from .config import get_graph_token
from .sharepoint_client import (
    get_site_id, fetch_validation_rules, download_file, upload_file,
    update_validation_status, update_drive_item_fields
)
from .report import generate_report
from .sharepoint_results import save_validation_result, update_document_metadata
from .word_validator import validate_word_document
from .visio_validator import validate_visio_document
from .excel_validator import validate_excel_document
from .powerpoint_validator import validate_powerpoint_document


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('=== STYLE VALIDATION FUNCTION TRIGGERED ===')

    try:
        # 1. Parse request
        req_body = req.get_json()
        logging.info(f"Request keys: {list(req_body.keys())}")

        item_id = req_body.get('itemId') or req_body.get('ID')
        file_name = (req_body.get('fileName') or
                     req_body.get('FileLeafRef') or
                     req_body.get('Name'))
        file_extension = os.path.splitext(file_name)[1].lower() if file_name else ''
        file_content_base64 = req_body.get('fileContent')
        file_url = (req_body.get('fileUrl') or
                    req_body.get('FileRef') or
                    req_body.get('ServerRelativeUrl') or
                    req_body.get('fileRef'))

        logging.info(f"File: {file_name}, ID: {item_id}, Has content: {bool(file_content_base64)}, URL: {file_url}")

        # 2. Get Graph API token
        token = get_graph_token()
        logging.info('Token acquired')

        # 3. Update status to "Validating..."
        try:
            update_validation_status(token, item_id, "Validating...", None)
        except Exception as e:
            logging.warning(f"Could not set initial status: {e}")

        # 4. Fetch validation rules
        rules = fetch_validation_rules(token)
        logging.info(f"Loaded {len(rules)} validation rules")

        # 5. Get file content
        if file_content_base64:
            file_bytes = base64.b64decode(file_content_base64)
            file_stream = BytesIO(file_bytes)
            logging.info(f"File decoded from base64, size: {len(file_bytes)} bytes")
        elif file_url:
            file_stream = download_file(token, file_url)
        else:
            raise ValueError("Either fileContent or fileUrl must be provided")

        # 6. Validate based on file type
        logging.info(f'Validating {file_extension} document...')

        if file_extension in ['.docx', '.doc', '.docm', '.dotx', '.dotm']:
            result = validate_word_document(file_stream, rules)
            fixed_stream = BytesIO()
            result['document'].save(fixed_stream)
            fixed_stream.seek(0)

        elif file_extension in ['.vsdx', '.vsd']:
            result = validate_visio_document(file_stream, rules)
            import tempfile
            with tempfile.NamedTemporaryFile(suffix='.vsdx', delete=False) as tmp_out:
                tmp_out_path = tmp_out.name
            result['document'].save_vsdx(tmp_out_path)
            fixed_stream = BytesIO(open(tmp_out_path, 'rb').read())
            os.unlink(tmp_out_path)

        elif file_extension in ['.xlsx', '.xls', '.xlsm']:
            result = validate_excel_document(file_stream, rules)
            fixed_stream = BytesIO()
            result['document'].save(fixed_stream)
            fixed_stream.seek(0)

        elif file_extension in ['.pptx', '.ppt', '.pptm', '.potx', '.potm']:
            result = validate_powerpoint_document(file_stream, rules)
            fixed_stream = BytesIO()
            result['document'].save(fixed_stream)
            fixed_stream.seek(0)

        else:
            return func.HttpResponse(
                json.dumps({"error": f"Unsupported file type: {file_extension}"}),
                mimetype="application/json",
                status_code=400
            )

        # 7. Upload fixed file if fixes were applied
        if file_url and result['fixes_applied']:
            logging.info(f'Uploading fixed document ({len(result["fixes_applied"])} fixes)...')
            _web_url, _item_id = upload_file(token, fixed_stream, file_url)
        elif not file_url:
            logging.info('Skipping upload (no file URL)')
        else:
            logging.info('No fixes to upload')

        # 8. Generate and upload report
        report_html = generate_report(file_name, result['issues'], result['fixes_applied'])
        report_url = None
        report_drive_item_id = None

        report_stream = BytesIO(report_html.encode('utf-8'))
        report_filename = f"{os.path.splitext(file_name)[0]}_ValidationReport.html"
        if file_url:
            report_folder = os.path.dirname(file_url)
            report_path = f"{report_folder}/{report_filename}" if report_folder else f"/{report_filename}"
        else:
            report_path = f"/Validation Reports/{report_filename}"

        try:
            report_url, report_drive_item_id = upload_file(token, report_stream, report_path)
            logging.info(f"Report uploaded: {report_url}")
        except Exception as e:
            logging.error(f"Failed to upload report: {e}")

        # 9. Save validation result to SharePoint list
        validation_result_info = None
        try:
            site_id = get_site_id(token)
            remaining = [i for i in result['issues'] if isinstance(i, dict)]
            if len(remaining) == 0:
                result_status = "Passed"
            elif len(result['fixes_applied']) > 0:
                result_status = "Review Required"
            else:
                result_status = "Failed"

            # Update report file metadata (ValidationStatus + counts)
            if report_drive_item_id:
                try:
                    update_drive_item_fields(token, report_drive_item_id, {
                        "ValidationStatus": result_status,
                        "IssuesFound": len(result['issues']),
                        "IssuesFixed": len(result['fixes_applied'])
                    })
                except Exception as e:
                    logging.warning(f"Could not update report metadata: {e}")

            validation_result_info = save_validation_result(
                token=token,
                site_id=site_id,
                filename=file_name,
                issues_count=len(result['issues']),
                fixes_count=len(result['fixes_applied']),
                status=result_status,
                html_report=report_html,
                report_url=report_url
            )
            logging.info(f"Validation result saved: {validation_result_info['list_item_url']}")

            # Update document metadata with link to validation result
            if file_url:
                update_document_metadata(token, site_id, file_url, validation_result_info['list_item_url'])
            elif item_id:
                _update_metadata_by_item_id(token, site_id, item_id, validation_result_info['list_item_url'])

        except Exception as e:
            logging.error(f"Failed to save validation result: {e}")

        # 10. Update final validation status
        remaining = [i for i in result['issues'] if isinstance(i, dict)]
        if len(remaining) == 0:
            final_status = "Passed"
        elif len(result['fixes_applied']) > 0:
            final_status = "Review Required"
        else:
            final_status = "Failed"
        try:
            update_validation_status(token, item_id, final_status, report_url)
        except Exception as e:
            logging.warning(f"Could not set final status: {e}")

        # 11. Return response
        logging.info(f'=== VALIDATION COMPLETE: {final_status} ===')

        issues_count = len(result['issues'])
        fixes_count = len(result['fixes_applied'])
        remaining_count = len(remaining)
        if fixes_count > 0 and remaining_count == 0:
            description = f"{final_status} — {fixes_count} issue{'s' if fixes_count != 1 else ''} auto-fixed"
        elif fixes_count > 0:
            description = f"{final_status} — {fixes_count} fixed, {remaining_count} remaining"
        elif issues_count > 0:
            description = f"{final_status} — {issues_count} issue{'s' if issues_count != 1 else ''} found"
        else:
            description = f"{final_status} — no issues found"

        response_data = {
            "status": final_status,
            "description": description,
            "issuesFound": issues_count,
            "issuesFixed": fixes_count,
            "reportUrl": report_url,
            "validationResultUrl": validation_result_info['list_item_url'] if validation_result_info else None,
            "reportLink": {
                "Description": "View HTML Report",
                "Url": report_url
            } if report_url else None,
            "validationResultLink": {
                "Description": "View Validation Result",
                "Url": validation_result_info['list_item_url']
            } if validation_result_info else None
        }

        if result['fixes_applied']:
            fixed_stream.seek(0)
            response_data["fixedFileContent"] = base64.b64encode(fixed_stream.read()).decode('utf-8')

        return func.HttpResponse(
            json.dumps(response_data),
            mimetype="application/json",
            status_code=200
        )

    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        logging.error(f"=== VALIDATION FAILED ===")
        logging.error(f"Error: {e}")
        logging.error(f"Traceback: {tb}")
        return func.HttpResponse(
            json.dumps({
                "error": str(e),
                "error_type": type(e).__name__,
                "traceback": tb
            }),
            mimetype="application/json",
            status_code=500
        )


def _update_metadata_by_item_id(token, site_id, item_id, validation_result_url):
    """Update document metadata using item_id when file_url is not available"""
    import requests
    from .config import DOC_LIBRARY_LIST_ID

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{DOC_LIBRARY_LIST_ID}/items/{item_id}/fields"
    data = {
        "ValidationResultLink": json.dumps({
            "Description": "View Validation Result",
            "Url": validation_result_url
        })
    }
    resp = requests.patch(url, headers=headers, json=data)
    if resp.status_code >= 400:
        logging.warning(f"Failed to update ValidationResultLink via item_id: {resp.status_code} {resp.text}")
    else:
        logging.info("Document metadata updated with validation result link")
