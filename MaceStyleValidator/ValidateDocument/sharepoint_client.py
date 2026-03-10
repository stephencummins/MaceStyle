"""SharePoint / Microsoft Graph API operations"""
import os
import logging
import requests
from io import BytesIO
from urllib.parse import quote
from datetime import datetime, timezone
from .config import get_site_info, DOC_LIBRARY_LIST_ID


def get_site_id(token):
    """Get SharePoint site ID"""
    site_info = get_site_info()
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_info['hostname']}:{site_info['site_path']}"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()["id"]


def fetch_validation_rules(token):
    """Fetch rules from SharePoint 'Style Rules' list"""
    site_id = get_site_id(token)
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/items?expand=fields"
    response = requests.get(list_url, headers=headers)
    response.raise_for_status()

    rules = []
    for item in response.json().get("value", []):
        fields = item.get("fields", {})
        rules.append({
            'title': fields.get('Title'),
            'rule_type': fields.get('RuleType') or fields.get('field_1'),
            'doc_type': fields.get('DocumentType') or fields.get('field_2'),
            'check_value': fields.get('CheckValue') or fields.get('field_3'),
            'expected_value': fields.get('ExpectedValue') or fields.get('field_4'),
            'auto_fix': fields.get('AutoFix') if fields.get('AutoFix') is not None else fields.get('field_5'),
            'use_ai': fields.get('UseAI') if fields.get('UseAI') is not None else fields.get('field_7', False),
            'priority': fields.get('Priority') or fields.get('field_6', 999)
        })

    rules.sort(key=lambda x: x['priority'])
    return rules


def download_file(token, file_path):
    """Download file from SharePoint using Graph API"""
    if not file_path:
        raise ValueError("file_path cannot be None or empty")

    logging.info(f"Downloading file: {file_path}")
    site_id = get_site_id(token)
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    drive_relative_path = file_path
    if "Shared Documents/" in file_path:
        drive_relative_path = "/" + file_path.split("Shared Documents/", 1)[1]

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{drive_relative_path}:/content"
    response = requests.get(url, headers=headers)
    response.raise_for_status()

    logging.info(f"File downloaded, size: {len(response.content)} bytes")
    return BytesIO(response.content)


def upload_file(token, file_stream, target_path):
    """Upload file to SharePoint using Graph API"""
    if not target_path:
        raise ValueError("target_path cannot be None or empty")

    logging.info(f"Uploading file to: {target_path}")
    site_id = get_site_id(token)
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream"}

    drive_relative_path = target_path
    if "Shared Documents/" in target_path:
        drive_relative_path = "/" + target_path.split("Shared Documents/", 1)[1]

    file_stream.seek(0)
    encoded_path = quote(drive_relative_path, safe='/')
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{encoded_path}:/content"
    response = requests.put(url, headers=headers, data=file_stream.read())
    response.raise_for_status()

    web_url = response.json().get("webUrl")
    logging.info(f"File uploaded: {web_url}")
    return web_url


def update_validation_status(token, item_id, status, report_url):
    """Update ValidationStatus column in document library"""
    site_id = get_site_id(token)
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{DOC_LIBRARY_LIST_ID}/items/{item_id}/fields"
    data = {
        "ValidationStatus": status,
        "LastValidated": datetime.now(timezone.utc).isoformat()
    }
    if report_url:
        data["ValidationReport"] = report_url

    response = requests.patch(url, headers=headers, json=data)
    if response.status_code >= 400:
        logging.warning(
            f"Could not update validation status (HTTP {response.status_code}): {response.text}. "
            f"Check that ValidationStatus, LastValidated, and ValidationReport columns exist."
        )
    else:
        logging.info(f"Validation status updated for item {item_id}")
