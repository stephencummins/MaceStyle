"""
Save validation results to SharePoint Validation Results list
"""
import os
import logging
import requests
from datetime import datetime

def save_validation_result(token, site_id, filename, issues_count, fixes_count, status, html_report, report_url=None):
    """
    Save validation result to SharePoint Validation Results list

    Args:
        token: Microsoft Graph API access token
        site_id: SharePoint site ID
        filename: Name of validated document
        issues_count: Number of issues found
        fixes_count: Number of fixes applied
        status: Validation status ("Passed", "Failed", "Warning")
        html_report: HTML report content (stored in field)
        report_url: Optional URL to uploaded HTML report file

    Returns:
        dict with 'item_id', 'report_url', and 'list_item_url'
    """
    logging.info(f"Saving validation result for {filename} to SharePoint...")

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }

    # Create list item
    # Use list ID for reliability (found via inspect_validation_results.py)
    list_id = "d4f4cc72-7f68-4009-a1eb-e86d9e67a4dd"
    list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"

    item_data = {
        "fields": {
            "Title": f"Validation: {filename}",
            "FileName": filename,
            "ValidationDate": datetime.utcnow().isoformat() + "Z",
            "Status": status,
            "IssuesFound": str(issues_count),
            "IssuesFixed": str(fixes_count)
        }
    }

    logging.info(f"Creating list item with data: {item_data}")
    response = requests.post(list_url, headers=headers, json=item_data)
    response.raise_for_status()

    item = response.json()
    item_id = item["id"]
    logging.info(f"✓ Created validation result item ID: {item_id}")

    # Update ReportLink field if report_url provided
    if report_url:
        logging.info(f"Updating ReportLink field with URL: {report_url}")
        update_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}"
        update_data = {
            "fields": {
                "ReportLink": {
                    "Description": "View HTML Report",
                    "Url": report_url
                }
            }
        }

        response = requests.patch(update_url, headers=headers, json=update_data)
        response.raise_for_status()
        logging.info(f"✓ Updated ReportLink field")

    list_item_url = f"https://0rxf2.sharepoint.com/sites/StyleValidation/Lists/Validation%20Results/DispForm.aspx?ID={item_id}"
    logging.info(f"✓ Validation result saved: {list_item_url}")

    return {
        "item_id": item_id,
        "report_url": report_url,
        "list_item_url": list_item_url
    }


def update_document_metadata(token, site_id, file_url, validation_result_url):
    """
    Update document metadata in Document Library with link to validation result

    Args:
        token: Microsoft Graph API access token
        site_id: SharePoint site ID
        file_url: Full SharePoint URL or path to the document
        validation_result_url: URL to the validation result
    """
    logging.info(f"Updating document metadata with validation result link...")
    logging.info(f"File URL: {file_url}")
    logging.info(f"Validation result URL: {validation_result_url}")

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    try:
        # Extract the file path from the URL
        # file_url format: /sites/StyleValidation/Shared Documents/filename.docx
        # We need to get the drive item by path

        # Get the drive first
        drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
        drive_response = requests.get(drive_url, headers=headers)
        drive_response.raise_for_status()
        drive_id = drive_response.json()["id"]

        logging.info(f"Drive ID: {drive_id}")

        # Parse the file path - file_url is like "/sites/StyleValidation/Shared Documents/test.docx"
        # We need the path relative to the drive root
        if "/Shared Documents/" in file_url:
            # Extract path after "Shared Documents"
            relative_path = file_url.split("/Shared Documents/", 1)[1]
            # URL encode the path
            import urllib.parse
            encoded_path = urllib.parse.quote(relative_path)

            # Get the drive item
            item_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{encoded_path}"
            logging.info(f"Getting drive item: {item_url}")

            item_response = requests.get(item_url, headers=headers)
            item_response.raise_for_status()
            item_data = item_response.json()

            list_item_id = item_data.get("listItem", {}).get("id")

            if not list_item_id:
                logging.warning("Could not find listItem ID for document")
                return False

            logging.info(f"Found list item ID: {list_item_id}")

            # Update the ValidationResultLink field
            # First, get the list ID from the drive
            list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/list"
            list_response = requests.get(list_url, headers=headers)
            list_response.raise_for_status()
            list_id = list_response.json()["id"]

            logging.info(f"List ID: {list_id}")

            # Update the list item with ValidationResultLink
            update_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{list_item_id}"

            update_data = {
                "fields": {
                    "ValidationResultLink": {
                        "Description": "View Validation Result",
                        "Url": validation_result_url
                    }
                }
            }

            logging.info(f"Updating list item with: {update_data}")
            update_response = requests.patch(update_url, headers=headers, json=update_data)
            update_response.raise_for_status()

            logging.info(f"✓ Successfully updated ValidationResultLink for document")
            return True

        else:
            logging.warning(f"File URL format not recognized: {file_url}")
            return False

    except Exception as e:
        logging.error(f"Failed to update document metadata: {str(e)}")
        import traceback
        logging.error(f"Traceback: {traceback.format_exc()}")
        return False
