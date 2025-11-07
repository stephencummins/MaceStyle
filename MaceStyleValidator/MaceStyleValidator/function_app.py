
@app.route(route="ValidateSimple", methods=["POST"])
def ValidateSimple(req: func.HttpRequest) -> func.HttpResponse:
    """Simple validation test without SharePoint list tracking"""
    try:
        from ValidateDocument import get_graph_token, download_file, validate_word_document, generate_report

        # Parse request
        req_body = req.get_json()
        file_path = req_body.get('filePath')  # e.g., "/Shared Documents/test.docx"

        if not file_path:
            return func.HttpResponse(
                json.dumps({"error": "filePath is required"}),
                mimetype="application/json",
                status_code=400
            )

        # Get token
        token = get_graph_token()

        # Download file
        file_stream = download_file(token, file_path)

        # Create mock rules for testing
        mock_rules = [
            {
                'title': 'Heading Font Check',
                'rule_type': 'Font',
                'doc_type': 'Word',
                'check_value': 'Heading1Font',
                'expected_value': 'Arial',
                'auto_fix': True,
                'priority': 1
            }
        ]

        # Validate
        result = validate_word_document(file_stream, mock_rules)

        # Generate report
        file_name = file_path.split('/')[-1]
        report_html = generate_report(file_name, result['issues'], result['fixes_applied'])

        return func.HttpResponse(
            json.dumps({
                "status": "success",
                "fileName": file_name,
                "issuesFound": len(result['issues']),
                "issuesFixed": len(result['fixes_applied']),
                "issues": result['issues'],
                "fixes": result['fixes_applied']
            }, indent=2),
            mimetype="application/json"
        )

    except Exception as e:
        import traceback
        return func.HttpResponse(
            json.dumps({
                "error": str(e),
                "traceback": traceback.format_exc()
            }),
            mimetype="application/json",
            status_code=500
        )

@app.route(route="SetupRules", methods=["POST"])
def SetupRules(req: func.HttpRequest) -> func.HttpResponse:
    """Populate Style Rules list with default validation rules"""
    try:
        # Get Graph API token
        tenant_id = os.environ.get("SHAREPOINT_TENANT_ID")
        client_id = os.environ.get("SHAREPOINT_CLIENT_ID")
        client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET")
        site_url = os.environ.get("SHAREPOINT_SITE_URL")
        
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        scope = ["https://graph.microsoft.com/.default"]
        
        app_client = msal.ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret
        )
        
        result = app_client.acquire_token_for_client(scopes=scope)
        if "access_token" not in result:
            return func.HttpResponse(
                json.dumps({"error": "Failed to acquire token"}),
                status_code=500
            )
        
        token = result["access_token"]
        
        # Get site ID
        parts = site_url.replace("https://", "").split("/")
        hostname = parts[0]
        site_path = "/" + "/".join(parts[1:]) if len(parts) > 1 else ""
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        graph_site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
        site_response = requests.get(graph_site_url, headers=headers)
        site_response.raise_for_status()
        site_id = site_response.json()["id"]
        
        # Default validation rules
        default_rules = [
            {
                "Title": "Heading 1 Font Check",
                "RuleType": "Font",
                "DocumentType": "Word",
                "CheckValue": "Heading1Font",
                "ExpectedValue": "Arial",
                "AutoFix": True,
                "Priority": 1
            },
            {
                "Title": "Heading 1 Color Check",
                "RuleType": "Color",
                "DocumentType": "Word",
                "CheckValue": "Heading1Color",
                "ExpectedValue": "0,51,153",
                "AutoFix": True,
                "Priority": 2
            }
        ]
        
        # Add rules to Style Rules list
        created_rules = []
        list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/Style Rules/items"
        
        for rule in default_rules:
            item_data = {"fields": rule}
            response = requests.post(list_url, headers=headers, json=item_data)
            if response.status_code == 201:
                created_rules.append(rule["Title"])
            else:
                return func.HttpResponse(
                    json.dumps({
                        "error": f"Failed to create rule: {rule['Title']}",
                        "details": response.text
                    }),
                    status_code=500
                )
        
        return func.HttpResponse(
            json.dumps({
                "status": "success",
                "message": f"Created {len(created_rules)} validation rules",
                "rules": created_rules
            }, indent=2),
            mimetype="application/json"
        )
        
    except Exception as e:
        import traceback
        return func.HttpResponse(
            json.dumps({
                "error": str(e),
                "traceback": traceback.format_exc()
            }),
            status_code=500
        )
