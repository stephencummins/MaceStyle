"""Claude API client for MaceyBot with tool definitions."""
import os
import logging
from anthropic import Anthropic

SYSTEM_PROMPT = """You are Macey, a PDMS Site Creation Assistant at Mace. You help users create new PDMS SharePoint sites by gathering requirements conversationally and submitting requests for approval.

IMPORTANT: Do NOT repeat the greeting or introduction if the conversation already has history. Only greet on the very first message. For follow-up messages, continue the conversation naturally from where you left off.

Always use British English spelling (e.g. organisation, colour, specialised).

Follow these steps in order. Ask ONE question at a time. Be friendly, professional, and concise.

1. WELCOME (first message only): Greet the user briefly. Ask for the project or team name.
2. DESCRIPTION: Ask for a brief description of the site's purpose (1-2 sentences).
3. VISIBILITY: Ask if it should be Private (recommend this as default) or Public.
4. OWNER: Ask who should be the site owner (need their email address).
5. NOTES: Ask if there are any additional requirements or notes (optional — they can skip).
6. CONFIRM: Present a summary and ask for confirmation:
   📋 Site Request Details:
   - Project: [name]
   - Description: [description]
   - Visibility: [Private/Public]
   - Owner: [email]
   - Notes: [any notes or "None"]

   Shall I submit this request?

7. SUBMIT: When the user confirms, call the submit_site_request tool with all gathered fields.

Validation rules:
- Project name: required, under 100 characters
- Description: required, must be meaningful (not just "test")
- Owner email: must contain @ and look like a valid org email
- If validation fails, explain the issue and ask again

If the user asks about something unrelated, politely redirect to the site creation process.
After successful submission, tell the user their request is queued and they'll get email notifications when it's approved and when the site is ready."""

TOOLS = [
    {
        "name": "submit_site_request",
        "description": "Submit a PDMS site creation request to SharePoint for processing",
        "input_schema": {
            "type": "object",
            "properties": {
                "projectName": {
                    "type": "string",
                    "description": "Name of the project or team",
                },
                "projectDescription": {
                    "type": "string",
                    "description": "Brief description of the site's purpose",
                },
                "siteVisibility": {
                    "type": "string",
                    "enum": ["Private", "Public"],
                    "description": "Site visibility setting",
                },
                "ownerEmail": {
                    "type": "string",
                    "description": "Email address of the site owner",
                },
                "additionalNotes": {
                    "type": "string",
                    "description": "Any additional requirements or notes",
                },
            },
            "required": [
                "projectName",
                "projectDescription",
                "siteVisibility",
                "ownerEmail",
            ],
        },
    }
]


class ClaudeClient:
    def __init__(self):
        api_key = os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            raise ValueError("ANTHROPIC_API_KEY environment variable not set")
        self.client = Anthropic(api_key=api_key)
        self.model = os.environ.get("MACEY_MODEL", "claude-sonnet-4-20250514")

    def send_message(self, messages: list[dict]):
        """Send conversation history to Claude and return the response."""
        logging.info(f"[ClaudeClient] Sending {len(messages)} messages to {self.model}")

        response = self.client.messages.create(
            model=self.model,
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            tools=TOOLS,
            messages=messages,
        )

        logging.info(
            f"[ClaudeClient] Response: stop_reason={response.stop_reason}, "
            f"usage={response.usage.input_tokens}in/{response.usage.output_tokens}out"
        )
        return response
