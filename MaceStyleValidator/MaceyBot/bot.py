"""MaceyBot — conversation handler using Claude API with tool use."""
import json
import logging
import os

from botbuilder.core import ActivityHandler, TurnContext
from botbuilder.schema import ChannelAccount

from .claude_client import ClaudeClient
from .sharepoint_client import submit_to_sharepoint

MAX_HISTORY = 20


def _blob_container():
    """Get or create the blob container client for conversation history."""
    from azure.storage.blob import ContainerClient
    conn_str = os.environ.get("AzureWebJobsStorage", "")
    container = ContainerClient.from_connection_string(conn_str, "maceybot-conversations")
    try:
        container.create_container()
    except Exception:
        pass  # already exists
    return container


def _blob_name(conversation_id: str) -> str:
    """Sanitise conversation ID for use as blob name."""
    safe = conversation_id.replace("/", "_").replace("\\", "_")
    return f"{safe}.json"


def _get_history(conversation_id: str) -> list[dict]:
    try:
        container = _blob_container()
        blob = container.get_blob_client(_blob_name(conversation_id))
        data = blob.download_blob().readall()
        history = json.loads(data)
        logging.info(f"[MaceyBot] Loaded {len(history)} messages from blob for {conversation_id}")
        return history
    except Exception:
        logging.info(f"[MaceyBot] No existing history for {conversation_id}")
        return []


def _save_history(conversation_id: str, history: list[dict]):
    try:
        container = _blob_container()
        blob = container.get_blob_client(_blob_name(conversation_id))

        # Claude content blocks aren't JSON-serialisable by default — convert them
        def _serialise(obj):
            if hasattr(obj, "model_dump"):
                return obj.model_dump()
            if hasattr(obj, "to_dict"):
                return obj.to_dict()
            if hasattr(obj, "__dict__"):
                return obj.__dict__
            return str(obj)

        blob.upload_blob(
            json.dumps(history, default=_serialise),
            overwrite=True,
        )
        logging.info(f"[MaceyBot] Saved {len(history)} messages to blob for {conversation_id}")
    except Exception as e:
        logging.error(f"[MaceyBot] Failed to save history: {e}", exc_info=True)


def _trim_history(history: list[dict]) -> list[dict]:
    """Keep the last MAX_HISTORY messages to control token usage."""
    if len(history) > MAX_HISTORY:
        return history[-MAX_HISTORY:]
    return history


class MaceyBot(ActivityHandler):
    def __init__(self):
        self.claude = ClaudeClient()

    async def on_message_activity(self, turn_context: TurnContext):
        conversation_id = turn_context.activity.conversation.id
        user_text = turn_context.activity.text or ""

        if not user_text.strip():
            return

        logging.info(f"[MaceyBot] Message from {conversation_id}: {user_text[:100]}")

        history = _get_history(conversation_id)
        history.append({"role": "user", "content": user_text})
        history = _trim_history(history)

        try:
            response = self.claude.send_message(history)

            # Check for tool use
            if response.stop_reason == "tool_use":
                reply_parts = []
                tool_results = []

                for block in response.content:
                    if block.type == "text" and block.text:
                        reply_parts.append(block.text)
                    elif block.type == "tool_use":
                        logging.info(f"[MaceyBot] Tool call: {block.name} with {block.input}")

                        if block.name == "submit_site_request":
                            result = submit_to_sharepoint(block.input)
                            tool_results.append({
                                "type": "tool_result",
                                "tool_use_id": block.id,
                                "content": result,
                            })
                        else:
                            tool_results.append({
                                "type": "tool_result",
                                "tool_use_id": block.id,
                                "content": f"Unknown tool: {block.name}",
                                "is_error": True,
                            })

                # Send any text before the tool call
                if reply_parts:
                    await turn_context.send_activity(" ".join(reply_parts))

                # Add assistant response (with tool_use) and tool results to history
                history.append({"role": "assistant", "content": response.content})
                history.append({"role": "user", "content": tool_results})

                # Get Claude's final response after tool execution
                final_response = self.claude.send_message(history)
                assistant_text = _extract_text(final_response)

                history.append({"role": "assistant", "content": assistant_text})
                _save_history(conversation_id, _trim_history(history))

                await turn_context.send_activity(assistant_text)

            else:
                # Normal text response
                assistant_text = _extract_text(response)
                history.append({"role": "assistant", "content": assistant_text})
                _save_history(conversation_id, _trim_history(history))

                await turn_context.send_activity(assistant_text)

        except Exception as e:
            logging.error(f"[MaceyBot] Error: {e}", exc_info=True)
            await turn_context.send_activity(
                "I'm having trouble processing your request right now. "
                "Please try again in a moment."
            )

    async def on_members_added_activity(self, members_added: list[ChannelAccount], turn_context: TurnContext):
        for member in members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity(
                    "Hello! I'm **Macey**, the PDMS Site Creation Assistant.\n\n"
                    "I can help you request a new PDMS SharePoint site. "
                    "Just tell me the project or team name to get started!"
                )


def _extract_text(response) -> str:
    """Extract text content from a Claude API response."""
    parts = []
    for block in response.content:
        if hasattr(block, "text") and block.text:
            parts.append(block.text)
    return " ".join(parts) if parts else "I'm sorry, I didn't get a response. Could you try again?"
