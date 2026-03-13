"""MaceyBot — Teams bot endpoint powered by Claude API."""
import azure.functions as func
import json
import logging
import os
import traceback
import asyncio

from botbuilder.core import (
    BotFrameworkAdapter,
    BotFrameworkAdapterSettings,
    TurnContext,
)
from botbuilder.schema import Activity

from .bot import MaceyBot

# Adapter (initialised lazily so env vars are read at runtime)
_adapter = None
_bot = None


def _get_adapter():
    global _adapter
    if _adapter is None:
        settings = BotFrameworkAdapterSettings(
            app_id=os.environ.get("BOT_APP_ID", ""),
            app_password=os.environ.get("BOT_APP_PASSWORD", ""),
            channel_auth_tenant=os.environ.get("BOT_TENANT_ID", ""),
        )
        _adapter = BotFrameworkAdapter(settings)

        async def on_error(context: TurnContext, error: Exception):
            logging.error(f"[MaceyBot] Unhandled error: {error}")
            logging.error(traceback.format_exc())
            try:
                await context.send_activity("Sorry, something went wrong. Please try again in a moment.")
            except Exception:
                pass

        _adapter.on_turn_error = on_error
    return _adapter


def _get_bot():
    global _bot
    if _bot is None:
        _bot = MaceyBot()
    return _bot


async def main(req: func.HttpRequest) -> func.HttpResponse:
    """HTTP trigger — Azure Bot Service messaging endpoint."""
    logging.info(f"[MaceyBot] Request received: {req.method}")

    if req.method != "POST":
        return func.HttpResponse("MaceyBot endpoint is running.", status_code=200)

    try:
        body = req.get_json()
        logging.info(f"[MaceyBot] Activity type: {body.get('type', 'unknown')}")
    except ValueError:
        logging.error("[MaceyBot] Invalid JSON body")
        return func.HttpResponse("Invalid JSON", status_code=400)

    activity = Activity.deserialize(body)
    auth_header = req.headers.get("Authorization", "")

    adapter = _get_adapter()
    bot = _get_bot()

    try:
        response = await adapter.process_activity(activity, auth_header, bot.on_turn)

        if response:
            return func.HttpResponse(
                json.dumps(response.body),
                status_code=response.status,
                mimetype="application/json",
            )
        return func.HttpResponse(status_code=200)

    except Exception as e:
        logging.error(f"[MaceyBot] process_activity error: {e}")
        logging.error(traceback.format_exc())
        return func.HttpResponse(
            json.dumps({"error": str(e)}),
            status_code=500,
            mimetype="application/json",
        )
