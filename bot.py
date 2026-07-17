"""
ZT Scribe Teams bot.

Receives an @mention + image attachment, calls the extract/render pipeline,
and posts the minutes back to the conversation.
"""

import base64
import os
import aiohttp
from botbuilder.core import ActivityHandler, TurnContext
from dotenv import load_dotenv

from pipeline import extract, render

load_dotenv()


class ScribeBot(ActivityHandler):
    async def on_message_activity(self, turn_context: TurnContext):
        attachments = turn_context.activity.attachments or []
        image_att = next(
            (a for a in attachments if a.content_type and a.content_type.startswith("image/")),
            None,
        )

        if not image_att:
            await turn_context.send_activity(
                "Attach a whiteboard photo and I'll turn it into meeting minutes."
            )
            return

        await turn_context.send_activity("Reading the board, give me a moment...")

        try:
            image_bytes = await _download_image(image_att)
            board = extract(image_bytes, image_att.content_type or "image/jpeg")
            minutes = render(board)
            await turn_context.send_activity(minutes)
        except Exception as exc:
            await turn_context.send_activity(f"Something went wrong: {exc}")


async def _download_image(attachment) -> bytes:
    url = attachment.content_url

    if url.startswith("data:"):
        # Inline base64 (Bot Framework Emulator sometimes sends this)
        _, encoded = url.split(",", 1)
        return base64.b64decode(encoded)

    headers = await _bot_service_headers()
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            if resp.status == 401 and headers:
                # Header didn't help — try without (shouldn't normally happen)
                async with session.get(url) as resp2:
                    resp2.raise_for_status()
                    return await resp2.read()
            resp.raise_for_status()
            return await resp.read()


async def _bot_service_headers() -> dict:
    """Exchange bot credentials for a Bot Framework service bearer token.

    Required to download attachment URLs from https://smba.trafficmanager.net/...
    in a real Teams deployment. Returns an empty dict when running against the
    Bot Framework Emulator (no credentials set).
    """
    app_id = os.getenv("MicrosoftAppId", "")
    app_password = os.getenv("MicrosoftAppPassword", "")
    if not (app_id and app_password):
        return {}

    async with aiohttp.ClientSession() as session:
        data = {
            "grant_type": "client_credentials",
            "client_id": app_id,
            "client_secret": app_password,
            "scope": "https://api.botframework.com/.default",
        }
        async with session.post(
            "https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token",
            data=data,
        ) as resp:
            if resp.status == 200:
                token_data = await resp.json()
                return {"Authorization": f"Bearer {token_data['access_token']}"}

    return {}
