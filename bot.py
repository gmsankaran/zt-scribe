"""
ZT Scribe Teams bot.

Receives an @mention + image, calls the extract/render pipeline,
and posts the minutes back to the conversation.

In Teams channels, file attachments are not delivered in the bot activity.
The bot fetches the image from the message via Microsoft Graph API instead.
Requires ChannelMessage.Read.All application permission on the App Registration.
"""

import base64
import json
import os
import re
import sys
import aiohttp
import msal
from botbuilder.core import ActivityHandler, TurnContext
from dotenv import load_dotenv

from pipeline import extract, render

load_dotenv()

_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}
_EXT_TO_MIME = {
    ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".png": "image/png", ".gif": "image/gif", ".webp": "image/webp",
}


class ScribeBot(ActivityHandler):
    async def on_message_activity(self, turn_context: TurnContext):
        activity = turn_context.activity

        # First: check if the activity itself carries an image attachment
        image_bytes, mime = _image_from_activity(activity)

        # Second: if not, try Graph API (channel messages never include file data)
        if image_bytes is None and _is_channel(activity):
            await turn_context.send_activity("One moment — fetching the image from the channel...")
            image_bytes, mime = await _image_from_graph(activity)

        if image_bytes is None:
            await turn_context.send_activity(
                "Please attach a whiteboard photo to your message and @mention me. "
                "Use the 📎 attach button so the file is part of the same message."
            )
            return

        await turn_context.send_activity("Reading the board, give me a moment...")
        try:
            board = extract(image_bytes, mime)
            minutes = render(board)
            await turn_context.send_activity(minutes)
        except Exception as exc:
            print(f"[pipeline error] {exc}", file=sys.stderr)
            await turn_context.send_activity(f"Something went wrong: {exc}")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _is_channel(activity) -> bool:
    return (activity.conversation or {}).get("conversationType") == "channel" or \
           getattr(activity.conversation, "conversation_type", None) == "channel"


def _image_from_activity(activity) -> tuple[bytes | None, str]:
    """Extract image bytes from an attachment that arrived in the activity."""
    for a in (activity.attachments or []):
        ct = a.content_type or ""
        if ct.startswith("image/"):
            return _download_inline(a), ct

        if ct == "application/vnd.microsoft.teams.file.download.info":
            name = (a.name or "").lower()
            if any(name.endswith(ext) for ext in _IMAGE_EXTS):
                mime = _EXT_TO_MIME.get(os.path.splitext(name)[1], "image/jpeg")
                content = a.content
                if isinstance(content, str):
                    content = json.loads(content)
                url = (content or {}).get("downloadUrl") or a.content_url
                if url:
                    import asyncio
                    data = asyncio.get_event_loop().run_until_complete(_get(url, {}))
                    return data, mime

        if ct == "text/html":
            html = a.content if isinstance(a.content, str) else ""
            m = re.search(r'<img\b[^>]+\bsrc=["\']([^"\']+)["\']', html, re.IGNORECASE)
            if m:
                src = m.group(1)
                if src.startswith("data:image/"):
                    _, enc = src.split(",", 1)
                    detected_mime = src.split(";")[0][5:]
                    return base64.b64decode(enc), detected_mime

    return None, "image/jpeg"


def _download_inline(attachment) -> bytes:
    url = attachment.content_url or ""
    if url.startswith("data:"):
        _, enc = url.split(",", 1)
        return base64.b64decode(enc)
    # Synchronous helper — only called for image/* attachments which are rare
    import asyncio
    return asyncio.get_event_loop().run_until_complete(_get(url, {}))


async def _image_from_graph(activity) -> tuple[bytes | None, str]:
    """Fetch the image attached to a Teams channel message via Graph API.

    Requires ChannelMessage.Read.All application permission + admin consent.
    """
    token = await _graph_token()
    if not token:
        print("[Graph] no token — check ChannelMessage.Read.All permission + admin consent", file=sys.stderr)
        return None, "image/jpeg"

    headers = {"Authorization": f"Bearer {token}"}
    cd = activity.channel_data or {}
    team_id    = cd.get("teamsTeamId") or (cd.get("team") or {}).get("id")
    channel_id = cd.get("teamsChannelId") or (cd.get("channel") or {}).get("id")
    message_id = activity.id

    if not all([team_id, channel_id, message_id]):
        print(f"[Graph] missing IDs team={team_id} channel={channel_id} msg={message_id}", file=sys.stderr)
        return None, "image/jpeg"

    url = (f"https://graph.microsoft.com/v1.0"
           f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}"
           f"?$expand=hostedContents")

    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            if resp.status != 200:
                body = await resp.text()
                print(f"[Graph] GET message → {resp.status}: {body[:300]}", file=sys.stderr)
                return None, "image/jpeg"
            msg = await resp.json()

    print(f"[Graph] attachments={len(msg.get('attachments', []))} "
          f"hostedContents={len(msg.get('hostedContents', []))}", file=sys.stderr)

    # Inline/pasted images arrive as hostedContents
    for hc in msg.get("hostedContents") or []:
        ct = hc.get("contentType", "")
        if ct.startswith("image/"):
            raw = hc.get("contentBytes", "")
            if raw:
                return base64.b64decode(raw), ct

    # File attachments uploaded via SharePoint
    for att in msg.get("attachments") or []:
        name = (att.get("name") or "").lower()
        content_url = att.get("contentUrl") or ""
        if content_url and any(name.endswith(ext) for ext in _IMAGE_EXTS):
            mime = _EXT_TO_MIME.get(os.path.splitext(name)[1], "image/jpeg")
            try:
                data = await _get(content_url, headers)
                return data, mime
            except Exception as exc:
                print(f"[Graph] file download failed: {exc}", file=sys.stderr)

    print("[Graph] no image found in message", file=sys.stderr)
    return None, "image/jpeg"


async def _graph_token() -> str | None:
    app_id        = os.getenv("MicrosoftAppId", "")
    tenant_id     = os.getenv("MicrosoftAppTenantId", "")
    thumbprint    = os.getenv("MicrosoftCertThumbprint", "")
    key_inline    = os.getenv("MicrosoftCertPrivateKey", "")
    key_file      = os.getenv("MicrosoftCertKeyFile", "")

    if not (app_id and tenant_id and thumbprint and (key_inline or key_file)):
        return None

    private_key = _load_key(key_inline, key_file)
    app = msal.ConfidentialClientApplication(
        client_id=app_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential={"private_key": private_key, "thumbprint": thumbprint},
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    print(f"[Graph] token error: {result.get('error_description')}", file=sys.stderr)
    return None


async def _get(url: str, headers: dict) -> bytes:
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            resp.raise_for_status()
            return await resp.read()


def _load_key(inline: str, key_file: str) -> str:
    raw = inline if inline else open(key_file).read()
    if "\\n" in raw and "\n" not in raw:
        raw = raw.replace("\\n", "\n")
    raw = raw.replace("\r\n", "\n").replace("\r", "\n").strip()
    if "-----BEGIN" not in raw:
        raw = "-----BEGIN PRIVATE KEY-----\n" + raw
    if "-----END" not in raw:
        raw = raw.rstrip("\n") + "\n-----END PRIVATE KEY-----"
    return raw if raw.endswith("\n") else raw + "\n"
