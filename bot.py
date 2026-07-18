"""
ZT Scribe Teams bot.

In Teams channels file attachments never arrive in the bot activity.
The bot fetches the image from the message via Microsoft Graph API.
Requires ChannelMessage.Read.All application permission + admin consent.
"""

import base64
import json
import os
import re
import sys
from urllib.parse import quote
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

        # Try attachments in the activity first
        image_bytes, mime = await _image_from_activity(activity)

        # Channel messages never include file data — fetch via Graph API
        if image_bytes is None and _is_channel(activity):
            await turn_context.send_activity("One moment — fetching the image from the channel...")
            image_bytes, mime = await _image_from_graph(activity)

        if image_bytes is None:
            await turn_context.send_activity(
                "Please attach a whiteboard photo to your message and @mention me."
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


def _is_channel(activity) -> bool:
    conv = activity.conversation
    return getattr(conv, "conversation_type", None) == "channel"


async def _image_from_activity(activity) -> tuple[bytes | None, str]:
    """Extract image from any attachment that arrived directly in the activity."""
    for a in (activity.attachments or []):
        ct = a.content_type or ""

        if ct.startswith("image/"):
            url = a.content_url or ""
            if url.startswith("data:"):
                _, enc = url.split(",", 1)
                return base64.b64decode(enc), ct
            try:
                return await _get(url, {}), ct
            except Exception as exc:
                print(f"[att] inline image download failed: {exc}", file=sys.stderr)

        if ct == "application/vnd.microsoft.teams.file.download.info":
            name = (a.name or "").lower()
            if any(name.endswith(ext) for ext in _IMAGE_EXTS):
                mime = _EXT_TO_MIME.get(os.path.splitext(name)[1], "image/jpeg")
                content = a.content
                if isinstance(content, str):
                    content = json.loads(content)
                url = (content or {}).get("downloadUrl") or a.content_url or ""
                if url:
                    try:
                        return await _get(url, {}), mime
                    except Exception as exc:
                        print(f"[att] file download failed: {exc}", file=sys.stderr)

        if ct == "text/html":
            html = a.content if isinstance(a.content, str) else ""
            m = re.search(r'<img\b[^>]+\bsrc=["\']([^"\']+)["\']', html, re.IGNORECASE)
            if m:
                src = m.group(1)
                if src.startswith("data:image/"):
                    _, enc = src.split(",", 1)
                    return base64.b64decode(enc), src.split(";")[0][5:]

    return None, "image/jpeg"


async def _image_from_graph(activity) -> tuple[bytes | None, str]:
    """Fetch the image from a Teams channel message via Microsoft Graph."""
    token = await _graph_token()
    if not token:
        print("[Graph] no token — add ChannelMessage.Read.All + grant admin consent", file=sys.stderr)
        return None, "image/jpeg"

    headers = {"Authorization": f"Bearer {token}"}
    cd = activity.channel_data or {}
    # Graph API needs the AAD group ID for the team, not the Teams thread ID
    team_id    = (cd.get("team") or {}).get("aadGroupId")
    channel_id = cd.get("teamsChannelId") or (cd.get("channel") or {}).get("id")
    message_id = activity.id

    if not all([team_id, channel_id, message_id]):
        print(f"[Graph] missing IDs team={team_id} channel={channel_id} msg={message_id}", file=sys.stderr)
        return None, "image/jpeg"

    channel_id_enc = quote(channel_id, safe="")
    base = (f"https://graph.microsoft.com/v1.0"
            f"/teams/{team_id}/channels/{channel_id_enc}/messages/{message_id}")

    # GET the message (no $expand — unsupported on this endpoint per Graph docs)
    async with aiohttp.ClientSession() as session:
        async with session.get(base, headers=headers) as resp:
            body = await resp.text()
            if resp.status != 200:
                print(f"[Graph] GET message → {resp.status}: {body[:400]}", file=sys.stderr)
                return None, "image/jpeg"
            msg = json.loads(body)

    atts = msg.get("attachments") or []
    print(f"[Graph] message OK, attachments={len(atts)}", file=sys.stderr)

    # File attachments have contentType "reference" and a SharePoint contentUrl.
    # Must download via Graph /shares API (requires Files.Read.All permission).
    for att in atts:
        if att.get("contentType") != "reference":
            continue
        name = (att.get("name") or "").lower()
        if not any(name.endswith(ext) for ext in _IMAGE_EXTS):
            continue
        content_url = att.get("contentUrl") or ""
        if not content_url:
            continue
        mime = _EXT_TO_MIME.get(os.path.splitext(name)[1], "image/jpeg")
        enc = base64.b64encode(content_url.encode()).decode()
        enc = enc.replace("+", "-").replace("/", "_").rstrip("=")
        download_url = f"https://graph.microsoft.com/v1.0/shares/u!{enc}/driveItem/content"
        print(f"[Graph] downloading file via /shares", file=sys.stderr)
        try:
            return await _get(download_url, headers), mime
        except Exception as exc:
            print(f"[Graph] /shares download failed: {exc}", file=sys.stderr)

    # Inline/pasted images arrive as hostedContents.
    # List returns null bytes — must fetch /$value per item individually.
    async with aiohttp.ClientSession() as session:
        async with session.get(f"{base}/hostedContents", headers=headers) as resp:
            if resp.status == 200:
                hc_list = (await resp.json()).get("value") or []
                print(f"[Graph] hostedContents={len(hc_list)}", file=sys.stderr)
                for hc in hc_list:
                    hc_id = hc.get("id", "")
                    async with session.get(
                        f"{base}/hostedContents/{hc_id}/$value", headers=headers
                    ) as hc_resp:
                        if hc_resp.status == 200:
                            ct = hc_resp.headers.get("content-type", "").split(";")[0]
                            if ct.startswith("image/"):
                                return await hc_resp.read(), ct
            else:
                txt = await resp.text()
                print(f"[Graph] hostedContents → {resp.status}: {txt[:200]}", file=sys.stderr)

    print("[Graph] no image found in message", file=sys.stderr)
    return None, "image/jpeg"


async def _graph_token() -> str | None:
    app_id     = os.getenv("MicrosoftAppId", "")
    tenant_id  = os.getenv("MicrosoftAppTenantId", "")
    thumbprint = os.getenv("MicrosoftCertThumbprint", "")
    key_inline = os.getenv("MicrosoftCertPrivateKey", "")
    key_file   = os.getenv("MicrosoftCertKeyFile", "")

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
