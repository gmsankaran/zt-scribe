"""
ZT Scribe Teams bot.

Receives an @mention + image attachment, calls the extract/render pipeline,
and posts the minutes back to the conversation.
"""

import base64
import json
import os
import re
import sys
import aiohttp
from botbuilder.core import ActivityHandler, TurnContext
from dotenv import load_dotenv

from pipeline import extract, render

load_dotenv()

_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}

# Anthropic vision only accepts these media types
_EXT_TO_MIME = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".gif": "image/gif",
    ".webp": "image/webp",
}


def _is_image_attachment(a) -> bool:
    if a.content_type and a.content_type.startswith("image/"):
        return True
    if a.content_type == "application/vnd.microsoft.teams.file.download.info":
        return any((a.name or "").lower().endswith(ext) for ext in _IMAGE_EXTS)
    if a.content_type == "text/html":
        # Image embedded in message body — check the HTML has an <img> tag
        html = a.content if isinstance(a.content, str) else ""
        return bool(re.search(r"<img\b", html, re.IGNORECASE))
    return False


def _mime_type(attachment) -> str:
    if attachment.content_type and attachment.content_type.startswith("image/"):
        return attachment.content_type
    if attachment.content_type == "text/html":
        html = attachment.content if isinstance(attachment.content, str) else ""
        m = re.search(r'<img\b[^>]+\bsrc=["\']([^"\']+)["\']', html, re.IGNORECASE)
        if m:
            src = m.group(1)
            if src.startswith("data:image/"):
                return src.split(";")[0][5:]  # "data:image/jpeg;..." → "image/jpeg"
            ext = os.path.splitext(src.split("?")[0].lower())[1]
            return _EXT_TO_MIME.get(ext, "image/jpeg")
    ext = os.path.splitext((attachment.name or "").lower())[1]
    return _EXT_TO_MIME.get(ext, "image/jpeg")


class ScribeBot(ActivityHandler):
    async def on_message_activity(self, turn_context: TurnContext):
        attachments = turn_context.activity.attachments or []
        print(f"[MSG] text={turn_context.activity.text!r:.80} attachments={len(attachments)}", file=sys.stderr)
        for i, a in enumerate(attachments):
            content_preview = str(a.content or "")[:150] if a.content_type == "text/html" else ""
        print(f"[ATT {i}] type={a.content_type!r} name={a.name!r} url={str(a.content_url or '')[:60]!r} html={content_preview!r}", file=sys.stderr)

        image_att = next((_a for _a in attachments if _is_image_attachment(_a)), None)

        if not image_att:
            await turn_context.send_activity(
                "Attach a whiteboard photo and I'll turn it into meeting minutes."
            )
            return

        await turn_context.send_activity("Reading the board, give me a moment...")

        try:
            image_bytes = await _download_image(image_att)
            board = extract(image_bytes, _mime_type(image_att))
            minutes = render(board)
            await turn_context.send_activity(minutes)
        except Exception as exc:
            print(f"[pipeline error] {exc}", file=sys.stderr)
            await turn_context.send_activity(f"Something went wrong: {exc}")


async def _download_image(attachment) -> bytes:
    # Teams file upload — pre-signed downloadUrl, no auth needed
    if attachment.content_type == "application/vnd.microsoft.teams.file.download.info":
        content = attachment.content
        if isinstance(content, str):
            content = json.loads(content)
        url = (content or {}).get("downloadUrl") or attachment.content_url
        async with aiohttp.ClientSession() as session:
            async with session.get(url) as resp:
                resp.raise_for_status()
                return await resp.read()

    # Image embedded in Teams message HTML body
    if attachment.content_type == "text/html":
        html = attachment.content if isinstance(attachment.content, str) else ""
        match = re.search(r'<img\b[^>]+\bsrc=["\']([^"\']+)["\']', html, re.IGNORECASE)
        if not match:
            raise ValueError("text/html attachment has no <img src>")
        url = match.group(1)
        print(f"[HTML-IMG] src={url[:80]!r}", file=sys.stderr)
        if url.startswith("data:"):
            _, encoded = url.split(",", 1)
            return base64.b64decode(encoded)
        headers = await _bot_service_headers()
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers) as resp:
                resp.raise_for_status()
                return await resp.read()

    url = attachment.content_url

    if url.startswith("data:"):
        _, encoded = url.split(",", 1)
        return base64.b64decode(encoded)

    # Inline image from Bot Framework attachment service — needs a bearer token
    headers = await _bot_service_headers()
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            resp.raise_for_status()
            return await resp.read()


async def _bot_service_headers() -> dict:
    """Get a Bearer token for the Bot Framework attachment service.

    Uses certificate auth when available (production), falls back to client
    secret (emulator / local dev).
    """
    app_id = os.getenv("MicrosoftAppId", "")
    if not app_id:
        return {}

    cert_thumbprint = os.getenv("MicrosoftCertThumbprint", "")
    tenant_id = os.getenv("MicrosoftAppTenantId", "")
    cert_key_inline = os.getenv("MicrosoftCertPrivateKey", "")
    cert_key_file = os.getenv("MicrosoftCertKeyFile", "")

    if cert_thumbprint and (cert_key_inline or cert_key_file):
        try:
            import msal
            private_key = _load_key(cert_key_inline, cert_key_file)
            msal_app = msal.ConfidentialClientApplication(
                client_id=app_id,
                authority=f"https://login.microsoftonline.com/{tenant_id}",
                client_credential={"private_key": private_key, "thumbprint": cert_thumbprint},
            )
            result = msal_app.acquire_token_for_client(
                scopes=["https://api.botframework.com/.default"]
            )
            if "access_token" in result:
                return {"Authorization": f"Bearer {result['access_token']}"}
            print(f"[warn] cert token failed: {result.get('error_description')}", file=sys.stderr)
        except Exception as exc:
            print(f"[warn] cert token exception: {exc}", file=sys.stderr)

    app_password = os.getenv("MicrosoftAppPassword", "")
    if app_password:
        async with aiohttp.ClientSession() as session:
            async with session.post(
                "https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token",
                data={
                    "grant_type": "client_credentials",
                    "client_id": app_id,
                    "client_secret": app_password,
                    "scope": "https://api.botframework.com/.default",
                },
            ) as resp:
                if resp.status == 200:
                    token_data = await resp.json()
                    return {"Authorization": f"Bearer {token_data['access_token']}"}

    return {}


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
