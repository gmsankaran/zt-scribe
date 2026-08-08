"""
ZT Scribe Teams bot.

In Teams channels file attachments never arrive in the bot activity.
The bot fetches the image from the message via Microsoft Graph API.
Requires ChannelMessage.Read.All application permission + admin consent.

Turn flow
---------
  Turn 1 — image message arrives:
    a) If multiple templates are configured: store image in memory, send template
       selection card.
    b) If only one template: go straight to extraction.

  Turn 2 (optional) — template card submitted (zt_action: "select_template"):
    Retrieve stored image, extract with chosen template, send draft + review card.

  Turn 3 (optional) — review card submitted (zt_action: "finalize_minutes"):
    Apply owner/text corrections deterministically (no LLM call), send corrected
    minutes.
    Owner free-text field (owner_text_N) takes precedence over the dropdown
    (owner_N) if both are filled.

  LLM is called only for extraction. Card responses are pure JSON surgery.
  The only server-side state is the image store (keyed by conversation ID);
  all other context is re-read from team_context.json at response time.
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
from botbuilder.schema import Activity, Attachment
from dotenv import load_dotenv

from pipeline import (
    build_clarification_card,
    build_template_card,
    extract,
    load_context,
    render,
    resolve_template,
)

load_dotenv()

_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}
_EXT_TO_MIME = {
    ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".png": "image/png", ".gif": "image/gif", ".webp": "image/webp",
}


class ScribeBot(ActivityHandler):

    # Raw image bytes stored between template-selection card and user response.
    # Keyed by conversation ID.  Lost on server restart (Render scale-to-zero);
    # handled gracefully with a "session expired" message.
    _image_store: dict[str, tuple[bytes, str]] = {}
    _MAX_STORE = 20   # prevent unbounded growth; evict oldest on overflow

    # ------------------------------------------------------------------
    # Turn 1 — image message
    # ------------------------------------------------------------------

    async def on_message_activity(self, turn_context: TurnContext):
        activity = turn_context.activity

        # Card submission (turns 2 / 3)
        value = getattr(activity, "value", None)
        if isinstance(value, dict) and "zt_action" in value:
            await self._handle_card_action(turn_context, value)
            return

        # Fetch image
        image_bytes, mime = await _image_from_activity(activity)
        if image_bytes is None and _is_channel(activity):
            await turn_context.send_activity(
                "One moment — fetching the image from the channel…"
            )
            image_bytes, mime = await _image_from_graph(activity)

        if image_bytes is None:
            await turn_context.send_activity(
                "Please attach a whiteboard photo to your message and @mention me."
            )
            return

        ctx = load_context()
        template_card = build_template_card(ctx)

        if template_card:
            # Multiple templates: ask first, extract after user picks
            conv_id = activity.conversation.id
            if len(self._image_store) >= self._MAX_STORE:
                oldest = next(iter(self._image_store))
                del self._image_store[oldest]
            self._image_store[conv_id] = (image_bytes, mime)
            await turn_context.send_activity(Activity(
                type="message",
                attachments=[Attachment(
                    content_type="application/vnd.microsoft.card.adaptive",
                    content=template_card,
                )]
            ))
        else:
            # Single template: extract immediately with default
            resolved_ctx = resolve_template(ctx)
            await self._run_pipeline(turn_context, image_bytes, mime, resolved_ctx)

    # ------------------------------------------------------------------
    # Card action dispatcher
    # ------------------------------------------------------------------

    async def _handle_card_action(self, turn_context: TurnContext, value: dict):
        action = value.get("zt_action")

        # ---- Template selection (Turn 2) --------------------------------
        if action == "select_template":
            conv_id  = turn_context.activity.conversation.id
            stored   = self._image_store.pop(conv_id, None)
            if stored is None:
                await turn_context.send_activity(
                    "The session expired — please resend the image and @mention me again."
                )
                return
            image_bytes, mime = stored
            template_key = value.get("template_key") or None
            ctx          = load_context()
            resolved_ctx = resolve_template(ctx, template_key)
            await self._run_pipeline(turn_context, image_bytes, mime, resolved_ctx)
            return

        # ---- Review card: skip ----------------------------------------
        if action == "skip_review":
            await turn_context.send_activity("✅ Draft minutes stand as sent.")
            return

        # ---- Review card: apply corrections (Turn 3) -------------------
        if action == "finalize_minutes":
            try:
                b64   = value.get("board_b64", "")
                board = json.loads(base64.b64decode(b64).decode("utf-8"))
                items = board.get("items", [])

                # Collect all keys once so we can detect free-text overrides
                owner_text_keys = {
                    int(k[11:]): v.strip()
                    for k, v in value.items()
                    if k.startswith("owner_text_") and isinstance(v, str) and v.strip()
                    and k[11:].isdigit()
                }
                owner_keys = {
                    int(k[6:]): v
                    for k, v in value.items()
                    if k.startswith("owner_") and not k.startswith("owner_text_")
                    and isinstance(v, str) and v and k[6:].isdigit()
                }

                for idx in set(owner_text_keys) | set(owner_keys):
                    if 0 <= idx < len(items):
                        # Free text wins over dropdown
                        new_owner = owner_text_keys.get(idx) or owner_keys.get(idx)
                        if new_owner:
                            items[idx]["owners"]       = [new_owner]
                            items[idx]["owner_source"] = "confirmed"

                for k, v in value.items():
                    if k.startswith("text_") and isinstance(v, str) and v and k[5:].isdigit():
                        idx = int(k[5:])
                        if 0 <= idx < len(items):
                            items[idx]["text"]       = v
                            items[idx]["confidence"] = 1.0

                board["items"] = items
                ctx     = load_context()
                minutes = render(board, ctx)
                await turn_context.send_activity("✅ **Corrected minutes:**\n\n" + minutes)

            except Exception as exc:
                print(f"[card finalize error] {exc}", file=sys.stderr)
                await turn_context.send_activity(f"Error applying corrections: {exc}")
            return

        print(f"[bot] unknown zt_action: {action}", file=sys.stderr)

    # ------------------------------------------------------------------
    # Extraction + draft send
    # ------------------------------------------------------------------

    async def _run_pipeline(
        self,
        turn_context: TurnContext,
        image_bytes: bytes,
        mime: str,
        resolved_ctx: dict,
    ):
        await turn_context.send_activity("Reading the board, give me a moment…")
        try:
            board = extract(image_bytes, mime, resolved_ctx)
            draft = render(board, resolved_ctx)
            await turn_context.send_activity(draft)

            card = build_clarification_card(board, resolved_ctx)
            if card:
                await turn_context.send_activity(Activity(
                    type="message",
                    attachments=[Attachment(
                        content_type="application/vnd.microsoft.card.adaptive",
                        content=card,
                    )]
                ))
        except Exception as exc:
            print(f"[pipeline error] {exc}", file=sys.stderr)
            await turn_context.send_activity(f"Something went wrong: {exc}")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

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
            # smba.trafficmanager.net / botframework.com URLs require a
            # connector-service token we don't have here — Graph handles it.
            if "smba.trafficmanager.net" in url or "botframework.com" in url:
                print("[att] skipping bot-service URL — Graph will handle", file=sys.stderr)
                continue
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
        print("[Graph] no token — check ChannelMessage.Read.All consent", file=sys.stderr)
        return None, "image/jpeg"

    headers = {"Authorization": f"Bearer {token}"}
    cd = activity.channel_data or {}
    team_id    = (cd.get("team") or {}).get("aadGroupId")
    channel_id = cd.get("teamsChannelId") or (cd.get("channel") or {}).get("id")
    message_id = activity.id

    if not all([team_id, channel_id, message_id]):
        print(f"[Graph] missing IDs team={team_id} channel={channel_id} msg={message_id}",
              file=sys.stderr)
        return None, "image/jpeg"

    channel_id_enc = quote(channel_id, safe="")
    base = (f"https://graph.microsoft.com/v1.0"
            f"/teams/{team_id}/channels/{channel_id_enc}/messages/{message_id}")

    async with aiohttp.ClientSession() as session:
        async with session.get(base, headers=headers) as resp:
            body = await resp.text()
            if resp.status != 200:
                print(f"[Graph] GET message → {resp.status}: {body[:400]}", file=sys.stderr)
                return None, "image/jpeg"
            msg = json.loads(body)

    atts = msg.get("attachments") or []
    print(f"[Graph] message OK, attachments={len(atts)}", file=sys.stderr)

    # File attachments (contentType: "reference") — download via /shares
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
        enc  = base64.b64encode(content_url.encode()).decode()
        enc  = enc.replace("+", "-").replace("/", "_").rstrip("=")
        dl   = f"https://graph.microsoft.com/v1.0/shares/u!{enc}/driveItem/content"
        print("[Graph] downloading file via /shares", file=sys.stderr)
        try:
            return await _get(dl, headers), mime
        except Exception as exc:
            print(f"[Graph] /shares download failed: {exc}", file=sys.stderr)

    # Inline / pasted images — hostedContents (must fetch /$value per item)
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
