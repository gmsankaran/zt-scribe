"""
ZT Scribe — aiohttp server entry point.

Runs on http://localhost:3978/api/messages
Point your Azure Bot registration (or Bot Framework Emulator) at that URL,
or tunnel it with: ngrok http 3978
"""

import os
import sys

from aiohttp import web
from aiohttp.web import Request, Response, json_response
from botbuilder.core import BotFrameworkAdapterSettings, BotFrameworkAdapter
from botbuilder.schema import Activity
from botframework.connector.auth import CertificateAppCredentials
from dotenv import load_dotenv

from bot import ScribeBot

load_dotenv()

_app_id = os.getenv("MicrosoftAppId", "")
_tenant_id = os.getenv("MicrosoftAppTenantId", "")
_cert_thumbprint = os.getenv("MicrosoftCertThumbprint", "")
_cert_key_file = os.getenv("MicrosoftCertKeyFile", "")
_cert_key_inline = os.getenv("MicrosoftCertPrivateKey", "")  # Azure: paste PEM here

def _normalize_pem(raw: str) -> str:
    """Normalize PEM key regardless of how env var storage mangled the newlines."""
    # Literal backslash-n sequences (some env systems JSON-encode newlines)
    if "\\n" in raw and "\n" not in raw:
        raw = raw.replace("\\n", "\n")
    # Windows-style line endings
    raw = raw.replace("\r\n", "\n").replace("\r", "\n").strip()
    # Completely flattened — no newlines at all; reconstruct 64-char lines
    if "\n" not in raw:
        for begin_tag in ("-----BEGIN PRIVATE KEY-----", "-----BEGIN RSA PRIVATE KEY-----"):
            end_tag = begin_tag.replace("BEGIN", "END")
            if begin_tag in raw:
                body = raw.replace(begin_tag, "").replace(end_tag, "").strip()
                lines = [body[i : i + 64] for i in range(0, len(body), 64)]
                raw = begin_tag + "\n" + "\n".join(lines) + "\n" + end_tag
                break
    # Missing header/footer (pasted without the -----BEGIN/END----- lines)
    if "-----BEGIN" not in raw:
        raw = "-----BEGIN PRIVATE KEY-----\n" + raw
    if "-----END" not in raw:
        raw = raw.rstrip("\n") + "\n-----END PRIVATE KEY-----"
    return raw if raw.endswith("\n") else raw + "\n"


if _cert_thumbprint and (_cert_key_file or _cert_key_inline):
    if _cert_key_inline:
        _private_key = _normalize_pem(_cert_key_inline)
    else:
        _private_key = open(_cert_key_file).read()
    first_line = _private_key.split("\n")[0]
    print(f"[CERT] key loaded — first line: {repr(first_line)}, length: {len(_private_key)}, newlines: {_private_key.count(chr(10))}", file=sys.stderr)
    _app_credential = CertificateAppCredentials(
        app_id=_app_id,
        certificate_thumbprint=_cert_thumbprint,
        certificate_private_key=_private_key,
        channel_auth_tenant=_tenant_id or None,
    )
    SETTINGS = BotFrameworkAdapterSettings(app_id=_app_id, app_credentials=_app_credential)
else:
    # Falls back to client secret (for emulator or tenants without the policy)
    SETTINGS = BotFrameworkAdapterSettings(
        app_id=_app_id,
        app_password=os.getenv("MicrosoftAppPassword", ""),
    )
ADAPTER = BotFrameworkAdapter(SETTINGS)


async def _on_error(context, error: Exception):
    print(f"[on_turn_error] {error}", file=sys.stderr)
    await context.send_activity("The bot hit an unexpected error. Try again.")


ADAPTER.on_turn_error = _on_error

BOT = ScribeBot()


async def messages(req: Request) -> Response:
    if req.content_type != "application/json":
        return Response(status=415)
    body = await req.json()
    import json as _json
    print(f"[RAW] {_json.dumps(body)[:4000]}", file=sys.stderr)
    activity = Activity().deserialize(body)
    auth_header = req.headers.get("Authorization", "")
    invoke_response = await ADAPTER.process_activity(activity, auth_header, BOT.on_turn)
    if invoke_response:
        return json_response(data=invoke_response.body, status=invoke_response.status)
    return Response(status=201)


async def health(req: Request) -> Response:
    return Response(status=200, text="ok")


app = web.Application()
app.router.add_post("/api/messages", messages)
app.router.add_get("/health", health)

if __name__ == "__main__":
    port = int(os.getenv("PORT", 3978))
    host = "0.0.0.0" if os.getenv("PORT") else "localhost"
    print(f"ZT Scribe bot listening on http://{host}:{port}/api/messages")
    web.run_app(app, host=host, port=port)
