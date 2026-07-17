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

if _cert_thumbprint and (_cert_key_file or _cert_key_inline):
    _private_key = (
        _cert_key_inline.replace("\\n", "\n")
        if _cert_key_inline
        else open(_cert_key_file).read()
    )
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
