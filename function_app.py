"""
ZT Scribe — Azure Functions entry point (production deployment).

For local development, use app.py (aiohttp) instead.
"""
import json
import os
import sys

import azure.functions as func
from botbuilder.core import BotFrameworkAdapterSettings, BotFrameworkAdapter
from botbuilder.schema import Activity
from botframework.connector.auth import CertificateAppCredentials
from dotenv import load_dotenv

from bot import ScribeBot

load_dotenv()

_app_id          = os.getenv("MicrosoftAppId", "")
_tenant_id       = os.getenv("MicrosoftAppTenantId", "")
_cert_thumbprint = os.getenv("MicrosoftCertThumbprint", "")
_cert_key_file   = os.getenv("MicrosoftCertKeyFile", "")
_cert_key_inline = os.getenv("MicrosoftCertPrivateKey", "")

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

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)


@app.route(route="api/messages", methods=["POST"])
async def messages(req: func.HttpRequest) -> func.HttpResponse:
    if "application/json" not in req.headers.get("content-type", "").lower():
        return func.HttpResponse(status_code=415)

    body = req.get_json()
    activity = Activity().deserialize(body)
    auth_header = req.headers.get("Authorization", "")

    invoke_response = await ADAPTER.process_activity(activity, auth_header, BOT.on_turn)

    if invoke_response:
        return func.HttpResponse(
            body=json.dumps(invoke_response.body),
            status_code=invoke_response.status,
            mimetype="application/json",
        )
    return func.HttpResponse(status_code=201)
