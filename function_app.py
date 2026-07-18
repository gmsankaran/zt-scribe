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

def _normalize_pem(raw: str) -> str:
    if "\\n" in raw and "\n" not in raw:
        raw = raw.replace("\\n", "\n")
    raw = raw.replace("\r\n", "\n").replace("\r", "\n").strip()
    if "\n" not in raw:
        for begin_tag in ("-----BEGIN PRIVATE KEY-----", "-----BEGIN RSA PRIVATE KEY-----"):
            end_tag = begin_tag.replace("BEGIN", "END")
            if begin_tag in raw:
                body = raw.replace(begin_tag, "").replace(end_tag, "").strip()
                lines = [body[i : i + 64] for i in range(0, len(body), 64)]
                raw = begin_tag + "\n" + "\n".join(lines) + "\n" + end_tag
                break
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
