# ZT Bot — Architecture & Setup Reference

A ready reference for building and deploying Teams bots under the ZweiThreads account.
ZT Scribe was the first bot built with this pattern.

---

## Architecture

### Production (Azure Functions)

```
┌─────────────────────────────────────────────────────────────────────┐
│  Microsoft Teams                                                    │
│                                                                     │
│  User posts "@ZT Scribe + photo" in a channel                      │
│        │                                                            │
│        ▼                                                            │
│  Teams Service  ◄──────────────── posts reply to thread ──────────┐│
└────────┼────────────────────────────────────────────────────────── ┘│
         │ POST /api/messages (activity JSON)                         │
         ▼                                                            │
┌─────────────────┐  validates cert auth  ┌───────────────────────┐  │
│ Azure Bot       │ ────────────────────► │ Azure Functions       │  │
│ Service         │                       │ (Flex Consumption)    │  │
│ (Bot Channels   │ ◄──────────────────── │ function_app.py       │──┘
│  Registration)  │   send_activity()     │ bot.py · pipeline.py  │
└────────┬────────┘                       └──────────┬────────────┘
         │                                           │
         │  token validation                         │ image bytes
         ▼                                           ▼
┌─────────────────┐                      ┌───────────────────────┐
│ Azure Entra ID  │                      │ Anthropic Claude      │
│ App Registration│                      │ Vision API            │
│ (App ID + Cert) │                      │ (claude-sonnet-5)     │
└─────────────────┘                      └───────────────────────┘
```

### Local development (ngrok)

Replace Azure Functions with:
```
ngrok tunnel (public HTTPS URL)  ←  update Azure Bot endpoint each session
      │
      ▼
localhost:3978  ←  python app.py
```

ngrok gives Teams a reachable URL during development. Every time ngrok restarts,
its URL changes — you must update the Azure Bot messaging endpoint each time.
This is why ngrok is only for development; Azure Functions runs permanently.

---

## Azure Components and Their Roles

| Component | What it is | Why you need it |
|---|---|---|
| **App Registration** (Entra ID) | An identity in your Azure AD tenant | Gives the bot an App ID and a way to authenticate to Azure Bot Service |
| **Certificate** (bot.cer + bot.key) | Public cert uploaded to Entra; private key stays on your server | Authenticates your code to Azure — used because ZweiThreads tenant policy blocks client secrets |
| **Azure Bot Service** | A managed routing layer | Receives messages from Teams, validates the bot's identity, forwards activity JSON to your endpoint, sends replies back to Teams |
| **Teams App Package** (ztscribe.zip) | manifest.json + color.png + outline.png | Tells Teams the bot's name, App ID, and scopes. Uploaded once via Teams → Apps → Upload a custom app |
| **Azure Functions (Flex Consumption)** | Serverless Python host | Runs your code on demand — no VM quota needed, scales to zero when idle, no always-on cost |
| **Storage Account** | Azure blob storage | Required by Azure Functions to store function state — created automatically in the portal wizard |

### The one GUID that ties everything together

The **App Registration App ID** (`425c4c34-45d2-407c-9796-8e55e37190a1` for ZT Scribe)
must match in **three places**:

1. `MicrosoftAppId` in your app's environment variables (Application Settings)
2. `botId` in `manifest.json` (Teams app package)
3. The "Microsoft App ID" field on the Azure Bot resource

If any one of these diverges, auth breaks silently.

---

## How Certificate Auth Works

```
Your code (function_app.py)
  reads private key content from MicrosoftCertPrivateKey env var
       │
       ▼
Azure Entra ID verifies the thumbprint matches bot.cer uploaded to App Registration
       │
       ▼
Issues an OAuth2 access token scoped to Bot Framework
       │
       ▼
Bot Framework validates the token on every incoming message
```

The `channel_auth_tenant` parameter in `CertificateAppCredentials` must be set to
your **tenant ID** — not the default `botframework.com`. Without it you get
`AADSTS700016: application not found in directory d6d49420` — that hex ID is
Microsoft's Bot Framework tenant, which is where the SDK looks by default.

---

## Project Files

```
ZT Scribe/
├── function_app.py      Azure Functions entry point (production)
├── app.py               aiohttp server entry point (local development only)
├── bot.py               ScribeBot(ActivityHandler) — Teams turn event handler
├── pipeline.py          Core logic — extract(image_bytes) → render(board)
├── host.json            Required Azure Functions runtime configuration
├── scribe.py            CLI tool (original PoC, still works standalone)
├── gen_icons.py         Generates manifest/color.png and manifest/outline.png (Pillow)
├── gen_cert.py          Generates bot.key + bot.cer (no OpenSSL — uses cryptography pkg)
├── build_package.py     Builds ztscribe.zip for Teams sideloading
├── requirements.txt     Python dependencies (includes azure-functions)
├── .env                 Secrets — NEVER commit (in .gitignore)
├── .env.example         Documents all env vars without values
├── bot.key              Private key — NEVER commit (in .gitignore)
├── bot.cer              Public cert — safe to keep locally; upload to Azure once
├── manifest/
│   ├── manifest.json    Teams app manifest
│   ├── color.png        192×192 bot icon (ZT two-swoosh robot face)
│   └── outline.png      32×32 white outline icon for Teams sidebar
└── ztscribe.zip         Teams app package — re-run build_package.py to regenerate
```

### Environment variables

| Variable | Local (.env) | Azure (Application Settings) | Purpose |
|---|---|---|---|
| `ANTHROPIC_API_KEY` | ✓ | ✓ | Claude Vision API key |
| `MicrosoftAppId` | ✓ | ✓ | App Registration App ID |
| `MicrosoftAppTenantId` | ✓ | ✓ | Your Azure tenant ID — required for cert auth |
| `MicrosoftCertThumbprint` | ✓ | ✓ | SHA-1 thumbprint from Azure cert upload |
| `MicrosoftCertKeyFile` | ✓ | — | Path to bot.key file (local only; file doesn't exist on Azure) |
| `MicrosoftCertPrivateKey` | — | ✓ | Full PEM content of bot.key (Azure only; paste in portal) |
| `PORT` | auto | auto | App listens on this port; set automatically by Azure |

---

## Setting Up a New ZT Bot (Steps)

### 1. Azure App Registration

1. Azure portal → Entra ID → App registrations → New registration
2. Name: `zt_<botname>`, Supported account type: **Single tenant**
3. No redirect URI needed
4. Note the **Application (client) ID** — this is your Bot App ID
5. Note the **Directory (tenant) ID** — this is `MicrosoftAppTenantId`

### 2. Generate and upload a certificate

```bash
python gen_cert.py
```

Produces `bot.key` (keep secret, never commit) and `bot.cer` (upload to Azure).

In the App Registration → Certificates & secrets → **Certificates** tab → Upload `bot.cer`.
Copy the **thumbprint** shown after upload.

> ZweiThreads tenant policy blocks client secrets — always use certificate auth.

### 3. Create the Azure Bot resource

1. Azure portal → Create a resource → **Azure Bot**
2. Bot handle: `zt-<botname>`
3. Microsoft App ID: paste your App Registration App ID
4. Select **Use existing app registration**
5. Messaging endpoint: leave blank for now (fill after deploy)
6. After creation → Channels → Add **Microsoft Teams** channel

### 4. Write the bot code

Copy `function_app.py`, `bot.py`, `pipeline.py`, `host.json` from ZT Scribe.
- `pipeline.py`: your domain logic (Claude prompt + rendering)
- `bot.py`: `ScribeBot(ActivityHandler)` with `on_message_activity`
- `function_app.py` and `app.py`: infrastructure wrappers — unchanged between bots

Key behavioural logic lives entirely in `bot.py` and `pipeline.py`.

### 5. Build the Teams app package

Update `manifest/manifest.json` with the new bot's name and App ID, then:

```bash
python gen_icons.py       # generates manifest/color.png and outline.png
python build_package.py   # produces ztscribe.zip (rename as needed)
```

Upload the zip in Teams → Apps → Manage your apps → Upload a custom app.

### 6. Test locally with ngrok

```bash
# Terminal 1
ngrok http 3978

# Terminal 2
python app.py
```

Update Azure Bot → Configuration → Messaging endpoint to your current ngrok URL:
`https://<your-subdomain>.ngrok-free.dev/api/messages`

---

## Deploying to Azure Functions (Flex Consumption)

**Why Flex Consumption:** Serverless — no VM quota required, scales to zero when idle
(no idle cost), pays only for execution time. No GitHub required; code is uploaded
directly via zip file.

**Two entry points:**
- `app.py` → local development only (aiohttp web server, `python app.py`)
- `function_app.py` → Azure deployment (HTTP-triggered Azure Function)

Both share `bot.py` and `pipeline.py` — no logic duplication.

### Step 1 — Create the Function App in the Portal

Portal → Create a resource → **Function App** → **Flex Consumption**

Fill in:
- Subscription / Resource group: your existing group
- Function App name: `ztscribe-fn` (must be globally unique)
- Runtime: **Python 3.11**
- Region: **East US** (or any available region)

**Deployment tab:** leave **Disable** selected — GitHub is for continuous deployment
(auto-deploy on every push) and is entirely optional. Skip it; code goes in via zip.

**Storage tab:** the wizard creates a storage account automatically — no separate step needed.

→ **Review + create**

### Step 2 — Set environment variables

**Via Cloud Shell** (portal.azure.com → `>_` icon) for most settings:

```bash
az functionapp config appsettings set \
  --name ztscribe-fn \
  --resource-group rg-zt_tools \
  --settings \
    ANTHROPIC_API_KEY="sk-ant-api03-..." \
    MicrosoftAppId="425c4c34-45d2-407c-9796-8e55e37190a1" \
    MicrosoftAppTenantId="6520c29a-d6e4-4465-ade7-8a119a49f3ab" \
    MicrosoftCertThumbprint="0F5C43432BAA08B8CF1226FA669AAF31D76F2D77"
```

**Via Portal** for the certificate key (multiline PEM is easier to paste in the UI):

Portal → Function App → **Configuration** → Application settings → **+ New application setting**
- Name: `MicrosoftCertPrivateKey`
- Value: paste the full contents of `bot.key` (including header/footer lines)
→ Save

### Step 3 — Create the deploy zip (on your laptop)

```powershell
$project = "C:\Users\zw_gasa0001\Documents\Projects\ZT Scribe"
Set-Location $project
Compress-Archive `
  -Path function_app.py, host.json, bot.py, pipeline.py, requirements.txt `
  -DestinationPath deploy_fn.zip -Force
```

The zip contains only the runtime code. Secrets stay in Application Settings.
The `.venv` folder, `bot.key`, `.env`, and dev tools are excluded.

### Step 4 — Upload and deploy the zip

**Via Cloud Shell:**

Upload `deploy_fn.zip` using the **Upload** button in Cloud Shell, then:

```bash
az functionapp deployment source config-zip \
  --name ztscribe-fn \
  --resource-group rg-zt_tools \
  --src deploy_fn.zip
```

**Via Portal (no Cloud Shell needed):**

Portal → Function App → **Deployment Center** → Source: **ZIP deploy** → upload `deploy_fn.zip`

### Step 5 — Update the Azure Bot messaging endpoint

Portal → Azure Bot resource → **Configuration** → Messaging endpoint:
```
https://ztscribe-fn.azurewebsites.net/api/messages
```
Save. Stop ngrok. The bot now runs 24/7 from Azure at no idle cost.

### Subsequent deploys

```powershell
# Laptop: re-zip changed files
Compress-Archive `
  -Path function_app.py, host.json, bot.py, pipeline.py, requirements.txt `
  -DestinationPath deploy_fn.zip -Force
```
Then upload and run the `config-zip` command again.

### Check logs

```bash
az functionapp log tail --name ztscribe-fn --resource-group rg-zt_tools
```

---

## Cold Starts

Flex Consumption scales to zero when idle — the first message after ~10 minutes of
inactivity triggers a cold start (~15 seconds for Python with heavy dependencies).
Teams may show a timeout error on that first message; subsequent messages are fast.

This is acceptable for a prototype. To eliminate cold starts in production, set
**Always Ready instances** to 1 in the Function App → Scale and concurrency settings
(incurs a small always-on cost).

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Bot replies with help text, no image processed | Attachment download failed | Check logs — usually auth token issue or ngrok dropped |
| `AADSTS700016: app not found in d6d49420` | `channel_auth_tenant` not set | Add `MicrosoftAppTenantId` to env — must be your tenant GUID |
| `SubscriptionNotFound` in Cloud Shell | Cloud Shell on wrong subscription | `az account list` then `az account set --subscription <id>` |
| App Service plan quota error (`Total VMs: 0`) | No VM quota in region | Use Azure Functions (Flex Consumption) — no VM quota needed |
| `app_credential` keyword error | Typo in BotFrameworkAdapterSettings | Use `app_credentials` (with s) |
| `packageName` validation failure in Teams | Teams manifest 1.17 rejects this field | Remove `packageName` from manifest.json |
| Bot doesn't respond at all | Messaging endpoint wrong or cold start | Verify endpoint in Azure Bot → Configuration; check Function App logs |
| Auth works locally but not on Azure | `bot.key` file not present on Azure | Set `MicrosoftCertPrivateKey` in Application Settings with PEM content |
| First message times out, rest work fine | Cold start (Flex Consumption scale to zero) | Expected behaviour — retry; or enable Always Ready instances |

---

## ZT Bot Icon System

All ZT bots use the same robot face, derived from the ZweiThreads two-swoosh motif:
- **Left half of face border**: ZT blue `#1E6DB7` (the Z stroke)
- **Right half of face border**: ZT red `#E63027` (the t stroke)
- Left antenna + left eye pupil: blue · Right antenna + right eye pupil: red

Same two-stroke DNA as the ZweiThreads circle, recast as a robot head.
The inner content can vary per bot; the face frame stays constant across all ZT bots.

```bash
python gen_icons.py    # regenerates manifest/color.png and manifest/outline.png
python build_package.py  # rebuilds ztscribe.zip with new icons
```

---

*ZweiThreads internal reference — last updated 2026-07-17*
