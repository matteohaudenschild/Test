# Exchange Mail Bridge

Cloud bridge for moving incoming Exchange/OWA mail metadata into a Google Sheet without keeping a local PC online.

Flow:

```text
Exchange EWS -> GitHub Actions schedule -> Python poller -> Apps Script Web App -> Google Sheet
```

## Why this path

The mailbox is reachable through classic OWA/EWS at `mail.wackerhagengruppe.de`, but IMAP/POP are not open and Microsoft Graph app registration / Power Automate are blocked for this account. Google Apps Script also cannot directly authenticate to EWS with NTLM. A small cloud poller fills that gap.

## Google setup

1. Create a Google Sheet and copy the Sheet ID from the URL.
2. Open the Apps Script project:
   `https://script.google.com/d/1kB5PLWbCJRpj2F5BGqbVC2armm5fUbAkVzXAyZO3VG5EkIhAU1iZtwU5/edit`
3. Paste `apps_script/Code.gs` into the project.
4. Set script properties:
   - `SHEET_ID`: your Google Sheet ID
   - `BRIDGE_TOKEN`: a long random secret
   - `SHEET_NAME`: `Exchange Mail`
5. Deploy the project:
   - Deploy > New deployment > Web app
   - Execute as: Me
   - Who has access: Anyone with the link / Anyone
   - Copy the `/exec` web app URL.

The web app is token-protected. Still, keep the URL and token private.

## GitHub setup

Put this folder in a GitHub repository so the workflow file is available at:

```text
.github/workflows/exchange-mail-bridge.yml
```

Add repository secrets:

```text
EWS_URL=https://mail.wackerhagengruppe.de/EWS/Exchange.asmx
EWS_EMAIL=matteo.merkle@sicherheit-nord.de
EWS_USERNAME=matteo.merkle@sicherheit-nord.de
EWS_PASSWORD=<mailbox password>
EWS_AUTH_TYPE=NTLM
EWS_VERIFY_TLS=true
APPS_SCRIPT_WEBAPP_URL=<Apps Script /exec URL>
BRIDGE_TOKEN=<same token as Apps Script>
```

Run the workflow once manually from GitHub Actions. After that it runs roughly every 10 minutes. GitHub scheduled workflows can be delayed, so this is near-real-time, not instant.

## Local smoke test

Create a local `.env` from `.env.example`, then run:

```powershell
python -m pip install -r .\exchange_mail_bridge\requirements.txt
$env:EWS_URL="https://mail.wackerhagengruppe.de/EWS/Exchange.asmx"
$env:EWS_EMAIL="matteo.merkle@sicherheit-nord.de"
$env:EWS_USERNAME="matteo.merkle@sicherheit-nord.de"
$env:EWS_PASSWORD="..."
$env:EWS_AUTH_TYPE="NTLM"
$env:MAIL_LOOKBACK_MINUTES="60"
python .\exchange_mail_bridge\exchange_to_apps_script.py --dry-run
```

For the real run, also set:

```powershell
$env:APPS_SCRIPT_WEBAPP_URL="https://script.google.com/macros/s/.../exec"
$env:BRIDGE_TOKEN="..."
python .\exchange_mail_bridge\exchange_to_apps_script.py
```

## Limits

- If Exchange blocks GitHub Actions IPs or denies NTLM from outside, use another small always-on host instead.
- The workflow polls recent mail and deduplicates by Exchange message ID in the Sheet.
- Store credentials only as GitHub Secrets, never in committed files or chat logs.
