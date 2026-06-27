# Plaid E*TRADE Google Sheets Dashboard

A Google Sheets + Apps Script dashboard that connects E*TRADE or other Plaid-supported brokerage accounts, stores the Plaid access token after first login, and refreshes holdings/transactions directly into Google Sheets.

## Current Google Sheet

```text
Plaid E*TRADE Holdings Dashboard
Spreadsheet ID: 1sgNWMAZEIdOBargwH8ILs_oB-1HQzUoVssaqnsujY2g
```

## Architecture

```text
Google Sheet Sidebar
  -> Plaid Link popup
  -> Apps Script backend
  -> Plaid Investments API
  -> Holdings / By Ticker / By Account / Portfolio Total / Transactions tabs
```

## What this app does

- Connect brokerage through Plaid Link.
- Exchange `public_token` for `access_token` server-side.
- Store token memory in Apps Script Properties so you do not relink every refresh.
- Pull brokerage holdings into a clean Sheet table.
- Summarize by ticker and account.
- Pull investment transactions.
- Support sandbox, development, and production Plaid environments.
- Keep secrets out of GitHub.

## Files

```text
apps-script/Code.js          Apps Script backend
apps-script/Sidebar.html     User-friendly dashboard UI
apps-script/appsscript.json  Apps Script manifest
.github/workflows/ci.yml     Basic CI syntax/security checks
scripts/push_to_apps_script.sh
SECURITY.md
SETUP.md
```

## Fast setup

1. Open the Google Sheet.
2. Go to `Extensions > Apps Script`.
3. Create these files in Apps Script:
   - `Code.js`
   - `Sidebar.html`
   - `appsscript.json`
4. Paste the matching repo file contents.
5. Reload the Sheet.
6. Open `Plaid Brokerage > Open Dashboard`.
7. Save Plaid credentials.
8. Click `Connect Brokerage`.
9. Click `Pull Holdings Now`.

## Secrets needed inside the Sheet app

Do **not** put these in GitHub.

```text
PLAID_CLIENT_ID
PLAID_SECRET
PLAID_ENV=sandbox
TOKEN_STORAGE_SCOPE=user
```

For real E*TRADE data later:

```text
PLAID_ENV=production
```

## GitHub + clasp workflow

```bash
npm install -g @google/clasp
clasp login
clasp clone-script YOUR_SCRIPT_ID --rootDir apps-script
clasp push
```

## Notes

Plaid is read-only here. This app does not place trades. It only pulls holdings, securities, balances, and investment transactions.
