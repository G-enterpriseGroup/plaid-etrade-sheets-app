# Setup Secrets

Do not put Plaid secrets in GitHub.

## In Google Sheet

Open:

```text
Plaid Brokerage > Open Dashboard
```

Enter:

```text
PLAID_CLIENT_ID = your Plaid Client ID
PLAID_SECRET = your Plaid sandbox or production secret
PLAID_ENV = sandbox
TOKEN_STORAGE_SCOPE = user
```

Use `sandbox` first. Change to `production` only after Plaid production approval for Investments.

## Token memory

Use:

```text
TOKEN_STORAGE_SCOPE = user
```

This stores the Plaid access token for your Google user so you do not have to relink every refresh.

Relink is only needed if Plaid/E*TRADE requires update mode, MFA refresh, password change, or the item breaks.

## GitHub secrets

For the current Apps Script build, GitHub secrets are not required.

If later we add GitHub Actions deployment with clasp, then add GitHub repo secrets:

```text
CLASPRC_JSON
SCRIPT_ID
```

Never add:

```text
PLAID_SECRET
PLAID_ACCESS_TOKEN
Secrets.Env
plaid_access_tokens.json
```
