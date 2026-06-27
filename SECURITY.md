# Security

## Never commit secrets

Do not commit:

```text
PLAID_SECRET
Plaid access tokens
Secrets.Env
.env
plaid_access_tokens.json
.clasprc.json
```

## Where secrets belong

For this personal Google Sheets app, secrets are stored in Apps Script Properties through the dashboard UI.

## Token memory

The app stores Plaid `access_token` values in Apps Script Properties so you do not relink every refresh.

Default setting:

```text
TOKEN_STORAGE_SCOPE=user
```

This stores token memory for the current Google user only.

## Production recommendation

For a real multi-user production app, move tokens into a backend database with encryption and a cloud secret manager. Apps Script Properties are fine for personal testing, but not ideal for a public multi-user product.

## Plaid data use

This app is read-only. It does not place trades. It pulls holdings and investment transactions through Plaid's Investments endpoints.
