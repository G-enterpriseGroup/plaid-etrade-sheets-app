# Subscription Architecture

This project can be used for personal testing inside Google Sheets, but a paid subscription version needs a real backend.

## Personal testing

For your own testing, save Plaid settings through the Sheet sidebar:

```text
Plaid Brokerage > Open Dashboard > Save Plaid Settings
```

Apps Script stores settings in Script Properties and stores each user's linked item memory in User Properties.

## Paid subscription version

Do not expose provider credentials or stored brokerage tokens inside customer-visible Apps Script.

Use this structure instead:

```text
Google Sheet Add-on / Sidebar
  -> Your backend API
  -> Secret Manager
  -> Encrypted database for customer item tokens
  -> Plaid API
  -> Return holdings summary to the Sheet
```

Recommended backend options:

```text
Google Cloud Run
Firebase Functions
Supabase Edge Functions
AWS Lambda
```

Recommended storage:

```text
Secret Manager for provider credentials
Encrypted database for customer item tokens
Stripe for subscription status
```

## Why not GitHub for secrets

GitHub is for source code. Real credentials and brokerage access tokens should not be committed to a repository. Even a private repo is not the right place for production secrets.

## Current safe path

The current app remains safe for local testing because the committed repo contains only source code. Secrets are entered through the dashboard and stored in Apps Script Properties.
