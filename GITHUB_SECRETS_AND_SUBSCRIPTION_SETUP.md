# GitHub Secrets and Subscription Setup

## Important answer

GitHub Secrets can store your Plaid Client ID and Plaid Secret for deployment, but Google Apps Script cannot directly read GitHub Secrets at runtime.

So GitHub Secrets are useful for deploying your backend, not for letting the Sheet directly call Plaid.

## Correct paid-product flow

```text
Customer Google Sheet
  -> Apps Script sidebar
  -> Your backend API
  -> Backend checks Stripe subscription
  -> Backend uses your Plaid credentials from secure server environment
  -> Backend stores each customer's encrypted Plaid access token
  -> Backend returns holdings/transactions to the Sheet
```

## What users should see

Users should not see or enter:

```text
PLAID_CLIENT_ID
PLAID_SECRET
PLAID_ENV
```

They should only see:

```text
Connect Brokerage
Manage Subscription
Pull Holdings
Pull Transactions
```

## Where to put secrets in GitHub

Use repository secrets for deployment automation:

```text
PLAID_CLIENT_ID
PLAID_SECRET
PLAID_ENV
STRIPE_SECRET_KEY
STRIPE_WEBHOOK_SECRET
DATABASE_URL
BACKEND_API_KEY
```

GitHub path:

```text
Repo > Settings > Secrets and variables > Actions > Secrets > New repository secret
```

## What GitHub Secrets are NOT

GitHub Secrets are not a live database.
GitHub Secrets are not readable by Apps Script.
GitHub Secrets should not store every customer's Plaid access token.

## Where customer Plaid access tokens go

Use an encrypted backend database table:

```text
customer_id
email
stripe_customer_id
subscription_status
plaid_item_id
encrypted_plaid_access_token
institution_name
created_at
last_successful_pull
```

## Current repo status

The current Apps Script sidebar has no Plaid secret input fields for customers.

The current prototype still has direct Plaid-call functions in Apps Script. Before selling this, replace direct Plaid calls with backend API calls.

## Next build target

Build backend endpoints:

```text
POST /api/create-link-token
POST /api/exchange-public-token
POST /api/holdings
POST /api/transactions
POST /api/check-subscription
```

Apps Script should call only your backend API, not Plaid directly.
