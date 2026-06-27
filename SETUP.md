# Setup Guide

## 1. Google Sheet

Created Sheet:

```text
Plaid E*TRADE Holdings Dashboard
Spreadsheet ID: 1sgNWMAZEIdOBargwH8ILs_oB-1HQzUoVssaqnsujY2g
```

Open the Sheet, then go to:

```text
Extensions > Apps Script
```

## 2. Paste Apps Script files

Create these files in the Apps Script editor:

```text
Code.js
Sidebar.html
appsscript.json
```

Paste the matching content from the `apps-script` folder.

If you do not see `appsscript.json`, go to:

```text
Project Settings > Show appsscript.json manifest file in editor
```

## 3. Reload the Sheet

After saving Apps Script, reload the Google Sheet.

A new menu should appear:

```text
Plaid Brokerage
```

## 4. Save Plaid secrets

Go to:

```text
Plaid Brokerage > Open Dashboard
```

Enter:

```text
Plaid Client ID
Plaid Secret
Environment: sandbox
Token Memory Scope: user
```

Click:

```text
Save Settings
```

## 5. Link brokerage

Click:

```text
Connect Brokerage
```

After successful Plaid Link, the app stores the Plaid access token in Apps Script Properties.

## 6. Pull data

Click:

```text
Pull Holdings Now
```

Tabs updated:

```text
Holdings
By Ticker
By Account
Portfolio Total
```

For transactions:

```text
Pull Transactions - 365 Days
```

## 7. Local GitHub + clasp sync

Install clasp:

```bash
npm install -g @google/clasp
clasp login
```

Clone the Apps Script project:

```bash
clasp clone-script YOUR_SCRIPT_ID --rootDir apps-script
```

Push changes:

```bash
npm run push:apps-script
```

## 8. Production switch

Only after Plaid approves production access:

```text
Environment: production
```

Then relink the brokerage once in production mode.
