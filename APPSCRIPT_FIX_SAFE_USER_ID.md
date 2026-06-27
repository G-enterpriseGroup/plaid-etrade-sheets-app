# Fix: Plaid `client_user_id` email error

Plaid rejects email addresses in `user.client_user_id`.

In `apps-script/Code.js`, replace this function:

```js
function getUserKey_() {
  const email = Session.getActiveUser().getEmail();
  return email || 'google-sheets-user';
}
```

or this version:

```js
function getUserKey_() {
  const email = Session.getActiveUser().getEmail();
  if (email) return email;
  return 'raj-sheets-user';
}
```

with this safe version:

```js
function getUserKey_() {
  const props = PropertiesService.getUserProperties();
  let userId = props.getProperty('PLAID_SAFE_CLIENT_USER_ID');

  if (!userId) {
    userId = 'sheets_user_' + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('PLAID_SAFE_CLIENT_USER_ID', userId);
  }

  return userId;
}
```

Then save Apps Script and try **Connect E*TRADE / Brokerage** again.
