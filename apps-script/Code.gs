/**
 * Portfolio Link - Google Apps Script backend
 * Paste this as Code.gs in Apps Script.
 * Version: simplified-two-tabs-account-selection-v4
 */

const APP_NAME = 'Portfolio Link';
const DEFAULT_SPREADSHEET_ID = '1sgNWMAZEIdOBargwH8ILs_oB-1HQzUoVssaqnsujY2g';
const ITD_START_DATE = '2000-01-01';

const SHEETS = {
  holdings: 'Holdings',
  transactions: 'Transactions'
};

const OLD_SHEETS_TO_REMOVE = ['Config', 'By Ticker', 'By Account', 'Portfolio Total'];

const PROPS = {
  clientId: 'PLAID_CLIENT_ID',
  secret: 'PLAID_SECRET',
  env: 'PLAID_ENV',
  tokenScope: 'TOKEN_STORAGE_SCOPE',
  itemsUser: 'PLAID_ITEMS_USER_JSON',
  itemsDocument: 'PLAID_ITEMS_DOCUMENT_JSON',
  safeClientUserId: 'PLAID_SAFE_CLIENT_USER_ID'
};

const HEADERS = {
  holdings: [
    'institution_name', 'account_name', 'account_official_name', 'account_mask',
    'account_type', 'account_subtype', 'ticker_symbol', 'security_name',
    'security_type', 'security_subtype', 'quantity', 'cost_basis',
    'institution_price', 'institution_value', 'calculated_market_value',
    'unrealized_gain_loss', 'unrealized_gain_loss_pct', 'portfolio_weight',
    'account_weight', 'close_price', 'close_price_as_of', 'pulled_at', 'item_id'
  ],
  transactions: [
    'institution_name', 'account_name', 'ticker_symbol', 'security_name',
    'type', 'subtype', 'date', 'name', 'quantity', 'price', 'amount', 'fees',
    'transaction_id', 'pulled_at'
  ]
};

function onOpen(e) {
  try {
    addPortfolioMenu_();
  } catch (err) {
    console.log('onOpen skipped outside Sheet UI: ' + err);
  }
}

function addPortfolioMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('Plaid Brokerage')
    .addItem('Open Dashboard', 'showDashboard')
    .addToUi();
}

function manualSetupFromEditor() {
  return setupDashboard();
}

function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Portfolio Link')
    .setWidth(430);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getDashboardState() {
  const cfg = getSafeConfig_();
  const items = getStoredItems_();
  return {
    appName: APP_NAME,
    spreadsheetId: getSpreadsheet_().getId(),
    env: cfg.env,
    hasClient: Boolean(cfg.clientId),
    hasSecret: Boolean(cfg.secret),
    tokenScope: getTokenScope_(),
    linkedCount: items.length,
    linkedItems: items.map(stripSecretItem_),
    defaultStartDate: ITD_START_DATE,
    today: ymd_(new Date())
  };
}

function setupDashboard() {
  const ss = getSpreadsheet_();
  removeOldSheets_(ss);
  createOrRepairSheet_(ss, SHEETS.holdings, HEADERS.holdings, 2000, 30);
  createOrRepairSheet_(ss, SHEETS.transactions, HEADERS.transactions, 5000, 40);
  formatFinanceTabs_();
  return 'Simplified tabs are ready: Holdings and Transactions only.';
}

function removeOldSheets_(ss) {
  OLD_SHEETS_TO_REMOVE.forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (sh && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}

function savePlaidSettings(form) {
  form = form || {};
  const clientId = String(form.clientId || '').trim();
  const secret = String(form.secret || '').trim();
  const env = String(form.env || 'sandbox').trim().toLowerCase();
  const tokenScope = String(form.tokenScope || 'user').trim().toLowerCase();

  if (!clientId) throw new Error('Missing service Client ID.');
  if (!secret) throw new Error('Missing service secret.');
  if (!['sandbox', 'development', 'production'].includes(env)) throw new Error('Invalid environment.');
  if (!['user', 'document'].includes(tokenScope)) throw new Error('Invalid token storage scope.');

  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROPS.clientId, clientId);
  props.setProperty(PROPS.secret, secret);
  props.setProperty(PROPS.env, env);
  props.setProperty(PROPS.tokenScope, tokenScope);

  setupDashboard();
  return getDashboardState();
}

function createPlaidLinkToken() {
  const cfg = getPlaidConfig_();
  const payload = {
    client_id: cfg.clientId,
    secret: cfg.secret,
    client_name: APP_NAME,
    user: { client_user_id: getUserKey_() },
    products: ['investments'],
    country_codes: ['US'],
    language: 'en'
  };
  return plaidPost_('/link/token/create', payload).link_token;
}

function createPlaidUpdateLinkToken(itemId) {
  const item = getStoredItems_().find(function(x) { return x.item_id === itemId; });
  if (!item) throw new Error('Item not found: ' + itemId);

  const cfg = getPlaidConfig_();
  const payload = {
    client_id: cfg.clientId,
    secret: cfg.secret,
    client_name: APP_NAME,
    user: { client_user_id: getUserKey_() },
    access_token: item.access_token,
    country_codes: ['US'],
    language: 'en'
  };
  return plaidPost_('/link/token/create', payload).link_token;
}

function exchangePublicToken(publicToken, metadata) {
  if (!publicToken) throw new Error('Missing public token.');
  const cfg = getPlaidConfig_();
  const res = plaidPost_('/item/public_token/exchange', {
    client_id: cfg.clientId,
    secret: cfg.secret,
    public_token: publicToken
  });

  const institution = metadata && metadata.institution ? metadata.institution : {};
  const accounts = normalizeMetadataAccounts_((metadata && metadata.accounts) || []);
  const record = {
    access_token: res.access_token,
    item_id: res.item_id,
    institution_name: institution.name || 'Unknown Institution',
    institution_id: institution.institution_id || '',
    linked_at: now_(),
    last_successful_pull: '',
    last_error: '',
    environment: cfg.env,
    metadata_accounts: accounts,
    selected_account_ids: accounts.map(function(a) { return a.account_id; })
  };

  upsertStoredItem_(record);
  return { status: 'success', item: stripSecretItem_(record), state: getDashboardState() };
}

function markUpdateModeSuccess(itemId) {
  const items = getStoredItems_();
  items.forEach(function(item) {
    if (item.item_id === itemId) item.last_update_mode = now_();
  });
  setStoredItems_(items);
  return getDashboardState();
}

function saveAccountSelection(payload) {
  payload = payload || {};
  const selectedByItem = payload.selectedByItem || {};
  const items = getStoredItems_();

  items.forEach(function(item) {
    const acctIds = Array.isArray(selectedByItem[item.item_id]) ? selectedByItem[item.item_id] : [];
    item.selected_account_ids = acctIds.map(String).filter(Boolean);
  });

  setStoredItems_(items);
  return getDashboardState();
}

function pullHoldingsToSheet() {
  setupDashboard();
  const items = getStoredItems_();
  if (!items.length) throw new Error('No linked brokerage items. Open Settings and connect first.');

  const allRows = [];
  const errors = [];

  items.forEach(function(item) {
    try {
      const rows = getHoldingsRowsForItem_(item);
      rows.forEach(function(r) { allRows.push(r); });
      item.last_successful_pull = now_();
      item.last_error = '';
    } catch (err) {
      item.last_error = err && err.message ? err.message : String(err);
      errors.push((item.institution_name || item.item_id) + ': ' + item.last_error);
    }
  });

  applyPortfolioWeights_(allRows);
  allRows.sort(function(a, b) { return safeNumber_(b[13]) - safeNumber_(a[13]); });

  replaceSheetData_(getSpreadsheet_().getSheetByName(SHEETS.holdings), HEADERS.holdings, allRows);
  setStoredItems_(items);

  return {
    status: errors.length ? 'partial' : 'success',
    rows: allRows.length,
    message: errors.length ? 'Holdings pulled with errors: ' + errors.join(' | ') : 'Holdings pulled successfully.'
  };
}

function getHoldingsRowsForItem_(item) {
  const cfg = getPlaidConfig_();
  const res = plaidPost_('/investments/holdings/get', {
    client_id: cfg.clientId,
    secret: cfg.secret,
    access_token: item.access_token
  });

  const accounts = mapBy_(res.accounts || [], 'account_id');
  const securities = mapBy_(res.securities || [], 'security_id');
  refreshItemAccountsFromPlaid_(item, res.accounts || []);

  const rows = [];
  (res.holdings || []).forEach(function(h) {
    if (!accountIsSelected_(item, h.account_id)) return;

    const acct = accounts[h.account_id] || {};
    const sec = securities[h.security_id] || {};
    const quantity = num_(h.quantity);
    const costBasis = num_(h.cost_basis);
    const price = num_(h.institution_price);
    const value = num_(h.institution_value);
    const calculatedValue = quantity !== '' && price !== '' ? safeNumber_(quantity) * safeNumber_(price) : '';
    const unrealized = value !== '' && costBasis !== '' ? safeNumber_(value) - safeNumber_(costBasis) : '';
    const unrealizedPct = costBasis ? safeNumber_(unrealized) / safeNumber_(costBasis) : '';

    rows.push([
      item.institution_name || '', acct.name || '', acct.official_name || '', acct.mask || '',
      acct.type || '', acct.subtype || '', sec.ticker_symbol || '', sec.name || '',
      sec.type || '', sec.subtype || '', quantity, costBasis, price, value,
      calculatedValue, unrealized, unrealizedPct, '', '', sec.close_price || '',
      sec.close_price_as_of || '', now_(), item.item_id || ''
    ]);
  });

  applyAccountWeights_(rows);
  return rows;
}

function applyPortfolioWeights_(rows) {
  const total = rows.reduce(function(sum, r) { return sum + safeNumber_(r[13]); }, 0);
  rows.forEach(function(r) { r[17] = total ? safeNumber_(r[13]) / total : ''; });
}

function applyAccountWeights_(rows) {
  const accountTotals = {};
  rows.forEach(function(r) {
    const key = r[0] + '|' + r[1] + '|' + r[3];
    accountTotals[key] = (accountTotals[key] || 0) + safeNumber_(r[13]);
  });
  rows.forEach(function(r) {
    const key = r[0] + '|' + r[1] + '|' + r[3];
    r[18] = accountTotals[key] ? safeNumber_(r[13]) / accountTotals[key] : '';
  });
}

function pullTransactions365() {
  return pullTransactionsFromSidebar({ dateMode: 'custom', startDate: daysAgoYmd_(365), endDate: ymd_(new Date()), writeMode: 'append' });
}

function pullTransactionsITDAppend() {
  return pullTransactionsFromSidebar({ dateMode: 'itd', writeMode: 'append' });
}

function pullTransactionsFromSidebar(form) {
  form = form || {};
  setupDashboard();

  const dateMode = String(form.dateMode || 'itd').toLowerCase();
  const writeMode = String(form.writeMode || 'append').toLowerCase();
  const replaceConfirmed = Boolean(form.replaceConfirmed);

  let startDate = dateMode === 'custom' ? String(form.startDate || '').trim() : ITD_START_DATE;
  let endDate = dateMode === 'custom' ? String(form.endDate || '').trim() : ymd_(new Date());

  if (!startDate) startDate = ITD_START_DATE;
  if (!endDate) endDate = ymd_(new Date());
  validateDate_(startDate, 'Start date');
  validateDate_(endDate, 'End date');
  if (startDate > endDate) throw new Error('Start date cannot be after end date.');
  if (writeMode === 'replace' && !replaceConfirmed) throw new Error('Replace blocked. User permission required.');

  const rows = getInvestmentTransactionRows_(startDate, endDate);
  const sheet = getSpreadsheet_().getSheetByName(SHEETS.transactions);

  if (writeMode === 'replace') {
    replaceSheetData_(sheet, HEADERS.transactions, rows);
    return { status: 'success', mode: 'replace', rowsPulled: rows.length, rowsAdded: rows.length, skippedDuplicates: 0, message: 'Transactions replaced for ' + startDate + ' to ' + endDate + '.' };
  }

  const result = appendTransactionRows_(sheet, rows);
  return {
    status: 'success',
    mode: 'append',
    rowsPulled: rows.length,
    rowsAdded: result.added,
    skippedDuplicates: result.duplicates,
    message: 'Transactions appended for ' + startDate + ' to ' + endDate + '. Added ' + result.added + ', skipped duplicates ' + result.duplicates + '.'
  };
}

function getInvestmentTransactionRows_(startDate, endDate) {
  const items = getStoredItems_();
  if (!items.length) throw new Error('No linked brokerage items. Open Settings and connect first.');

  const allRows = [];
  items.forEach(function(item) {
    getTransactionRowsForItem_(item, startDate, endDate).forEach(function(r) { allRows.push(r); });
  });

  allRows.sort(function(a, b) { return String(b[6]).localeCompare(String(a[6])); });
  return allRows;
}

function getTransactionRowsForItem_(item, startDate, endDate) {
  const cfg = getPlaidConfig_();
  const allTx = [];
  let offset = 0;
  const count = 500;
  let total = null;
  let lastRes = null;

  while (total === null || allTx.length < total) {
    const res = plaidPost_('/investments/transactions/get', {
      client_id: cfg.clientId,
      secret: cfg.secret,
      access_token: item.access_token,
      start_date: startDate,
      end_date: endDate,
      options: { count: count, offset: offset }
    });
    lastRes = res;
    const batch = res.investment_transactions || [];
    batch.forEach(function(tx) { allTx.push(tx); });
    total = res.total_investment_transactions || allTx.length;
    if (!batch.length) break;
    offset += count;
  }

  const accounts = mapBy_((lastRes && lastRes.accounts) || [], 'account_id');
  const securities = mapBy_((lastRes && lastRes.securities) || [], 'security_id');
  refreshItemAccountsFromPlaid_(item, (lastRes && lastRes.accounts) || []);

  return allTx.filter(function(t) {
    return accountIsSelected_(item, t.account_id);
  }).map(function(t) {
    const acct = accounts[t.account_id] || {};
    const sec = securities[t.security_id] || {};
    return [item.institution_name || '', acct.name || '', sec.ticker_symbol || '', sec.name || '', t.type || '', t.subtype || '', t.date || '', t.name || '', t.quantity || '', t.price || '', t.amount || '', t.fees || '', t.investment_transaction_id || transactionFallbackKey_(item, t), now_()];
  });
}

function appendTransactionRows_(sheet, rows) {
  createOrRepairSheet_(getSpreadsheet_(), SHEETS.transactions, HEADERS.transactions, 5000, 40);
  const existing = getExistingTransactionKeys_(sheet);
  const toAppend = [];
  let duplicates = 0;

  rows.forEach(function(row) {
    const key = String(row[12] || '').trim();
    if (key && existing[key]) {
      duplicates += 1;
      return;
    }
    if (key) existing[key] = true;
    toAppend.push(row);
  });

  if (toAppend.length) {
    const startRow = Math.max(sheet.getLastRow() + 1, 2);
    ensureSize_(sheet, startRow + toAppend.length + 20, Math.max(HEADERS.transactions.length, sheet.getMaxColumns()));
    sheet.getRange(startRow, 1, toAppend.length, HEADERS.transactions.length).setValues(toAppend);
    sheet.autoResizeColumns(1, HEADERS.transactions.length);
  }

  return { added: toAppend.length, duplicates: duplicates };
}

function getExistingTransactionKeys_(sheet) {
  const keys = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;
  const values = sheet.getRange(2, 13, lastRow - 1, 1).getValues();
  values.forEach(function(r) {
    const key = String(r[0] || '').trim();
    if (key) keys[key] = true;
  });
  return keys;
}

function transactionFallbackKey_(item, t) {
  return [item.item_id || '', t.account_id || '', t.date || '', t.name || '', t.amount || '', t.quantity || '', t.price || ''].join('|');
}

function accountIsSelected_(item, accountId) {
  const selected = Array.isArray(item.selected_account_ids) ? item.selected_account_ids.map(String) : [];
  if (!selected.length) return true;
  return selected.indexOf(String(accountId || '')) !== -1;
}

function normalizeMetadataAccounts_(accounts) {
  return (accounts || []).map(function(a) {
    return {
      account_id: String(a.id || a.account_id || ''),
      name: a.name || '',
      mask: a.mask || '',
      type: a.type || '',
      subtype: a.subtype || ''
    };
  }).filter(function(a) { return a.account_id; });
}

function refreshItemAccountsFromPlaid_(item, accounts) {
  const normalized = normalizePlaidAccounts_(accounts || []);
  if (!normalized.length) return;
  item.metadata_accounts = normalized;
  if (!Array.isArray(item.selected_account_ids)) {
    item.selected_account_ids = normalized.map(function(a) { return a.account_id; });
  }
}

function normalizePlaidAccounts_(accounts) {
  return (accounts || []).map(function(a) {
    return {
      account_id: String(a.account_id || a.id || ''),
      name: a.name || '',
      official_name: a.official_name || '',
      mask: a.mask || '',
      type: a.type || '',
      subtype: a.subtype || ''
    };
  }).filter(function(a) { return a.account_id; });
}

function plaidPost_(path, payload) {
  const cfg = getPlaidConfig_();
  const response = UrlFetchApp.fetch(cfg.baseUrl + path, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  let data;
  try {
    data = JSON.parse(text);
  } catch (err) {
    throw new Error('Provider returned non-JSON response: ' + text);
  }

  if (code < 200 || code >= 300) {
    throw new Error('Provider API error ' + code + ': ' + (data.error_message || data.display_message || data.error_code || text));
  }

  return data;
}

function getPlaidConfig_() {
  const cfg = getSafeConfig_();
  if (!cfg.clientId || !cfg.secret) throw new Error('Service settings missing. For prototype, admin must configure Script Properties first.');
  return cfg;
}

function getSafeConfig_() {
  const props = PropertiesService.getScriptProperties();
  const env = props.getProperty(PROPS.env) || 'sandbox';
  return { clientId: props.getProperty(PROPS.clientId), secret: props.getProperty(PROPS.secret), env: env, baseUrl: plaidBaseUrl_(env) };
}

function plaidBaseUrl_(env) {
  if (env === 'production') return 'https://production.plaid.com';
  if (env === 'development') return 'https://development.plaid.com';
  return 'https://sandbox.plaid.com';
}

function getPlaidEnv_() {
  return PropertiesService.getScriptProperties().getProperty(PROPS.env) || 'sandbox';
}

function getTokenScope_() {
  return PropertiesService.getScriptProperties().getProperty(PROPS.tokenScope) || 'user';
}

function getStorage_() {
  return getTokenScope_() === 'document' ? PropertiesService.getDocumentProperties() : PropertiesService.getUserProperties();
}

function getItemsKey_() {
  return getTokenScope_() === 'document' ? PROPS.itemsDocument : PROPS.itemsUser;
}

function getStoredItems_() {
  const raw = getStorage_().getProperty(getItemsKey_());
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function setStoredItems_(items) {
  getStorage_().setProperty(getItemsKey_(), JSON.stringify(items || []));
}

function upsertStoredItem_(item) {
  const items = getStoredItems_();
  const idx = items.findIndex(function(x) { return x.item_id === item.item_id; });
  if (idx >= 0) items[idx] = Object.assign({}, items[idx], item);
  else items.push(item);
  setStoredItems_(items);
}

function deleteLinkedItem(itemId) {
  setStoredItems_(getStoredItems_().filter(function(x) { return x.item_id !== itemId; }));
  return getDashboardState();
}

function resetTokenMemory() {
  getStorage_().deleteProperty(getItemsKey_());
  return getDashboardState();
}

function showLinkedItemsAlert() {
  SpreadsheetApp.getUi().alert(JSON.stringify(getStoredItems_().map(stripSecretItem_), null, 2));
}

function resetTokenMemoryConfirm() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('Reset Token Memory', 'This removes saved account tokens. Continue?', ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    resetTokenMemory();
    ui.alert('Token memory reset.');
  }
}

function stripSecretItem_(item) {
  const accounts = Array.isArray(item.metadata_accounts) ? item.metadata_accounts : [];
  const selected = Array.isArray(item.selected_account_ids) ? item.selected_account_ids.map(String) : accounts.map(function(a) { return String(a.account_id || ''); });
  return {
    item_id: item.item_id || '',
    institution_name: item.institution_name || '',
    institution_id: item.institution_id || '',
    linked_at: item.linked_at || '',
    last_successful_pull: item.last_successful_pull || '',
    last_update_mode: item.last_update_mode || '',
    environment: item.environment || '',
    last_error: item.last_error || '',
    accounts: accounts,
    selected_account_ids: selected
  };
}

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(DEFAULT_SPREADSHEET_ID);
}

function createOrRepairSheet_(ss, name, headers, minRows, minCols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  ensureSize_(sh, minRows, minCols);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9e8ff');
  sh.autoResizeColumns(1, headers.length);
  return sh;
}

function replaceSheetData_(sheet, headers, rows) {
  ensureSize_(sheet, Math.max((rows || []).length + 1, 100), Math.max(headers.length, sheet.getMaxColumns()));
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9e8ff');
  sheet.autoResizeColumns(1, headers.length);
}

function ensureSize_(sheet, minRows, minCols) {
  if (sheet.getMaxRows() < minRows) sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < minCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), minCols - sheet.getMaxColumns());
}

function formatFinanceTabs_() {
  const ss = getSpreadsheet_();
  formatColumns_(ss.getSheetByName(SHEETS.holdings), [12, 13, 14, 15, 16, 20], '$#,##0.00');
  formatColumns_(ss.getSheetByName(SHEETS.holdings), [17, 18, 19], '0.00%');
  formatColumns_(ss.getSheetByName(SHEETS.transactions), [9, 10, 11, 12], '$#,##0.00');
}

function formatColumns_(sheet, cols, numberFormat) {
  if (!sheet) return;
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  cols.forEach(function(col) { sheet.getRange(2, col, rows, 1).setNumberFormat(numberFormat); });
}

function mapBy_(arr, key) {
  const out = {};
  arr.forEach(function(x) { if (x && x[key]) out[x[key]] = x; });
  return out;
}

function getUserKey_() {
  const props = PropertiesService.getUserProperties();
  let userId = props.getProperty(PROPS.safeClientUserId);
  if (!userId) {
    userId = 'sheets_user_' + Utilities.getUuid().replace(/-/g, '');
    props.setProperty(PROPS.safeClientUserId, userId);
  }
  return userId;
}

function validateDate_(value, label) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error(label + ' must be YYYY-MM-DD.');
}

function daysAgoYmd_(days) {
  const d = new Date();
  d.setDate(d.getDate() - Number(days || 0));
  return ymd_(d);
}

function now_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function ymd_(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function num_(x) {
  if (x === null || x === undefined || x === '') return '';
  const n = Number(x);
  return isNaN(n) ? '' : n;
}

function safeNumber_(x) {
  const n = Number(x);
  return isNaN(n) ? 0 : n;
}
