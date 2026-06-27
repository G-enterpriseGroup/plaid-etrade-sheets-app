/**
 * Plaid E*TRADE Google Sheets Dashboard
 * Apps Script backend
 *
 * Secrets are stored in Apps Script Properties from the sidebar UI.
 * Do not hardcode Plaid secrets in this repo.
 *
 * Version: safe-client-user-id-fix
 */

const APP_NAME = 'Plaid E*TRADE Holdings Dashboard';
const DEFAULT_SPREADSHEET_ID = '1sgNWMAZEIdOBargwH8ILs_oB-1HQzUoVssaqnsujY2g';

const SHEETS = {
  holdings: 'Holdings',
  byTicker: 'By Ticker',
  byAccount: 'By Account',
  total: 'Portfolio Total',
  transactions: 'Transactions',
  config: 'Config'
};

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
  byTicker: [
    'ticker_symbol', 'security_name', 'total_quantity', 'total_market_value',
    'total_cost_basis', 'total_unrealized_gain_loss', 'unrealized_gain_loss_pct',
    'portfolio_weight', 'accounts'
  ],
  byAccount: [
    'institution_name', 'account_name', 'total_market_value', 'total_cost_basis',
    'total_unrealized_gain_loss', 'unrealized_gain_loss_pct', 'portfolio_weight', 'positions'
  ],
  total: [
    'total_market_value', 'total_cost_basis', 'total_unrealized_gain_loss',
    'total_unrealized_gain_loss_pct', 'positions', 'unique_tickers', 'linked_items', 'last_refresh'
  ],
  transactions: [
    'institution_name', 'account_name', 'ticker_symbol', 'security_name',
    'type', 'subtype', 'date', 'name', 'quantity', 'price', 'amount', 'fees',
    'transaction_id', 'pulled_at'
  ],
  config: ['Setting', 'Value', 'Notes']
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Plaid Brokerage')
    .addItem('Open Dashboard', 'showDashboard')
    .addSeparator()
    .addItem('Setup / Repair Tabs', 'setupDashboard')
    .addItem('Pull Holdings Now', 'pullHoldingsToSheet')
    .addItem('Pull Transactions - 365 Days', 'pullTransactions365')
    .addSeparator()
    .addItem('Show Linked Items', 'showLinkedItemsAlert')
    .addItem('Reset Token Memory', 'resetTokenMemoryConfirm')
    .addToUi();
}

function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Plaid Brokerage Dashboard')
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
    linkedItems: items.map(stripSecretItem_)
  };
}

function setupDashboard() {
  const ss = getSpreadsheet_();
  createOrRepairSheet_(ss, SHEETS.holdings, HEADERS.holdings, 2000, 30);
  createOrRepairSheet_(ss, SHEETS.byTicker, HEADERS.byTicker, 1000, 20);
  createOrRepairSheet_(ss, SHEETS.byAccount, HEADERS.byAccount, 1000, 20);
  createOrRepairSheet_(ss, SHEETS.total, HEADERS.total, 200, 12);
  createOrRepairSheet_(ss, SHEETS.transactions, HEADERS.transactions, 5000, 40);
  createOrRepairSheet_(ss, SHEETS.config, HEADERS.config, 200, 10);

  const configSheet = ss.getSheetByName(SHEETS.config);
  configSheet.getRange(2, 1, 8, 3).setValues([
    ['Spreadsheet ID', ss.getId(), 'Used by the Apps Script backend.'],
    ['Plaid Environment', getPlaidEnv_(), 'sandbox first; production later.'],
    ['Token Memory Scope', getTokenScope_(), 'user = per Google user; document = shared in this spreadsheet script.'],
    ['GitHub Repo', 'G-enterpriseGroup/plaid-etrade-sheets-app', 'Source controlled app code.'],
    ['Secret Location', 'Apps Script Properties', 'Do not commit Plaid secrets to GitHub.'],
    ['Relink Rule', 'Only when Plaid update mode is needed', 'Normal refresh uses stored access_token.'],
    ['Client User ID Rule', 'Random safe ID', 'Plaid rejects emails in client_user_id.'],
    ['Last Setup', now_(), '']
  ]);

  return 'Dashboard tabs are ready.';
}

function savePlaidSettings(form) {
  form = form || {};
  const clientId = String(form.clientId || '').trim();
  const secret = String(form.secret || '').trim();
  const env = String(form.env || 'sandbox').trim().toLowerCase();
  const tokenScope = String(form.tokenScope || 'user').trim().toLowerCase();

  if (!clientId) throw new Error('Missing Plaid Client ID.');
  if (!secret) throw new Error('Missing Plaid Secret.');
  if (!['sandbox', 'development', 'production'].includes(env)) throw new Error('Invalid Plaid environment.');
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
  const item = getStoredItems_().find(x => x.item_id === itemId);
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
  const record = {
    access_token: res.access_token,
    item_id: res.item_id,
    institution_name: institution.name || 'Unknown Institution',
    institution_id: institution.institution_id || '',
    linked_at: now_(),
    last_successful_pull: '',
    last_error: '',
    environment: cfg.env,
    metadata_accounts: metadata && metadata.accounts ? metadata.accounts : []
  };

  upsertStoredItem_(record);
  return { status: 'success', item: stripSecretItem_(record), state: getDashboardState() };
}

function markUpdateModeSuccess(itemId) {
  const items = getStoredItems_();
  items.forEach(item => {
    if (item.item_id === itemId) item.last_update_mode = now_();
  });
  setStoredItems_(items);
  return getDashboardState();
}

function pullHoldingsToSheet() {
  setupDashboard();
  const items = getStoredItems_();
  if (!items.length) throw new Error('No linked brokerage items. Open dashboard and connect E*TRADE first.');

  const allRows = [];
  const errors = [];

  items.forEach(item => {
    try {
      const rows = getHoldingsRowsForItem_(item);
      rows.forEach(r => allRows.push(r));
      item.last_successful_pull = now_();
      item.last_error = '';
    } catch (err) {
      item.last_error = err && err.message ? err.message : String(err);
      errors.push((item.institution_name || item.item_id) + ': ' + item.last_error);
    }
  });

  applyPortfolioWeights_(allRows);
  allRows.sort((a, b) => safeNumber_(b[13]) - safeNumber_(a[13]));

  replaceSheetData_(getSpreadsheet_().getSheetByName(SHEETS.holdings), HEADERS.holdings, allRows);
  writeSummaryTabs_(allRows, items.length);
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
  const rows = [];

  (res.holdings || []).forEach(h => {
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
  const total = rows.reduce((sum, r) => sum + safeNumber_(r[13]), 0);
  rows.forEach(r => { r[17] = total ? safeNumber_(r[13]) / total : ''; });
}

function applyAccountWeights_(rows) {
  const accountTotals = {};
  rows.forEach(r => {
    const key = r[0] + '|' + r[1] + '|' + r[3];
    accountTotals[key] = (accountTotals[key] || 0) + safeNumber_(r[13]);
  });
  rows.forEach(r => {
    const key = r[0] + '|' + r[1] + '|' + r[3];
    r[18] = accountTotals[key] ? safeNumber_(r[13]) / accountTotals[key] : '';
  });
}

function writeSummaryTabs_(holdingsRows, linkedItems) {
  const ss = getSpreadsheet_();
  const tickerMap = {};
  const accountMap = {};
  const uniqueTickers = {};
  let totalValue = 0;
  let totalCost = 0;
  let totalUnrealized = 0;

  holdingsRows.forEach(r => {
    const institution = r[0] || '';
    const account = r[1] || '';
    const ticker = r[6] || '(no ticker)';
    const security = r[7] || '';
    const qty = safeNumber_(r[10]);
    const cost = safeNumber_(r[11]);
    const value = safeNumber_(r[13]);
    const unrealized = safeNumber_(r[15]);

    totalValue += value;
    totalCost += cost;
    totalUnrealized += unrealized;
    uniqueTickers[ticker] = true;

    if (!tickerMap[ticker]) tickerMap[ticker] = { ticker, security, qty: 0, value: 0, cost: 0, unrealized: 0, accounts: {} };
    tickerMap[ticker].qty += qty;
    tickerMap[ticker].value += value;
    tickerMap[ticker].cost += cost;
    tickerMap[ticker].unrealized += unrealized;
    tickerMap[ticker].accounts[account] = true;

    const acctKey = institution + '|' + account;
    if (!accountMap[acctKey]) accountMap[acctKey] = { institution, account, value: 0, cost: 0, unrealized: 0, positions: 0 };
    accountMap[acctKey].value += value;
    accountMap[acctKey].cost += cost;
    accountMap[acctKey].unrealized += unrealized;
    accountMap[acctKey].positions += 1;
  });

  const byTickerRows = Object.keys(tickerMap).map(k => {
    const x = tickerMap[k];
    return [
      x.ticker,
      x.security,
      x.qty,
      x.value,
      x.cost,
      x.unrealized,
      x.cost ? x.unrealized / x.cost : '',
      totalValue ? x.value / totalValue : '',
      Object.keys(x.accounts).length
    ];
  }).sort((a, b) => safeNumber_(b[3]) - safeNumber_(a[3]));

  const byAccountRows = Object.keys(accountMap).map(k => {
    const x = accountMap[k];
    return [
      x.institution,
      x.account,
      x.value,
      x.cost,
      x.unrealized,
      x.cost ? x.unrealized / x.cost : '',
      totalValue ? x.value / totalValue : '',
      x.positions
    ];
  }).sort((a, b) => safeNumber_(b[2]) - safeNumber_(a[2]));

  const totalRows = [[
    totalValue,
    totalCost,
    totalUnrealized,
    totalCost ? totalUnrealized / totalCost : '',
    holdingsRows.length,
    Object.keys(uniqueTickers).length,
    linkedItems,
    now_()
  ]];

  replaceSheetData_(ss.getSheetByName(SHEETS.byTicker), HEADERS.byTicker, byTickerRows);
  replaceSheetData_(ss.getSheetByName(SHEETS.byAccount), HEADERS.byAccount, byAccountRows);
  replaceSheetData_(ss.getSheetByName(SHEETS.total), HEADERS.total, totalRows);
  formatFinanceTabs_();
}

function pullTransactions365() {
  return pullInvestmentTransactionsToSheet(365);
}

function pullInvestmentTransactionsToSheet(daysBack) {
  setupDashboard();
  const items = getStoredItems_();
  if (!items.length) throw new Error('No linked brokerage items.');

  const end = new Date();
  const start = new Date();
  start.setDate(start.getDate() - Number(daysBack || 365));
  const allRows = [];

  items.forEach(item => {
    getTransactionRowsForItem_(item, ymd_(start), ymd_(end)).forEach(r => allRows.push(r));
  });

  allRows.sort((a, b) => String(b[6]).localeCompare(String(a[6])));
  replaceSheetData_(getSpreadsheet_().getSheetByName(SHEETS.transactions), HEADERS.transactions, allRows);
  return { status: 'success', rows: allRows.length, message: 'Transactions pulled successfully.' };
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
      options: { count, offset }
    });
    lastRes = res;
    const batch = res.investment_transactions || [];
    batch.forEach(tx => allTx.push(tx));
    total = res.total_investment_transactions || allTx.length;
    if (!batch.length) break;
    offset += count;
  }

  const accounts = mapBy_((lastRes && lastRes.accounts) || [], 'account_id');
  const securities = mapBy_((lastRes && lastRes.securities) || [], 'security_id');

  return allTx.map(t => {
    const acct = accounts[t.account_id] || {};
    const sec = securities[t.security_id] || {};
    return [
      item.institution_name || '', acct.name || '', sec.ticker_symbol || '', sec.name || '',
      t.type || '', t.subtype || '', t.date || '', t.name || '', t.quantity || '',
      t.price || '', t.amount || '', t.fees || '', t.investment_transaction_id || '', now_()
    ];
  });
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
    throw new Error('Plaid returned non-JSON response: ' + text);
  }

  if (code < 200 || code >= 300) {
    throw new Error('Plaid API error ' + code + ': ' + (data.error_message || data.display_message || data.error_code || text));
  }

  return data;
}

function getPlaidConfig_() {
  const cfg = getSafeConfig_();
  if (!cfg.clientId || !cfg.secret) throw new Error('Plaid settings missing. Open dashboard and save Plaid Client ID + Secret first.');
  return cfg;
}

function getSafeConfig_() {
  const props = PropertiesService.getScriptProperties();
  const env = props.getProperty(PROPS.env) || 'sandbox';
  return {
    clientId: props.getProperty(PROPS.clientId),
    secret: props.getProperty(PROPS.secret),
    env,
    baseUrl: plaidBaseUrl_(env)
  };
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
  const idx = items.findIndex(x => x.item_id === item.item_id);
  if (idx >= 0) items[idx] = Object.assign({}, items[idx], item);
  else items.push(item);
  setStoredItems_(items);
}

function deleteLinkedItem(itemId) {
  setStoredItems_(getStoredItems_().filter(x => x.item_id !== itemId));
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
  const res = ui.alert('Reset Token Memory', 'This removes saved Plaid tokens. You will need to relink. Continue?', ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    resetTokenMemory();
    ui.alert('Token memory reset.');
  }
}

function stripSecretItem_(item) {
  return {
    item_id: item.item_id || '',
    institution_name: item.institution_name || '',
    institution_id: item.institution_id || '',
    linked_at: item.linked_at || '',
    last_successful_pull: item.last_successful_pull || '',
    last_update_mode: item.last_update_mode || '',
    environment: item.environment || '',
    last_error: item.last_error || ''
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
  formatColumns_(ss.getSheetByName(SHEETS.byTicker), [4, 5, 6], '$#,##0.00');
  formatColumns_(ss.getSheetByName(SHEETS.byTicker), [7, 8], '0.00%');
  formatColumns_(ss.getSheetByName(SHEETS.byAccount), [3, 4, 5], '$#,##0.00');
  formatColumns_(ss.getSheetByName(SHEETS.byAccount), [6, 7], '0.00%');
  formatColumns_(ss.getSheetByName(SHEETS.total), [1, 2, 3], '$#,##0.00');
  formatColumns_(ss.getSheetByName(SHEETS.total), [4], '0.00%');
}

function formatColumns_(sheet, cols, numberFormat) {
  if (!sheet) return;
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  cols.forEach(col => sheet.getRange(2, col, rows, 1).setNumberFormat(numberFormat));
}

function mapBy_(arr, key) {
  const out = {};
  arr.forEach(x => { if (x && x[key]) out[x[key]] = x; });
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
