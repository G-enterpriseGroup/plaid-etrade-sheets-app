/**
 * Market Analysis Report module
 * Reads Holdings and creates Report Market Analysis.
 * Run: buildMarketAnalysisReport()
 * Debug: debugMarketAnalysisHoldingsHeaders()
 * Version: robust-holdings-header-scan-v2
 */

var MARKET_REPORT = {
  holdingsSheet: 'Holdings',
  reportSheet: 'Report Market Analysis',
  font: 'Times New Roman'
};

var MARKET_SECTOR_ETFS = [
  ['XLC', 'Communication Services'],
  ['XLY', 'Consumer Discretionary'],
  ['XLP', 'Consumer Staples'],
  ['XLE', 'Energy'],
  ['XLF', 'Financials'],
  ['XLV', 'Healthcare'],
  ['XLI', 'Industrials'],
  ['XLB', 'Materials'],
  ['XLRE', 'Real Estate'],
  ['XLK', 'Technology'],
  ['XLU', 'Utilities']
];

var MARKET_TICKER_MAP = {
  SPYM: ['Core Equity', 'SPY', 'S&P 500 core exposure'],
  DIA: ['Core Equity', 'DIA', 'Dow blue-chip exposure'],
  SCHG: ['Growth Equity', 'XLK', 'Large-cap growth tilt'],
  SPMO: ['Momentum Equity', 'SPY', 'Momentum factor exposure'],
  BAC: ['Financials', 'XLF', 'Bank exposure'],
  MS: ['Financials', 'XLF', 'Capital markets exposure'],
  STT: ['Financials', 'XLF', 'Custody bank exposure'],
  SONY: ['Consumer / ADR', 'XLY', 'Consumer technology and media ADR'],
  LMT: ['Defense / Industrials', 'XLI', 'Defense industrial exposure'],
  HTD: ['Income Equity', 'XLU', 'Dividend income sleeve'],
  SGOL: ['Gold / Alternative', 'GLD', 'Gold hedge exposure'],
  JPST: ['Short Duration Safety', 'BIL', 'Ultra-short income stabilizer'],
  VRIG: ['Floating Rate Safety', 'BIL', 'Floating-rate income stabilizer'],
  CLOZ: ['Credit Income', 'BIL', 'CLO credit income sleeve']
};

function buildMarketAnalysisReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var holdings = mrReadHoldings_(ss);
  if (!holdings.length) {
    throw new Error('No usable holdings found. The Holdings tab exists, but no rows matched after header scan/exclusions. Run debugMarketAnalysisHoldingsHeaders() and send the result.');
  }

  var sh = ss.getSheetByName(MARKET_REPORT.reportSheet) || ss.insertSheet(MARKET_REPORT.reportSheet);
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), Math.min(13, sh.getMaxColumns()))
    .setFontFamily(MARKET_REPORT.font)
    .setFontSize(9)
    .setWrap(true)
    .setVerticalAlignment('top');

  var portfolioRows = mrBuildPortfolioRows_(holdings);
  var row = 1;

  row = mrTitle_(sh, row, 'Raj Market Rotation Report', 'Built ' + mrNow_() + '. Reads connected E*TRADE Holdings, excluding cash and money market. This foundation tab prepares the structure for the later live macro/news engine.');

  row = mrSection_(sh, row, 'Macro Risk Dashboard');
  row = mrTable_(sh, row, ['Risk Area', 'Current Read', 'Risk Level', 'Portfolio Meaning'], [
    ['War / Geopolitics', 'Manual live-news check required', 'Review', 'Affects oil, defense, gold, and risk appetite'],
    ['Oil / Energy shock', 'Check oil trend and XLE', 'Review', 'Affects inflation and energy rotation'],
    ['Crisis / Liquidity', 'Check SPY trend, credit, and TLT', 'Review', 'Affects risk-on versus safety sleeves'],
    ['Fed / Yields', 'Check FOMC, TLT, and 10Y yield', 'Review', 'Affects growth, banks, real estate, and bonds'],
    ['Inflation / Jobs', 'Check CPI, PCE, payrolls', 'Review', 'Affects duration and equity multiple risk']
  ]);
  row = mrParagraph_(sh, row, 'Macro summary: this Apps Script version builds the report shell and portfolio mapping from Holdings. The final paid/backend version should inject live macro, news, earnings, and source citations before using the report for decisions.');

  row = mrSection_(sh, row, 'Sector Rotation - SPDR Map');
  row = mrTable_(sh, row, ['ETF', 'Sector', 'Rotation Read', 'Flow Pressure', 'Action Bias'], mrBuildSectorRows_());
  row = mrParagraph_(sh, row, 'Sector summary: SPDR sectors are used as the rotation map. SPY is the benchmark for relative strength. Confirm sector leadership with both macro thesis and technical trend.');

  row = mrSection_(sh, row, 'Technical Confirmation Snapshot');
  row = mrTable_(sh, row, ['ETF', 'Sector', 'Price vs 20 EMA', '50 SMA', '200 SMA', '20/50 Crossover', '50/200 Trend', 'MACD vs Signal', 'MACD Zero', 'RS vs SPY'], mrBuildTechnicalRows_());

  row = mrSection_(sh, row, 'Prior Week Recap + Forward Trend');
  row = mrParagraph_(sh, row, 'Prior-week recap should summarize what led, what lagged, and whether the forward setup favors growth, income, safety, or defense. This tab is ready for the live macro/news layer.');

  row = mrSection_(sh, row, 'Raj Portfolio Impact - Exact Actions');
  row = mrTable_(sh, row, ['Ticker', 'Qty Held', 'Cost Basis', 'Current Value', 'Unrealized $', 'P&L %', 'Tolerance Status', 'Thesis', 'Exact Action', 'Action Qty', 'Est. $ Value', 'Reason / Price Source', 'Sleeve'], portfolioRows);

  row = mrSection_(sh, row, 'Total Suggested Trims and Primary Goal');
  row = mrTable_(sh, row, ['Metric', 'Value'], [
    ['Total suggested trims', '$0.00 pending live confirmation'],
    ['Primary goal', 'Use macro + sector + technical confirmation before changing holdings. Keep hold names at the bottom and action names at the top once action engine is enabled.']
  ]);

  row = mrSection_(sh, row, 'Aggressive Growth Setup With Risk Controls');
  row = mrTable_(sh, row, ['Setup', 'Trigger', 'Risk Control', 'Action Bias'], [
    ['Core growth add', 'SPY/XLK trend confirmed', 'Starter size only', 'Add only if macro confirms'],
    ['Profit protection', 'Large gain or overweight sleeve', 'Trim small amount, not full exit', 'Protect gains'],
    ['Safety redeployment', 'Safety sleeve heavy and growth confirmed', 'Keep liquidity buffer', 'Gradual shift only'],
    ['Weak trend', 'Below key averages with weak momentum', 'No averaging down without confirmation', 'Wait / avoid']
  ]);

  row = mrSection_(sh, row, 'Key Catalysts to Watch');
  row = mrTable_(sh, row, ['Catalyst', 'Why It Matters'], [
    ['Fed/FOMC and Treasury yields', 'Affects growth multiples, banks, real estate, and fixed income'],
    ['Oil and geopolitical headlines', 'Affects energy, inflation pressure, defense, and gold'],
    ['Earnings guidance', 'Can override sector trend'],
    ['Credit spreads and liquidity', 'Important for CLOZ, banks, and risk appetite'],
    ['SPY breadth and sector relative strength', 'Confirms whether rotation is broadening or narrowing']
  ]);

  row = mrSection_(sh, row, 'Chicago-Style Source List');
  row = mrTable_(sh, row, ['Source', 'Use'], [
    ['Connected E*TRADE Holdings tab', 'Portfolio quantities, prices, values, weights, and P&L'],
    ['SPDR sector ETF map', '11-sector rotation framework'],
    ['SPY benchmark', 'Relative-strength benchmark'],
    ['Manual live-news layer', 'Macro, Fed, inflation, earnings, and geopolitical catalysts']
  ]);

  mrFinalize_(sh, row);
  return 'Report Market Analysis tab built from Holdings. Rows analyzed: ' + holdings.length;
}

function debugMarketAnalysisHoldingsHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MARKET_REPORT.holdingsSheet);
  if (!sh) return 'No Holdings tab found.';
  var values = sh.getDataRange().getValues();
  var found = mrFindHeaderRow_(values);
  if (!found) {
    return 'Could not find Holdings header row. First row values: ' + JSON.stringify(values[0] || []);
  }
  var rows = mrReadHoldings_(ss);
  return 'Header row detected at sheet row ' + (found.headerRowIndex + 1) + '. Holdings rows included after exclusions: ' + rows.length + '. Headers detected: ' + JSON.stringify(found.headers);
}

function mrReadHoldings_(ss) {
  var sh = ss.getSheetByName(MARKET_REPORT.holdingsSheet);
  if (!sh) throw new Error('Missing Holdings tab.');
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  var found = mrFindHeaderRow_(values);
  if (!found) {
    throw new Error('Could not find the Holdings header row. Make sure the sheet has columns like ticker_symbol, security_name, quantity, institution_value.');
  }

  var ix = found.index;
  var rows = [];
  for (var r = found.headerRowIndex + 1; r < values.length; r++) {
    var row = values[r];
    var ticker = mrCellByKey_(row, ix, 'ticker_symbol').toUpperCase();
    var name = mrCellByKey_(row, ix, 'security_name');
    var type = mrCellByKey_(row, ix, 'security_type').toLowerCase();
    if (!ticker && !name) continue;
    if (mrExclude_(ticker, name, type)) continue;

    rows.push({
      ticker: ticker || '(no ticker)',
      name: name,
      type: type,
      subtype: mrCellByKey_(row, ix, 'security_subtype'),
      qty: mrNum_(mrValueByKey_(row, ix, 'quantity')),
      cost: mrNum_(mrValueByKey_(row, ix, 'cost_basis')),
      price: mrNum_(mrValueByKey_(row, ix, 'institution_price')),
      value: mrNum_(mrValueByKey_(row, ix, 'institution_value')) || mrNum_(mrValueByKey_(row, ix, 'calculated_market_value')),
      pnl: mrNum_(mrValueByKey_(row, ix, 'unrealized_gain_loss')),
      pnlPct: mrPct_(mrValueByKey_(row, ix, 'unrealized_gain_loss_pct')),
      weight: mrPct_(mrValueByKey_(row, ix, 'portfolio_weight'))
    });
  }
  return rows;
}

function mrFindHeaderRow_(values) {
  var scanRows = Math.min(values.length, 25);
  for (var r = 0; r < scanRows; r++) {
    var headers = values[r].map(function(h) { return mrNormHeader_(h); });
    var index = {};
    headers.forEach(function(h, i) { if (h) index[h] = i; });
    var score = 0;
    ['ticker_symbol', 'security_name', 'quantity', 'institution_value'].forEach(function(k) {
      if (index[k] !== undefined) score++;
    });
    if (score >= 3) {
      return { headerRowIndex: r, headers: headers, index: index };
    }
  }
  return null;
}

function mrNormHeader_(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
    .replace(/[^a-z0-9_]/g, '')
    .replace(/^_+|_+$/g, '');
}

function mrCellByKey_(row, ix, key) {
  var i = ix[key];
  if (i === undefined || i < 0) return '';
  var v = row[i];
  return String(v === null || v === undefined ? '' : v).trim();
}

function mrValueByKey_(row, ix, key) {
  var i = ix[key];
  if (i === undefined || i < 0) return '';
  return row[i];
}

function mrExclude_(ticker, name, type) {
  var n = String(name || '').toLowerCase();
  if (type === 'cash') return true;
  if (ticker.indexOf('CUR:') === 0) return true;
  if (n.indexOf('money market') >= 0) return true;
  if (n.indexOf('sweep') >= 0) return true;
  if (['VMFXX','SWVXX','SPAXX','FDRXX'].indexOf(ticker) >= 0) return true;
  return false;
}

function mrBuildSectorRows_() {
  return MARKET_SECTOR_ETFS.map(function(x) {
    return [x[0], x[1], 'Pending live technical feed', 'Pending RS vs SPY', 'Hold / wait for confirmation'];
  });
}

function mrBuildTechnicalRows_() {
  return MARKET_SECTOR_ETFS.map(function(x) {
    return [x[0], x[1], 'Pending', 'Pending', 'Pending', 'Pending', 'Pending', 'Pending', 'Pending', 'Pending'];
  });
}

function mrBuildPortfolioRows_(holdings) {
  return holdings.map(function(h) {
    var meta = mrMeta_(h);
    return [
      h.ticker,
      h.qty,
      mrMoney_(h.cost),
      mrMoney_(h.value),
      mrMoney_(h.pnl),
      mrPctText_(h.pnlPct),
      mrTolerance_(h, meta),
      meta[2],
      'Hold',
      0,
      mrMoney_(0),
      'Portfolio row built from E*TRADE institution price/value. Action engine requires live macro + sector + technical confirmation.',
      meta[0]
    ];
  }).sort(function(a, b) {
    return mrParseMoney_(b[3]) - mrParseMoney_(a[3]);
  });
}

function mrMeta_(h) {
  if (MARKET_TICKER_MAP[h.ticker]) return MARKET_TICKER_MAP[h.ticker];
  if (h.type.indexOf('fixed') >= 0 || h.name.toLowerCase().indexOf('fdic') >= 0) return ['Fixed Income Safety', 'BIL', 'Principal/income stabilizer; watch safety overweight'];
  if (h.type.indexOf('etf') >= 0) return ['ETF / Unmapped', 'SPY', 'ETF exposure; map sleeve manually if material'];
  return ['Equity / Unmapped', 'SPY', 'Single-stock exposure; validate thesis manually'];
}

function mrTolerance_(h, meta) {
  if (meta[0].indexOf('Safety') >= 0 && h.weight > 0.25) return 'Overweight safety';
  if (meta[0].indexOf('Safety') >= 0) return 'Income tilt';
  if (h.pnlPct > 0.50) return 'Profit outlier';
  if (h.pnlPct < -0.10) return 'Weak';
  if (h.weight < 0.015) return 'Too small';
  if (h.weight > 0.08) return 'Near limit';
  return 'In tolerance';
}

function mrTitle_(sh, row, title, subtitle) {
  sh.getRange(row, 1, 1, 13).merge().setValue(title).setFontSize(18).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff');
  row++;
  sh.getRange(row, 1, 1, 13).merge().setValue(subtitle).setFontSize(9).setFontColor('#374151').setBackground('#eef2ff');
  return row + 2;
}
function mrSection_(sh, row, title) {
  sh.getRange(row, 1, 1, 13).merge().setValue(title).setFontSize(13).setFontWeight('bold').setBackground('#dbeafe').setFontColor('#111827');
  return row + 1;
}
function mrParagraph_(sh, row, text) {
  sh.getRange(row, 1, 1, 13).merge().setValue(text).setFontSize(9).setBackground('#ffffff').setWrap(true);
  sh.setRowHeight(row, 42);
  return row + 2;
}
function mrTable_(sh, row, headers, rows) {
  var data = [headers].concat(rows || []);
  var range = sh.getRange(row, 1, data.length, headers.length);
  range.setValues(data).setWrap(true).setVerticalAlignment('top').setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, 1, 1, headers.length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
  if (data.length > 1) sh.getRange(row + 1, 1, data.length - 1, headers.length).setBackground('#ffffff').setFontColor('#111827');
  return row + data.length + 2;
}
function mrFinalize_(sh, lastRow) {
  for (var c = 1; c <= 13; c++) sh.setColumnWidth(c, c === 8 || c === 12 ? 210 : 115);
  sh.getRange(1, 1, Math.max(1, lastRow), 13).setFontFamily(MARKET_REPORT.font);
  sh.autoResizeRows(1, Math.max(1, lastRow));
}

function mrNum_(x) { if (x === null || x === undefined || x === '') return 0; if (typeof x === 'number') return x; var n = Number(String(x).replace(/[$,%()\s]/g, '').replace(/,/g, '')); return isNaN(n) ? 0 : n; }
function mrPct_(x) { if (x === null || x === undefined || x === '') return 0; if (typeof x === 'number') return Math.abs(x) > 1 ? x / 100 : x; var s = String(x); var n = mrNum_(s); return s.indexOf('%') >= 0 ? n / 100 : n; }
function mrMoney_(n) { n = Number(n || 0); return '$' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function mrParseMoney_(s) { return mrNum_(s); }
function mrPctText_(x) { return (Number(x || 0) * 100).toFixed(2) + '%'; }
function mrNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
