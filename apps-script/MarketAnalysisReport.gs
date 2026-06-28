/**
 * Market Analysis Report module
 * Reads Holdings and creates Report Market Analysis.
 * Run: buildMarketAnalysisReport()
 * Debug: debugMarketAnalysisHoldingsHeaders()
 * Version: advanced-technical-engine-v4
 * Note: Google Apps Script is JavaScript, so the Python-style technical algorithm is implemented here in .gs.
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
  if (!holdings.length) throw new Error('No usable holdings found. Run debugMarketAnalysisHoldingsHeaders() and send the result.');

  var market = mrBuildMarketEngine_();
  var macroRows = mrBuildMacroRows_(market);
  var sectorRows = mrBuildSectorRows_(market);
  var technicalRows = mrBuildTechnicalRows_(market);
  var portfolioRows = mrBuildPortfolioRows_(holdings, market);

  var sh = ss.getSheetByName(MARKET_REPORT.reportSheet) || ss.insertSheet(MARKET_REPORT.reportSheet);
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), Math.min(13, sh.getMaxColumns()))
    .setFontFamily(MARKET_REPORT.font)
    .setFontSize(9)
    .setWrap(true)
    .setVerticalAlignment('top');

  var row = 1;
  row = mrTitle_(sh, row, 'Raj Market Rotation Report', 'Built ' + mrNow_() + '. Reads connected E*TRADE Holdings, excluding cash and money market. Technical engine uses Stooq daily prices, 20 EMA, 50 SMA, 200 SMA, MACD, and relative strength vs SPY.');

  row = mrSection_(sh, row, 'Macro Risk Dashboard');
  row = mrTable_(sh, row, ['Risk Area', 'Current Read', 'Risk Level', 'Portfolio Meaning'], macroRows);
  row = mrParagraph_(sh, row, 'Macro summary: the script uses market proxies for oil, bonds/yields, dollar, gold, and SPY trend. News, Fed calendar, CPI/jobs releases, and company-specific earnings still need live review before trading.');

  row = mrSection_(sh, row, 'Sector Rotation - SPDR Map');
  row = mrTable_(sh, row, ['ETF', 'Sector', 'Rotation Read', 'Flow Pressure', 'Action Bias'], sectorRows);
  row = mrParagraph_(sh, row, 'Sector summary: leadership requires trend confirmation plus relative strength versus SPY. A sector can be positive but still not lead if it is underperforming SPY.');

  row = mrSection_(sh, row, 'Technical Confirmation Snapshot');
  row = mrTable_(sh, row, ['ETF', 'Sector', 'Price vs 20 EMA', 'Price vs 50 SMA', 'Price vs 200 SMA', '20/50 Crossover', '50/200 Trend', 'MACD vs Signal', 'MACD Zero', 'RS vs SPY'], technicalRows);

  row = mrSection_(sh, row, 'Prior Week Recap + Forward Trend');
  row = mrParagraph_(sh, row, mrPriorWeekText_(market));

  row = mrSection_(sh, row, 'Raj Portfolio Impact - Exact Actions');
  row = mrTable_(sh, row, ['Ticker', 'Qty Held', 'Cost Basis', 'Current Value', 'Unrealized $', 'P&L %', 'Tolerance Status', 'Thesis', 'Exact Action', 'Action Qty', 'Est. $ Value', 'Reason / Price Source', 'Sleeve'], portfolioRows);

  row = mrSection_(sh, row, 'Total Suggested Trims and Primary Goal');
  row = mrTable_(sh, row, ['Metric', 'Value'], [
    ['Total suggested trims', '$0.00 pending full live macro/news confirmation'],
    ['Total suggested sells', '$0.00 pending full live macro/news confirmation'],
    ['Total suggested adds', '$0.00 pending full live macro/news confirmation'],
    ['Primary goal', 'Use macro + sector + technical confirmation before changing holdings. This engine now confirms trend/rotation; exact add/trim/sell sizing should remain controlled and reviewed.']
  ]);

  row = mrSection_(sh, row, 'Aggressive Growth Setup With Risk Controls');
  row = mrTable_(sh, row, ['Setup', 'Trigger', 'Risk Control', 'Action Bias'], [
    ['Core growth add', 'SPY/XLK above 20 EMA and 50 SMA with MACD > signal', 'Starter size only; do not add into weak macro', 'Watchlist / controlled add only after confirmation'],
    ['Profit protection', 'Large P&L outlier or overweight sleeve', 'Trim 5% to 15% only after live confirmation', 'Protect gains without forcing full exit'],
    ['Safety redeployment', 'Safety sleeve overweight and growth leadership confirmed', 'Keep liquidity buffer', 'Gradual shift only'],
    ['Avoid weak trend', 'Below 20 EMA/50 SMA with MACD < signal', 'No averaging down without macro confirmation', 'Avoid or hold-no-add']
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
    ['Stooq daily price data', 'Daily OHLC close data for SPY, sector ETFs, GLD, USO, TLT, UUP, and BIL'],
    ['SPDR sector ETF map', '11-sector rotation framework'],
    ['SPY benchmark', 'Relative-strength benchmark'],
    ['Manual live-news layer', 'Macro, Fed, inflation, earnings, and geopolitical catalysts']
  ]);

  mrFinalize_(sh, row);
  return 'Report Market Analysis tab built. Holdings analyzed: ' + holdings.length + '. Technical snapshot updated with EMA/SMA/MACD/RS.';
}

function mrBuildMarketEngine_() {
  var symbols = ['SPY', 'DIA', 'GLD', 'USO', 'TLT', 'UUP', 'BIL'];
  MARKET_SECTOR_ETFS.forEach(function(x) { symbols.push(x[0]); });
  var engine = {};
  symbols.forEach(function(sym) { engine[sym] = mrTechnicalForSymbol_(sym); });
  return engine;
}

function mrTechnicalForSymbol_(ticker) {
  try {
    var bars = mrFetchStooqDaily_(ticker);
    if (bars.length < 80) return mrBlankTech_(ticker, 'Not enough history');
    var closes = bars.map(function(b) { return b.close; });
    var price = closes[closes.length - 1];
    var ema20 = mrLast_(mrEmaSeries_(closes, 20));
    var sma50 = mrSma_(closes, 50);
    var sma200 = mrSma_(closes, 200);
    var macd = mrMacd_(closes);
    var chg5 = mrChange_(closes, 5);
    var chg20 = mrChange_(closes, 20);
    var score = 0;
    if (price >= ema20) score++;
    if (price >= sma50) score++;
    if (price >= sma200) score++;
    if (ema20 >= sma50) score++;
    if (sma50 >= sma200) score++;
    if (macd.macd >= macd.signal) score++;
    if (macd.macd >= 0) score++;
    if (chg20 > 0) score++;
    var read = score >= 7 ? 'Bullish confirmed' : score >= 5 ? 'Improving' : score >= 3 ? 'Mixed' : 'Weak';
    return {
      ticker: ticker,
      price: price,
      ema20: ema20,
      sma50: sma50,
      sma200: sma200,
      macd: macd.macd,
      signal: macd.signal,
      hist: macd.hist,
      change5d: chg5,
      change20d: chg20,
      score: score,
      priceVs20: price >= ema20 ? 'Above' : 'Below',
      priceVs50: price >= sma50 ? 'Above' : 'Below',
      priceVs200: price >= sma200 ? 'Above' : 'Below',
      cross20_50: ema20 >= sma50 ? '20 EMA > 50 SMA' : '20 EMA < 50 SMA',
      trend50_200: sma50 >= sma200 ? '50 SMA > 200 SMA' : '50 SMA < 200 SMA',
      macdSignal: macd.macd >= macd.signal ? 'MACD > Signal' : 'MACD < Signal',
      macdZero: macd.macd >= 0 ? 'Above 0' : 'Below 0',
      read: read,
      source: 'Stooq daily close'
    };
  } catch (err) {
    return mrBlankTech_(ticker, err.message || String(err));
  }
}

function mrFetchStooqDaily_(ticker) {
  var symbol = String(ticker || '').toLowerCase() + '.us';
  var url = 'https://stooq.com/q/d/l/?s=' + encodeURIComponent(symbol) + '&i=d';
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var text = res.getContentText();
  var lines = text.trim().split(/\r?\n/);
  if (lines.length < 3 || String(lines[0]).toLowerCase().indexOf('date') < 0) throw new Error('No Stooq data');
  var out = [];
  for (var i = 1; i < lines.length; i++) {
    var p = lines[i].split(',');
    if (p.length < 5) continue;
    var close = Number(p[4]);
    if (!isNaN(close)) out.push({ date: p[0], close: close });
  }
  return out;
}

function mrBlankTech_(ticker, note) {
  return { ticker: ticker, price: '', ema20: '', sma50: '', sma200: '', macd: '', signal: '', hist: '', change5d: '', change20d: '', score: 0, priceVs20: 'n/a', priceVs50: 'n/a', priceVs200: 'n/a', cross20_50: 'n/a', trend50_200: 'n/a', macdSignal: 'n/a', macdZero: 'n/a', read: 'Needs data', source: note || 'No data' };
}

function mrBuildMacroRows_(m) {
  var spy = m.SPY || mrBlankTech_('SPY', '');
  var uso = m.USO || mrBlankTech_('USO', '');
  var tlt = m.TLT || mrBlankTech_('TLT', '');
  var uup = m.UUP || mrBlankTech_('UUP', '');
  var gld = m.GLD || mrBlankTech_('GLD', '');
  return [
    ['War / Geopolitics', 'GLD 20d ' + mrPctText_(gld.change20d) + '; USO 20d ' + mrPctText_(uso.change20d), (gld.change20d > 0.03 || uso.change20d > 0.05) ? 'Elevated watch' : 'Normal watch', 'Gold/defense can matter if risk shock rises'],
    ['Oil / Energy shock', 'USO 20d ' + mrPctText_(uso.change20d), uso.change20d > 0.05 ? 'Oil pressure rising' : uso.change20d < -0.05 ? 'Oil easing' : 'Neutral', 'Affects XLE, inflation expectations, and consumer pressure'],
    ['Crisis / Liquidity', 'SPY ' + spy.priceVs50 + ' 50 SMA; SPY 20d ' + mrPctText_(spy.change20d), spy.priceVs50 === 'Below' ? 'Risk-off watch' : 'Risk-on acceptable', 'Controls how aggressive adds should be'],
    ['Fed / Yields', 'TLT 20d ' + mrPctText_(tlt.change20d), tlt.change20d < -0.04 ? 'Yield pressure rising' : tlt.change20d > 0.04 ? 'Yield pressure easing' : 'Neutral', 'Affects growth stocks, banks, real estate, and bond sleeves'],
    ['Inflation / Dollar', 'UUP 20d ' + mrPctText_(uup.change20d) + '; USO 20d ' + mrPctText_(uso.change20d), (uup.change20d > 0.03 || uso.change20d > 0.05) ? 'Tighter impulse' : 'Contained', 'Affects international ADRs, gold, and valuation multiples'],
    ['Jobs / Growth', 'SPY technical read: ' + spy.read, spy.score >= 5 ? 'Growth trend constructive' : 'Growth trend mixed/weak', 'Confirms whether rotation supports risk assets']
  ];
}

function mrBuildSectorRows_(m) {
  var spy = m.SPY || mrBlankTech_('SPY', '');
  return MARKET_SECTOR_ETFS.map(function(x) {
    var t = x[0], sector = x[1], tech = m[t] || mrBlankTech_(t, '');
    var rs = mrRsRead_(tech, spy);
    var rotation = mrRotationRead_(tech, spy);
    var flow = mrFlowRead_(tech, spy);
    var bias = mrActionBias_(rotation, tech);
    return [t, sector, rotation, flow + ' / ' + rs, bias];
  });
}

function mrBuildTechnicalRows_(m) {
  var spy = m.SPY || mrBlankTech_('SPY', '');
  return MARKET_SECTOR_ETFS.map(function(x) {
    var t = x[0], sector = x[1], tech = m[t] || mrBlankTech_(t, '');
    return [t, sector, tech.priceVs20, tech.priceVs50, tech.priceVs200, tech.cross20_50, tech.trend50_200, tech.macdSignal, tech.macdZero, mrRsRead_(tech, spy)];
  });
}

function mrRotationRead_(tech, spy) {
  if (tech.read === 'Bullish confirmed' && tech.change20d > spy.change20d) return 'Leadership';
  if (tech.score >= 6) return 'Positive trend';
  if (tech.score >= 4 && tech.change20d >= spy.change20d) return 'Improving';
  if (tech.score <= 2) return 'Lagging / weak';
  return 'Mixed';
}

function mrFlowRead_(tech, spy) {
  if (tech.change20d > spy.change20d + 0.02 && tech.score >= 5) return 'Positive pressure';
  if (tech.change20d < spy.change20d - 0.02 && tech.score <= 4) return 'Negative pressure';
  return 'Mixed pressure';
}

function mrRsRead_(tech, spy) {
  if (tech.change20d === '' || spy.change20d === '') return 'n/a';
  var diff = Number(tech.change20d) - Number(spy.change20d);
  if (diff > 0.02) return 'Outperforming SPY';
  if (diff < -0.02) return 'Underperforming SPY';
  return 'In line with SPY';
}

function mrActionBias_(rotation, tech) {
  if (rotation === 'Leadership') return 'Best add/hold candidates after macro confirmation';
  if (rotation === 'Positive trend') return 'Hold / selective add only';
  if (rotation === 'Improving') return 'Watchlist / starter only';
  if (rotation === 'Lagging / weak') return 'Avoid adds / review trims';
  return 'Hold / wait for confirmation';
}

function mrPriorWeekText_(m) {
  var spy = m.SPY || mrBlankTech_('SPY', '');
  var leaders = [];
  MARKET_SECTOR_ETFS.forEach(function(x) {
    var tech = m[x[0]] || mrBlankTech_(x[0], '');
    if (mrRotationRead_(tech, spy) === 'Leadership') leaders.push(x[0] + ' ' + x[1]);
  });
  return 'Forward trend: SPY is ' + spy.read + ' with 20-day change of ' + mrPctText_(spy.change20d) + '. Current sector leadership from the technical engine: ' + (leaders.length ? leaders.join(', ') : 'none confirmed') + '. Use this with live macro/news before taking action.';
}

function mrBuildPortfolioRows_(holdings, market) {
  return holdings.map(function(h) {
    var meta = mrMeta_(h);
    var sectorTicker = meta[1];
    var tech = market[sectorTicker] || market.SPY || mrBlankTech_(sectorTicker, 'No mapped technical data');
    var action = 'Hold';
    var tolerance = mrTolerance_(h, meta, tech);
    var reason = 'Mapped to ' + sectorTicker + ' (' + tech.read + ', score ' + tech.score + '/8). Price source: E*TRADE institution price/value. No add/trim/sell without live macro/news confirmation.';
    return [
      h.ticker,
      h.qty,
      mrMoney_(h.cost),
      mrMoney_(h.value),
      mrMoney_(h.pnl),
      mrPctText_(h.pnlPct),
      tolerance,
      meta[2],
      action,
      0,
      mrMoney_(0),
      reason,
      meta[0]
    ];
  }).sort(function(a, b) {
    return mrParseMoney_(b[3]) - mrParseMoney_(a[3]);
  });
}

function mrTolerance_(h, meta, tech) {
  if (meta[0].indexOf('Safety') >= 0 && h.weight > 0.25) return 'Overweight safety';
  if (meta[0].indexOf('Safety') >= 0) return 'Income tilt';
  if (h.pnlPct > 0.50) return 'Profit outlier';
  if (h.pnlPct < -0.10 && tech.score <= 3) return 'Weak';
  if (h.weight < 0.015) return 'Too small';
  if (h.weight > 0.08) return 'Near limit';
  if (tech.score <= 3) return 'Near limit';
  return 'In tolerance';
}

function debugMarketAnalysisHoldingsHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MARKET_REPORT.holdingsSheet);
  if (!sh) return 'No Holdings tab found.';
  var display = sh.getDataRange().getDisplayValues();
  var found = mrFindHeaderRow_(display);
  var rows = [];
  for (var i = 0; i < Math.min(8, display.length); i++) rows.push('Row ' + (i + 1) + ': ' + JSON.stringify(display[i].slice(0, 12)));
  var included = 0;
  try { included = mrReadHoldings_(ss).length; } catch (e) { included = 'ERROR: ' + e.message; }
  return 'Header found: ' + (found ? 'YES row ' + (found.headerRowIndex + 1) : 'NO, using fixed fallback') + '\nIncluded rows: ' + included + '\nFirst rows:\n' + rows.join('\n');
}

function mrReadHoldings_(ss) {
  var sh = ss.getSheetByName(MARKET_REPORT.holdingsSheet);
  if (!sh) throw new Error('Missing Holdings tab.');
  var raw = sh.getDataRange().getValues();
  var display = sh.getDataRange().getDisplayValues();
  if (display.length < 1) return [];
  var found = mrFindHeaderRow_(display);
  var ix = found ? found.index : mrStandardIndex_();
  var startRow = found ? found.headerRowIndex + 1 : 0;
  var rows = [];
  for (var r = startRow; r < display.length; r++) {
    var drow = mrNormalizeRowShape_(display[r]);
    var rrow = mrNormalizeRowShape_(raw[r]);
    if (mrLooksLikeHeaderRow_(drow)) continue;
    var ticker = mrCellByKey_(drow, ix, 'ticker_symbol').toUpperCase();
    var name = mrCellByKey_(drow, ix, 'security_name');
    var type = mrCellByKey_(drow, ix, 'security_type').toLowerCase();
    if (!ticker && !name) continue;
    if (mrExclude_(ticker, name, type)) continue;
    rows.push({
      ticker: ticker || '(no ticker)',
      name: name,
      type: type,
      subtype: mrCellByKey_(drow, ix, 'security_subtype'),
      qty: mrNum_(mrValueByKey_(rrow, drow, ix, 'quantity')),
      cost: mrNum_(mrValueByKey_(rrow, drow, ix, 'cost_basis')),
      price: mrNum_(mrValueByKey_(rrow, drow, ix, 'institution_price')),
      value: mrNum_(mrValueByKey_(rrow, drow, ix, 'institution_value')) || mrNum_(mrValueByKey_(rrow, drow, ix, 'calculated_market_value')),
      pnl: mrNum_(mrValueByKey_(rrow, drow, ix, 'unrealized_gain_loss')),
      pnlPct: mrPct_(mrValueByKey_(rrow, drow, ix, 'unrealized_gain_loss_pct')),
      weight: mrPct_(mrValueByKey_(rrow, drow, ix, 'portfolio_weight'))
    });
  }
  return rows;
}

function mrNormalizeRowShape_(row) {
  row = row || [];
  if (row.length && String(row[0] || '').indexOf('\t') >= 0) {
    var first = String(row[0] || '');
    var restBlank = true;
    for (var i = 1; i < row.length; i++) if (String(row[i] || '').trim()) restBlank = false;
    if (restBlank) return first.split('\t');
  }
  return row;
}

function mrFindHeaderRow_(displayValues) {
  var scanRows = Math.min(displayValues.length, 75);
  for (var r = 0; r < scanRows; r++) {
    var row = mrNormalizeRowShape_(displayValues[r]);
    var index = {};
    for (var c = 0; c < row.length; c++) {
      var h = mrNormHeader_(row[c]);
      if (h) index[h] = c;
    }
    var aliased = mrAliasIndex_(index);
    var score = 0;
    ['ticker_symbol', 'security_name', 'quantity', 'institution_value'].forEach(function(k) { if (aliased[k] !== undefined) score++; });
    if (score >= 2 && (aliased.ticker_symbol !== undefined || aliased.security_name !== undefined)) return { headerRowIndex: r, index: aliased };
  }
  return null;
}

function mrAliasIndex_(index) {
  var out = {};
  function pick(target, names) { for (var i = 0; i < names.length; i++) if (index[names[i]] !== undefined) { out[target] = index[names[i]]; return; } }
  pick('institution_name', ['institution_name', 'institution', 'brokerage']);
  pick('account_name', ['account_name', 'account']);
  pick('ticker_symbol', ['ticker_symbol', 'ticker', 'symbol']);
  pick('security_name', ['security_name', 'security', 'name', 'description']);
  pick('security_type', ['security_type', 'type']);
  pick('security_subtype', ['security_subtype', 'subtype']);
  pick('quantity', ['quantity', 'qty', 'shares']);
  pick('cost_basis', ['cost_basis', 'cost', 'basis']);
  pick('institution_price', ['institution_price', 'price', 'last_price', 'current_price']);
  pick('institution_value', ['institution_value', 'value', 'market_value', 'current_value']);
  pick('calculated_market_value', ['calculated_market_value', 'calculated_value']);
  pick('unrealized_gain_loss', ['unrealized_gain_loss', 'unrealized', 'gain_loss', 'p_l']);
  pick('unrealized_gain_loss_pct', ['unrealized_gain_loss_pct', 'p_l_pct', 'gain_loss_pct']);
  pick('portfolio_weight', ['portfolio_weight', 'weight']);
  return out;
}

function mrStandardIndex_() {
  return { institution_name: 0, account_name: 1, ticker_symbol: 6, security_name: 7, security_type: 8, security_subtype: 9, quantity: 10, cost_basis: 11, institution_price: 12, institution_value: 13, calculated_market_value: 14, unrealized_gain_loss: 15, unrealized_gain_loss_pct: 16, portfolio_weight: 17 };
}

function mrLooksLikeHeaderRow_(row) {
  var joined = row.map(function(x) { return mrNormHeader_(x); }).join('|');
  return joined.indexOf('ticker_symbol') >= 0 || joined.indexOf('security_name') >= 0 || joined.indexOf('institution_name') >= 0;
}

function mrNormHeader_(h) { return String(h || '').trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '').replace(/^_+|_+$/g, ''); }
function mrCellByKey_(row, ix, key) { var i = ix[key]; if (i === undefined || i < 0) return ''; var v = row[i]; return String(v === null || v === undefined ? '' : v).trim(); }
function mrValueByKey_(rawRow, displayRow, ix, key) { var i = ix[key]; if (i === undefined || i < 0) return ''; var raw = rawRow[i]; if (raw !== null && raw !== undefined && raw !== '') return raw; return displayRow[i]; }
function mrExclude_(ticker, name, type) { var n = String(name || '').toLowerCase(); if (type === 'cash') return true; if (ticker.indexOf('CUR:') === 0) return true; if (n.indexOf('money market') >= 0) return true; if (n.indexOf('sweep') >= 0) return true; if (['VMFXX','SWVXX','SPAXX','FDRXX'].indexOf(ticker) >= 0) return true; return false; }
function mrMeta_(h) { if (MARKET_TICKER_MAP[h.ticker]) return MARKET_TICKER_MAP[h.ticker]; if (h.type.indexOf('fixed') >= 0 || h.name.toLowerCase().indexOf('fdic') >= 0) return ['Fixed Income Safety', 'BIL', 'Principal/income stabilizer; watch safety overweight']; if (h.type.indexOf('etf') >= 0) return ['ETF / Unmapped', 'SPY', 'ETF exposure; map sleeve manually if material']; return ['Equity / Unmapped', 'SPY', 'Single-stock exposure; validate thesis manually']; }

function mrTitle_(sh, row, title, subtitle) { sh.getRange(row, 1, 1, 13).merge().setValue(title).setFontSize(18).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff'); row++; sh.getRange(row, 1, 1, 13).merge().setValue(subtitle).setFontSize(9).setFontColor('#374151').setBackground('#eef2ff'); return row + 2; }
function mrSection_(sh, row, title) { sh.getRange(row, 1, 1, 13).merge().setValue(title).setFontSize(13).setFontWeight('bold').setBackground('#dbeafe').setFontColor('#111827'); return row + 1; }
function mrParagraph_(sh, row, text) { sh.getRange(row, 1, 1, 13).merge().setValue(text).setFontSize(9).setBackground('#ffffff').setWrap(true); sh.setRowHeight(row, 42); return row + 2; }
function mrTable_(sh, row, headers, rows) { var data = [headers].concat(rows || []); var range = sh.getRange(row, 1, data.length, headers.length); range.setValues(data).setWrap(true).setVerticalAlignment('top').setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID); sh.getRange(row, 1, 1, headers.length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff'); if (data.length > 1) sh.getRange(row + 1, 1, data.length - 1, headers.length).setBackground('#ffffff').setFontColor('#111827'); return row + data.length + 2; }
function mrFinalize_(sh, lastRow) { for (var c = 1; c <= 13; c++) sh.setColumnWidth(c, c === 8 || c === 12 ? 210 : 115); sh.getRange(1, 1, Math.max(1, lastRow), 13).setFontFamily(MARKET_REPORT.font); sh.autoResizeRows(1, Math.max(1, lastRow)); }

function mrEmaSeries_(values, period) { var k = 2 / (period + 1), out = [], ema = values[0]; for (var i = 0; i < values.length; i++) { ema = i === 0 ? values[i] : values[i] * k + ema * (1 - k); out.push(ema); } return out; }
function mrMacd_(closes) { var e12 = mrEmaSeries_(closes, 12), e26 = mrEmaSeries_(closes, 26), macd = []; for (var i = 0; i < closes.length; i++) macd.push(e12[i] - e26[i]); var sig = mrEmaSeries_(macd, 9); var m = mrLast_(macd), s = mrLast_(sig); return { macd: m, signal: s, hist: m - s }; }
function mrSma_(arr, period) { var n = Math.min(arr.length, period), s = 0; for (var i = arr.length - n; i < arr.length; i++) s += arr[i]; return s / n; }
function mrLast_(arr) { return arr[arr.length - 1]; }
function mrChange_(arr, lookback) { if (arr.length <= lookback) return ''; var old = arr[arr.length - 1 - lookback], last = mrLast_(arr); return old ? last / old - 1 : ''; }
function mrNum_(x) { if (x === null || x === undefined || x === '') return 0; if (typeof x === 'number') return x; var s = String(x).replace(/[,$%\s]/g, '').replace(/[()]/g, ''); var n = Number(s); return isNaN(n) ? 0 : n; }
function mrPct_(x) { if (x === null || x === undefined || x === '') return 0; if (typeof x === 'number') return Math.abs(x) > 1 ? x / 100 : x; var s = String(x); var n = mrNum_(s); return s.indexOf('%') >= 0 ? n / 100 : n; }
function mrMoney_(n) { n = Number(n || 0); return '$' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function mrParseMoney_(s) { return mrNum_(s); }
function mrPctText_(x) { if (x === '' || x === null || x === undefined || isNaN(Number(x))) return 'n/a'; return (Number(x) * 100).toFixed(2) + '%'; }
function mrNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
