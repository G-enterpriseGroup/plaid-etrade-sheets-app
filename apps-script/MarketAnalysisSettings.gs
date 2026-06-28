/**
 * Portfolio Link Market Analysis settings helper.
 * This file does NOT rebuild the market report.
 * It only switches the Est. P&L method on the existing Report Market Analysis table.
 */

var MARKET_TRIM_PNL_METHOD_PROP = 'MARKET_TRIM_PNL_METHOD';
var MARKET_TRANSACTIONS_SHEET = 'Transactions';
var MARKET_REPORT_SHEET = 'Report Market Analysis';
var MARKET_TX_MAX_ROWS = 5000;

function getMarketTrimPnlMethod() {
  return mrMarketNormalizeTrimPnlMethod_(PropertiesService.getDocumentProperties().getProperty(MARKET_TRIM_PNL_METHOD_PROP) || 'AVERAGE');
}

function setMarketTrimPnlMethod(method) {
  var value = mrMarketNormalizeTrimPnlMethod_(method);
  PropertiesService.getDocumentProperties().setProperty(MARKET_TRIM_PNL_METHOD_PROP, value);
  return value;
}

// Old sidebar compatibility. Important: this no longer builds the slow full market report.
function buildMarketAnalysisReportFromSidebar(method) {
  return applyMarketTrimPnlMethodFastV3(method);
}

// New fast sidebar entrypoint. It only edits Est. P&L on the existing report.
function applyMarketTrimPnlMethodFastV3(method) {
  setMarketTrimPnlMethod(method);
  var methodNow = getMarketTrimPnlMethod();
  var result = mrMarketApplyTrimPnlMethodFast_(methodNow);
  if (!result.tableFound) {
    return 'No Raj Portfolio Impact table found with Est. P&L. Run the base Market Analysis report once first, then use FIFO/Average.';
  }
  if (!result.trimRows) {
    return 'Method saved as ' + mrMarketTrimPnlMethodLabel_() + ', but there are no Trim-QTY rows to update.';
  }
  return 'Done. Method: ' + mrMarketTrimPnlMethodLabel_() + '. Trim rows updated: ' + result.updated + '. This did not rebuild the full market report.';
}

function mrMarketNormalizeTrimPnlMethod_(method) {
  var v = String(method || '').toUpperCase().trim();
  return v === 'FIFO' ? 'FIFO' : 'AVERAGE';
}

function mrMarketTrimPnlMethodLabel_() {
  return getMarketTrimPnlMethod() === 'FIFO' ? 'FIFO using Transactions tax lots when available' : 'Evenly spread average-cost estimate';
}

function mrMarketApplyTrimPnlMethodFast_(method) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MARKET_REPORT_SHEET);
  if (!sh) return {tableFound:false, trimRows:0, updated:0};

  var lastRow = Math.min(sh.getLastRow(), 300);
  var lastCol = Math.min(sh.getLastColumn(), 13);
  if (lastRow < 2 || lastCol < 2) return {tableFound:false, trimRows:0, updated:0};

  var display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headerRow = -1;
  for (var r = 0; r < display.length; r++) {
    if (display[r].indexOf('Ticker') >= 0 && display[r].indexOf('Exact Action') >= 0) {
      headerRow = r + 1;
      break;
    }
  }
  if (headerRow < 1) return {tableFound:false, trimRows:0, updated:0};

  var headers = display[headerRow - 1];
  var idx = {};
  headers.forEach(function(h, i) { idx[String(h || '').trim()] = i + 1; });

  var actionCol = idx['Exact Action'];
  var qtyCol = idx['Qty Held'];
  var totalPnlCol = idx['Unrealized $'];
  var estPnlCol = idx['Est. P&L'];
  var estValueCol = idx['Est. $ Value'];
  var reasonCol = idx['Reason / Price Source'];
  if (!actionCol || !qtyCol || !totalPnlCol || !estPnlCol || !estValueCol) return {tableFound:false, trimRows:0, updated:0};

  var jobs = [];
  var neededTickers = {};
  for (var row = headerRow + 1; row <= lastRow; row++) {
    var rowValues = display[row - 1] || [];
    var ticker = String(rowValues[0] || '').trim().toUpperCase();
    var action = String(rowValues[actionCol - 1] || '').trim();

    if (!ticker) break;
    if (ticker === 'METRIC' || ticker.indexOf('TOTAL') === 0 || ticker.indexOf('AGGRESSIVE') === 0) break;
    if (action.indexOf('Trim-QTY') !== 0) continue;

    var trimQty = mrMarketActionQty_(action);
    if (!trimQty) continue;

    var heldQty = mrMarketNum_(rowValues[qtyCol - 1]);
    var totalPnl = mrMarketNum_(rowValues[totalPnlCol - 1]);
    var estValue = mrMarketNum_(rowValues[estValueCol - 1]);
    var sellPrice = trimQty ? estValue / trimQty : 0;
    var avgPnl = heldQty ? (totalPnl / heldQty) * trimQty : 0;

    jobs.push({row: row, ticker: ticker, trimQty: trimQty, sellPrice: sellPrice, avgPnl: avgPnl, reason: reasonCol ? String(rowValues[reasonCol - 1] || '') : ''});
    neededTickers[ticker] = true;
  }

  if (!jobs.length) return {tableFound:true, trimRows:0, updated:0};

  var lotsByTicker = {};
  if (method === 'FIFO') lotsByTicker = mrMarketBuildOpenLotsByTickerFast_(neededTickers);

  var estPnlValues = [];
  var reasonValues = [];
  jobs.forEach(function(job) {
    var result = {pnl: job.avgPnl, note: 'Average method: estimated by spreading current unrealized P&L evenly across held shares.'};
    if (method === 'FIFO') result = mrMarketEstimateFifoTrimPnl_(job.ticker, job.trimQty, job.sellPrice, job.avgPnl, lotsByTicker[job.ticker] || []);
    estPnlValues.push({row: job.row, value: mrMarketMoney_(result.pnl)});
    if (reasonCol) reasonValues.push({row: job.row, value: mrMarketStripMethodNotes_(job.reason) + ' ' + result.note});
  });

  estPnlValues.forEach(function(x) { sh.getRange(x.row, estPnlCol).setValue(x.value); });
  if (reasonCol) reasonValues.forEach(function(x) { sh.getRange(x.row, reasonCol).setValue(x.value); });

  return {tableFound:true, trimRows:jobs.length, updated:jobs.length};
}

function mrMarketBuildOpenLotsByTickerFast_(neededTickers) {
  var txs = mrMarketReadTransactionsForTickersFast_(neededTickers);
  var lotsByTicker = {};
  txs.forEach(function(tx) {
    if (!lotsByTicker[tx.ticker]) lotsByTicker[tx.ticker] = [];
    var lots = lotsByTicker[tx.ticker];
    if (tx.isBuy && tx.qtyAbs > 0 && tx.costPerShare > 0) {
      lots.push({qty: tx.qtyAbs, costPerShare: tx.costPerShare});
    } else if (tx.isSell && tx.qtyAbs > 0) {
      var sellQty = tx.qtyAbs;
      while (sellQty > 0 && lots.length) {
        var take = Math.min(sellQty, lots[0].qty);
        lots[0].qty -= take;
        sellQty -= take;
        if (lots[0].qty <= 0.00001) lots.shift();
      }
    }
  });
  return lotsByTicker;
}

function mrMarketReadTransactionsForTickersFast_(neededTickers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MARKET_TRANSACTIONS_SHEET);
  if (!sh) return [];

  var lastRow = Math.min(sh.getLastRow(), MARKET_TX_MAX_ROWS + 1);
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return [];

  var header = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var headers = header.map(mrMarketNorm_);
  var ix = {};
  headers.forEach(function(h, i) { ix[h] = i; });

  function col(keys) {
    for (var k = 0; k < keys.length; k++) if (ix[keys[k]] !== undefined) return ix[keys[k]] + 1;
    return 0;
  }

  var cols = {
    ticker: col(['ticker_symbol','ticker','symbol']),
    type: col(['type','transaction_type']),
    subtype: col(['subtype','transaction_subtype']),
    name: col(['name','description','transaction_name']),
    qty: col(['quantity','qty','units','shares']),
    price: col(['price','unit_price','security_price','institution_price']),
    amount: col(['amount','total_amount','value']),
    fees: col(['fees','fee','commission','commissions']),
    date: col(['date','transaction_date','posted_date']),
    account: col(['account_name','account','account_id'])
  };

  if (!cols.ticker || !cols.qty) return [];

  var width = lastCol;
  var values = sh.getRange(2, 1, lastRow - 1, width).getDisplayValues();
  var out = [];

  values.forEach(function(row, n) {
    var ticker = String(row[cols.ticker - 1] || '').trim().toUpperCase();
    if (!ticker || !neededTickers[ticker]) return;
    if (cols.account && !mrMarketIncludeAccount_(row[cols.account - 1])) return;

    var type = cols.type ? String(row[cols.type - 1] || '').toLowerCase() : '';
    var subtype = cols.subtype ? String(row[cols.subtype - 1] || '').toLowerCase() : '';
    var name = cols.name ? String(row[cols.name - 1] || '').toLowerCase() : '';
    var qty = mrMarketNum_(row[cols.qty - 1]);
    var qtyAbs = Math.abs(qty);
    if (!qtyAbs) return;

    var price = cols.price ? mrMarketNum_(row[cols.price - 1]) : 0;
    var amount = cols.amount ? mrMarketNum_(row[cols.amount - 1]) : 0;
    var fees = cols.fees ? Math.abs(mrMarketNum_(row[cols.fees - 1])) : 0;
    var date = cols.date ? mrMarketDate_(row[cols.date - 1]) : new Date(0);
    var text = type + ' ' + subtype + ' ' + name;
    var looksBuy = /\bbuy\b|\bbought\b|\bbot\b/.test(text) || qty > 0;
    var looksSell = /\bsell\b|\bsold\b/.test(text) || qty < 0;
    var isBuy = looksBuy && qtyAbs > 0;
    var isSell = looksSell && qtyAbs > 0 && !isBuy;
    var costPerShare = isBuy ? (price > 0 ? price + (fees / qtyAbs) : (Math.abs(amount) / qtyAbs)) : 0;

    out.push({date: date, seq: n, ticker: ticker, qtyAbs: qtyAbs, isBuy: isBuy, isSell: isSell, costPerShare: costPerShare});
  });

  out.sort(function(a, b) { return a.date.getTime() - b.date.getTime() || a.seq - b.seq; });
  return out;
}

function mrMarketEstimateFifoTrimPnl_(ticker, trimQty, sellPrice, fallbackPnl, lots) {
  var q = Number(trimQty || 0), pnl = 0, matched = 0;
  var fifoLots = (lots || []).map(function(lot) { return {qty: lot.qty, costPerShare: lot.costPerShare}; });
  for (var i = 0; i < fifoLots.length && q > 0; i++) {
    var take = Math.min(q, fifoLots[i].qty);
    pnl += take * (sellPrice - fifoLots[i].costPerShare);
    matched += take;
    q -= take;
  }
  if (matched >= trimQty - 0.00001 && trimQty > 0) {
    return {pnl: pnl, note: 'FIFO method: Est. P&L uses oldest remaining priced buy lots from the Transactions tab first.'};
  }
  return {pnl: fallbackPnl, note: 'FIFO selected, but only ' + mrMarketQty_(matched) + ' of ' + mrMarketQty_(trimQty) + ' shares had priced FIFO lots in Transactions; used aggregate average estimate for this row.'};
}

function mrMarketIncludeAccount_(accountName) {
  var s = String(accountName || '').toLowerCase();
  return !(s.indexOf('baljinder') >= 0 || s.indexOf('parminder') >= 0 || s.indexOf('b + g') >= 0);
}

function mrMarketActionQty_(action) {
  var m = String(action || '').match(/Trim-QTY\s+([0-9.]+)/i);
  return m ? Number(m[1]) : 0;
}

function mrMarketStripMethodNotes_(reason) {
  return String(reason || '')
    .replace(/\s*FIFO method: Est\. P&L uses oldest remaining priced buy lots from the Transactions tab first\./g, '')
    .replace(/\s*FIFO selected, but only [\s\S]*?used aggregate average estimate for this row\./g, '')
    .replace(/\s*Average method: estimated by spreading current unrealized P&L evenly across held shares\./g, '')
    .trim();
}

function mrMarketNorm_(x) { return String(x || '').trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, ''); }
function mrMarketDate_(x) { var d = new Date(x); return isNaN(d.getTime()) ? new Date(0) : d; }
function mrMarketNum_(x) {
  if (x === null || x === undefined || x === '') return 0;
  if (typeof x === 'number') return x;
  var s = String(x), neg = s.indexOf('(') >= 0 && s.indexOf(')') >= 0;
  s = s.replace(/[,$%\s()]/g, '');
  var n = Number(s);
  return isNaN(n) ? 0 : (neg ? -n : n);
}
function mrMarketMoney_(n) { var sign = Number(n || 0) < 0 ? '-$' : '$'; return sign + Math.abs(Number(n || 0)).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}); }
function mrMarketQty_(n) { return Number(n || 0).toLocaleString('en-US', {maximumFractionDigits: 4}); }
