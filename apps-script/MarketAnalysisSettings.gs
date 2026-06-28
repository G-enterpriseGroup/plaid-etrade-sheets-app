/**
 * Portfolio Link Market Analysis settings helper.
 * Keeps the core MarketAnalysisReport.gs engine untouched.
 * Adds sidebar-controlled trim P&L method handling.
 * FIFO now applies to the existing report first, so changing the method does not rebuild the slow market cache.
 */

var MARKET_TRIM_PNL_METHOD_PROP = 'MARKET_TRIM_PNL_METHOD';
var MARKET_TRANSACTIONS_SHEET = 'Transactions';
var MARKET_REPORT_SHEET = 'Report Market Analysis';

function getMarketTrimPnlMethod() {
  return mrMarketNormalizeTrimPnlMethod_(PropertiesService.getDocumentProperties().getProperty(MARKET_TRIM_PNL_METHOD_PROP) || 'AVERAGE');
}

function setMarketTrimPnlMethod(method) {
  var value = mrMarketNormalizeTrimPnlMethod_(method);
  PropertiesService.getDocumentProperties().setProperty(MARKET_TRIM_PNL_METHOD_PROP, value);
  return value;
}

function buildMarketAnalysisReportFromSidebar(method) {
  setMarketTrimPnlMethod(method);

  // If a report already exists, changing FIFO/average only needs to update Est. P&L.
  // This avoids rerunning the slow GoogleFinance history engine just to switch P&L method.
  if (mrMarketHasPortfolioTable_()) {
    var updated = mrMarketApplyTrimPnlMethod_();
    return 'Trim P&L method applied to existing report. Method: ' + mrMarketTrimPnlMethodLabel_() + '. Rows updated: ' + updated + '. Use a full market refresh only when you need new market data.';
  }

  var message = buildMarketAnalysisReport();
  var updatedAfterBuild = mrMarketApplyTrimPnlMethod_();
  return message + ' Trim P&L method: ' + mrMarketTrimPnlMethodLabel_() + '. Rows updated: ' + updatedAfterBuild + '.';
}

function mrMarketNormalizeTrimPnlMethod_(method) {
  var v = String(method || '').toUpperCase().trim();
  return v === 'FIFO' ? 'FIFO' : 'AVERAGE';
}

function mrMarketTrimPnlMethodLabel_() {
  return getMarketTrimPnlMethod() === 'FIFO' ? 'FIFO using Transactions tax lots when available' : 'Evenly spread average-cost estimate';
}

function mrMarketHasPortfolioTable_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MARKET_REPORT_SHEET);
  if (!sh || sh.getLastRow() < 2) return false;
  var scanRows = Math.min(sh.getLastRow(), 220);
  var display = sh.getRange(1, 1, scanRows, Math.min(sh.getLastColumn(), 13)).getDisplayValues();
  for (var r = 0; r < display.length; r++) {
    if (display[r].indexOf('Ticker') >= 0 && display[r].indexOf('Exact Action') >= 0 && display[r].indexOf('Est. P&L') >= 0) return true;
  }
  return false;
}

function mrMarketApplyTrimPnlMethod_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MARKET_REPORT_SHEET);
  if (!sh) return 0;

  var lastRow = sh.getLastRow();
  var lastCol = Math.min(sh.getLastColumn(), 13);
  if (lastRow < 2 || lastCol < 2) return 0;

  var display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headerRow = -1;
  for (var r = 0; r < display.length; r++) {
    if (display[r].indexOf('Ticker') >= 0 && display[r].indexOf('Exact Action') >= 0 && display[r].indexOf('Est. P&L') >= 0) {
      headerRow = r + 1;
      break;
    }
  }
  if (headerRow < 1) return 0;

  var headers = display[headerRow - 1];
  var idx = {};
  headers.forEach(function(h, i) { idx[String(h || '').trim()] = i + 1; });
  var actionCol = idx['Exact Action'];
  var qtyCol = idx['Qty Held'];
  var totalPnlCol = idx['Unrealized $'];
  var estPnlCol = idx['Est. P&L'];
  var estValueCol = idx['Est. $ Value'];
  var reasonCol = idx['Reason / Price Source'];
  if (!actionCol || !qtyCol || !totalPnlCol || !estPnlCol || !estValueCol) return 0;

  var jobs = [];
  var neededTickers = {};
  for (var row = headerRow + 1; row <= lastRow; row++) {
    var rowValues = display[row - 1] || [];
    var ticker = String(rowValues[0] || '').trim().toUpperCase();
    var action = String(rowValues[actionCol - 1] || '').trim();
    if (!ticker || ticker.indexOf('TOTAL') === 0 || ticker.indexOf('AGGRESSIVE') === 0) break;
    if (action.indexOf('Trim-QTY') !== 0) continue;

    var trimQty = mrMarketActionQty_(action);
    if (!trimQty) continue;

    var heldQty = mrMarketNum_(rowValues[qtyCol - 1]);
    var totalPnl = mrMarketNum_(rowValues[totalPnlCol - 1]);
    var estValue = mrMarketNum_(rowValues[estValueCol - 1]);
    var sellPrice = trimQty ? estValue / trimQty : 0;
    var avgPnl = heldQty ? (totalPnl / heldQty) * trimQty : 0;

    jobs.push({row: row, ticker: ticker, trimQty: trimQty, sellPrice: sellPrice, avgPnl: avgPnl, reason: String(rowValues[reasonCol - 1] || '')});
    neededTickers[ticker] = true;
  }
  if (!jobs.length) return 0;

  var method = getMarketTrimPnlMethod();
  var lotsByTicker = method === 'FIFO' ? mrMarketBuildOpenLotsByTicker_(neededTickers) : {};
  var updated = 0;

  jobs.forEach(function(job) {
    var result = {pnl: job.avgPnl, note: 'Average method: estimated by spreading current unrealized P&L evenly across held shares.'};
    if (method === 'FIFO') result = mrMarketEstimateFifoTrimPnl_(job.ticker, job.trimQty, job.sellPrice, job.avgPnl, lotsByTicker[job.ticker] || []);

    sh.getRange(job.row, estPnlCol).setValue(mrMarketMoney_(result.pnl));
    if (reasonCol) {
      var cleanReason = mrMarketStripMethodNotes_(job.reason);
      sh.getRange(job.row, reasonCol).setValue(cleanReason + ' ' + result.note);
    }
    updated++;
  });

  return updated;
}

function mrMarketEstimateFifoTrimPnl_(ticker, trimQty, sellPrice, fallbackPnl, lots) {
  var q = Number(trimQty || 0);
  var pnl = 0;
  var matched = 0;
  var fifoLots = (lots || []).map(function(lot) { return {qty: lot.qty, costPerShare: lot.costPerShare}; });

  for (var i = 0; i < fifoLots.length && q > 0; i++) {
    var take = Math.min(q, fifoLots[i].qty);
    pnl += take * (sellPrice - fifoLots[i].costPerShare);
    matched += take;
    q -= take;
  }

  if (matched >= trimQty - 0.00001 && trimQty > 0) {
    return {
      pnl: pnl,
      note: 'FIFO method: Est. P&L uses oldest remaining priced buy lots from the Transactions tab first.'
    };
  }

  return {
    pnl: fallbackPnl,
    note: 'FIFO selected, but only ' + mrMarketQty_(matched) + ' of ' + mrMarketQty_(trimQty) + ' shares had priced FIFO lots in Transactions; used aggregate average estimate for this row.'
  };
}

function mrMarketBuildOpenLotsByTicker_(neededTickers) {
  var txs = mrMarketReadTransactionsForTickers_(neededTickers);
  var lotsByTicker = {};

  txs.forEach(function(tx) {
    if (!lotsByTicker[tx.ticker]) lotsByTicker[tx.ticker] = [];
    var lots = lotsByTicker[tx.ticker];

    if (tx.isBuy && tx.qtyAbs > 0 && tx.costPerShare > 0) {
      lots.push({qty: tx.qtyAbs, costPerShare: tx.costPerShare, date: tx.date});
      return;
    }

    if (tx.isSell && tx.qtyAbs > 0) {
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

function mrMarketReadTransactionsForTickers_(neededTickers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MARKET_TRANSACTIONS_SHEET);
  if (!sh) return [];

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return [];

  // One read only. This was the main FIFO performance fix.
  var values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = values[0].map(mrMarketNorm_);
  var ix = {};
  headers.forEach(function(h, i) { ix[h] = i; });

  function get(row, keys) {
    for (var k = 0; k < keys.length; k++) {
      var i = ix[keys[k]];
      if (i !== undefined) return row[i];
    }
    return '';
  }

  var out = [];
  values.slice(1).forEach(function(row, n) {
    var ticker = String(get(row, ['ticker_symbol','ticker','symbol']) || '').trim().toUpperCase();
    if (!ticker || !neededTickers[ticker]) return;
    if (!mrMarketIncludeAccount_(get(row, ['account_name','account','account_id']))) return;

    var type = String(get(row, ['type','transaction_type']) || '').toLowerCase();
    var subtype = String(get(row, ['subtype','transaction_subtype']) || '').toLowerCase();
    var name = String(get(row, ['name','description','transaction_name']) || '').toLowerCase();
    var qty = mrMarketNum_(get(row, ['quantity','qty','units','shares']));
    var qtyAbs = Math.abs(qty);
    if (!qtyAbs) return;

    var price = mrMarketNum_(get(row, ['price','unit_price','security_price','institution_price']));
    var amount = mrMarketNum_(get(row, ['amount','total_amount','value']));
    var fees = Math.abs(mrMarketNum_(get(row, ['fees','fee','commission','commissions'])));
    var date = mrMarketDate_(get(row, ['date','transaction_date','posted_date']));

    var text = type + ' ' + subtype + ' ' + name;
    var looksBuy = /\bbuy\b|\bbought\b|\bbot\b/.test(text);
    var looksSell = /\bsell\b|\bsold\b/.test(text);
    var isBuy = looksBuy && qtyAbs > 0;
    var isSell = looksSell && qtyAbs > 0;
    var costPerShare = 0;

    if (isBuy) {
      costPerShare = price > 0 ? price + (fees / qtyAbs) : (Math.abs(amount) / qtyAbs);
    }

    out.push({
      date: date,
      seq: n,
      ticker: ticker,
      qty: qty,
      qtyAbs: qtyAbs,
      price: price,
      amount: amount,
      fees: fees,
      isBuy: isBuy,
      isSell: isSell,
      costPerShare: costPerShare
    });
  });

  out.sort(function(a, b) {
    return a.date.getTime() - b.date.getTime() || a.seq - b.seq;
  });
  return out;
}

function mrMarketIncludeAccount_(accountName) {
  var s = String(accountName || '').toLowerCase();
  if (!s) return true;
  if (s.indexOf('baljinder') >= 0) return false;
  if (s.indexOf('parminder') >= 0) return false;
  if (s.indexOf('b + g') >= 0) return false;
  return true;
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

function mrMarketNorm_(x) {
  return String(x || '').trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
}

function mrMarketDate_(x) {
  var d = new Date(x);
  return isNaN(d.getTime()) ? new Date(0) : d;
}

function mrMarketNum_(x) {
  if (x === null || x === undefined || x === '') return 0;
  if (typeof x === 'number') return x;
  var s = String(x);
  var neg = s.indexOf('(') >= 0 && s.indexOf(')') >= 0;
  s = s.replace(/[,$%\s()]/g, '');
  var n = Number(s);
  if (isNaN(n)) return 0;
  return neg ? -n : n;
}

function mrMarketMoney_(n) {
  var sign = Number(n || 0) < 0 ? '-$' : '$';
  return sign + Math.abs(Number(n || 0)).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

function mrMarketQty_(n) {
  return Number(n || 0).toLocaleString('en-US', {maximumFractionDigits: 4});
}
