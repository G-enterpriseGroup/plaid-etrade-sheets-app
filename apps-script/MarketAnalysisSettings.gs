/**
 * Portfolio Link Market Analysis settings helper.
 * Keeps the core MarketAnalysisReport.gs engine untouched.
 * Adds sidebar-controlled trim P&L method handling.
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
  var message = buildMarketAnalysisReport();
  mrMarketApplyTrimPnlMethod_();
  return message + ' Trim P&L method: ' + mrMarketTrimPnlMethodLabel_() + '.';
}

function mrMarketNormalizeTrimPnlMethod_(method) {
  var v = String(method || '').toUpperCase().trim();
  return v === 'FIFO' ? 'FIFO' : 'AVERAGE';
}

function mrMarketTrimPnlMethodLabel_() {
  return getMarketTrimPnlMethod() === 'FIFO' ? 'FIFO using Transactions tax lots when available' : 'Evenly spread average-cost estimate';
}

function mrMarketApplyTrimPnlMethod_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MARKET_REPORT_SHEET);
  if (!sh) return;

  var lastRow = sh.getLastRow();
  var lastCol = Math.min(sh.getLastColumn(), 13);
  if (lastRow < 2 || lastCol < 2) return;

  var display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headerRow = -1;
  for (var r = 0; r < display.length; r++) {
    if (display[r].indexOf('Ticker') >= 0 && display[r].indexOf('Exact Action') >= 0 && display[r].indexOf('Est. P&L') >= 0) {
      headerRow = r + 1;
      break;
    }
  }
  if (headerRow < 1) return;

  var headers = display[headerRow - 1];
  var idx = {};
  headers.forEach(function(h, i) { idx[String(h || '').trim()] = i + 1; });
  var actionCol = idx['Exact Action'];
  var qtyCol = idx['Qty Held'];
  var totalPnlCol = idx['Unrealized $'];
  var estPnlCol = idx['Est. P&L'];
  var estValueCol = idx['Est. $ Value'];
  var reasonCol = idx['Reason / Price Source'];
  if (!actionCol || !qtyCol || !totalPnlCol || !estPnlCol || !estValueCol) return;

  var method = getMarketTrimPnlMethod();
  for (var row = headerRow + 1; row <= lastRow; row++) {
    var ticker = String(sh.getRange(row, 1).getDisplayValue() || '').trim().toUpperCase();
    var action = String(sh.getRange(row, actionCol).getDisplayValue() || '').trim();
    if (!ticker || ticker.indexOf('TOTAL') === 0 || ticker.indexOf('AGGRESSIVE') === 0) break;
    if (action.indexOf('Trim-QTY') !== 0) continue;

    var trimQty = mrMarketActionQty_(action);
    if (!trimQty) continue;

    var heldQty = mrMarketNum_(sh.getRange(row, qtyCol).getDisplayValue());
    var totalPnl = mrMarketNum_(sh.getRange(row, totalPnlCol).getDisplayValue());
    var estValue = mrMarketNum_(sh.getRange(row, estValueCol).getDisplayValue());
    var sellPrice = trimQty ? estValue / trimQty : 0;
    var avgPnl = heldQty ? (totalPnl / heldQty) * trimQty : 0;
    var result = {pnl: avgPnl, note: 'Average method: estimated by spreading current unrealized P&L evenly across held shares.'};

    if (method === 'FIFO') {
      result = mrMarketEstimateFifoTrimPnl_(ticker, trimQty, sellPrice, avgPnl);
    }

    sh.getRange(row, estPnlCol).setValue(mrMarketMoney_(result.pnl));
    if (reasonCol) {
      var reason = String(sh.getRange(row, reasonCol).getDisplayValue() || '');
      reason = mrMarketStripMethodNotes_(reason);
      sh.getRange(row, reasonCol).setValue(reason + ' ' + result.note);
    }
  }
}

function mrMarketEstimateFifoTrimPnl_(ticker, trimQty, sellPrice, fallbackPnl) {
  var lotsData = mrMarketBuildOpenLots_(ticker);
  var lots = lotsData.lots || [];
  var q = Number(trimQty || 0);
  var pnl = 0;
  var matched = 0;

  for (var i = 0; i < lots.length && q > 0; i++) {
    var take = Math.min(q, lots[i].qty);
    pnl += take * (sellPrice - lots[i].costPerShare);
    matched += take;
    q -= take;
  }

  if (matched >= trimQty - 0.00001 && trimQty > 0) {
    return {
      pnl: pnl,
      note: 'FIFO method: Est. P&L uses oldest remaining priced buy lots from the Transactions tab first.'
    };
  }

  var missing = trimQty - matched;
  return {
    pnl: fallbackPnl,
    note: 'FIFO selected, but only ' + mrMarketQty_(matched) + ' of ' + mrMarketQty_(trimQty) + ' shares had priced FIFO lots in Transactions; used aggregate average estimate for this row.'
  };
}

function mrMarketBuildOpenLots_(ticker) {
  var txs = mrMarketReadTransactions_(ticker);
  var lots = [];
  txs.forEach(function(tx) {
    if (tx.isBuy && tx.qty > 0 && tx.costPerShare > 0) {
      lots.push({qty: tx.qty, costPerShare: tx.costPerShare, date: tx.date});
      return;
    }
    if (tx.isSell && tx.qty < 0) {
      var sellQty = Math.abs(tx.qty);
      while (sellQty > 0 && lots.length) {
        var take = Math.min(sellQty, lots[0].qty);
        lots[0].qty -= take;
        sellQty -= take;
        if (lots[0].qty <= 0.00001) lots.shift();
      }
    }
  });
  return {lots: lots, transactions: txs.length};
}

function mrMarketReadTransactions_(ticker) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MARKET_TRANSACTIONS_SHEET);
  if (!sh) return [];
  var values = sh.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return [];

  var headers = values[0].map(mrMarketNorm_);
  var ix = {};
  headers.forEach(function(h, i) { ix[h] = i; });
  function get(row, key) { var i = ix[key]; return i === undefined ? '' : row[i]; }

  var out = [];
  values.slice(1).forEach(function(row, n) {
    var t = String(get(row, 'ticker_symbol') || '').trim().toUpperCase();
    if (t !== ticker) return;
    if (!mrMarketIncludeAccount_(get(row, 'account_name'))) return;

    var type = String(get(row, 'type') || '').toLowerCase();
    var subtype = String(get(row, 'subtype') || '').toLowerCase();
    var name = String(get(row, 'name') || '').toLowerCase();
    var qty = mrMarketNum_(get(row, 'quantity'));
    var price = mrMarketNum_(get(row, 'price'));
    var amount = mrMarketNum_(get(row, 'amount'));
    var fees = Math.abs(mrMarketNum_(get(row, 'fees')));
    var date = mrMarketDate_(get(row, 'date'));

    var looksBuy = type === 'buy' || subtype === 'buy' || name.indexOf('bot ') === 0 || name.indexOf('bought ') === 0;
    var looksSell = type === 'sell' || subtype === 'sell' || name.indexOf('sold ') === 0;
    var isBuy = qty > 0 && price > 0 && looksBuy;
    var isSell = qty < 0 && looksSell;
    var costPerShare = 0;

    if (isBuy) {
      costPerShare = amount > 0 && qty > 0 ? amount / qty : price + (qty ? fees / qty : 0);
    }

    out.push({
      date: date,
      seq: n,
      ticker: t,
      qty: qty,
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
