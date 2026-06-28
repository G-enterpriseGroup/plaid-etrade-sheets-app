/**
 * Portfolio Link Market Analysis settings helper.
 * Keeps the core MarketAnalysisReport.gs engine untouched.
 * Adds sidebar-controlled trim P&L method handling.
 */

var MARKET_TRIM_PNL_METHOD_PROP = 'MARKET_TRIM_PNL_METHOD';

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
  return getMarketTrimPnlMethod() === 'FIFO' ? 'FIFO / E*TRADE default selected' : 'Evenly spread average-cost estimate';
}

function mrMarketApplyTrimPnlMethod_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Report Market Analysis');
  if (!sh) return;

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return;

  var display = sh.getRange(1, 1, lastRow, Math.min(lastCol, 13)).getDisplayValues();
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
  var reasonCol = idx['Reason / Price Source'];
  if (!actionCol || !qtyCol || !totalPnlCol || !estPnlCol) return;

  var method = getMarketTrimPnlMethod();
  var methodNote = method === 'FIFO'
    ? ' Method: FIFO selected in sidebar. Exact FIFO tax-lot P&L requires lot-level data; current report uses available aggregate holdings estimate.'
    : ' Method: evenly spread average-cost estimate.';

  for (var row = headerRow + 1; row <= lastRow; row++) {
    var ticker = String(sh.getRange(row, 1).getDisplayValue() || '').trim();
    var action = String(sh.getRange(row, actionCol).getDisplayValue() || '').trim();
    if (!ticker || ticker.indexOf('Total') === 0 || ticker.indexOf('Aggressive') === 0) break;
    if (action.indexOf('Trim-QTY') !== 0) continue;

    var qtyMatch = action.match(/Trim-QTY\s+(\d+)/i);
    var trimQty = qtyMatch ? Number(qtyMatch[1]) : 0;
    var heldQty = mrMarketNum_(sh.getRange(row, qtyCol).getDisplayValue());
    var totalPnl = mrMarketNum_(sh.getRange(row, totalPnlCol).getDisplayValue());
    var estPnl = heldQty ? (totalPnl / heldQty) * trimQty : 0;

    sh.getRange(row, estPnlCol).setValue(mrMarketMoney_(estPnl));
    if (reasonCol) {
      var reason = String(sh.getRange(row, reasonCol).getDisplayValue() || '');
      reason = reason.replace(/ Method: FIFO selected in sidebar\.[\s\S]*?aggregate holdings estimate\./g, '').replace(/ Method: evenly spread average-cost estimate\./g, '');
      sh.getRange(row, reasonCol).setValue(reason + methodNote);
    }
  }
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
  return '$' + Number(n || 0).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
}
