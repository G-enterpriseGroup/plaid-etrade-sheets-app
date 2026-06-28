/**
 * Portfolio Link Market Analysis settings helper.
 * Provides three sidebar actions:
 * 1) Save FIFO/Average method
 * 2) Apply method to existing report only
 * 3) Build full report, then apply saved method
 */

var MARKET_TRIM_PNL_METHOD_PROP = 'MARKET_TRIM_PNL_METHOD';
var MARKET_TRANSACTIONS_SHEET = 'Transactions';
var MARKET_REPORT_SHEET = 'Report Market Analysis';
var MARKET_TX_MAX_ROWS = 12000;

function getMarketTrimPnlMethod() {
  return mrMarketNormalizeTrimPnlMethod_(PropertiesService.getDocumentProperties().getProperty(MARKET_TRIM_PNL_METHOD_PROP) || 'AVERAGE');
}

function setMarketTrimPnlMethod(method) {
  var value = mrMarketNormalizeTrimPnlMethod_(method);
  PropertiesService.getDocumentProperties().setProperty(MARKET_TRIM_PNL_METHOD_PROP, value);
  return value;
}

function saveMarketTrimPnlMethod(method) {
  return 'Saved method: ' + mrMarketLabelForMethod_(setMarketTrimPnlMethod(method)) + '.';
}

// Backward-compatible old sidebar name. Keep it fast and apply-only.
function buildMarketAnalysisReportFromSidebar(method) {
  return applyMarketTrimPnlMethodFastV3(method);
}

function applyMarketTrimPnlMethodFastV3(method) {
  var saved = setMarketTrimPnlMethod(method);
  var result = mrMarketApplyTrimPnlMethodFast_(saved);
  if (!result.tableFound) return mrMarketResponse_('No Raj Portfolio Impact table found with Est. P&L. Use Build Full Report once first, then apply FIFO/Average.', result);
  if (!result.trimRows) return mrMarketResponse_('Method saved as ' + mrMarketLabelForMethod_(saved) + ', but there are no Trim-QTY rows to update.', result);
  return mrMarketResponse_('Done. Method: ' + mrMarketLabelForMethod_(saved) + '. Trim rows updated: ' + result.updated + '. This did not rebuild market data.', result);
}

function buildFullMarketAnalysisWithSavedMethod(method) {
  var saved = setMarketTrimPnlMethod(method);
  var msg = buildMarketAnalysisReport();
  var result = mrMarketApplyTrimPnlMethodFast_(saved);
  return mrMarketResponse_(msg + ' Method: ' + mrMarketLabelForMethod_(saved) + '. Trim rows updated: ' + result.updated + '.', result);
}

function mrMarketResponse_(message, result) {
  result = result || {};
  var details = result.fallbackDetails || [];
  return {
    message: message,
    fallbackCount: details.length,
    fallbackTickers: details.map(function(x) { return x.ticker; }).join(', '),
    fallbackDetails: details,
    logText: details.length ? mrMarketFallbackLogText_(details) : ''
  };
}

function mrMarketFallbackLogText_(details) {
  var lines = ['FIFO fallback explanation:'];
  details.forEach(function(d) {
    lines.push('- ' + d.ticker + ': used average estimate because ' + d.reason + ' Matched ' + mrMarketQty_(d.matchedQty) + ' of ' + mrMarketQty_(d.neededQty) + ' trim shares from priced FIFO lots.');
  });
  return lines.join('\n');
}

function mrMarketNormalizeTrimPnlMethod_(method) {
  var v = String(method || '').toUpperCase().trim();
  return v === 'FIFO' ? 'FIFO' : 'AVERAGE';
}

function mrMarketLabelForMethod_(method) {
  return mrMarketNormalizeTrimPnlMethod_(method) === 'FIFO' ? 'FIFO using Transactions tax lots' : 'Evenly spread average-cost estimate';
}

function mrMarketApplyTrimPnlMethodFast_(method) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MARKET_REPORT_SHEET);
  if (!sh) return {tableFound:false, trimRows:0, updated:0, fallbackDetails:[]};

  var lastRow = Math.min(sh.getLastRow(), 400);
  var lastCol = Math.min(sh.getLastColumn(), 13);
  if (lastRow < 2 || lastCol < 2) return {tableFound:false, trimRows:0, updated:0, fallbackDetails:[]};

  var display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headerRow = -1;
  for (var r = 0; r < display.length; r++) {
    if (display[r].indexOf('Ticker') >= 0 && display[r].indexOf('Exact Action') >= 0 && display[r].indexOf('Est. P&L') >= 0) {
      headerRow = r + 1;
      break;
    }
  }
  if (headerRow < 1) return {tableFound:false, trimRows:0, updated:0, fallbackDetails:[]};

  var headers = display[headerRow - 1];
  var idx = {};
  headers.forEach(function(h, i) { idx[String(h || '').trim()] = i + 1; });

  var actionCol = idx['Exact Action'];
  var qtyCol = idx['Qty Held'];
  var totalPnlCol = idx['Unrealized $'];
  var estPnlCol = idx['Est. P&L'];
  var estValueCol = idx['Est. $ Value'];
  var reasonCol = idx['Reason / Price Source'];
  if (!actionCol || !qtyCol || !totalPnlCol || !estPnlCol || !estValueCol) return {tableFound:false, trimRows:0, updated:0, fallbackDetails:[]};

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

  if (!jobs.length) return {tableFound:true, trimRows:0, updated:0, fallbackDetails:[]};

  var lotsByTicker = {};
  if (method === 'FIFO') lotsByTicker = mrMarketBuildOpenLotsByTickerFast_(neededTickers);

  var fallbackDetails = [];
  jobs.forEach(function(job) {
    var result = {pnl: job.avgPnl, note: 'Average method: estimated by spreading current unrealized P&L evenly across held shares.', fallback:false};
    if (method === 'FIFO') result = mrMarketEstimateFifoTrimPnl_(job.ticker, job.trimQty, job.sellPrice, job.avgPnl, lotsByTicker[job.ticker] || []);
    sh.getRange(job.row, estPnlCol).setValue(mrMarketMoney_(result.pnl));
    if (reasonCol) sh.getRange(job.row, reasonCol).setValue(mrMarketStripMethodNotes_(job.reason) + ' ' + result.note);
    if (result.fallback) {
      fallbackDetails.push({
        ticker: job.ticker,
        neededQty: job.trimQty,
        matchedQty: result.matchedQty || 0,
        reason: result.reason || 'there were not enough priced FIFO buy lots in Transactions.'
      });
    }
  });

  return {tableFound:true, trimRows:jobs.length, updated:jobs.length, fallbackDetails:fallbackDetails};
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

  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  var out = [];

  values.forEach(function(row, n) {
    var ticker = String(row[cols.ticker - 1] || '').trim().toUpperCase();
    if (!ticker || !neededTickers[ticker]) return;
    if (cols.account && !mrMarketIncludeAccount_(row[cols.account - 1])) return;

    var type = cols.type ? String(row[cols.type - 1] || '').toLowerCase() : '';
    var subtype = cols.subtype ? String(row[cols.subtype - 1] || '').toLowerCase() : '';
    var name = cols.name ? String(row[cols.name - 1] || '').toLowerCase() : '';
    var text = type + ' ' + subtype + ' ' + name;
    var qtyRaw = mrMarketNum_(row[cols.qty - 1]);
    var qtyAbs = Math.abs(qtyRaw);
    if (!qtyAbs) return;

    var hasBuyText = /\bbuy\b|\bbought\b|\bbot\b|\bpurchase\b|\breinvest/.test(text);
    var hasSellText = /\bsell\b|\bsold\b|\bsld\b/.test(text);
    var isBuy = hasBuyText && !hasSellText;
    var isSell = hasSellText && !hasBuyText;
    if (!isBuy && !isSell) {
      isBuy = qtyRaw > 0;
      isSell = qtyRaw < 0;
    }
    if (!isBuy && !isSell) return;

    var price = cols.price ? mrMarketNum_(row[cols.price - 1]) : 0;
    var amount = cols.amount ? mrMarketNum_(row[cols.amount - 1]) : 0;
    var fees = cols.fees ? Math.abs(mrMarketNum_(row[cols.fees - 1])) : 0;
    var date = cols.date ? mrMarketDate_(row[cols.date - 1]) : new Date(0);
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
    return {pnl: pnl, note: 'FIFO method: Est. P&L uses oldest remaining priced buy lots from the Transactions tab first.', fallback:false};
  }
  var reason = matched <= 0 ? 'no priced FIFO buy lots were found for this ticker in the Transactions tab.' : 'not enough priced FIFO buy lots were found for the requested trim quantity.';
  return {
    pnl: fallbackPnl,
    note: 'FIFO fallback: used average estimate because ' + reason + ' Matched ' + mrMarketQty_(matched) + ' of ' + mrMarketQty_(trimQty) + ' shares.',
    fallback:true,
    matchedQty: matched,
    reason: reason
  };
}

function mrMarketIncludeAccount_(accountName) {
  var s = String(accountName || '').toLowerCase();
  return !(s.indexOf('baljinder') >= 0 || s.indexOf('parminder') >= 0 || s.indexOf('b + g') >= 0);
}
function mrMarketActionQty_(action) { var m = String(action || '').match(/Trim-QTY\s+([0-9.]+)/i); return m ? Number(m[1]) : 0; }
function mrMarketStripMethodNotes_(reason) {
  return String(reason || '')
    .replace(/\s*FIFO method: Est\. P&L uses oldest remaining priced buy lots from the Transactions tab first\./g, '')
    .replace(/\s*FIFO fallback: used average estimate because [\s\S]*? shares\./g, '')
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
