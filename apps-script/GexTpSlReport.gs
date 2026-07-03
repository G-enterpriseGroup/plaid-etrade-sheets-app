/**
 * Portfolio Link - GEX Take Profit & Stop Limit
 * Builds a new sheet named "GEX Take Profit & Stop Limit" from Holdings tickers.
 * Pulls delayed option-chain data from Cboe, estimates dealer gamma exposure,
 * and maps Call Wall / Put Wall / Gamma Flip into TP and SL guide prices.
 */

var GEX_TPSL = {
  holdingsSheet: 'Holdings',
  reportSheet: 'GEX Take Profit & Stop Limit',
  font: 'Times New Roman',
  maxTickers: 40,
  maxDte: 45,
  minOi: 1,
  cboeBase: 'https://cdn.cboe.com/api/global/delayed_quotes/options/'
};

function buildGexTakeProfitStopLimit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var holdings = gexReadHoldings_(ss).slice(0, GEX_TPSL.maxTickers);
  if (!holdings.length) throw new Error('No usable holdings tickers found. Pull Holdings first.');

  var sh = ss.getSheetByName(GEX_TPSL.reportSheet) || ss.insertSheet(GEX_TPSL.reportSheet);
  gexPrepareSheet_(sh);

  var rows = [];
  var errors = [];
  holdings.forEach(function(h) {
    try {
      var result = gexAnalyzeTicker_(h.ticker, h.price);
      rows.push(gexBuildOutputRow_(h, result));
      Utilities.sleep(250);
    } catch (e) {
      errors.push(h.ticker + ': ' + e.message);
      rows.push([h.ticker, h.qty, gexMoney_(h.cost), '', 'n/a', 'n/a', 'n/a', 'n/a', '', 'n/a', '']);
    }
  });

  var row = 1;
  row = gexTitle_(sh, row, 'GEX Take Profit & Stop Limit', 'Built ' + gexNow_() + '. Spot uses live GOOGLEFINANCE formulas. GEX levels use delayed Cboe option-chain data where available.');
  row = gexSection_(sh, row, 'How To Read This');
  row = gexParagraph_(sh, row, 'Call Wall is the largest call-open-interest strike and is treated as the main upside magnet/resistance. Put Wall is the largest put-open-interest strike and is treated as downside support/risk. Gamma Flip is where cumulative estimated dealer GEX crosses zero. Take Profit leans toward the call wall. Stop Limit leans toward the put wall or gamma flip below spot. TP % and SL % are live formulas measured from the GOOGLEFINANCE spot cell.');
  row = gexSection_(sh, row, 'Portfolio GEX TP / SL Map');
  row = gexTable_(sh, row, ['Ticker','Qty','Cost Basis','Live Spot','Call Wall','Put Wall','Gamma Flip','Take Profit','TP %','Stop Limit','SL %'], rows);
  if (errors.length) {
    row = gexSection_(sh, row, 'Data Issues');
    row = gexParagraph_(sh, row, errors.join('\n'));
  }
  gexFinalize_(sh, row);
  return 'GEX TP & SL complete. Rows: ' + rows.length + (errors.length ? '. Issues: ' + errors.length + '.' : '.');
}

function gexAnalyzeTicker_(ticker, fallbackSpot) {
  var chain = gexFetchCboeChain_(ticker);
  var spot = Number(chain.spot || fallbackSpot || 0);
  if (!spot) throw new Error('missing spot price');
  var options = chain.options || [];
  if (!options.length) throw new Error('empty option chain');

  var byStrike = {};
  var callOi = {}, putOi = {};
  options.forEach(function(o) {
    var parsed = gexParseOption_(o, ticker);
    if (!parsed || parsed.dte < 0 || parsed.dte > GEX_TPSL.maxDte || parsed.oi < GEX_TPSL.minOi) return;
    var gamma = parsed.gamma || gexBsGamma_(spot, parsed.strike, parsed.dte, parsed.iv || 0.35);
    var gex = gamma * parsed.oi * 100 * spot * spot * 0.01;
    if (parsed.type === 'P') gex = -Math.abs(gex);
    else gex = Math.abs(gex);
    var k = String(parsed.strike);
    byStrike[k] = (byStrike[k] || 0) + gex;
    if (parsed.type === 'C') callOi[k] = (callOi[k] || 0) + parsed.oi;
    if (parsed.type === 'P') putOi[k] = (putOi[k] || 0) + parsed.oi;
  });

  var strikes = Object.keys(byStrike).map(Number).sort(function(a,b){return a-b;});
  if (!strikes.length) throw new Error('no option rows inside ' + GEX_TPSL.maxDte + ' DTE');
  var callWall = gexMaxOiStrike_(callOi);
  var putWall = gexMaxOiStrike_(putOi);
  var gammaFlip = gexGammaFlip_(strikes, byStrike, spot);
  var netGex = strikes.reduce(function(s,k){return s + byStrike[String(k)];}, 0);
  var regime = netGex >= 0 ? 'Positive GEX' : 'Negative GEX';

  return {spot: spot, callWall: callWall, putWall: putWall, gammaFlip: gammaFlip, netGex: netGex, regime: regime, rowsUsed: strikes.length, source: chain.source};
}

function gexBuildOutputRow_(h, r) {
  var spot = r.spot || h.price || 0;
  var tp = r.callWall && r.callWall > spot ? r.callWall : gexNearestAbove_(spot, [r.callWall, r.gammaFlip]);
  var slCandidates = [r.putWall, r.gammaFlip].filter(function(x){ return x && x < spot; });
  var sl = slCandidates.length ? Math.max.apply(null, slCandidates) : r.putWall || r.gammaFlip || '';
  return [h.ticker, h.qty, gexMoney_(h.cost), '', gexPrice_(r.callWall), gexPrice_(r.putWall), gexPrice_(r.gammaFlip), gexPrice_(tp), '', gexPrice_(sl), ''];
}

function gexFetchCboeChain_(ticker) {
  var sym = String(ticker || '').trim().toUpperCase().replace(/[^A-Z0-9.]/g, '');
  if (!sym) throw new Error('blank ticker');
  var tries = [sym, sym.replace('.', '-')];
  var lastErr = '';
  for (var i = 0; i < tries.length; i++) {
    var url = GEX_TPSL.cboeBase + encodeURIComponent(tries[i]) + '.json';
    try {
      var res = UrlFetchApp.fetch(url, {muteHttpExceptions: true, followRedirects: true, headers: {'User-Agent': 'Mozilla/5.0'}});
      if (res.getResponseCode() !== 200) { lastErr = 'Cboe HTTP ' + res.getResponseCode(); continue; }
      var data = JSON.parse(res.getContentText());
      var options = data.options || (data.data && data.data.options) || [];
      var current = data.current_price || data.currentPrice || (data.data && (data.data.current_price || data.data.currentPrice));
      if (!options.length) { lastErr = 'no options in Cboe response'; continue; }
      return {spot: Number(current || 0), options: options, source: 'Cboe delayed quotes'};
    } catch(e) { lastErr = e.message; }
  }
  throw new Error(lastErr || 'Cboe fetch failed');
}

function gexParseOption_(o, ticker) {
  var option = String(o.option || o.option_symbol || o.symbol || '').toUpperCase();
  var type = String(o.option_type || o.type || '').toUpperCase().charAt(0);
  var strike = Number(o.strike || o.strike_price || 0);
  var exp = o.expiration_date || o.expiration || o.expiry || '';
  if ((!type || !strike || !exp) && option) {
    var m = option.match(/(\d{6})([CP])(\d{8})$/);
    if (m) {
      type = m[2];
      strike = Number(m[3]) / 1000;
      exp = '20' + m[1].slice(0,2) + '-' + m[1].slice(2,4) + '-' + m[1].slice(4,6);
    }
  }
  if (type !== 'C' && type !== 'P') return null;
  if (!strike) return null;
  var dte = gexDte_(exp);
  var oi = Number(o.open_interest || o.openInterest || o.oi || 0);
  var gamma = Number(o.gamma || (o.greeks && o.greeks.gamma) || 0);
  var iv = Number(o.iv || o.implied_volatility || (o.greeks && o.greeks.iv) || 0);
  if (iv > 3) iv = iv / 100;
  return {type: type, strike: strike, dte: dte, oi: oi, gamma: gamma, iv: iv};
}

function gexBsGamma_(s, k, dte, iv) {
  if (!s || !k || !dte) return 0;
  var t = Math.max(dte / 365, 1 / 365);
  var v = Math.max(Number(iv || 0.35), 0.05);
  var d1 = (Math.log(s / k) + (0.5 * v * v) * t) / (v * Math.sqrt(t));
  return gexNormPdf_(d1) / (s * v * Math.sqrt(t));
}

function gexGammaFlip_(strikes, byStrike, spot) {
  var cum = 0, best = null, bestAbs = Infinity, prevCum = null;
  for (var i = 0; i < strikes.length; i++) {
    var k = strikes[i];
    cum += byStrike[String(k)] || 0;
    if (Math.abs(cum) < bestAbs) { bestAbs = Math.abs(cum); best = k; }
    if (prevCum !== null && ((prevCum <= 0 && cum >= 0) || (prevCum >= 0 && cum <= 0))) return k;
    prevCum = cum;
  }
  return best;
}

function gexMaxOiStrike_(map) {
  var best = '', bestOi = -1;
  Object.keys(map || {}).forEach(function(k){ if (map[k] > bestOi) { bestOi = map[k]; best = Number(k); } });
  return best;
}
function gexNearestAbove_(spot, arr){ var a=arr.filter(function(x){return x&&x>spot;}).sort(function(a,b){return a-b;}); return a.length?a[0]:''; }
function gexNormPdf_(x){ return Math.exp(-0.5*x*x)/Math.sqrt(2*Math.PI); }
function gexDte_(exp){ var d = new Date(exp); if (isNaN(d.getTime())) return 9999; var today = new Date(); return Math.ceil((new Date(d.getFullYear(), d.getMonth(), d.getDate()) - new Date(today.getFullYear(), today.getMonth(), today.getDate())) / 86400000); }
function gexNow_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
function gexMoney_(n){ return n || n === 0 ? '$' + Number(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}) : 'n/a'; }
function gexPrice_(n){ return n ? Number(n).toFixed(2) : 'n/a'; }

function gexReadHoldings_(ss) {
  var sh = ss.getSheetByName(GEX_TPSL.holdingsSheet);
  if (!sh) throw new Error('Missing Holdings tab.');
  var values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  var headers = values[0].map(gexNorm_);
  var ix = {}; headers.forEach(function(h,i){ix[h]=i;});
  function col(names, fallback){ for(var i=0;i<names.length;i++){ if(ix[names[i]]!==undefined) return ix[names[i]]; } return fallback; }
  var cTicker = col(['ticker_symbol','ticker','symbol'], 6);
  var cQty = col(['quantity','qty','shares'], 10);
  var cCost = col(['cost_basis','cost','basis'], 11);
  var cPrice = col(['institution_price','price','current_price'], 12);
  var out = [], seen = {};
  values.slice(1).forEach(function(r){
    var ticker = String(r[cTicker] || '').trim().toUpperCase();
    if (!ticker || seen[ticker] || gexExcludeTicker_(ticker)) return;
    seen[ticker] = true;
    out.push({ticker: ticker, qty: gexNum_(r[cQty]), cost: gexNum_(r[cCost]), price: gexNum_(r[cPrice])});
  });
  return out;
}
function gexExcludeTicker_(t){ return t.indexOf('CUR:')===0 || ['VMFXX','SWVXX','SPAXX','FDRXX'].indexOf(t)>=0; }
function gexNorm_(x){ return String(x||'').trim().toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,''); }
function gexNum_(x){ if(typeof x==='number') return x; var s=String(x||''); var neg=s.indexOf('(')>=0; s=s.replace(/[,$%\s()]/g,''); var n=Number(s); return isNaN(n)?0:(neg?-n:n); }
function gexGoogleFinanceFormula_(ticker, fallback) {
  var t = String(ticker || '').replace(/"/g, '');
  var fb = Number(fallback || 0);
  return '=IFERROR(GOOGLEFINANCE("' + t + '","price"),' + fb + ')';
}

function gexPrepareSheet_(sh){
  try { sh.showColumns(1, sh.getMaxColumns()); } catch(e) {}
  sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.setHiddenGridlines(true);
  if (sh.getMaxColumns() < 11) sh.insertColumnsAfter(sh.getMaxColumns(), 11 - sh.getMaxColumns());
  sh.getRange(1,1,sh.getMaxRows(),11).setFontFamily(GEX_TPSL.font).setFontSize(10).setWrap(true).setVerticalAlignment('middle');
}
function gexTitle_(sh,row,title,sub){ sh.getRange(row,1,1,11).merge().setValue(title).setFontSize(18).setFontWeight('bold').setBackground('#111827').setFontColor('#fff').setHorizontalAlignment('center'); sh.setRowHeight(row,34); row++; sh.getRange(row,1,1,11).merge().setValue(sub).setBackground('#eef2ff').setFontColor('#374151'); sh.setRowHeight(row,28); return row+2; }
function gexSection_(sh,row,title){ sh.getRange(row,1,1,11).merge().setValue(title).setFontSize(13).setFontWeight('bold').setBackground('#dbeafe').setFontColor('#111827'); sh.setRowHeight(row,24); return row+1; }
function gexParagraph_(sh,row,text){ sh.getRange(row,1,1,11).merge().setValue(text).setWrap(true).setBorder(true,true,true,true,null,null,'#e5e7eb',SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(row,58); return row+2; }
function gexTable_(sh,row,headers,rows){
  var data=[headers].concat(rows);
  var range=sh.getRange(row,1,data.length,headers.length);
  range.setValues(data).setWrap(true).setVerticalAlignment('middle').setBorder(true,true,true,true,true,true,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row,1,1,headers.length).setFontWeight('bold').setFontSize(9).setBackground('#1f2937').setFontColor('#fff').setHorizontalAlignment('center');
  for(var r=1;r<data.length;r++) {
    var rr = row + r;
    var ticker = String(rows[r-1][0] || '');
    var fallbackSpot = gexNum_(rows[r-1][3]);
    sh.getRange(rr,4).setFormula(gexGoogleFinanceFormula_(ticker, fallbackSpot)).setNumberFormat('$#,##0.00');
    sh.getRange(rr,9).setFormula('=IFERROR(H' + rr + '/D' + rr + '-1,"")').setNumberFormat('0.00%');
    sh.getRange(rr,11).setFormula('=IFERROR(J' + rr + '/D' + rr + '-1,"")').setNumberFormat('0.00%');
    sh.setRowHeight(rr,28);
  }
  sh.setRowHeight(row,26);
  return row+data.length+2;
}
function gexFinalize_(sh,lastRow){ var widths=[62,52,86,84,78,78,78,84,64,84,64]; for(var c=1;c<=11;c++) sh.setColumnWidth(c,widths[c-1]); if(sh.getMaxColumns()>11) sh.hideColumns(12, sh.getMaxColumns()-11); sh.setFrozenRows(0); }
