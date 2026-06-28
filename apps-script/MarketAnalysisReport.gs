/**
 * Portfolio Link Market Analysis - Local GoogleFinance Engine
 * No GitHub Actions. No Python. No Yahoo/Stooq scraping.
 * Flow: Holdings tab -> hidden Market Data Cache tab using GOOGLEFINANCE -> technical calculations -> Report Market Analysis.
 * Run: buildMarketAnalysisReport()
 * Version: local-googlefinance-v1
 */

var MR_LOCAL = {
  holdingsSheet: 'Holdings',
  reportSheet: 'Report Market Analysis',
  cacheSheet: 'Market Data Cache',
  font: 'Times New Roman',
  lookbackDays: 760
};

var MR_SECTORS = [
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

var MR_PROXIES = ['SPY','QQQ','DIA','IWM','GLD','USO','TLT','UUP','HYG','BIL'];

var MR_GF_TICKERS = {
  SPY:'NYSEARCA:SPY', QQQ:'NASDAQ:QQQ', DIA:'NYSEARCA:DIA', IWM:'NYSEARCA:IWM',
  GLD:'NYSEARCA:GLD', USO:'NYSEARCA:USO', TLT:'NASDAQ:TLT', UUP:'NYSEARCA:UUP',
  HYG:'NYSEARCA:HYG', BIL:'NYSEARCA:BIL', XLC:'NYSEARCA:XLC', XLY:'NYSEARCA:XLY',
  XLP:'NYSEARCA:XLP', XLE:'NYSEARCA:XLE', XLF:'NYSEARCA:XLF', XLV:'NYSEARCA:XLV',
  XLI:'NYSEARCA:XLI', XLB:'NYSEARCA:XLB', XLRE:'NYSEARCA:XLRE', XLK:'NYSEARCA:XLK', XLU:'NYSEARCA:XLU'
};

var MR_TICKER_MAP = {
  SPYM: ['Core Equity', 'SPY', 'Low-cost S&P 500 core exposure.'],
  DIA: ['Core Equity', 'DIA', 'Dow blue-chip exposure.'],
  SCHG: ['Growth Equity', 'XLK', 'Large-cap growth tilt; sensitive to rates and tech leadership.'],
  SPMO: ['Momentum Equity', 'SPY', 'Momentum factor exposure tied to risk appetite.'],
  BAC: ['Financials', 'XLF', 'Bank exposure; sensitive to rates, credit, and yield curve.'],
  MS: ['Financials', 'XLF', 'Capital markets and wealth-management exposure.'],
  STT: ['Financials', 'XLF', 'Custody bank / asset-servicing exposure.'],
  SONY: ['Consumer / ADR', 'XLY', 'Consumer technology, gaming, media, and ADR exposure.'],
  LMT: ['Defense / Industrials', 'XLI', 'Defense industrial; can act as geopolitical hedge.'],
  HTD: ['Income Equity', 'XLU', 'Dividend-income sleeve with utility/financial income profile.'],
  SGOL: ['Gold / Alternative', 'GLD', 'Gold hedge against real-rate, dollar, and geopolitical stress.'],
  JPST: ['Short Duration Safety', 'BIL', 'Ultra-short income stabilizer.'],
  VRIG: ['Floating Rate Safety', 'BIL', 'Floating-rate investment-grade income stabilizer.'],
  CLOZ: ['Credit Income', 'HYG', 'CLO credit-income sleeve; sensitive to credit spreads.']
};

function buildMarketAnalysisReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var holdings = mrReadHoldings_(ss);
  if (!holdings.length) throw new Error('No usable holdings found. Pull Holdings first.');

  var symbols = mrUnique_(MR_PROXIES.concat(MR_SECTORS.map(function(x){return x[0];})));
  var histories = mrLoadGoogleFinanceHistories_(ss, symbols);
  var market = mrBuildMarketState_(symbols, histories);
  var portfolioRows = mrBuildPortfolioRows_(holdings, market);

  var sh = ss.getSheetByName(MR_LOCAL.reportSheet) || ss.insertSheet(MR_LOCAL.reportSheet);
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), Math.min(13, sh.getMaxColumns()))
    .setFontFamily(MR_LOCAL.font).setFontSize(9).setWrap(true).setVerticalAlignment('top');

  var row = 1;
  row = mrTitle_(sh, row, 'Raj Market Rotation Report', 'Built ' + mrNow_() + '. Powered by local Google Sheets GOOGLEFINANCE cache, not GitHub scraping.');
  row = mrParagraph_(sh, row, 'Uses connected E*TRADE holdings excluding cash and money market. Technical engine calculates 20 EMA, 50 SMA, 200 SMA, MACD, RSI, volatility, drawdown, sector rotation, and relative strength versus SPY.');

  row = mrSection_(sh, row, 'Data Quality Check');
  row = mrTable_(sh, row, ['Ticker','Status','Source / Error','Last Price'], mrDataQualityRows_(symbols, histories, market));

  row = mrSection_(sh, row, 'Macro Risk Dashboard');
  row = mrTable_(sh, row, ['Risk Area','Current Read','Risk Level','Portfolio Meaning'], mrMacroRows_(market));
  row = mrSection_(sh, row, 'Macro Summary');
  row = mrParagraph_(sh, row, 'Macro reads use market proxies for oil, yields/bonds, dollar, gold, high-yield credit, and SPY trend. Headlines, Fed calendar, CPI/jobs, and earnings still need manual/live review before trading.');

  row = mrSection_(sh, row, 'Sector Rotation - SPDR Map');
  row = mrTable_(sh, row, ['ETF','Sector','Rotation Read','Flow Pressure','Action Bias'], mrSectorRows_(market));
  row = mrSection_(sh, row, 'Sector Summary');
  row = mrParagraph_(sh, row, 'Leadership requires both trend confirmation and relative strength versus SPY. Mixed or lagging sectors are not add candidates without a separate macro catalyst.');

  row = mrSection_(sh, row, 'Technical Confirmation Snapshot');
  row = mrTable_(sh, row, ['ETF','Sector','Price vs 20 EMA','Price vs 50 SMA','Price vs 200 SMA','20/50 Crossover','50/200 Trend','MACD vs Signal','MACD Zero','RS vs SPY'], mrTechnicalRows_(market));

  row = mrSection_(sh, row, 'Prior Week Recap + Forward Trend');
  row = mrParagraph_(sh, row, mrPriorWeekText_(market));

  row = mrSection_(sh, row, 'Raj Portfolio Impact - Exact Actions');
  row = mrTable_(sh, row, ['Ticker','Qty Held','Cost Basis','Current Value','Unrealized $','P&L %','Tolerance Status','Thesis','Exact Action','Action Qty','Est. $ Value','Reason / Price Source','Sleeve'], portfolioRows);

  row = mrSection_(sh, row, 'Total Suggested Trims and Primary Goal');
  row = mrTable_(sh, row, ['Metric','Value'], [
    ['Total suggested trims', mrMoney_(mrTotalAction_(portfolioRows, 'Trim-QTY'))],
    ['Total suggested sells', mrMoney_(mrTotalAction_(portfolioRows, 'Sell-QTY'))],
    ['Total suggested adds', mrMoney_(mrTotalAction_(portfolioRows, 'Add-QTY'))],
    ['Primary goal', 'Keep actions small, reduce overlap, protect profit outliers, and add only when macro/sector thesis confirms technical trend.']
  ]);

  row = mrSection_(sh, row, 'Aggressive Growth Setup With Risk Controls');
  row = mrTable_(sh, row, ['Setup','Trigger','Risk Control','Action Bias'], [
    ['Core growth add','SPY/XLK above 20 EMA and 50 SMA with MACD > signal','Starter size only; do not add into weak macro','Add-QTY only when confirmed'],
    ['Profit protection','Large P&L outlier or overweight sleeve','Trim 5% to 15%, not full exit','Trim-QTY'],
    ['Safety redeployment','Safety sleeve overweight and growth leadership confirmed','Keep liquidity buffer','Gradual shift only'],
    ['Avoid weak trend','Below 20 EMA/50 SMA with MACD < signal','No averaging down without macro confirmation','Avoid or Hold']
  ]);

  row = mrSection_(sh, row, 'Key Catalysts to Watch');
  row = mrTable_(sh, row, ['Catalyst','Why It Matters'], [
    ['Fed/FOMC and Treasury yields','Affects growth multiples, banks, real estate, and fixed income.'],
    ['Oil and geopolitical headlines','Affects energy, inflation pressure, defense, and gold.'],
    ['Earnings guidance','Can override sector trend.'],
    ['Credit spreads / liquidity','Important for CLOZ, banks, and risk appetite.'],
    ['SPY breadth and sector relative strength','Confirms whether rotation is broadening or narrowing.']
  ]);

  row = mrSection_(sh, row, 'Chicago-Style Source List');
  row = mrTable_(sh, row, ['Source','Use'], [
    ['Google Sheets GOOGLEFINANCE','Daily price history for SPY, sectors, and market proxies.'],
    ['Connected E*TRADE Holdings tab','Portfolio quantities, prices, values, weights, and P&L.'],
    ['SPDR sector ETF map','11-sector rotation framework.'],
    ['SPY benchmark','Relative-strength benchmark.'],
    ['Manual live-news layer','Macro, Fed, inflation, earnings, and geopolitical catalysts.']
  ]);

  mrFinalize_(sh, row);
  return 'Market Analysis complete. Report written to ' + MR_LOCAL.reportSheet + '. Data source: GOOGLEFINANCE cache.';
}

function startMarketAnalysisReport(){ return buildMarketAnalysisReport(); }
function fetchMarketAnalysisReport(){ return 'No fetch needed. Local GOOGLEFINANCE engine writes the report immediately.'; }
function fetchLatestMarketAnalysisReport(){ return 'No fetch needed. Local GOOGLEFINANCE engine writes the report immediately.'; }
function tryFetchMarketAnalysisReport(){ return {ready:true, message:'No polling needed. Local report already writes directly.'}; }
function testMarketGitHubConnection(){ return 'GitHub is no longer used for market analysis. Use buildMarketAnalysisReport.'; }
function checkMarketGitHubConfig(){ return 'GitHub market config is no longer needed. Active engine is local GOOGLEFINANCE.'; }

function mrLoadGoogleFinanceHistories_(ss, symbols) {
  var sh = ss.getSheetByName(MR_LOCAL.cacheSheet) || ss.insertSheet(MR_LOCAL.cacheSheet);
  sh.clear();
  sh.setHiddenGridlines(true);
  var neededCols = symbols.length * 4;
  if (sh.getMaxColumns() < neededCols) sh.insertColumnsAfter(sh.getMaxColumns(), neededCols - sh.getMaxColumns());
  if (sh.getMaxRows() < 620) sh.insertRowsAfter(sh.getMaxRows(), 620 - sh.getMaxRows());

  symbols.forEach(function(sym, i){
    var col = 1 + i * 4;
    sh.getRange(1, col).setValue(sym).setFontWeight('bold');
    var gf = MR_GF_TICKERS[sym] || sym;
    var formula = '=GOOGLEFINANCE("' + gf + '","close",TODAY()-' + MR_LOCAL.lookbackDays + ',TODAY(),"DAILY")';
    sh.getRange(2, col).setFormula(formula);
  });
  SpreadsheetApp.flush();

  var histories = {};
  for (var wait = 0; wait < 10; wait++) {
    Utilities.sleep(wait === 0 ? 6000 : 3000);
    SpreadsheetApp.flush();
    histories = mrReadCacheHistories_(sh, symbols);
    var okCount = symbols.filter(function(s){return histories[s] && histories[s].closes.length >= 60;}).length;
    if (okCount >= Math.max(5, Math.floor(symbols.length * 0.75))) break;
  }

  try { sh.hideSheet(); } catch(e) {}
  return histories;
}

function mrReadCacheHistories_(sh, symbols) {
  var histories = {};
  symbols.forEach(function(sym, i){
    var col = 1 + i * 4;
    var values = sh.getRange(2, col, Math.min(650, sh.getMaxRows()-1), 2).getValues();
    var display = sh.getRange(2, col, Math.min(650, sh.getMaxRows()-1), 2).getDisplayValues();
    var dates = [], closes = [], error = '';
    for (var r = 0; r < values.length; r++) {
      var a = values[r][0], b = values[r][1];
      var da = display[r][0], db = display[r][1];
      if (!error && (String(da).indexOf('#') === 0 || String(db).indexOf('#') === 0)) error = da || db;
      if (a instanceof Date && typeof b === 'number' && !isNaN(b)) {
        dates.push(a);
        closes.push(b);
      }
    }
    histories[sym] = {dates: dates, closes: closes, error: error, source: 'GOOGLEFINANCE ' + (MR_GF_TICKERS[sym] || sym)};
  });
  return histories;
}

function mrBuildMarketState_(symbols, histories) {
  var out = {};
  symbols.forEach(function(sym){
    var h = histories[sym];
    if (!h || h.closes.length < 60) out[sym] = mrBlankTech_(sym, h && h.error ? h.error : 'GOOGLEFINANCE not loaded / insufficient rows');
    else out[sym] = mrTech_(sym, h);
  });
  return out;
}

function mrTech_(sym, hist) {
  var c = hist.closes;
  var price = mrLast_(c);
  var ema20s = mrEmaSeries_(c, 20);
  var ema20 = mrLast_(ema20s);
  var sma50 = mrSma_(c, 50);
  var sma200 = mrSma_(c, Math.min(200, c.length));
  var macd = mrMacd_(c);
  var rsi14 = mrRsi_(c, 14);
  var ret5 = mrRet_(c, 5), ret20 = mrRet_(c, 20), ret60 = mrRet_(c, 60);
  var vol20 = mrVol_(c, 20);
  var dd63 = mrDrawdown_(c, 63);
  var trendPts = 0;
  if (price >= ema20) trendPts++;
  if (price >= sma50) trendPts++;
  if (price >= sma200) trendPts++;
  if (ema20 >= sma50) trendPts++;
  if (sma50 >= sma200) trendPts++;
  var momPts = 0;
  if (macd.macd >= macd.signal) momPts++;
  if (macd.macd >= 0) momPts++;
  if (ret20 > 0) momPts++;
  if (ret60 > 0) momPts++;
  if (rsi14 >= 50) momPts++;
  var trendScore = trendPts / 5;
  var momentumScore = momPts / 5;
  var riskScore = Math.max(0, Math.min(1, 1 - (vol20 / 0.45) + (dd63 / 0.35)));
  var composite = 0.45 * trendScore + 0.40 * momentumScore + 0.15 * riskScore;
  var read = composite >= 0.78 ? 'Bullish confirmed' : composite >= 0.58 ? 'Improving' : composite >= 0.38 ? 'Mixed' : 'Weak';
  return {
    ticker: sym, price: price, ema20: ema20, sma50: sma50, sma200: sma200, macd: macd.macd,
    signal: macd.signal, hist: macd.hist, rsi14: rsi14, ret5d: ret5, ret20d: ret20, ret60d: ret60,
    vol20: vol20, drawdown63d: dd63, trendScore: trendScore, momentumScore: momentumScore,
    riskScore: riskScore, compositeScore: composite, source: hist.source, rows: c.length,
    priceVs20: price >= ema20 ? 'Above' : 'Below', priceVs50: price >= sma50 ? 'Above' : 'Below',
    priceVs200: price >= sma200 ? 'Above' : 'Below', cross20_50: ema20 >= sma50 ? '20 EMA > 50 SMA' : '20 EMA < 50 SMA',
    trend50_200: sma50 >= sma200 ? '50 SMA > 200 SMA' : '50 SMA < 200 SMA',
    macdSignal: macd.macd >= macd.signal ? 'MACD > Signal' : 'MACD < Signal', macdZero: macd.macd >= 0 ? 'Above 0' : 'Below 0',
    read: read
  };
}

function mrBlankTech_(sym, reason) {
  return {ticker:sym, price:0, ema20:0, sma50:0, sma200:0, macd:0, signal:0, hist:0, rsi14:50, ret5d:0, ret20d:0, ret60d:0, vol20:0, drawdown63d:0, trendScore:0, momentumScore:0, riskScore:0, compositeScore:0, source:reason, rows:0, priceVs20:'n/a', priceVs50:'n/a', priceVs200:'n/a', cross20_50:'n/a', trend50_200:'n/a', macdSignal:'n/a', macdZero:'n/a', read:'Needs data'};
}

function mrDataQualityRows_(symbols, histories, market) {
  return symbols.map(function(s){
    var t = market[s];
    var h = histories[s] || {closes:[], source:'', error:''};
    return [s, t.read === 'Needs data' ? 'Failed' : 'OK', t.read === 'Needs data' ? t.source : (h.source + ', rows=' + h.closes.length), t.price ? t.price.toFixed(2) : 'n/a'];
  });
}

function mrMacroRows_(m) {
  var spy=m.SPY, uso=m.USO, tlt=m.TLT, uup=m.UUP, gld=m.GLD, hyg=m.HYG;
  return [
    ['War / Geopolitics','GLD 20d ' + mrPctText_(gld.ret20d) + '; USO 20d ' + mrPctText_(uso.ret20d), (gld.ret20d>0.03 || uso.ret20d>0.05) ? 'Elevated watch' : 'Normal watch','Gold/defense matter more if risk shock rises.'],
    ['Oil / Energy shock','USO 20d ' + mrPctText_(uso.ret20d), uso.ret20d>0.05 ? 'Oil pressure rising' : uso.ret20d<-0.05 ? 'Oil easing' : 'Neutral','Affects XLE, inflation expectations, and consumer pressure.'],
    ['Crisis / Liquidity','SPY ' + spy.priceVs50 + ' 50 SMA; HYG 20d ' + mrPctText_(hyg.ret20d), spy.priceVs50==='Below' || hyg.ret20d<-0.03 ? 'Risk-off watch' : 'Risk-on acceptable','Controls how aggressive adds should be.'],
    ['Fed / Yields','TLT 20d ' + mrPctText_(tlt.ret20d), tlt.ret20d<-0.04 ? 'Yield pressure rising' : tlt.ret20d>0.04 ? 'Yield pressure easing' : 'Neutral','Affects growth stocks, banks, real estate, and bond sleeves.'],
    ['Inflation / Dollar','UUP 20d ' + mrPctText_(uup.ret20d) + '; USO 20d ' + mrPctText_(uso.ret20d), uup.ret20d>0.03 || uso.ret20d>0.05 ? 'Tighter impulse' : 'Contained','Affects ADRs, gold, and valuation multiples.'],
    ['Jobs / Growth','SPY technical read: ' + spy.read, spy.compositeScore>=0.58 ? 'Growth trend constructive' : 'Growth trend mixed/weak','Confirms whether rotation supports risk assets.']
  ];
}

function mrSectorRows_(m) {
  var spy = m.SPY;
  return MR_SECTORS.map(function(x){ var t=m[x[0]], rot=mrRotation_(t, spy); return [x[0], x[1], rot, mrFlow_(t, spy) + ' / ' + mrRS_(t, spy), mrBias_(rot)]; });
}

function mrTechnicalRows_(m) {
  var spy = m.SPY;
  return MR_SECTORS.map(function(x){ var t=m[x[0]]; return [x[0], x[1], t.priceVs20, t.priceVs50, t.priceVs200, t.cross20_50, t.trend50_200, t.macdSignal, t.macdZero, mrRS_(t, spy)]; });
}

function mrRotation_(t, spy){ if(t.read==='Needs data') return 'Needs data'; if(t.compositeScore>=0.75 && t.ret20d>spy.ret20d) return 'Leadership'; if(t.compositeScore>=0.60) return 'Positive trend'; if(t.compositeScore>=0.45 && t.ret20d>=spy.ret20d) return 'Improving'; if(t.compositeScore<0.35) return 'Lagging / weak'; return 'Mixed'; }
function mrFlow_(t, spy){ if(t.read==='Needs data'||spy.read==='Needs data') return 'Needs data'; if(t.ret20d>spy.ret20d+0.02 && t.compositeScore>=0.55) return 'Positive pressure'; if(t.ret20d<spy.ret20d-0.02 && t.compositeScore<=0.50) return 'Negative pressure'; return 'Mixed pressure'; }
function mrRS_(t, spy){ if(t.read==='Needs data'||spy.read==='Needs data') return 'Needs data'; var d=t.ret20d-spy.ret20d; if(d>0.02) return 'Outperforming SPY'; if(d<-0.02) return 'Underperforming SPY'; return 'In line with SPY'; }
function mrBias_(r){ if(r==='Leadership') return 'Best add/hold candidates after macro confirmation'; if(r==='Positive trend') return 'Hold / selective add only'; if(r==='Improving') return 'Watchlist / starter only'; if(r==='Lagging / weak') return 'Avoid adds / review trims'; return 'Data needed / wait'; }
function mrPriorWeekText_(m){ var spy=m.SPY, leaders=[], lag=[]; MR_SECTORS.forEach(function(x){ var rot=mrRotation_(m[x[0]], spy); if(rot==='Leadership') leaders.push(x[0]+' '+x[1]); if(rot==='Lagging / weak') lag.push(x[0]+' '+x[1]); }); return 'Forward trend: SPY is ' + spy.read + ' with 20-day change of ' + mrPctText_(spy.ret20d) + '. Confirmed technical leadership: ' + (leaders.length?leaders.join(', '):'none') + '. Weak/lagging sectors: ' + (lag.length?lag.slice(0,4).join(', '):'none') + '. Use live macro/news and earnings checks before trading.'; }

function mrBuildPortfolioRows_(holdings, market) {
  var weights = {};
  holdings.forEach(function(h){ var meta=mrMeta_(h); weights[meta[0]]=(weights[meta[0]]||0)+h.weight; });
  return holdings.map(function(h){
    var meta=mrMeta_(h), tech=market[meta[1]]||market.SPY, tol=mrTolerance_(h, meta, tech, weights), act=mrAction_(h, meta, tech, tol);
    return [h.ticker, h.qty, mrMoney_(h.cost), mrMoney_(h.value), mrMoney_(h.pnl), mrPctText_(h.pnlPct), tol, meta[2], act[0], act[1], mrMoney_(act[2]), act[3] + ' Sector proxy ' + meta[1] + ' = ' + tech.read + ' (' + tech.compositeScore.toFixed(2) + ').', meta[0]];
  }).sort(function(a,b){ return mrActionPriority_(a[8]) - mrActionPriority_(b[8]) || mrNum_(b[3]) - mrNum_(a[3]); });
}

function mrMeta_(h){ if(MR_TICKER_MAP[h.ticker]) return MR_TICKER_MAP[h.ticker]; var n=String(h.name||'').toLowerCase(); if(String(h.type).indexOf('fixed')>=0||n.indexOf('fdic')>=0||n.indexOf(' cd ')>=0) return ['Fixed Income Safety','BIL','Principal/income stabilizer; watch safety overweight.']; if(String(h.type).indexOf('etf')>=0) return ['ETF / Unmapped','SPY','ETF exposure; map sleeve manually if material.']; return ['Equity / Unmapped','SPY','Single-stock exposure; validate thesis manually.']; }
function mrTolerance_(h, meta, t, weights){ var sleeve=meta[0], sw=weights[sleeve]||0, safety=sleeve.indexOf('Safety')>=0||sleeve.indexOf('Income')>=0||sleeve.indexOf('Credit')>=0; if(safety&&sw>0.45) return 'Overweight safety'; if(safety&&sw>0.18) return 'Safety overlap'; if(safety) return 'Income tilt'; if(h.pnlPct>0.50) return 'Profit outlier'; if(h.pnlPct<-0.10&&t.compositeScore<0.45) return 'Weak'; if(h.weight<0.015) return 'Too small'; if(h.weight>0.08||sw>0.18) return 'Near limit'; if(t.compositeScore<0.35&&t.read!=='Needs data') return 'Out of tolerance'; return 'In tolerance'; }
function mrAction_(h, meta, t, tol){ var price=h.price||(h.value/h.qty)||0, qty=Math.max(0,Math.floor(h.qty)), safety=meta[0].indexOf('Safety')>=0||meta[0].indexOf('Income')>=0||meta[0].indexOf('Credit')>=0; if(qty<=0) return ['Hold',0,0,'No share quantity available. Price source: E*TRADE institution price/value.']; if(t.read==='Needs data') return ['Hold',0,0,'No market technical data yet. Price source: E*TRADE institution price/value.']; if(tol==='Profit outlier'&&t.compositeScore<0.65){var q=Math.max(1,Math.floor(qty*0.10)); return ['Trim-QTY '+q,q,q*price,'Profit outlier and mapped trend is not leadership. Estimate uses E*TRADE institution price/value.'];} if((tol==='Weak'||tol==='Out of tolerance')&&t.compositeScore<0.40&&!safety){var q2=Math.max(1,Math.floor(qty*0.15)); return ['Sell-QTY '+q2,q2,q2*price,'Weak mapped trend plus tolerance pressure. Estimate uses E*TRADE institution price/value.'];} if((tol==='Overweight safety'||tol==='Safety overlap')&&safety){var q3=Math.max(1,Math.floor(qty*0.05)); return ['Trim-QTY '+q3,q3,q3*price,'Safety sleeve overlap; trim only if reallocating to confirmed leadership. Estimate uses E*TRADE institution price/value.'];} if(tol==='Too small'&&t.compositeScore>=0.75&&!safety){var q4=price>=100?1:Math.max(1,Math.floor(500/Math.max(price,1))); return ['Add-QTY '+q4,q4,q4*price,'Small position with mapped technical leadership. Estimate uses E*TRADE institution price/value.'];} return ['Hold',0,0,'Hold-no-add until macro/news confirms. Price source: E*TRADE institution price/value.']; }
function mrActionPriority_(s){ if(String(s).indexOf('Add-QTY')===0) return 1; if(String(s).indexOf('Trim-QTY')===0) return 2; if(String(s).indexOf('Sell-QTY')===0) return 3; if(String(s).indexOf('Avoid')===0) return 4; return 9; }
function mrTotalAction_(rows, prefix){ return rows.reduce(function(sum,r){ return String(r[8]).indexOf(prefix)===0 ? sum + mrNum_(r[10]) : sum; },0); }

function mrReadHoldings_(ss) {
  var sh = ss.getSheetByName(MR_LOCAL.holdingsSheet); if(!sh) throw new Error('Missing Holdings tab.');
  var values = sh.getDataRange().getDisplayValues(); if(values.length<2) return [];
  var headers = values[0].map(mrNorm_); var ix={}; headers.forEach(function(h,i){ix[h]=i;});
  var fallback = {ticker_symbol:6, security_name:7, security_type:8, security_subtype:9, quantity:10, cost_basis:11, institution_price:12, institution_value:13, calculated_market_value:14, unrealized_gain_loss:15, unrealized_gain_loss_pct:16, portfolio_weight:17};
  function col(k){ return ix[k]!==undefined?ix[k]:fallback[k]; }
  function get(row,k){ var i=col(k); return i===undefined||i>=row.length?'':row[i]; }
  var out=[];
  values.slice(1).forEach(function(row){
    var ticker=String(get(row,'ticker_symbol')||'').trim().toUpperCase(), name=String(get(row,'security_name')||'').trim(), type=String(get(row,'security_type')||'').trim().toLowerCase();
    if(!ticker&&!name) return; if(mrExclude_(ticker,name,type)) return;
    var qty=mrNum_(get(row,'quantity')), value=mrNum_(get(row,'institution_value'))||mrNum_(get(row,'calculated_market_value')), price=mrNum_(get(row,'institution_price'));
    if(!price&&value&&qty) price=value/qty;
    out.push({ticker:ticker||'(NO TICKER)', name:name, type:type, subtype:get(row,'security_subtype'), qty:qty, cost:mrNum_(get(row,'cost_basis')), price:price, value:value, pnl:mrNum_(get(row,'unrealized_gain_loss')), pnlPct:mrPctNum_(get(row,'unrealized_gain_loss_pct')), weight:mrPctNum_(get(row,'portfolio_weight'))});
  });
  return out;
}
function mrExclude_(ticker,name,type){ var n=String(name||'').toLowerCase(); return type==='cash'||ticker.indexOf('CUR:')===0||['VMFXX','SWVXX','SPAXX','FDRXX'].indexOf(ticker)>=0||n.indexOf('money market')>=0||n.indexOf('sweep')>=0; }

function mrEmaSeries_(arr, p){ var k=2/(p+1), out=[], ema=arr[0]; for(var i=0;i<arr.length;i++){ ema=i===0?arr[i]:arr[i]*k+ema*(1-k); out.push(ema);} return out; }
function mrMacd_(arr){ var e12=mrEmaSeries_(arr,12), e26=mrEmaSeries_(arr,26), m=[]; for(var i=0;i<arr.length;i++) m.push(e12[i]-e26[i]); var s=mrEmaSeries_(m,9), macd=mrLast_(m), sig=mrLast_(s); return {macd:macd, signal:sig, hist:macd-sig}; }
function mrRsi_(arr,p){ var gains=[],losses=[]; for(var i=1;i<arr.length;i++){var d=arr[i]-arr[i-1]; gains.push(Math.max(0,d)); losses.push(Math.max(0,-d));} var ag=mrSma_(gains.slice(-p),p)||0, al=mrSma_(losses.slice(-p),p)||0; if(al===0) return 100; var rs=ag/al; return 100-(100/(1+rs)); }
function mrSma_(arr,p){ if(!arr.length) return 0; var n=Math.min(arr.length,p), s=0; for(var i=arr.length-n;i<arr.length;i++) s+=arr[i]; return s/n; }
function mrRet_(arr,n){ return arr.length>n&&arr[arr.length-1-n] ? arr[arr.length-1]/arr[arr.length-1-n]-1 : 0; }
function mrVol_(arr,n){ var rets=[]; for(var i=Math.max(1,arr.length-n);i<arr.length;i++) rets.push(arr[i]/arr[i-1]-1); if(!rets.length) return 0; var mean=rets.reduce(function(a,b){return a+b;},0)/rets.length; var v=rets.reduce(function(a,b){return a+Math.pow(b-mean,2);},0)/rets.length; return Math.sqrt(v)*Math.sqrt(252); }
function mrDrawdown_(arr,n){ var slice=arr.slice(-n), peak=0, min=0; slice.forEach(function(x){ peak=Math.max(peak,x); if(peak) min=Math.min(min,x/peak-1); }); return min; }
function mrLast_(arr){ return arr[arr.length-1]; }
function mrUnique_(arr){ var m={},out=[]; arr.forEach(function(x){ if(!m[x]){m[x]=true; out.push(x);} }); return out; }
function mrNorm_(x){ return String(x||'').trim().toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,''); }
function mrNum_(x){ if(x===null||x===undefined||x==='') return 0; if(typeof x==='number') return x; var s=String(x).replace(/[,$%\s]/g,'').replace(/[()]/g,''); var n=Number(s); return isNaN(n)?0:n; }
function mrPctNum_(x){ if(typeof x==='number') return Math.abs(x)>1?x/100:x; var s=String(x||''); var n=mrNum_(s); return s.indexOf('%')>=0?n/100:n; }
function mrMoney_(n){ return '$' + Number(n||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}); }
function mrPctText_(n){ return n===undefined||n===null||isNaN(Number(n))?'n/a':(Number(n)*100).toFixed(2)+'%'; }
function mrNow_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }

function mrTitle_(sh,row,title,sub){ sh.getRange(row,1,1,13).merge().setValue(title).setFontSize(18).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff'); row++; sh.getRange(row,1,1,13).merge().setValue(sub).setFontSize(9).setFontColor('#374151').setBackground('#eef2ff'); return row+2; }
function mrSection_(sh,row,title){ sh.getRange(row,1,1,13).merge().setValue(title).setFontSize(13).setFontWeight('bold').setBackground('#dbeafe').setFontColor('#111827'); return row+1; }
function mrParagraph_(sh,row,text){ sh.getRange(row,1,1,13).merge().setValue(text).setFontSize(9).setBackground('#ffffff').setWrap(true).setVerticalAlignment('top'); sh.setRowHeight(row,42); return row+2; }
function mrTable_(sh,row,headers,rows){ var width=Math.max(headers.length,1), data=[headers].concat(rows||[]); var clean=data.map(function(r){var o=[]; for(var i=0;i<width;i++) o.push(r[i]===undefined||r[i]===null?'':r[i]); return o;}); var range=sh.getRange(row,1,clean.length,width); range.setValues(clean).setWrap(true).setVerticalAlignment('top').setBorder(true,true,true,true,true,true,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID); sh.getRange(row,1,1,width).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff'); if(clean.length>1) sh.getRange(row+1,1,clean.length-1,width).setBackground('#ffffff').setFontColor('#111827'); return row+clean.length+2; }
function mrFinalize_(sh,lastRow){ for(var c=1;c<=13;c++){ var w=110; if(c===8||c===12) w=220; if(c===13) w=135; sh.setColumnWidth(c,w);} sh.getRange(1,1,Math.max(1,lastRow),Math.min(13,sh.getMaxColumns())).setFontFamily(MR_LOCAL.font); sh.autoResizeRows(1,Math.max(1,lastRow)); }
