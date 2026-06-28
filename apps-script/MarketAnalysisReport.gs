/**
 * Portfolio Link Market Analysis - Portrait Print Engine
 * Local only: Holdings tab -> GOOGLEFINANCE cache -> Report Market Analysis.
 * Designed to print cleanly as portrait PDF.
 * Version: portrait-print-v3
 */

var MR_LOCAL = {
  holdingsSheet: 'Holdings',
  reportSheet: 'Report Market Analysis',
  cacheSheet: 'Market Data Cache',
  font: 'Times New Roman',
  lookbackDays: 760,
  macroDaysForward: 120,
  reportCols: 6,
  newsMaxItems: 4
};

var MR_SECTORS = [
  ['XLC', 'Communication Services'], ['XLY', 'Consumer Discretionary'], ['XLP', 'Consumer Staples'],
  ['XLE', 'Energy'], ['XLF', 'Financials'], ['XLV', 'Healthcare'], ['XLI', 'Industrials'],
  ['XLB', 'Materials'], ['XLRE', 'Real Estate'], ['XLK', 'Technology'], ['XLU', 'Utilities']
];
var MR_PROXIES = ['SPY','QQQ','DIA','IWM','GLD','USO','TLT','UUP','HYG','BIL'];
var MR_GF_TICKERS = {
  SPY:'NYSEARCA:SPY', QQQ:'NASDAQ:QQQ', DIA:'NYSEARCA:DIA', IWM:'NYSEARCA:IWM', GLD:'NYSEARCA:GLD',
  USO:'NYSEARCA:USO', TLT:'NASDAQ:TLT', UUP:'NYSEARCA:UUP', HYG:'NYSEARCA:HYG', BIL:'NYSEARCA:BIL',
  XLC:'NYSEARCA:XLC', XLY:'NYSEARCA:XLY', XLP:'NYSEARCA:XLP', XLE:'NYSEARCA:XLE', XLF:'NYSEARCA:XLF',
  XLV:'NYSEARCA:XLV', XLI:'NYSEARCA:XLI', XLB:'NYSEARCA:XLB', XLRE:'NYSEARCA:XLRE', XLK:'NYSEARCA:XLK', XLU:'NYSEARCA:XLU'
};
var MR_TICKER_MAP = {
  SPYM:['Core Equity','SPY','Low-cost S&P 500 core exposure.'], DIA:['Core Equity','DIA','Dow blue-chip exposure.'],
  SCHG:['Growth Equity','XLK','Large-cap growth tilt; sensitive to rates and tech leadership.'], SPMO:['Momentum Equity','SPY','Momentum factor exposure tied to risk appetite.'],
  BAC:['Financials','XLF','Bank exposure; sensitive to rates, credit, and yield curve.'], MS:['Financials','XLF','Capital markets and wealth-management exposure.'],
  STT:['Financials','XLF','Custody bank / asset-servicing exposure.'], SONY:['Consumer / ADR','XLY','Consumer technology, gaming, media, and ADR exposure.'],
  LMT:['Defense / Industrials','XLI','Defense industrial; can act as geopolitical hedge.'], HTD:['Income Equity','XLU','Dividend-income sleeve with utility/financial income profile.'],
  SGOL:['Gold / Alternative','GLD','Gold hedge against real-rate, dollar, and geopolitical stress.'], JPST:['Short Duration Safety','BIL','Ultra-short income stabilizer.'],
  VRIG:['Floating Rate Safety','BIL','Floating-rate investment-grade income stabilizer.'], CLOZ:['Credit Income','HYG','CLO credit-income sleeve; sensitive to credit spreads.']
};

function buildMarketAnalysisReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var holdings = mrReadHoldings_(ss);
  if (!holdings.length) throw new Error('No usable ticker holdings found. Pull Holdings first, or make sure rows have ticker_symbol.');

  var symbols = mrUnique_(MR_PROXIES.concat(MR_SECTORS.map(function(x){ return x[0]; })));
  var histories = mrLoadGoogleFinanceHistories_(ss, symbols);
  var market = mrBuildMarketState_(symbols, histories);
  var quality = mrQualitySummary_(symbols, market);
  var portfolioRows = mrBuildPortfolioRows_(holdings, market);
  var macroEvents = mrBuildMacroEvents_();
  var newsRows = mrNewsCatalystRows_(market);

  var sh = ss.getSheetByName(MR_LOCAL.reportSheet) || ss.insertSheet(MR_LOCAL.reportSheet);
  mrResetReportSheet_(sh);

  var row = 1;
  row = mrTitle_(sh, row, 'Raj Market Rotation Report', 'Built ' + mrNow_() + ' | Portrait PDF layout | Ticker holdings only | Source: E*TRADE + GOOGLEFINANCE');

  row = mrSection_(sh, row, 'Executive Read');
  row = mrParagraph_(sh, row, mrExecutiveRead_(market, portfolioRows), 64);

  row = mrSection_(sh, row, 'Current Market Pressure Rating');
  row = mrTable_(sh, row, ['Pressure Area','Evidence','Severity','Portfolio Meaning'], mrPressureRows_(market), [1,1,1,3]);

  row = mrSection_(sh, row, 'Market Rotation Narrative');
  row = mrParagraph_(sh, row, mrRotationNarrative_(market), 76);

  row = mrSection_(sh, row, 'Macro Risk Dashboard');
  row = mrTable_(sh, row, ['Risk Area','Current Read','Risk Level','Portfolio Meaning'], mrMacroRows_(market), [1,1,1,3]);

  row = mrSection_(sh, row, 'Sector Rotation - SPDR Map');
  row = mrTable_(sh, row, ['ETF','Sector','Rotation','Pressure','Bias'], mrSectorRows_(market), [1,1,1,1,2]);

  row = mrSection_(sh, row, 'Technical Confirmation Snapshot');
  row = mrTable_(sh, row, ['ETF','Sector','Trend Read','MA Stack','MACD','RS vs SPY'], mrTechnicalRowsCompact_(market), [1,1,1,1,1,1]);

  row = mrSection_(sh, row, 'Recent News Catalysts');
  row = mrTable_(sh, row, ['Date','Headline / Catalyst','Severity','Why It Matters'], newsRows, [1,2,1,2]);

  row = mrSection_(sh, row, 'Upcoming Macro Catalysts to Watch');
  row = mrTable_(sh, row, ['Date','Time','Event','Impact','Days','Why It Matters'], mrMacroCatalystRows_(macroEvents), [1,1,1,1,1,1]);

  row = mrSection_(sh, row, 'Prior Week Recap + Forward Trend');
  row = mrParagraph_(sh, row, mrPriorWeekText_(market), 66);

  row = mrSection_(sh, row, 'Raj Portfolio Impact - Exact Actions');
  row = mrPortfolioCards_(sh, row, portfolioRows);

  row = mrSection_(sh, row, 'Total Suggested Actions + Primary Goal');
  row = mrTable_(sh, row, ['Metric','Value'], [
    ['Total suggested trims', mrMoney_(mrTotalAction_(portfolioRows, 'Trim-QTY'))],
    ['Total suggested sells', mrMoney_(mrTotalAction_(portfolioRows, 'Sell-QTY'))],
    ['Total suggested adds', mrMoney_(mrTotalAction_(portfolioRows, 'Add-QTY'))],
    ['Primary goal', 'Keep actions small; reduce overlap; protect weak positions; add only when macro, sector, and technical confirmation line up.']
  ], [1,5]);

  row = mrSection_(sh, row, 'Aggressive Growth Setup With Risk Controls');
  row = mrTable_(sh, row, ['Setup','Trigger','Risk Control','Bias'], [
    ['Core growth add','SPY/XLK above key averages and MACD > signal','Starter size only; do not add into weak macro','Add only when confirmed'],
    ['Profit protection','Large P&L outlier or overweight sleeve','Trim 5% to 15%, not full exit','Trim-QTY'],
    ['Safety redeployment','Safety sleeve overweight and growth leadership confirmed','Keep liquidity buffer','Gradual shift only'],
    ['Avoid weak trend','Below key averages with MACD < signal','No averaging down without macro confirmation','Avoid / Hold']
  ], [1,2,2,1]);

  row = mrSection_(sh, row, 'Source List');
  row = mrTable_(sh, row, ['Source','Use'], [
    ['Google Sheets GOOGLEFINANCE','Daily price history for SPY, sector ETFs, and market proxies.'],
    ['Google News RSS','Recent market catalyst headlines when available.'],
    ['Connected E*TRADE Holdings tab','Portfolio quantities, prices, values, weights, and P&L.'],
    ['Macro schedule logic','CPI, PPI, jobs, JOLTS, PCE, GDP, retail sales, ISM, and FOMC dates.'],
    ['SPY benchmark','Relative-strength benchmark for sector rotation.']
  ], [1,5]);

  mrFinalize_(sh, row);
  return 'Market Analysis complete. Portrait report written to ' + MR_LOCAL.reportSheet + '. Market data: ' + quality.ok + '/' + quality.total + ' tickers loaded. ' + quality.note;
}

function startMarketAnalysisReport(){ return buildMarketAnalysisReport(); }
function fetchMarketAnalysisReport(){ return 'No fetch needed. Local engine writes the report directly.'; }
function fetchLatestMarketAnalysisReport(){ return 'No fetch needed. Local engine writes the report directly.'; }
function tryFetchMarketAnalysisReport(){ return {ready:true, message:'Local report writes directly.'}; }
function testMarketGitHubConnection(){ return 'GitHub is no longer used for market analysis. Use buildMarketAnalysisReport.'; }
function checkMarketGitHubConfig(){ return 'GitHub market config is no longer needed. Active engine is local GOOGLEFINANCE.'; }

function mrResetReportSheet_(sh) {
  try { sh.showColumns(1, sh.getMaxColumns()); } catch(e) {}
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.setHiddenGridlines(true);
  var maxCols = sh.getMaxColumns();
  if (maxCols < MR_LOCAL.reportCols) sh.insertColumnsAfter(maxCols, MR_LOCAL.reportCols - maxCols);
  if (sh.getMaxRows() < 260) sh.insertRowsAfter(sh.getMaxRows(), 260 - sh.getMaxRows());
  var rg = sh.getRange(1, 1, sh.getMaxRows(), MR_LOCAL.reportCols);
  rg.setFontFamily(MR_LOCAL.font).setFontSize(10).setWrap(true).setVerticalAlignment('top').setNumberFormat('@');
  var widths = [80, 105, 105, 90, 86, 235];
  for (var c = 1; c <= MR_LOCAL.reportCols; c++) sh.setColumnWidth(c, widths[c-1]);
  if (sh.getMaxColumns() > MR_LOCAL.reportCols) {
    try { sh.hideColumns(MR_LOCAL.reportCols + 1, sh.getMaxColumns() - MR_LOCAL.reportCols); } catch(e2) {}
  }
}

function mrLoadGoogleFinanceHistories_(ss, symbols) {
  var sh = ss.getSheetByName(MR_LOCAL.cacheSheet) || ss.insertSheet(MR_LOCAL.cacheSheet);
  try { sh.showSheet(); } catch(e) {}
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.setHiddenGridlines(true);
  var neededCols = symbols.length * 4;
  if (sh.getMaxColumns() < neededCols) sh.insertColumnsAfter(sh.getMaxColumns(), neededCols - sh.getMaxColumns());
  if (sh.getMaxRows() < 720) sh.insertRowsAfter(sh.getMaxRows(), 720 - sh.getMaxRows());
  symbols.forEach(function(sym, i){
    var col = 1 + i * 4;
    sh.getRange(1, col).setValue(sym).setFontWeight('bold');
    var gf = MR_GF_TICKERS[sym] || sym;
    sh.getRange(2, col).setFormula('=GOOGLEFINANCE("' + gf + '","close",TODAY()-' + MR_LOCAL.lookbackDays + ',TODAY(),"DAILY")');
  });
  SpreadsheetApp.flush();
  var histories = {};
  for (var wait = 0; wait < 12; wait++) {
    Utilities.sleep(wait === 0 ? 7000 : 3000);
    SpreadsheetApp.flush();
    histories = mrReadCacheHistories_(sh, symbols);
    var okCount = symbols.filter(function(s){ return histories[s] && histories[s].closes.length >= 60; }).length;
    if (okCount >= Math.max(5, Math.floor(symbols.length * 0.85))) break;
  }
  try { sh.hideSheet(); } catch(e3) {}
  return histories;
}

function mrReadCacheHistories_(sh, symbols) {
  var histories = {};
  symbols.forEach(function(sym, i){
    var col = 1 + i * 4;
    var n = Math.min(700, sh.getMaxRows()-1);
    var values = sh.getRange(2, col, n, 2).getValues();
    var display = sh.getRange(2, col, n, 2).getDisplayValues();
    var dates = [], closes = [], error = '';
    for (var r = 0; r < values.length; r++) {
      var a = values[r][0], b = values[r][1], da = display[r][0], db = display[r][1];
      if (!error && (String(da).indexOf('#') === 0 || String(db).indexOf('#') === 0)) error = da || db;
      if (a instanceof Date && typeof b === 'number' && !isNaN(b)) { dates.push(a); closes.push(b); }
    }
    histories[sym] = {dates: dates, closes: closes, error: error, source: 'GOOGLEFINANCE ' + (MR_GF_TICKERS[sym] || sym)};
  });
  return histories;
}

function mrBuildMarketState_(symbols, histories) {
  var out = {};
  symbols.forEach(function(sym){
    var h = histories[sym];
    out[sym] = (!h || h.closes.length < 60) ? mrBlankTech_(sym, h && h.error ? h.error : 'GOOGLEFINANCE insufficient rows') : mrTech_(sym, h);
  });
  return out;
}
function mrTech_(sym, hist) {
  var c = hist.closes, price = mrLast_(c), ema20s = mrEmaSeries_(c, 20), ema20 = mrLast_(ema20s), sma50 = mrSma_(c, 50), sma200 = mrSma_(c, Math.min(200, c.length));
  var macd = mrMacd_(c), rsi14 = mrRsi_(c, 14), ret20 = mrRet_(c, 20), ret60 = mrRet_(c, 60), vol20 = mrVol_(c, 20), dd63 = mrDrawdown_(c, 63);
  var trendPts = (price>=ema20?1:0)+(price>=sma50?1:0)+(price>=sma200?1:0)+(ema20>=sma50?1:0)+(sma50>=sma200?1:0);
  var momPts = (macd.macd>=macd.signal?1:0)+(macd.macd>=0?1:0)+(ret20>0?1:0)+(ret60>0?1:0)+(rsi14>=50?1:0);
  var trendScore = trendPts/5, momentumScore = momPts/5, riskScore = Math.max(0, Math.min(1, 1 - (vol20/0.45) + (dd63/0.35)));
  var composite = 0.45*trendScore + 0.40*momentumScore + 0.15*riskScore;
  var read = composite>=0.78 ? 'Bullish confirmed' : composite>=0.58 ? 'Improving' : composite>=0.38 ? 'Mixed' : 'Weak';
  return {ticker:sym, price:price, ema20:ema20, sma50:sma50, sma200:sma200, macd:macd.macd, signal:macd.signal, ret20d:ret20, ret60d:ret60, vol20:vol20, drawdown63d:dd63, trendScore:trendScore, momentumScore:momentumScore, riskScore:riskScore, compositeScore:composite, source:hist.source, rows:c.length, priceVs20:price>=ema20?'Above 20':'Below 20', priceVs50:price>=sma50?'Above 50':'Below 50', priceVs200:price>=sma200?'Above 200':'Below 200', cross20_50:ema20>=sma50?'20>50':'20<50', trend50_200:sma50>=sma200?'50>200':'50<200', macdSignal:macd.macd>=macd.signal?'MACD>Sig':'MACD<Sig', macdZero:macd.macd>=0?'Above 0':'Below 0', read:read};
}
function mrBlankTech_(sym, reason) { return {ticker:sym, price:0, ema20:0, sma50:0, sma200:0, macd:0, signal:0, ret20d:0, ret60d:0, vol20:0, drawdown63d:0, trendScore:0, momentumScore:0, riskScore:0, compositeScore:0, source:reason, rows:0, priceVs20:'n/a', priceVs50:'n/a', priceVs200:'n/a', cross20_50:'n/a', trend50_200:'n/a', macdSignal:'n/a', macdZero:'n/a', read:'Needs data'}; }
function mrQualitySummary_(symbols, market){ var ok=0, failed=[]; symbols.forEach(function(s){ if(market[s] && market[s].read!=='Needs data') ok++; else failed.push(s); }); return {ok:ok, total:symbols.length, failed:failed, note: failed.length ? 'Missing: ' + failed.join(', ') : 'All sector/proxy tickers loaded.'}; }

function mrExecutiveRead_(m, rows){ var p=mrOverallPressure_(m), adds=mrCountAction_(rows,'Add-QTY'), trims=mrCountAction_(rows,'Trim-QTY'), sells=mrCountAction_(rows,'Sell-QTY'); return 'Market pressure is ' + p.rating + ' because ' + p.reason + '. Portfolio flags show ' + adds + ' add candidates, ' + trims + ' trim candidates, and ' + sells + ' sell/reduce candidates. Treat this as a risk review layer, not an automatic trade instruction.'; }
function mrOverallPressure_(m){ var score=0, reasons=[]; if(m.SPY.priceVs50==='Below 50'){score+=2; reasons.push('SPY is below its 50 SMA');} if(m.HYG.priceVs50==='Below 50'||m.HYG.ret20d<-0.03){score+=2; reasons.push('HYG credit proxy is weak');} if(m.TLT.ret20d<-0.04){score+=2; reasons.push('long bonds are falling');} if(m.USO.ret20d>0.06){score+=1; reasons.push('oil is rising');} if(m.UUP.ret20d>0.03){score+=1; reasons.push('dollar is firming');} if(m.SPY.read==='Bullish confirmed') score-=1; return {rating:score>=5?'HIGH':score>=3?'MEDIUM':score>=1?'WATCH':'LOW', score:score, reason:reasons.length?reasons.join('; '):'SPY trend and credit conditions are not showing major stress'}; }
function mrPressureRows_(m){ var p=mrOverallPressure_(m); return [
  ['Broad trend','SPY ' + m.SPY.priceVs50 + '; ' + m.SPY.read, p.rating, 'Controls add aggressiveness.'],
  ['Rates','TLT 20d ' + mrPctText_(m.TLT.ret20d), m.TLT.ret20d<-0.04?'HIGH':m.TLT.ret20d<-0.02?'MEDIUM':'LOW', 'Higher yields pressure growth.'],
  ['Credit','HYG 20d ' + mrPctText_(m.HYG.ret20d) + '; ' + m.HYG.priceVs50, (m.HYG.priceVs50==='Below 50'||m.HYG.ret20d<-0.03)?'HIGH':m.HYG.ret20d<0?'MEDIUM':'LOW', 'Credit weakness lowers risk appetite.'],
  ['Oil','USO 20d ' + mrPctText_(m.USO.ret20d), m.USO.ret20d>0.06?'HIGH':m.USO.ret20d>0.03?'MEDIUM':'LOW', 'Oil can revive inflation pressure.'],
  ['Dollar','UUP 20d ' + mrPctText_(m.UUP.ret20d), m.UUP.ret20d>0.03?'MEDIUM':'LOW', 'Strong dollar pressures ADRs and multinationals.']
]; }
function mrMacroRows_(m) { return [
  ['Geopolitics','GLD 20d ' + mrPctText_(m.GLD.ret20d) + '; USO 20d ' + mrPctText_(m.USO.ret20d), (m.GLD.ret20d>0.03||m.USO.ret20d>0.05)?'Elevated':'Normal', 'Gold/defense matter more if shocks rise.'],
  ['Energy','USO 20d ' + mrPctText_(m.USO.ret20d), m.USO.ret20d>0.05?'Oil rising':m.USO.ret20d<-0.05?'Oil easing':'Neutral', 'Affects inflation and consumer sectors.'],
  ['Liquidity','SPY ' + m.SPY.priceVs50 + '; HYG 20d ' + mrPctText_(m.HYG.ret20d), m.SPY.priceVs50==='Below 50'||m.HYG.ret20d<-0.03?'Risk-off':'Acceptable', 'Controls add size.'],
  ['Fed / Yields','TLT 20d ' + mrPctText_(m.TLT.ret20d), m.TLT.ret20d<-0.04?'Yield pressure':m.TLT.ret20d>0.04?'Easing':'Neutral', 'Affects growth, banks, bonds.'],
  ['Inflation / $','UUP 20d ' + mrPctText_(m.UUP.ret20d), m.UUP.ret20d>0.03?'Tighter':'Contained', 'Affects ADRs and valuation multiples.'],
  ['Growth','SPY read: ' + m.SPY.read, m.SPY.compositeScore>=0.58?'Constructive':'Mixed/weak', 'Confirms risk appetite.']
]; }
function mrSectorRows_(m) { var spy=m.SPY; return MR_SECTORS.map(function(x){ var t=m[x[0]], rot=mrRotation_(t, spy); return [x[0], x[1], rot, mrFlow_(t, spy), mrBiasShort_(rot)]; }); }
function mrTechnicalRowsCompact_(m) { var spy=m.SPY; return MR_SECTORS.map(function(x){ var t=m[x[0]]; return [x[0], x[1], t.read, t.priceVs20 + ' | ' + t.priceVs50 + ' | ' + t.priceVs200, t.macdSignal + ' / ' + t.macdZero, mrRS_(t, spy)]; }); }
function mrRotationNarrative_(m){ var spy=m.SPY, entering=[], exiting=[], mixed=[]; MR_SECTORS.forEach(function(x){ var t=m[x[0]], rot=mrRotation_(t, spy), rs=mrRS_(t, spy), label=x[0]+' '+x[1]; if((rot==='Leadership'||rot==='Positive trend') && rs!=='Underperforming') entering.push(label); else if(rot==='Lagging / weak'||rs==='Underperforming') exiting.push(label); else mixed.push(label); }); return 'Rotation is favoring: ' + (entering.length?entering.join(', '):'no clear leadership') + '. It is avoiding or exiting: ' + (exiting.length?exiting.join(', '):'no clear exit group') + '. Mixed: ' + (mixed.length?mixed.join(', '):'none') + '. This is based on moving averages, MACD, and relative strength versus SPY.'; }
function mrRotation_(t, spy){ if(t.read==='Needs data') return 'Needs data'; if(t.compositeScore>=0.75 && t.ret20d>spy.ret20d) return 'Leadership'; if(t.compositeScore>=0.60) return 'Positive trend'; if(t.compositeScore>=0.45 && t.ret20d>=spy.ret20d) return 'Improving'; if(t.compositeScore<0.35) return 'Lagging / weak'; return 'Mixed'; }
function mrFlow_(t, spy){ if(t.read==='Needs data'||spy.read==='Needs data') return 'Needs data'; if(t.ret20d>spy.ret20d+0.02&&t.compositeScore>=0.55) return 'Positive / Outperforming'; if(t.ret20d<spy.ret20d-0.02&&t.compositeScore<=0.50) return 'Negative / Underperforming'; return 'Mixed / In-line'; }
function mrRS_(t, spy){ if(t.read==='Needs data'||spy.read==='Needs data') return 'Needs data'; var d=t.ret20d-spy.ret20d; if(d>0.02) return 'Outperforming'; if(d<-0.02) return 'Underperforming'; return 'In-line'; }
function mrBiasShort_(r){ if(r==='Leadership') return 'Best hold/add'; if(r==='Positive trend') return 'Selective add'; if(r==='Improving') return 'Watch'; if(r==='Lagging / weak') return 'Avoid/trim'; return 'Wait'; }
function mrPriorWeekText_(m){ var spy=m.SPY, leaders=[], lag=[]; MR_SECTORS.forEach(function(x){ var rot=mrRotation_(m[x[0]], spy); if(rot==='Leadership') leaders.push(x[0]); if(rot==='Lagging / weak') lag.push(x[0]); }); return 'SPY is ' + spy.read + ' with 20-day change of ' + mrPctText_(spy.ret20d) + '. Leadership: ' + (leaders.length?leaders.join(', '):'none') + '. Weak/lagging: ' + (lag.length?lag.join(', '):'none') + '. Use the macro calendar and news catalysts before changing position sizes.'; }

function mrBuildPortfolioRows_(holdings, market) { var weights={}; holdings.forEach(function(h){ var meta=mrMeta_(h); weights[meta[0]]=(weights[meta[0]]||0)+h.weight; }); return holdings.map(function(h){ var meta=mrMeta_(h), tech=market[meta[1]]||market.SPY, tol=mrTolerance_(h, meta, tech, weights), act=mrAction_(h, meta, tech, tol); return {ticker:h.ticker, qty:h.qty, cost:mrMoney_(h.cost), value:mrMoney_(h.value), pnl:mrMoney_(h.pnl), pnlPct:mrPctText_(h.pnlPct), tolerance:tol, thesis:meta[2], action:act[0], actionQty:String(act[1]), est:mrMoney_(act[2]), reason:act[3] + ' Proxy ' + meta[1] + ': ' + tech.read + ' (' + tech.compositeScore.toFixed(2) + ').', sleeve:meta[0], priority:mrActionPriority_(act[0]), sortValue:h.value}; }).sort(function(a,b){ return a.priority-b.priority || b.sortValue-a.sortValue; }); }
function mrPortfolioCards_(sh, row, rows) {
  rows.forEach(function(r){
    var bg = r.action.indexOf('Add-QTY')===0 ? '#dcfce7' : r.action.indexOf('Sell-QTY')===0 ? '#fee2e2' : r.action.indexOf('Trim-QTY')===0 ? '#fef3c7' : '#f8fafc';
    sh.getRange(row,1,1,6).merge().setValue(r.ticker + ' | ' + r.action + ' | ' + r.sleeve).setFontWeight('bold').setFontSize(11).setBackground(bg).setBorder(true,true,true,true,null,null,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(row, 25); row++;
    row = mrTable_(sh, row, ['Qty','Value','P&L','P&L %','Tolerance','Est. Action $'], [[r.qty, r.value, r.pnl, r.pnlPct, r.tolerance, r.est]], [1,1,1,1,1,1]);
    row = mrMiniParagraph_(sh, row, 'Thesis: ' + r.thesis, 38);
    row = mrMiniParagraph_(sh, row, 'Reason: ' + r.reason, 46);
  });
  return row;
}
function mrMeta_(h){ if(MR_TICKER_MAP[h.ticker]) return MR_TICKER_MAP[h.ticker]; if(String(h.type).indexOf('etf')>=0) return ['ETF / Unmapped','SPY','ETF exposure; map sleeve manually if material.']; return ['Equity / Unmapped','SPY','Single-stock exposure; validate thesis manually.']; }
function mrTolerance_(h, meta, t, weights){ var sleeve=meta[0], sw=weights[sleeve]||0, safety=sleeve.indexOf('Safety')>=0||sleeve.indexOf('Income')>=0||sleeve.indexOf('Credit')>=0; if(safety&&sw>0.45) return 'Overweight safety'; if(safety&&sw>0.18) return 'Safety overlap'; if(safety) return 'Income tilt'; if(h.pnlPct>0.50) return 'Profit outlier'; if(h.pnlPct<-0.10&&t.compositeScore<0.45) return 'Weak'; if(h.weight<0.015) return 'Too small'; if(h.weight>0.08||sw>0.18) return 'Near limit'; if(t.compositeScore<0.35&&t.read!=='Needs data') return 'Out of tolerance'; return 'In tolerance'; }
function mrAction_(h, meta, t, tol){ var price=h.price||(h.value/h.qty)||0, qty=Math.max(0,Math.floor(h.qty)), safety=meta[0].indexOf('Safety')>=0||meta[0].indexOf('Income')>=0||meta[0].indexOf('Credit')>=0; if(qty<=0) return ['Hold',0,0,'No share quantity available.']; if(t.read==='Needs data') return ['Hold',0,0,'No market technical data yet.']; if(tol==='Profit outlier'&&t.compositeScore<0.65){var q=Math.max(1,Math.floor(qty*0.10)); return ['Trim-QTY '+q,q,q*price,'Profit outlier and mapped trend is not leadership.'];} if((tol==='Weak'||tol==='Out of tolerance')&&t.compositeScore<0.40&&!safety){var q2=Math.max(1,Math.floor(qty*0.15)); return ['Sell-QTY '+q2,q2,q2*price,'Weak mapped trend plus tolerance pressure.'];} if((tol==='Overweight safety'||tol==='Safety overlap')&&safety){var q3=Math.max(1,Math.floor(qty*0.05)); return ['Trim-QTY '+q3,q3,q3*price,'Safety sleeve overlap; trim only if reallocating to confirmed leadership.'];} if(tol==='Too small'&&t.compositeScore>=0.75&&!safety){var q4=price>=100?1:Math.max(1,Math.floor(500/Math.max(price,1))); return ['Add-QTY '+q4,q4,q4*price,'Small position with mapped technical leadership.'];} return ['Hold',0,0,'Hold-no-add until macro/news confirms.']; }
function mrActionPriority_(s){ if(String(s).indexOf('Add-QTY')===0) return 1; if(String(s).indexOf('Trim-QTY')===0) return 2; if(String(s).indexOf('Sell-QTY')===0) return 3; if(String(s).indexOf('Avoid')===0) return 4; return 9; }
function mrTotalAction_(rows, prefix){ return rows.reduce(function(sum,r){ return String(r.action).indexOf(prefix)===0 ? sum + mrNum_(r.est) : sum; },0); }
function mrCountAction_(rows, prefix){ return rows.filter(function(r){ return String(r.action).indexOf(prefix)===0; }).length; }

function mrBuildMacroEvents_(){ return mrNormalizeMacroEvents_(mrBlsMacroEvents_().concat(mrFomcEvents_()).concat(mrIsmEvents_()).concat(mrEstimatedMacroEvents_())); }
function mrMacroCatalystRows_(events){ return events.slice(0,12).map(function(e){ return [e.dateText, e.time, e.event, e.impact, String(e.daysUntil), e.why]; }); }
function mrCreateMacro_(event, dateObj, time, impact, why){ return {event:event, dateObj:mrStripTime_(dateObj), time:time, impact:impact, why:why}; }
function mrBlsMacroEvents_(){ var rows=[[2026,6,2,'8:30 AM ET','Jobs Report','HIGH','Moves Fed-rate expectations and broad indexes.'],[2026,6,14,'8:30 AM ET','CPI Inflation','HIGH','Major inflation report for yields, QQQ, and equities.'],[2026,6,15,'8:30 AM ET','PPI Inflation','HIGH','Shows producer inflation pressure.'],[2026,7,4,'10:00 AM ET','JOLTS','MEDIUM','Shows labor demand and rate-cut pressure.'],[2026,7,7,'8:30 AM ET','Jobs Report','HIGH','Moves Fed-rate expectations and broad indexes.'],[2026,7,12,'8:30 AM ET','CPI Inflation','HIGH','Major inflation report for yields, QQQ, and equities.'],[2026,7,13,'8:30 AM ET','PPI Inflation','HIGH','Shows producer inflation pressure.'],[2026,8,4,'8:30 AM ET','Jobs Report','HIGH','Moves Fed-rate expectations and broad indexes.'],[2026,8,11,'8:30 AM ET','CPI Inflation','HIGH','Major inflation report for yields, QQQ, and equities.'],[2026,9,2,'8:30 AM ET','Jobs Report','HIGH','Moves Fed-rate expectations and broad indexes.'],[2026,9,14,'8:30 AM ET','CPI Inflation','HIGH','Major inflation report for yields, QQQ, and equities.']]; return rows.map(function(x){ return mrCreateMacro_(x[4], new Date(x[0],x[1],x[2]), x[3], x[5], x[6]); }); }
function mrFomcEvents_(){ return [[2026,6,29,'FOMC Rate Decision'],[2026,8,16,'FOMC + SEP'],[2026,9,28,'FOMC Rate Decision'],[2026,11,9,'FOMC + SEP']].map(function(x){ return mrCreateMacro_(x[3], new Date(x[0],x[1],x[2]), '2:00 PM ET', 'HIGH', 'Fed decision can move yields, QQQ, and volatility.'); }); }
function mrIsmEvents_(){ var out=[], today=mrToday_(), start=new Date(today.getFullYear(), today.getMonth(), 1); for(var i=0;i<6;i++){ var d=new Date(start.getFullYear(), start.getMonth()+i, 1); out.push(mrCreateMacro_('ISM Manufacturing PMI', mrFirstBusinessDay_(d.getFullYear(),d.getMonth()), '10:00 AM ET', 'MEDIUM', 'Early growth signal for cyclicals.')); out.push(mrCreateMacro_('ISM Services PMI', mrNthBusinessDay_(d.getFullYear(),d.getMonth(),3), '10:00 AM ET', 'MEDIUM', 'Services surprise can move yields and growth stocks.')); } return out; }
function mrEstimatedMacroEvents_(){ var out=[], today=mrToday_(), start=new Date(today.getFullYear(), today.getMonth(), 1); for(var i=0;i<6;i++){ var d=new Date(start.getFullYear(), start.getMonth()+i, 1), y=d.getFullYear(), m=d.getMonth(); out.push(mrCreateMacro_('Retail Sales', mrNthWeekday_(y,m,4,2), '8:30 AM ET', 'MEDIUM', 'Consumer strength can support or weaken risk-on moves.')); out.push(mrCreateMacro_('PCE Inflation / Personal Income', mrLastBusinessDay_(y,m), '8:30 AM ET', 'HIGH', 'Fed preferred inflation gauge; affects yields and growth.')); if([0,3,6,9].indexOf(m)!==-1) out.push(mrCreateMacro_('GDP Report', mrNthWeekday_(y,m,4,4), '8:30 AM ET', 'HIGH', 'Confirms growth or slowdown risk.')); } return out; }
function mrNormalizeMacroEvents_(events){ var today=mrToday_(), end=mrAddDays_(today, MR_LOCAL.macroDaysForward), seen={}, out=[]; events.forEach(function(e){ if(!e||!(e.dateObj instanceof Date)||isNaN(e.dateObj.getTime())) return; if(e.dateObj<today||e.dateObj>end) return; e.daysUntil=mrDaysBetween_(today,e.dateObj); e.dateText=mrFormatDate_(e.dateObj); var key=e.event+'|'+e.dateText; if(!seen[key]){ seen[key]=true; out.push(e); } }); out.sort(function(a,b){ return a.dateObj-b.dateObj || mrImpactRank_(a.impact)-mrImpactRank_(b.impact); }); return out; }
function mrImpactRank_(x){ x=String(x||'').toUpperCase(); return x==='HIGH'?1:x==='MEDIUM'?2:3; }

function mrNewsCatalystRows_(market){ var rows=[]; try{ var url='https://news.google.com/rss/search?q=' + encodeURIComponent('stock market Fed inflation yields oil earnings') + '&hl=en-US&gl=US&ceid=US:en'; var xml=UrlFetchApp.fetch(url,{muteHttpExceptions:true,followRedirects:true}).getContentText(); var items=xml.match(/<item>[\s\S]*?<\/item>/g)||[]; items.slice(0,MR_LOCAL.newsMaxItems).forEach(function(item){ var title=mrClip_(mrXml_(item,'title'),105), date=mrXml_(item,'pubDate'), sev=mrHeadlineSeverity_(title); if(title) rows.push([mrNewsDate_(date), title, sev, mrNewsWhy_(sev)]); }); }catch(e){} if(!rows.length){ var p=mrOverallPressure_(market); rows.push([mrFormatDate_(new Date()), 'Live news feed unavailable; using market-proxy pressure model.', p.rating, mrClip_(p.reason,115)]); } return rows; }
function mrXml_(xml, tag){ var re=new RegExp('<'+tag+'>([\\s\\S]*?)<\\/'+tag+'>','i'), m=String(xml||'').match(re); return m?mrDecode_(m[1]).replace(/<!\[CDATA\[|\]\]>/g,'').trim():''; }
function mrHeadlineSeverity_(title){ var s=String(title||'').toLowerCase(); if(/fed|fomc|cpi|inflation|jobs|payroll|yields|treasury|oil|war|geopolitical|tariff|recession|credit/.test(s)) return 'HIGH'; if(/earnings|guidance|dollar|consumer|retail|gdp|pce|ppi/.test(s)) return 'MEDIUM'; return 'WATCH'; }
function mrNewsWhy_(sev){ return sev==='HIGH'?'Can directly affect yields, inflation expectations, credit risk, or broad equity multiples.':sev==='MEDIUM'?'Can affect sector rotation, earnings expectations, or risk appetite.':'Context item; not enough alone to drive portfolio action.'; }
function mrNewsDate_(d){ var x=new Date(d); return isNaN(x.getTime()) ? mrFormatDate_(new Date()) : mrFormatDate_(x); }
function mrDecode_(s){ return String(s||'').replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>'); }
function mrClip_(s,n){ s=String(s||''); return s.length>n?s.slice(0,n-1)+'…':s; }

function mrReadHoldings_(ss) { var sh=ss.getSheetByName(MR_LOCAL.holdingsSheet); if(!sh) throw new Error('Missing Holdings tab.'); var values=sh.getDataRange().getDisplayValues(); if(values.length<2) return []; var headers=values[0].map(mrNorm_), ix={}; headers.forEach(function(h,i){ix[h]=i;}); var fallback={ticker_symbol:6,security_name:7,security_type:8,security_subtype:9,quantity:10,cost_basis:11,institution_price:12,institution_value:13,calculated_market_value:14,unrealized_gain_loss:15,unrealized_gain_loss_pct:16,portfolio_weight:17}; function col(k){return ix[k]!==undefined?ix[k]:fallback[k];} function get(row,k){var i=col(k); return i===undefined||i>=row.length?'':row[i];} var out=[]; values.slice(1).forEach(function(row){ var ticker=String(get(row,'ticker_symbol')||'').trim().toUpperCase(), name=String(get(row,'security_name')||'').trim(), type=String(get(row,'security_type')||'').trim().toLowerCase(); if(!ticker) return; if(mrExclude_(ticker,name,type)) return; var qty=mrNum_(get(row,'quantity')), value=mrNum_(get(row,'institution_value'))||mrNum_(get(row,'calculated_market_value')), price=mrNum_(get(row,'institution_price')); if(!price&&value&&qty) price=value/qty; out.push({ticker:ticker, name:name, type:type, subtype:get(row,'security_subtype'), qty:qty, cost:mrNum_(get(row,'cost_basis')), price:price, value:value, pnl:mrNum_(get(row,'unrealized_gain_loss')), pnlPct:mrPctNum_(get(row,'unrealized_gain_loss_pct')), weight:mrPctNum_(get(row,'portfolio_weight'))}); }); return out; }
function mrExclude_(ticker,name,type){ var n=String(name||'').toLowerCase(); return type==='cash'||ticker.indexOf('CUR:')===0||['VMFXX','SWVXX','SPAXX','FDRXX'].indexOf(ticker)>=0||n.indexOf('money market')>=0||n.indexOf('sweep')>=0; }
function mrEmaSeries_(arr,p){ var k=2/(p+1), out=[], ema=arr[0]; for(var i=0;i<arr.length;i++){ ema=i===0?arr[i]:arr[i]*k+ema*(1-k); out.push(ema);} return out; }
function mrMacd_(arr){ var e12=mrEmaSeries_(arr,12), e26=mrEmaSeries_(arr,26), m=[]; for(var i=0;i<arr.length;i++) m.push(e12[i]-e26[i]); var s=mrEmaSeries_(m,9), macd=mrLast_(m), sig=mrLast_(s); return {macd:macd, signal:sig, hist:macd-sig}; }
function mrRsi_(arr,p){ var gains=[],losses=[]; for(var i=1;i<arr.length;i++){var d=arr[i]-arr[i-1]; gains.push(Math.max(0,d)); losses.push(Math.max(0,-d));} var ag=mrSma_(gains.slice(-p),p)||0, al=mrSma_(losses.slice(-p),p)||0; if(al===0) return 100; var rs=ag/al; return 100-(100/(1+rs)); }
function mrSma_(arr,p){ if(!arr.length) return 0; var n=Math.min(arr.length,p), s=0; for(var i=arr.length-n;i<arr.length;i++) s+=arr[i]; return s/n; }
function mrRet_(arr,n){ return arr.length>n&&arr[arr.length-1-n] ? arr[arr.length-1]/arr[arr.length-1-n]-1 : 0; }
function mrVol_(arr,n){ var rets=[]; for(var i=Math.max(1,arr.length-n);i<arr.length;i++) rets.push(arr[i]/arr[i-1]-1); if(!rets.length) return 0; var mean=rets.reduce(function(a,b){return a+b;},0)/rets.length; var v=rets.reduce(function(a,b){return a+Math.pow(b-mean,2);},0)/rets.length; return Math.sqrt(v)*Math.sqrt(252); }
function mrDrawdown_(arr,n){ var slice=arr.slice(-n), peak=0, min=0; slice.forEach(function(x){ peak=Math.max(peak,x); if(peak) min=Math.min(min,x/peak-1); }); return min; }
function mrLast_(arr){ return arr[arr.length-1]; }
function mrUnique_(arr){ var m={},out=[]; arr.forEach(function(x){ if(!m[x]){m[x]=true; out.push(x);} }); return out; }
function mrNorm_(x){ return String(x||'').trim().toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,''); }
function mrNum_(x){ if(x===null||x===undefined||x==='') return 0; if(typeof x==='number') return x; var s=String(x), neg=s.indexOf('(')>=0&&s.indexOf(')')>=0; s=s.replace(/[,$%\s()]/g,''); var n=Number(s); if(isNaN(n)) return 0; return neg?-n:n; }
function mrPctNum_(x){ if(typeof x==='number') return Math.abs(x)>1?x/100:x; var s=String(x||''), n=mrNum_(s); return s.indexOf('%')>=0?n/100:n; }
function mrMoney_(n){ return '$' + Number(n||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}); }
function mrPctText_(n){ return n===undefined||n===null||isNaN(Number(n))?'n/a':(Number(n)*100).toFixed(2)+'%'; }
function mrNow_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
function mrToday_(){ var n=new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }
function mrStripTime_(d){ return new Date(d.getFullYear(),d.getMonth(),d.getDate()); }
function mrAddDays_(d,days){ var x=new Date(d.getFullYear(),d.getMonth(),d.getDate()); x.setDate(x.getDate()+days); return x; }
function mrDaysBetween_(a,b){ return Math.round((mrStripTime_(b).getTime()-mrStripTime_(a).getTime())/86400000); }
function mrFormatDate_(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, yyyy'); }
function mrIsWeekend_(d){ var x=d.getDay(); return x===0||x===6; }
function mrFirstBusinessDay_(y,m){ var d=new Date(y,m,1); while(mrIsWeekend_(d)) d.setDate(d.getDate()+1); return mrStripTime_(d); }
function mrNthBusinessDay_(y,m,n){ var d=new Date(y,m,1), c=0; while(d.getMonth()===m){ if(!mrIsWeekend_(d)) c++; if(c===n) return mrStripTime_(d); d.setDate(d.getDate()+1); } return mrLastBusinessDay_(y,m); }
function mrNthWeekday_(y,m,weekday,n){ var d=new Date(y,m,1), c=0; while(d.getMonth()===m){ if(d.getDay()===weekday) c++; if(c===n) return mrStripTime_(d); d.setDate(d.getDate()+1); } return mrLastBusinessDay_(y,m); }
function mrLastBusinessDay_(y,m){ var d=new Date(y,m+1,0); while(mrIsWeekend_(d)) d.setDate(d.getDate()-1); return mrStripTime_(d); }

function mrTitle_(sh,row,title,sub){ sh.getRange(row,1,1,6).merge().setValue(title).setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#111827').setFontColor('#ffffff'); sh.setRowHeight(row,30); row++; sh.getRange(row,1,1,6).merge().setValue(sub).setFontSize(9).setFontColor('#374151').setBackground('#eef2ff').setWrap(true).setVerticalAlignment('middle'); sh.setRowHeight(row,30); return row+2; }
function mrSection_(sh,row,title){ sh.getRange(row,1,1,6).merge().setValue(title).setFontSize(12).setFontWeight('bold').setBackground('#dbeafe').setFontColor('#111827').setHorizontalAlignment('left').setVerticalAlignment('middle'); sh.setRowHeight(row,24); return row+1; }
function mrParagraph_(sh,row,text,h){ sh.getRange(row,1,1,6).merge().setValue(text).setFontSize(10).setBackground('#ffffff').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('left').setBorder(true,true,true,true,null,null,'#e5e7eb',SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(row,h||58); return row+2; }
function mrMiniParagraph_(sh,row,text,h){ sh.getRange(row,1,1,6).merge().setValue(text).setFontSize(9).setBackground('#ffffff').setWrap(true).setVerticalAlignment('middle').setBorder(true,true,true,true,null,null,'#e5e7eb',SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(row,h||38); return row+1; }
function mrTable_(sh,row,headers,rows,spans){ rows=rows||[]; spans=spans||headers.map(function(){return 1;}); var start=row, maxRows=1+rows.length; var col=1; for(var h=0;h<headers.length;h++){ var span=spans[h]||1; sh.getRange(row,col,1,span).merge().setValue(headers[h]).setFontWeight('bold').setFontSize(8).setBackground('#1f2937').setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true,true,true,true,true,true,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID); col+=span; } sh.setRowHeight(row,22); row++; rows.forEach(function(r){ col=1; for(var c=0;c<headers.length;c++){ var span2=spans[c]||1; sh.getRange(row,col,1,span2).merge().setValue(r[c]===undefined?'':String(r[c])).setFontSize(9).setBackground('#ffffff').setFontColor('#111827').setWrap(true).setVerticalAlignment('middle').setBorder(true,true,true,true,true,true,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID); col+=span2; } sh.setRowHeight(row, mrRowHeight_(r)); row++; }); return row+1; }
function mrRowHeight_(r){ var joined=String((r||[]).join(' ')); if(joined.length>230) return 60; if(joined.length>145) return 48; return 34; }
function mrFinalize_(sh,lastRow){ sh.getRange(1,1,Math.max(1,lastRow),6).setFontFamily(MR_LOCAL.font).setNumberFormat('@'); for(var r=1;r<=lastRow;r++){ if(sh.getRowHeight(r)>82) sh.setRowHeight(r,82); } sh.setFrozenRows(0); }
