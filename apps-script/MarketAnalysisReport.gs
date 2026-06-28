/**
 * Portfolio Link Market Analysis - Local GoogleFinance Engine
 * No GitHub Actions. No Python. No Yahoo/Stooq scraping.
 * Flow: Holdings tab -> hidden Market Data Cache tab using GOOGLEFINANCE -> technical calculations -> Report Market Analysis.
 * Version: local-googlefinance-v2-handcrafted-report
 */

var MR_LOCAL = {
  holdingsSheet: 'Holdings',
  reportSheet: 'Report Market Analysis',
  cacheSheet: 'Market Data Cache',
  font: 'Times New Roman',
  lookbackDays: 760,
  macroDaysForward: 120,
  newsMaxItems: 5
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
  if (!holdings.length) throw new Error('No usable ticker holdings found. Pull Holdings first, or make sure rows have ticker_symbol.');

  var symbols = mrUnique_(MR_PROXIES.concat(MR_SECTORS.map(function(x){return x[0];})));
  var histories = mrLoadGoogleFinanceHistories_(ss, symbols);
  var market = mrBuildMarketState_(symbols, histories);
  var quality = mrQualitySummary_(symbols, histories, market);
  var macroEvents = mrBuildMacroEvents_();
  var newsRows = mrNewsCatalystRows_(market);
  var portfolioRows = mrBuildPortfolioRows_(holdings, market);

  var sh = ss.getSheetByName(MR_LOCAL.reportSheet) || ss.insertSheet(MR_LOCAL.reportSheet);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), Math.min(13, sh.getMaxColumns()))
    .setFontFamily(MR_LOCAL.font)
    .setFontSize(12)
    .setWrap(true)
    .setVerticalAlignment('top');

  var row = 1;
  row = mrTitle_(sh, row, 'Raj Market Rotation Report', 'Built ' + mrNow_() + '. Local GoogleFinance technical engine. Cash, money market, and no-ticker rows excluded.');

  row = mrSection_(sh, row, 'Executive Read');
  row = mrParagraph_(sh, row, mrExecutiveRead_(market, portfolioRows));

  row = mrSection_(sh, row, 'Current Market Pressure Rating');
  row = mrTable_(sh, row, ['Pressure Area','Current Evidence','Severity','Portfolio Meaning'], mrPressureRows_(market));

  row = mrSection_(sh, row, 'Market Rotation Narrative');
  row = mrParagraph_(sh, row, mrRotationNarrative_(market));

  row = mrSection_(sh, row, 'Macro Risk Dashboard');
  row = mrTable_(sh, row, ['Risk Area','Current Read','Risk Level','Portfolio Meaning'], mrMacroRows_(market));

  row = mrSection_(sh, row, 'Sector Rotation - SPDR Map');
  row = mrTable_(sh, row, ['ETF','Sector','Rotation Read','Flow Pressure','Action Bias'], mrSectorRows_(market));

  row = mrSection_(sh, row, 'Technical Confirmation Snapshot');
  row = mrTable_(sh, row, ['ETF','Sector','Price vs 20 EMA','Price vs 50 SMA','Price vs 200 SMA','20/50 Crossover','50/200 Trend','MACD vs Signal','MACD Zero','RS vs SPY'], mrTechnicalRows_(market));

  row = mrSection_(sh, row, 'Recent News Catalysts');
  row = mrTable_(sh, row, ['Date','Headline / Catalyst','Severity','Why It Matters'], newsRows);

  row = mrSection_(sh, row, 'Upcoming Macro Catalysts to Watch');
  row = mrTable_(sh, row, ['Date','Time ET','Event','Impact','Days Until','Why It Matters'], mrMacroCatalystRows_(macroEvents));

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

  row = mrSection_(sh, row, 'Chicago-Style Source List');
  row = mrTable_(sh, row, ['Source','Use'], [
    ['Google Sheets GOOGLEFINANCE','Daily price history for SPY, sectors, and market proxies.'],
    ['Google News RSS','Recent market catalyst headlines when available.'],
    ['Connected E*TRADE Holdings tab','Portfolio quantities, prices, values, weights, and P&L.'],
    ['Macro event schedule logic','CPI, PPI, jobs, JOLTS, PCE, GDP, retail sales, ISM, and FOMC calendar items.'],
    ['SPDR sector ETF map','11-sector rotation framework.'],
    ['SPY benchmark','Relative-strength benchmark.']
  ]);

  mrFinalize_(sh, row);
  return 'Market Analysis complete. Report written to ' + MR_LOCAL.reportSheet + '. Market data status: ' + quality.ok + '/' + quality.total + ' tickers loaded. ' + quality.note;
}

function startMarketAnalysisReport(){ return buildMarketAnalysisReport(); }
function fetchMarketAnalysisReport(){ return 'No fetch needed. Local GOOGLEFINANCE engine writes the report immediately.'; }
function fetchLatestMarketAnalysisReport(){ return 'No fetch needed. Local GOOGLEFINANCE engine writes the report immediately.'; }
function tryFetchMarketAnalysisReport(){ return {ready:true, message:'No polling needed. Local report already writes directly.'}; }
function testMarketGitHubConnection(){ return 'GitHub is no longer used for market analysis. Use buildMarketAnalysisReport.'; }
function checkMarketGitHubConfig(){ return 'GitHub market config is no longer needed. Active engine is local GOOGLEFINANCE.'; }

function mrLoadGoogleFinanceHistories_(ss, symbols) {
  var sh = ss.getSheetByName(MR_LOCAL.cacheSheet) || ss.insertSheet(MR_LOCAL.cacheSheet);
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
    var okCount = symbols.filter(function(s){return histories[s] && histories[s].closes.length >= 60;}).length;
    if (okCount >= Math.max(5, Math.floor(symbols.length * 0.85))) break;
  }
  try { sh.hideSheet(); } catch(e) {}
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
    if (!h || h.closes.length < 60) out[sym] = mrBlankTech_(sym, h && h.error ? h.error : 'GOOGLEFINANCE not loaded / insufficient rows');
    else out[sym] = mrTech_(sym, h);
  });
  return out;
}

function mrTech_(sym, hist) {
  var c = hist.closes, price = mrLast_(c), ema20s = mrEmaSeries_(c, 20), ema20 = mrLast_(ema20s), sma50 = mrSma_(c, 50), sma200 = mrSma_(c, Math.min(200, c.length));
  var macd = mrMacd_(c), rsi14 = mrRsi_(c, 14), ret5 = mrRet_(c, 5), ret20 = mrRet_(c, 20), ret60 = mrRet_(c, 60), vol20 = mrVol_(c, 20), dd63 = mrDrawdown_(c, 63);
  var trendPts = (price>=ema20?1:0)+(price>=sma50?1:0)+(price>=sma200?1:0)+(ema20>=sma50?1:0)+(sma50>=sma200?1:0);
  var momPts = (macd.macd>=macd.signal?1:0)+(macd.macd>=0?1:0)+(ret20>0?1:0)+(ret60>0?1:0)+(rsi14>=50?1:0);
  var trendScore = trendPts/5, momentumScore = momPts/5, riskScore = Math.max(0, Math.min(1, 1 - (vol20/0.45) + (dd63/0.35)));
  var composite = 0.45*trendScore + 0.40*momentumScore + 0.15*riskScore;
  var read = composite>=0.78 ? 'Bullish confirmed' : composite>=0.58 ? 'Improving' : composite>=0.38 ? 'Mixed' : 'Weak';
  return {ticker:sym, price:price, ema20:ema20, sma50:sma50, sma200:sma200, macd:macd.macd, signal:macd.signal, hist:macd.hist, rsi14:rsi14, ret5d:ret5, ret20d:ret20, ret60d:ret60, vol20:vol20, drawdown63d:dd63, trendScore:trendScore, momentumScore:momentumScore, riskScore:riskScore, compositeScore:composite, source:hist.source, rows:c.length, priceVs20:price>=ema20?'Above':'Below', priceVs50:price>=sma50?'Above':'Below', priceVs200:price>=sma200?'Above':'Below', cross20_50:ema20>=sma50?'20 EMA > 50 SMA':'20 EMA < 50 SMA', trend50_200:sma50>=sma200?'50 SMA > 200 SMA':'50 SMA < 200 SMA', macdSignal:macd.macd>=macd.signal?'MACD > Signal':'MACD < Signal', macdZero:macd.macd>=0?'Above 0':'Below 0', read:read};
}

function mrBlankTech_(sym, reason) { return {ticker:sym, price:0, ema20:0, sma50:0, sma200:0, macd:0, signal:0, hist:0, rsi14:50, ret5d:0, ret20d:0, ret60d:0, vol20:0, drawdown63d:0, trendScore:0, momentumScore:0, riskScore:0, compositeScore:0, source:reason, rows:0, priceVs20:'n/a', priceVs50:'n/a', priceVs200:'n/a', cross20_50:'n/a', trend50_200:'n/a', macdSignal:'n/a', macdZero:'n/a', read:'Needs data'}; }
function mrQualitySummary_(symbols, histories, market){ var ok=0, failed=[]; symbols.forEach(function(s){ if(market[s] && market[s].read!=='Needs data') ok++; else failed.push(s); }); return {ok:ok, total:symbols.length, failed:failed, note: failed.length ? 'Missing: ' + failed.join(', ') : 'All sector/proxy tickers loaded.'}; }

function mrExecutiveRead_(m, rows){ var p=mrOverallPressure_(m), adds=mrCountAction_(rows,'Add-QTY'), trims=mrCountAction_(rows,'Trim-QTY'), sells=mrCountAction_(rows,'Sell-QTY'); return 'Current market pressure is rated ' + p.rating + ' because ' + p.reason + ' Portfolio action flags: ' + adds + ' add candidates, ' + trims + ' trim candidates, and ' + sells + ' sell/reduce candidates. The report is not an auto-trade instruction; it is a risk-managed review layer combining sector rotation, macro pressure, and your E*TRADE position data.'; }
function mrOverallPressure_(m){ var score=0, reasons=[]; if(m.SPY.priceVs50==='Below'){score+=2; reasons.push('SPY is below its 50 SMA');} if(m.HYG.priceVs50==='Below'||m.HYG.ret20d<-0.03){score+=2; reasons.push('credit proxy HYG is weak');} if(m.TLT.ret20d<-0.04){score+=2; reasons.push('long bonds are falling, pointing to yield pressure');} if(m.USO.ret20d>0.06){score+=1; reasons.push('oil is rising');} if(m.UUP.ret20d>0.03){score+=1; reasons.push('the dollar is firming');} if(m.SPY.read==='Bullish confirmed') score-=1; var rating=score>=5?'HIGH':score>=3?'MEDIUM':score>=1?'WATCH':'LOW'; return {rating:rating, score:score, reason:reasons.length?reasons.join('; '):'SPY trend and credit conditions are not showing major stress'}; }
function mrPressureRows_(m){ var p=mrOverallPressure_(m); return [
  ['Broad equity trend','SPY ' + m.SPY.priceVs50 + ' 50 SMA; SPY read: ' + m.SPY.read, p.rating, 'Controls how aggressive new equity adds should be.'],
  ['Rates / yield pressure','TLT 20d ' + mrPctText_(m.TLT.ret20d) + '; TLT read: ' + m.TLT.read, m.TLT.ret20d<-0.04?'HIGH':m.TLT.ret20d<-0.02?'MEDIUM':'LOW', 'Higher yield pressure can hurt growth and long-duration equities.'],
  ['Credit / liquidity','HYG 20d ' + mrPctText_(m.HYG.ret20d) + '; HYG ' + m.HYG.priceVs50 + ' 50 SMA', (m.HYG.priceVs50==='Below'||m.HYG.ret20d<-0.03)?'HIGH':m.HYG.ret20d<0?'MEDIUM':'LOW', 'Credit weakness usually lowers risk appetite.'],
  ['Oil / inflation impulse','USO 20d ' + mrPctText_(m.USO.ret20d), m.USO.ret20d>0.06?'HIGH':m.USO.ret20d>0.03?'MEDIUM':'LOW', 'Oil strength can revive inflation concern and hurt consumer sectors.'],
  ['Dollar pressure','UUP 20d ' + mrPctText_(m.UUP.ret20d), m.UUP.ret20d>0.03?'MEDIUM':'LOW', 'A stronger dollar can pressure ADRs, commodities, and multinational earnings.']
]; }
function mrMacroRows_(m) { return [
  ['War / Geopolitics','GLD 20d ' + mrPctText_(m.GLD.ret20d) + '; USO 20d ' + mrPctText_(m.USO.ret20d), (m.GLD.ret20d>0.03||m.USO.ret20d>0.05)?'Elevated watch':'Normal watch','Gold/defense matter more if risk shock rises.'],
  ['Oil / Energy shock','USO 20d ' + mrPctText_(m.USO.ret20d), m.USO.ret20d>0.05?'Oil pressure rising':m.USO.ret20d<-0.05?'Oil easing':'Neutral','Affects XLE, inflation expectations, and consumer pressure.'],
  ['Crisis / Liquidity','SPY ' + m.SPY.priceVs50 + ' 50 SMA; HYG 20d ' + mrPctText_(m.HYG.ret20d), m.SPY.priceVs50==='Below'||m.HYG.ret20d<-0.03?'Risk-off watch':'Risk-on acceptable','Controls how aggressive adds should be.'],
  ['Fed / Yields','TLT 20d ' + mrPctText_(m.TLT.ret20d), m.TLT.ret20d<-0.04?'Yield pressure rising':m.TLT.ret20d>0.04?'Yield pressure easing':'Neutral','Affects growth stocks, banks, real estate, and bond sleeves.'],
  ['Inflation / Dollar','UUP 20d ' + mrPctText_(m.UUP.ret20d) + '; USO 20d ' + mrPctText_(m.USO.ret20d), m.UUP.ret20d>0.03||m.USO.ret20d>0.05?'Tighter impulse':'Contained','Affects ADRs, gold, and valuation multiples.'],
  ['Jobs / Growth','SPY technical read: ' + m.SPY.read, m.SPY.compositeScore>=0.58?'Growth trend constructive':'Growth trend mixed/weak','Confirms whether rotation supports risk assets.']
]; }

function mrSectorRows_(m) { var spy=m.SPY; return MR_SECTORS.map(function(x){ var t=m[x[0]], rot=mrRotation_(t, spy); return [x[0], x[1], rot, mrFlow_(t, spy) + ' / ' + mrRS_(t, spy), mrBias_(rot)]; }); }
function mrTechnicalRows_(m) { var spy=m.SPY; return MR_SECTORS.map(function(x){ var t=m[x[0]]; return [x[0], x[1], t.priceVs20, t.priceVs50, t.priceVs200, t.cross20_50, t.trend50_200, t.macdSignal, t.macdZero, mrRS_(t, spy)]; }); }
function mrRotationNarrative_(m){ var spy=m.SPY, entering=[], exiting=[], mixed=[]; MR_SECTORS.forEach(function(x){ var t=m[x[0]], rot=mrRotation_(t, spy), rs=mrRS_(t, spy); var label=x[0]+' ('+x[1]+')'; if((rot==='Leadership'||rot==='Positive trend') && rs!=='Underperforming SPY') entering.push(label); else if(rot==='Lagging / weak'||rs==='Underperforming SPY') exiting.push(label); else mixed.push(label); }); return 'Based on the 11 SPDR sector ETFs, market rotation appears to be entering or favoring: ' + (entering.length?entering.join(', '):'no clear sector leadership') + '. It appears to be exiting, avoiding, or underweighting: ' + (exiting.length?exiting.join(', '):'no clear exit group') + '. Mixed/neutral areas: ' + (mixed.length?mixed.join(', '):'none') + '. This paragraph is built from price versus moving averages, MACD confirmation, and relative strength versus SPY.'; }
function mrRotation_(t, spy){ if(t.read==='Needs data') return 'Needs data'; if(t.compositeScore>=0.75 && t.ret20d>spy.ret20d) return 'Leadership'; if(t.compositeScore>=0.60) return 'Positive trend'; if(t.compositeScore>=0.45 && t.ret20d>=spy.ret20d) return 'Improving'; if(t.compositeScore<0.35) return 'Lagging / weak'; return 'Mixed'; }
function mrFlow_(t, spy){ if(t.read==='Needs data'||spy.read==='Needs data') return 'Needs data'; if(t.ret20d>spy.ret20d+0.02&&t.compositeScore>=0.55) return 'Positive pressure'; if(t.ret20d<spy.ret20d-0.02&&t.compositeScore<=0.50) return 'Negative pressure'; return 'Mixed pressure'; }
function mrRS_(t, spy){ if(t.read==='Needs data'||spy.read==='Needs data') return 'Needs data'; var d=t.ret20d-spy.ret20d; if(d>0.02) return 'Outperforming SPY'; if(d<-0.02) return 'Underperforming SPY'; return 'In line with SPY'; }
function mrBias_(r){ if(r==='Leadership') return 'Best add/hold candidates after macro confirmation'; if(r==='Positive trend') return 'Hold / selective add only'; if(r==='Improving') return 'Watchlist / starter only'; if(r==='Lagging / weak') return 'Avoid adds / review trims'; return 'Data needed / wait'; }
function mrPriorWeekText_(m){ var spy=m.SPY, leaders=[], lag=[]; MR_SECTORS.forEach(function(x){ var rot=mrRotation_(m[x[0]], spy); if(rot==='Leadership') leaders.push(x[0]+' '+x[1]); if(rot==='Lagging / weak') lag.push(x[0]+' '+x[1]); }); return 'Forward trend: SPY is ' + spy.read + ' with 20-day change of ' + mrPctText_(spy.ret20d) + '. Confirmed technical leadership: ' + (leaders.length?leaders.join(', '):'none') + '. Weak/lagging sectors: ' + (lag.length?lag.slice(0,4).join(', '):'none') + '. Use macro dates and news catalysts above before changing position sizes.'; }

function mrBuildPortfolioRows_(holdings, market) { var weights={}; holdings.forEach(function(h){ var meta=mrMeta_(h); weights[meta[0]]=(weights[meta[0]]||0)+h.weight; }); return holdings.map(function(h){ var meta=mrMeta_(h), tech=market[meta[1]]||market.SPY, tol=mrTolerance_(h, meta, tech, weights), act=mrAction_(h, meta, tech, tol); return [h.ticker, h.qty, mrMoney_(h.cost), mrMoney_(h.value), mrMoney_(h.pnl), mrPctText_(h.pnlPct), tol, meta[2], act[0], act[1], mrMoney_(act[2]), act[3] + ' Sector proxy ' + meta[1] + ' = ' + tech.read + ' (' + tech.compositeScore.toFixed(2) + ').', meta[0]]; }).sort(function(a,b){ return mrActionPriority_(a[8]) - mrActionPriority_(b[8]) || mrNum_(b[3]) - mrNum_(a[3]); }); }
function mrMeta_(h){ if(MR_TICKER_MAP[h.ticker]) return MR_TICKER_MAP[h.ticker]; if(String(h.type).indexOf('etf')>=0) return ['ETF / Unmapped','SPY','ETF exposure; map sleeve manually if material.']; return ['Equity / Unmapped','SPY','Single-stock exposure; validate thesis manually.']; }
function mrTolerance_(h, meta, t, weights){ var sleeve=meta[0], sw=weights[sleeve]||0, safety=sleeve.indexOf('Safety')>=0||sleeve.indexOf('Income')>=0||sleeve.indexOf('Credit')>=0; if(safety&&sw>0.45) return 'Overweight safety'; if(safety&&sw>0.18) return 'Safety overlap'; if(safety) return 'Income tilt'; if(h.pnlPct>0.50) return 'Profit outlier'; if(h.pnlPct<-0.10&&t.compositeScore<0.45) return 'Weak'; if(h.weight<0.015) return 'Too small'; if(h.weight>0.08||sw>0.18) return 'Near limit'; if(t.compositeScore<0.35&&t.read!=='Needs data') return 'Out of tolerance'; return 'In tolerance'; }
function mrAction_(h, meta, t, tol){ var price=h.price||(h.value/h.qty)||0, qty=Math.max(0,Math.floor(h.qty)), safety=meta[0].indexOf('Safety')>=0||meta[0].indexOf('Income')>=0||meta[0].indexOf('Credit')>=0; if(qty<=0) return ['Hold',0,0,'No share quantity available. Price source: E*TRADE institution price/value.']; if(t.read==='Needs data') return ['Hold',0,0,'No market technical data yet. Price source: E*TRADE institution price/value.']; if(tol==='Profit outlier'&&t.compositeScore<0.65){var q=Math.max(1,Math.floor(qty*0.10)); return ['Trim-QTY '+q,q,q*price,'Profit outlier and mapped trend is not leadership. Estimate uses E*TRADE institution price/value.'];} if((tol==='Weak'||tol==='Out of tolerance')&&t.compositeScore<0.40&&!safety){var q2=Math.max(1,Math.floor(qty*0.15)); return ['Sell-QTY '+q2,q2,q2*price,'Weak mapped trend plus tolerance pressure. Estimate uses E*TRADE institution price/value.'];} if((tol==='Overweight safety'||tol==='Safety overlap')&&safety){var q3=Math.max(1,Math.floor(qty*0.05)); return ['Trim-QTY '+q3,q3,q3*price,'Safety sleeve overlap; trim only if reallocating to confirmed leadership. Estimate uses E*TRADE institution price/value.'];} if(tol==='Too small'&&t.compositeScore>=0.75&&!safety){var q4=price>=100?1:Math.max(1,Math.floor(500/Math.max(price,1))); return ['Add-QTY '+q4,q4,q4*price,'Small position with mapped technical leadership. Estimate uses E*TRADE institution price/value.'];} return ['Hold',0,0,'Hold-no-add until macro/news confirms. Price source: E*TRADE institution price/value.']; }
function mrActionPriority_(s){ if(String(s).indexOf('Add-QTY')===0) return 1; if(String(s).indexOf('Trim-QTY')===0) return 2; if(String(s).indexOf('Sell-QTY')===0) return 3; if(String(s).indexOf('Avoid')===0) return 4; return 9; }
function mrTotalAction_(rows, prefix){ return rows.reduce(function(sum,r){ return String(r[8]).indexOf(prefix)===0 ? sum + mrNum_(r[10]) : sum; },0); }
function mrCountAction_(rows, prefix){ return rows.filter(function(r){return String(r[8]).indexOf(prefix)===0;}).length; }

function mrBuildMacroEvents_(){ var events=[]; events=events.concat(mrBlsMacroEvents_()).concat(mrFomcEvents_()).concat(mrIsmEvents_()).concat(mrEstimatedMacroEvents_()); return mrNormalizeMacroEvents_(events); }
function mrMacroCatalystRows_(events){ return events.slice(0,14).map(function(e){ return [e.dateText, e.time, e.event, e.impact, e.daysUntil, e.why]; }); }
function mrCreateMacro_(event, dateObj, time, impact, category, source, why){ return {event:event, dateObj:mrStripTime_(dateObj), time:time, impact:impact, category:category, source:source, why:why}; }
function mrBlsMacroEvents_(){ var rows=[
  [2026,6,2,'8:30 AM ET','Jobs Report / Nonfarm Payrolls','HIGH','Labor','Labor strength changes Fed-rate expectations and can cause large index moves.'],
  [2026,6,14,'8:30 AM ET','CPI Inflation','HIGH','Inflation','CPI is one of the biggest inflation reports for yields, QQQ, and broad equities.'],
  [2026,6,15,'8:30 AM ET','PPI Inflation','HIGH','Inflation','PPI shows producer inflation pressure and can move rate-cut expectations.'],
  [2026,7,4,'10:00 AM ET','JOLTS','MEDIUM','Labor','JOLTS shows labor demand; hot openings can pressure rate-cut expectations.'],
  [2026,7,7,'8:30 AM ET','Jobs Report / Nonfarm Payrolls','HIGH','Labor','Labor strength changes Fed-rate expectations and can cause large index moves.'],
  [2026,7,12,'8:30 AM ET','CPI Inflation','HIGH','Inflation','CPI is one of the biggest inflation reports for yields, QQQ, and broad equities.'],
  [2026,7,13,'8:30 AM ET','PPI Inflation','HIGH','Inflation','PPI shows producer inflation pressure and can move rate-cut expectations.'],
  [2026,8,1,'10:00 AM ET','JOLTS','MEDIUM','Labor','JOLTS shows labor demand; hot openings can pressure rate-cut expectations.'],
  [2026,8,4,'8:30 AM ET','Jobs Report / Nonfarm Payrolls','HIGH','Labor','Labor strength changes Fed-rate expectations and can cause large index moves.'],
  [2026,8,10,'8:30 AM ET','PPI Inflation','HIGH','Inflation','PPI shows producer inflation pressure and can move rate-cut expectations.'],
  [2026,8,11,'8:30 AM ET','CPI Inflation','HIGH','Inflation','CPI is one of the biggest inflation reports for yields, QQQ, and broad equities.'],
  [2026,8,29,'10:00 AM ET','JOLTS','MEDIUM','Labor','JOLTS shows labor demand; hot openings can pressure rate-cut expectations.'],
  [2026,9,2,'8:30 AM ET','Jobs Report / Nonfarm Payrolls','HIGH','Labor','Labor strength changes Fed-rate expectations and can cause large index moves.'],
  [2026,9,14,'8:30 AM ET','CPI Inflation','HIGH','Inflation','CPI is one of the biggest inflation reports for yields, QQQ, and broad equities.'],
  [2026,9,15,'8:30 AM ET','PPI Inflation','HIGH','Inflation','PPI shows producer inflation pressure and can move rate-cut expectations.']
]; return rows.map(function(x){ return mrCreateMacro_(x[4], new Date(x[0],x[1],x[2]), x[3], x[5], x[6], 'BLS schedule logic', x[7]); }); }
function mrFomcEvents_(){ var rows=[[2026,6,29,'FOMC Rate Decision'],[2026,8,16,'FOMC Rate Decision + SEP'],[2026,9,28,'FOMC Rate Decision'],[2026,11,9,'FOMC Rate Decision + SEP']]; return rows.map(function(x){ return mrCreateMacro_(x[3], new Date(x[0],x[1],x[2]), '2:00 PM ET', 'HIGH', 'Fed / Rates', 'Fed schedule logic', x[3].indexOf('SEP')>=0?'Fed decision plus dot plot/SEP can move yields, QQQ, and volatility.':'Fed decision can move yields, QQQ, and volatility.'); }); }
function mrIsmEvents_(){ var events=[], today=mrToday_(), end=mrAddDays_(today, MR_LOCAL.macroDaysForward), start=new Date(today.getFullYear(), today.getMonth(), 1); for(var i=0;i<8;i++){ var d0=new Date(start.getFullYear(), start.getMonth()+i, 1); if(d0>end) break; events.push(mrCreateMacro_('ISM Manufacturing PMI', mrFirstBusinessDay_(d0.getFullYear(),d0.getMonth()), '10:00 AM ET', 'MEDIUM', 'Growth / PMI', 'Rule-based ISM schedule', 'Early manufacturing growth signal; weak prints can pressure cyclicals and risk appetite.')); events.push(mrCreateMacro_('ISM Services PMI', mrNthBusinessDay_(d0.getFullYear(),d0.getMonth(),3), '10:00 AM ET', 'MEDIUM', 'Growth / PMI', 'Rule-based ISM schedule', 'Services are a major part of the U.S. economy; surprise prints can move yields and growth stocks.')); } return events; }
function mrEstimatedMacroEvents_(){ var events=[], today=mrToday_(), end=mrAddDays_(today, MR_LOCAL.macroDaysForward), start=new Date(today.getFullYear(), today.getMonth(), 1); for(var i=0;i<8;i++){ var mdate=new Date(start.getFullYear(), start.getMonth()+i, 1); if(mdate>end) break; var y=mdate.getFullYear(), m=mdate.getMonth(); events.push(mrCreateMacro_('Retail Sales', mrNthWeekday_(y,m,4,2), '8:30 AM ET', 'MEDIUM', 'Consumer', 'Estimated macro schedule', 'Consumer spending strength can support or weaken risk-on equity moves.')); events.push(mrCreateMacro_('PCE Inflation / Personal Income', mrLastBusinessDay_(y,m), '8:30 AM ET', 'HIGH', 'Inflation', 'Estimated macro schedule', 'PCE is the Fed preferred inflation gauge and can move yields and growth stocks.')); if([0,3,6,9].indexOf(m)!==-1) events.push(mrCreateMacro_('GDP Report', mrNthWeekday_(y,m,4,4), '8:30 AM ET', 'HIGH', 'Growth', 'Estimated macro schedule', 'GDP confirms growth or slowdown risk and can change broad-market sentiment.')); } return events; }
function mrNormalizeMacroEvents_(events){ var today=mrToday_(), end=mrAddDays_(today, MR_LOCAL.macroDaysForward), kept={}, out=[]; events.forEach(function(e){ if(!e||!(e.dateObj instanceof Date)||isNaN(e.dateObj.getTime())) return; if(e.dateObj<today||e.dateObj>end) return; e.daysUntil=mrDaysBetween_(today,e.dateObj); e.dateText=mrFormatDate_(e.dateObj); var key=e.event+'|'+Utilities.formatDate(e.dateObj, Session.getScriptTimeZone(), 'yyyy-MM'); if(!kept[key]){ kept[key]=true; out.push(e); } }); out.sort(function(a,b){ return a.dateObj-b.dateObj || mrImpactRank_(a.impact)-mrImpactRank_(b.impact); }); return out; }
function mrImpactRank_(x){ x=String(x||'').toUpperCase(); return x==='HIGH'?1:x==='MEDIUM'?2:x==='WATCH'?3:4; }

function mrNewsCatalystRows_(market){ var rows=[]; try{ var url='https://news.google.com/rss/search?q=' + encodeURIComponent('stock market Fed inflation yields oil earnings') + '&hl=en-US&gl=US&ceid=US:en'; var xml=UrlFetchApp.fetch(url,{muteHttpExceptions:true,followRedirects:true}).getContentText(); var items=xml.match(/<item>[\s\S]*?<\/item>/g)||[]; items.slice(0,MR_LOCAL.newsMaxItems).forEach(function(item){ var title=mrXml_(item,'title'), date=mrXml_(item,'pubDate'), sev=mrHeadlineSeverity_(title); if(title) rows.push([mrNewsDate_(date), title, sev, mrNewsWhy_(title, sev)]); }); }catch(e){} if(!rows.length){ var p=mrOverallPressure_(market); rows.push([mrFormatDate_(new Date()), 'Live news feed unavailable inside Apps Script; using market-proxy pressure model instead.', p.rating, p.reason]); } return rows; }
function mrXml_(xml, tag){ var re=new RegExp('<'+tag+'>([\\s\\S]*?)<\\/'+tag+'>','i'), m=String(xml||'').match(re); return m?mrDecode_(m[1]).replace(/<!\[CDATA\[|\]\]>/g,'').trim():''; }
function mrHeadlineSeverity_(title){ var s=String(title||'').toLowerCase(); if(/fed|fomc|cpi|inflation|jobs|payroll|yields|treasury|oil|war|geopolitical|tariff|recession|credit/.test(s)) return 'HIGH'; if(/earnings|guidance|dollar|consumer|retail|gdp|pce|ppi/.test(s)) return 'MEDIUM'; return 'WATCH'; }
function mrNewsWhy_(title, sev){ if(sev==='HIGH') return 'High-severity headline because it can directly affect yields, inflation expectations, credit risk, or broad equity multiples.'; if(sev==='MEDIUM') return 'Medium-severity headline because it can affect sector rotation, earnings expectations, or risk appetite.'; return 'Watch item; useful context but not enough alone to drive a portfolio action.'; }
function mrNewsDate_(d){ var x=new Date(d); return isNaN(x.getTime()) ? mrFormatDate_(new Date()) : mrFormatDate_(x); }
function mrDecode_(s){ return String(s||'').replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>'); }

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
function mrNum_(x){ if(x===null||x===undefined||x==='') return 0; if(typeof x==='number') return x; var s=String(x); var neg=s.indexOf('(')>=0&&s.indexOf(')')>=0; s=s.replace(/[,$%\s()]/g,''); var n=Number(s); if(isNaN(n)) return 0; return neg?-n:n; }
function mrPctNum_(x){ if(typeof x==='number') return Math.abs(x)>1?x/100:x; var s=String(x||''); var n=mrNum_(s); return s.indexOf('%')>=0?n/100:n; }
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

function mrTitle_(sh,row,title,sub){ sh.getRange(row,1,1,13).merge().setValue(title).setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#111827').setFontColor('#ffffff'); sh.setRowHeight(row,38); row++; sh.getRange(row,1,1,13).merge().setValue(sub).setFontSize(12).setFontColor('#374151').setBackground('#eef2ff').setWrap(true); sh.setRowHeight(row,34); return row+2; }
function mrSection_(sh,row,title){ sh.getRange(row,1,1,13).merge().setValue(title).setFontSize(15).setFontWeight('bold').setBackground('#dbeafe').setFontColor('#111827').setHorizontalAlignment('left').setVerticalAlignment('middle'); sh.setRowHeight(row,30); return row+1; }
function mrParagraph_(sh,row,text){ sh.getRange(row,1,1,13).merge().setValue(text).setFontSize(12).setBackground('#ffffff').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('left'); sh.setRowHeight(row,74); return row+2; }
function mrTable_(sh,row,headers,rows){ rows=rows||[]; var width=Math.max(headers.length,1), data=[headers].concat(rows); var clean=data.map(function(r){var o=[]; for(var i=0;i<width;i++) o.push(r[i]===undefined||r[i]===null?'':r[i]); return o;}); var range=sh.getRange(row,1,clean.length,width); range.setValues(clean).setWrap(true).setVerticalAlignment('top').setBorder(true,true,true,true,true,true,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID); sh.getRange(row,1,1,width).setFontWeight('bold').setFontSize(11).setBackground('#1f2937').setFontColor('#ffffff').setHorizontalAlignment('center'); if(clean.length>1) sh.getRange(row+1,1,clean.length-1,width).setBackground('#ffffff').setFontColor('#111827').setFontSize(12).setHorizontalAlignment('left'); for(var r=0;r<clean.length;r++) sh.setRowHeight(row+r, r===0?28:46); return row+clean.length+2; }
function mrFinalize_(sh,lastRow){ var widths=[140,155,130,130,130,105,160,320,150,95,130,380,180]; for(var c=1;c<=13;c++) sh.setColumnWidth(c,widths[c-1]); sh.getRange(1,1,Math.max(1,lastRow),Math.min(13,sh.getMaxColumns())).setFontFamily(MR_LOCAL.font); sh.autoResizeRows(1,Math.max(1,lastRow)); sh.setFrozenRows(0); }
