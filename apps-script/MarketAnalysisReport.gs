/**
 * Portfolio Link Market Analysis - Fast Portrait Engine
 * Local only: Holdings tab -> cached GOOGLEFINANCE data -> portrait report.
 * Keeps Google News RSS, but removes the slow extra cleanup pass.
 */

var MR = {
  holdingsSheet: 'Holdings',
  reportSheet: 'Report Market Analysis',
  cacheSheet: 'Market Data Cache',
  font: 'Times New Roman',
  lookbackDays: 420,
  minRows: 60,
  reportCols: 6,
  newsMaxItems: 4
};

var MR_SECTORS = [
  ['XLC','Communication Services'], ['XLY','Consumer Discretionary'], ['XLP','Consumer Staples'],
  ['XLE','Energy'], ['XLF','Financials'], ['XLV','Healthcare'], ['XLI','Industrials'],
  ['XLB','Materials'], ['XLRE','Real Estate'], ['XLK','Technology'], ['XLU','Utilities']
];

var MR_PROXIES = ['SPY','QQQ','DIA','IWM','GLD','USO','TLT','UUP','HYG','BIL'];

var MR_GF = {
  SPY:'NYSEARCA:SPY', QQQ:'NASDAQ:QQQ', DIA:'NYSEARCA:DIA', IWM:'NYSEARCA:IWM', GLD:'NYSEARCA:GLD',
  USO:'NYSEARCA:USO', TLT:'NASDAQ:TLT', UUP:'NYSEARCA:UUP', HYG:'NYSEARCA:HYG', BIL:'NYSEARCA:BIL',
  XLC:'NYSEARCA:XLC', XLY:'NYSEARCA:XLY', XLP:'NYSEARCA:XLP', XLE:'NYSEARCA:XLE', XLF:'NYSEARCA:XLF',
  XLV:'NYSEARCA:XLV', XLI:'NYSEARCA:XLI', XLB:'NYSEARCA:XLB', XLRE:'NYSEARCA:XLRE', XLK:'NYSEARCA:XLK', XLU:'NYSEARCA:XLU'
};

var MR_MAP = {
  SPYM:['Core Equity','SPY','Low-cost S&P 500 core exposure.'], DIA:['Core Equity','DIA','Dow blue-chip exposure.'],
  SCHG:['Growth Equity','XLK','Large-cap growth tilt; sensitive to rates and technology leadership.'],
  SPMO:['Momentum Equity','SPY','Momentum factor exposure tied to risk appetite.'],
  BAC:['Financials','XLF','Bank exposure; sensitive to rates, credit, and yield curve.'],
  MS:['Financials','XLF','Capital markets and wealth-management exposure.'],
  STT:['Financials','XLF','Custody bank / asset-servicing exposure.'],
  SONY:['Consumer / ADR','XLY','Consumer technology, gaming, media, and ADR exposure.'],
  LMT:['Defense / Industrials','XLI','Defense industrial; can act as geopolitical hedge.'],
  HTD:['Income Equity','XLU','Dividend-income sleeve with utility/financial income profile.'],
  SGOL:['Gold / Alternative','GLD','Gold hedge against real-rate, dollar, and geopolitical stress.'],
  JPST:['Short Duration Safety','BIL','Ultra-short income stabilizer.'],
  VRIG:['Floating Rate Safety','BIL','Floating-rate investment-grade income stabilizer.'],
  CLOZ:['Credit Income','HYG','CLO credit-income sleeve; sensitive to credit spreads.']
};

function buildMarketAnalysisReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var holdings = mrReadHoldings_(ss);
  if (!holdings.length) throw new Error('No usable ticker holdings found. Pull Holdings first.');

  var symbols = mrUnique_(MR_PROXIES.concat(MR_SECTORS.map(function(x){ return x[0]; })));
  var histories = mrGetHistoriesFast_(ss, symbols);
  var market = mrBuildMarket_(symbols, histories);
  var portfolio = mrPortfolio_(holdings, market);
  var quality = mrQuality_(symbols, market);
  var newsRows = mrNewsRows_(market);

  var sh = ss.getSheetByName(MR.reportSheet) || ss.insertSheet(MR.reportSheet);
  mrReset_(sh);

  var row = 1;
  row = mrTitle_(sh, row, 'Raj Market Rotation Report', 'Built ' + mrNow_() + ' | Fast portrait layout | Cache-aware market engine');
  row = mrSection_(sh, row, 'Executive Read');
  row = mrPara_(sh, row, mrExecutive_(market, portfolio), 58);
  row = mrSection_(sh, row, 'Current News Catalysts');
  row = mrTable_(sh, row, ['Date','Headline / Catalyst','Severity','Why It Matters'], newsRows, [1,2,1,2]);
  row = mrSection_(sh, row, 'Market Pressure Rating');
  row = mrTable_(sh, row, ['Area','Evidence','Severity','Meaning'], mrPressureRows_(market), [1,1,1,3]);
  row = mrSection_(sh, row, 'Market Rotation Narrative');
  row = mrPara_(sh, row, mrRotationText_(market), 66);
  row = mrSection_(sh, row, 'Macro Risk Dashboard');
  row = mrTable_(sh, row, ['Risk','Read','Level','Portfolio Meaning'], mrMacroRows_(market), [1,1,1,3]);
  row = mrSection_(sh, row, 'Sector Rotation - SPDR Map');
  row = mrTable_(sh, row, ['ETF','Sector','Rotation','Pressure','Bias'], mrSectorRows_(market), [1,1,1,1,2]);
  row = mrSection_(sh, row, 'Technical Confirmation Snapshot');
  row = mrTable_(sh, row, ['ETF','Sector','Trend','MA Stack','MACD','RS'], mrTechRows_(market), [1,1,1,1,1,1]);
  row = mrSection_(sh, row, 'Upcoming Macro Catalysts to Watch');
  row = mrTable_(sh, row, ['Date','Time','Event','Impact','Days','Why It Matters'], mrMacroCatalystRows_(), [1,1,1,1,1,1]);
  row = mrSection_(sh, row, 'Prior Week Recap + Forward Trend');
  row = mrPara_(sh, row, mrForwardText_(market), 58);
  row = mrSection_(sh, row, 'Raj Portfolio Impact - Exact Actions');
  row = mrPortfolioCards_(sh, row, portfolio);
  row = mrSection_(sh, row, 'Total Suggested Actions + Primary Goal');
  row = mrTable_(sh, row, ['Metric','Value'], [
    ['Total suggested trims', mrMoney_(mrTotalAction_(portfolio, 'Trim-QTY'))],
    ['Total estimated trim P&L', mrMoney_(mrTotalTrimPnl_(portfolio))],
    ['Total suggested adds', mrMoney_(mrTotalAction_(portfolio, 'Add-QTY'))],
    ['Primary goal', 'Keep actions small; reduce overlap; add only when macro, sector, and technical confirmation line up.']
  ], [1,5]);

  mrFinal_(sh, row);
  return 'Market Analysis complete. Fast report written. Market data loaded: ' + quality.ok + '/' + quality.total + '. ' + quality.note;
}

function buildMarketAnalysisReportPrint(){ return buildMarketAnalysisReport(); }
function startMarketAnalysisReport(){ return buildMarketAnalysisReport(); }
function fetchMarketAnalysisReport(){ return 'No fetch needed. Fast local engine writes the report directly.'; }
function fetchLatestMarketAnalysisReport(){ return 'No fetch needed. Fast local engine writes the report directly.'; }
function tryFetchMarketAnalysisReport(){ return {ready:true, message:'Fast local report writes directly.'}; }
function testMarketGitHubConnection(){ return 'GitHub is no longer used for market analysis. Use buildMarketAnalysisReport.'; }
function checkMarketGitHubConfig(){ return 'GitHub market config is no longer needed. Active engine is local GOOGLEFINANCE.'; }

function mrGetHistoriesFast_(ss, symbols) {
  var sh = ss.getSheetByName(MR.cacheSheet) || ss.insertSheet(MR.cacheSheet);
  try { sh.showSheet(); } catch(e) {}
  var existing = mrReadCache_(sh, symbols);
  if (mrCountOk_(existing, symbols) >= Math.max(6, Math.floor(symbols.length * 0.70))) {
    try { sh.hideSheet(); } catch(e1) {}
    return existing;
  }
  mrRebuildCache_(sh, symbols);
  var histories = {};
  var waits = [5000, 4000, 3000, 3000, 3000];
  for (var i = 0; i < waits.length; i++) {
    Utilities.sleep(waits[i]);
    SpreadsheetApp.flush();
    histories = mrReadCache_(sh, symbols);
    if (mrCountOk_(histories, symbols) >= Math.max(6, Math.floor(symbols.length * 0.70))) break;
  }
  try { sh.hideSheet(); } catch(e2) {}
  return histories;
}

function mrRebuildCache_(sh, symbols) {
  sh.clear();
  sh.setHiddenGridlines(true);
  var neededCols = symbols.length * 3;
  if (sh.getMaxColumns() < neededCols) sh.insertColumnsAfter(sh.getMaxColumns(), neededCols - sh.getMaxColumns());
  if (sh.getMaxRows() < 460) sh.insertRowsAfter(sh.getMaxRows(), 460 - sh.getMaxRows());
  symbols.forEach(function(sym, i){
    var col = 1 + i * 3;
    sh.getRange(1, col).setValue(sym);
    sh.getRange(2, col).setFormula('=GOOGLEFINANCE("' + (MR_GF[sym] || sym) + '","close",TODAY()-' + MR.lookbackDays + ',TODAY(),"DAILY")');
  });
  SpreadsheetApp.flush();
}

function mrReadCache_(sh, symbols) {
  var out = {};
  var maxRows = sh.getMaxRows();
  symbols.forEach(function(sym, i){
    var col = 1 + i * 3;
    var tag = String(sh.getRange(1, col).getDisplayValue() || '').trim().toUpperCase();
    if (tag !== sym) { out[sym] = {closes:[], error:'cache missing'}; return; }
    var n = Math.min(450, Math.max(1, maxRows - 1));
    var values = sh.getRange(2, col, n, 2).getValues();
    var display = sh.getRange(2, col, n, 2).getDisplayValues();
    var closes = [], error = '';
    for (var r = 0; r < values.length; r++) {
      if (!error && (String(display[r][0]).indexOf('#') === 0 || String(display[r][1]).indexOf('#') === 0)) error = display[r][0] || display[r][1];
      if (values[r][0] instanceof Date && typeof values[r][1] === 'number' && !isNaN(values[r][1])) closes.push(values[r][1]);
    }
    out[sym] = {closes:closes, error:error || '', source:'GOOGLEFINANCE cache'};
  });
  return out;
}
function mrCountOk_(histories, symbols){ return symbols.filter(function(s){ return histories[s] && histories[s].closes && histories[s].closes.length >= MR.minRows; }).length; }

function mrBuildMarket_(symbols, histories){ var out={}; symbols.forEach(function(s){ var h=histories[s]; out[s]=(!h||!h.closes||h.closes.length<MR.minRows)?mrBlank_(s,h&&h.error?h.error:'not enough data'):mrTech_(s,h.closes); }); return out; }
function mrTech_(sym,c){ var price=mrLast_(c), ema20=mrLast_(mrEma_(c,20)), sma50=mrSma_(c,50), sma200=mrSma_(c,Math.min(200,c.length)), macd=mrMacd_(c), rsi=mrRsi_(c,14), r20=mrRet_(c,20), r60=mrRet_(c,60), vol=mrVol_(c,20), dd=mrDrawdown_(c,63); var trend=(price>=ema20?1:0)+(price>=sma50?1:0)+(price>=sma200?1:0)+(ema20>=sma50?1:0)+(sma50>=sma200?1:0); var mom=(macd.macd>=macd.signal?1:0)+(macd.macd>=0?1:0)+(r20>0?1:0)+(r60>0?1:0)+(rsi>=50?1:0); var ts=trend/5, ms=mom/5, rs=Math.max(0,Math.min(1,1-(vol/0.45)+(dd/0.35))), comp=0.45*ts+0.40*ms+0.15*rs; return {ticker:sym, price:price, ret20d:r20, ret60d:r60, compositeScore:comp, read:comp>=0.78?'Bullish confirmed':comp>=0.58?'Improving':comp>=0.38?'Mixed':'Weak', priceVs20:price>=ema20?'Above 20':'Below 20', priceVs50:price>=sma50?'Above 50':'Below 50', priceVs200:price>=sma200?'Above 200':'Below 200', macdSignal:macd.macd>=macd.signal?'MACD>Sig':'MACD<Sig', macdZero:macd.macd>=0?'Above 0':'Below 0'}; }
function mrBlank_(sym,err){ return {ticker:sym, price:0, ret20d:0, ret60d:0, compositeScore:0, read:'Needs data', priceVs20:'n/a', priceVs50:'n/a', priceVs200:'n/a', macdSignal:'n/a', macdZero:'n/a', error:err}; }
function mrQuality_(symbols,market){ var ok=0, missing=[]; symbols.forEach(function(s){ if(market[s] && market[s].read !== 'Needs data') ok++; else missing.push(s); }); return {ok:ok,total:symbols.length,note:missing.length?'Missing/slow: '+missing.join(', '):'All proxy data loaded.'}; }

function mrNewsRows_(market){
  var rows = [];
  try {
    var q = 'stock market Fed inflation yields oil earnings OR stocks';
    var url = 'https://news.google.com/rss/search?q=' + encodeURIComponent(q) + '&hl=en-US&gl=US&ceid=US:en';
    var xml = UrlFetchApp.fetch(url, {muteHttpExceptions:true, followRedirects:true}).getContentText();
    var items = xml.match(/<item>[\s\S]*?<\/item>/g) || [];
    items.slice(0, MR.newsMaxItems).forEach(function(item){
      var title = mrClip_(mrXml_(item,'title'), 105);
      var pub = mrXml_(item,'pubDate');
      var sev = mrHeadlineSeverity_(title);
      if (title) rows.push([mrNewsDate_(pub), title, sev, mrNewsWhy_(sev)]);
    });
  } catch(e) {}
  if (!rows.length) {
    var p = mrPressure_(market);
    rows.push([mrDate_(new Date()), 'Google News RSS unavailable during this run; using market-proxy pressure model.', p.rating, mrClip_(p.reason, 115)]);
  }
  return rows;
}
function mrXml_(xml, tag){ var re=new RegExp('<'+tag+'>([\\s\\S]*?)<\\/'+tag+'>','i'), m=String(xml||'').match(re); return m?mrDecode_(m[1]).replace(/<!\[CDATA\[|\]\]>/g,'').trim():''; }
function mrHeadlineSeverity_(title){ var s=String(title||'').toLowerCase(); if(/fed|fomc|cpi|inflation|jobs|payroll|yields|treasury|oil|war|geopolitical|tariff|recession|credit/.test(s)) return 'HIGH'; if(/earnings|guidance|dollar|consumer|retail|gdp|pce|ppi/.test(s)) return 'MEDIUM'; return 'WATCH'; }
function mrNewsWhy_(sev){ return sev==='HIGH'?'Can directly affect yields, inflation expectations, credit risk, or broad equity multiples.':sev==='MEDIUM'?'Can affect sector rotation, earnings expectations, or risk appetite.':'Context item; not enough alone to drive portfolio action.'; }
function mrNewsDate_(d){ var x=new Date(d); return isNaN(x.getTime())?mrDate_(new Date()):mrDate_(x); }
function mrDecode_(s){ return String(s||'').replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>'); }
function mrClip_(s,n){ s=String(s||''); return s.length>n?s.slice(0,n-1)+'…':s; }

function mrExecutive_(m,p){ var pressure=mrPressure_(m), adds=mrCountAction_(p,'Add-QTY'), trims=mrCountAction_(p,'Trim-QTY'); return 'Market pressure is ' + pressure.rating + ' because ' + pressure.reason + '. Portfolio flags show ' + adds + ' add candidates and ' + trims + ' trim candidates. This is a risk review layer, not an automatic trade instruction.'; }
function mrPressure_(m){ var s=0, r=[]; if(m.SPY.priceVs50==='Below 50'){s+=2;r.push('SPY below 50 SMA');} if(m.HYG.priceVs50==='Below 50'||m.HYG.ret20d<-0.03){s+=2;r.push('credit proxy weak');} if(m.TLT.ret20d<-0.04){s+=2;r.push('long bonds falling / yield pressure');} if(m.USO.ret20d>0.06){s+=1;r.push('oil rising');} if(m.UUP.ret20d>0.03){s+=1;r.push('dollar firming');} if(m.SPY.read==='Bullish confirmed') s-=1; return {rating:s>=5?'HIGH':s>=3?'MEDIUM':s>=1?'WATCH':'LOW', reason:r.length?r.join('; '):'SPY trend and credit are not showing major stress'}; }
function mrPressureRows_(m){ var p=mrPressure_(m); return [['Broad trend','SPY ' + m.SPY.priceVs50 + '; ' + m.SPY.read,p.rating,'Controls add aggressiveness.'],['Rates','TLT 20d ' + mrPct_(m.TLT.ret20d),m.TLT.ret20d<-0.04?'HIGH':m.TLT.ret20d<-0.02?'MEDIUM':'LOW','Higher yields pressure growth.'],['Credit','HYG 20d ' + mrPct_(m.HYG.ret20d),m.HYG.ret20d<-0.03?'HIGH':m.HYG.ret20d<0?'MEDIUM':'LOW','Credit weakness lowers risk appetite.'],['Oil','USO 20d ' + mrPct_(m.USO.ret20d),m.USO.ret20d>0.06?'HIGH':m.USO.ret20d>0.03?'MEDIUM':'LOW','Oil can revive inflation pressure.'],['Dollar','UUP 20d ' + mrPct_(m.UUP.ret20d),m.UUP.ret20d>0.03?'MEDIUM':'LOW','Strong dollar can pressure ADRs.']]; }
function mrMacroRows_(m){ return [['Geopolitics','GLD 20d '+mrPct_(m.GLD.ret20d)+'; USO '+mrPct_(m.USO.ret20d),(m.GLD.ret20d>0.03||m.USO.ret20d>0.05)?'Elevated':'Normal','Gold/defense matter more if shocks rise.'],['Energy','USO 20d '+mrPct_(m.USO.ret20d),m.USO.ret20d>0.05?'Oil rising':m.USO.ret20d<-0.05?'Oil easing':'Neutral','Affects inflation and consumer sectors.'],['Liquidity','SPY '+m.SPY.priceVs50+'; HYG '+mrPct_(m.HYG.ret20d),m.SPY.priceVs50==='Below 50'||m.HYG.ret20d<-0.03?'Risk-off':'Acceptable','Controls add size.'],['Fed / Yields','TLT 20d '+mrPct_(m.TLT.ret20d),m.TLT.ret20d<-0.04?'Yield pressure':m.TLT.ret20d>0.04?'Easing':'Neutral','Affects growth, banks, bonds.'],['Growth','SPY read: '+m.SPY.read,m.SPY.compositeScore>=0.58?'Constructive':'Mixed/weak','Confirms risk appetite.']]; }
function mrSectorRows_(m){ var spy=m.SPY; return MR_SECTORS.map(function(x){ var t=m[x[0]], rot=mrRotation_(t,spy); return [x[0],x[1],rot,mrFlow_(t,spy),mrBias_(rot)]; }); }
function mrTechRows_(m){ var spy=m.SPY; return MR_SECTORS.map(function(x){ var t=m[x[0]]; return [x[0],x[1],t.read,t.priceVs20+' | '+t.priceVs50+' | '+t.priceVs200,t.macdSignal+' / '+t.macdZero,mrRS_(t,spy)]; }); }
function mrRotationText_(m){ var spy=m.SPY, inA=[], outA=[], mix=[]; MR_SECTORS.forEach(function(x){ var t=m[x[0]], rot=mrRotation_(t,spy), rs=mrRS_(t,spy), label=x[0]+' '+x[1]; if((rot==='Leadership'||rot==='Positive trend')&&rs!=='Underperforming') inA.push(label); else if(rot==='Lagging / weak'||rs==='Underperforming') outA.push(label); else mix.push(label); }); return 'Rotation is favoring: ' + (inA.length?inA.join(', '):'no clear leadership') + '. It is avoiding or trimming exposure to: ' + (outA.length?outA.join(', '):'no clear trim group') + '. Mixed: ' + (mix.length?mix.join(', '):'none') + '.'; }
function mrForwardText_(m){ var spy=m.SPY, leaders=[], weak=[]; MR_SECTORS.forEach(function(x){ var rot=mrRotation_(m[x[0]],spy); if(rot==='Leadership') leaders.push(x[0]); if(rot==='Lagging / weak') weak.push(x[0]); }); return 'SPY is ' + spy.read + ' with 20-day change of ' + mrPct_(spy.ret20d) + '. Leadership: ' + (leaders.length?leaders.join(', '):'none') + '. Weak/lagging: ' + (weak.length?weak.join(', '):'none') + '. Use macro dates before changing position sizes.'; }
function mrRotation_(t,spy){ if(t.read==='Needs data') return 'Needs data'; if(t.compositeScore>=0.75&&t.ret20d>spy.ret20d) return 'Leadership'; if(t.compositeScore>=0.60) return 'Positive trend'; if(t.compositeScore>=0.45&&t.ret20d>=spy.ret20d) return 'Improving'; if(t.compositeScore<0.35) return 'Lagging / weak'; return 'Mixed'; }
function mrFlow_(t,spy){ if(t.read==='Needs data'||spy.read==='Needs data') return 'Needs data'; if(t.ret20d>spy.ret20d+0.02&&t.compositeScore>=0.55) return 'Positive / Outperforming'; if(t.ret20d<spy.ret20d-0.02&&t.compositeScore<=0.50) return 'Negative / Underperforming'; return 'Mixed / In-line'; }
function mrRS_(t,spy){ if(t.read==='Needs data'||spy.read==='Needs data') return 'Needs data'; var d=t.ret20d-spy.ret20d; return d>0.02?'Outperforming':d<-0.02?'Underperforming':'In-line'; }
function mrBias_(r){ return r==='Leadership'?'Best hold/add':r==='Positive trend'?'Selective add':r==='Improving'?'Watch':r==='Lagging / weak'?'Avoid/trim':'Wait'; }

function mrPortfolio_(holdings,market){ var weights={}; holdings.forEach(function(h){ var meta=mrMeta_(h); weights[meta[0]]=(weights[meta[0]]||0)+h.weight; }); return holdings.map(function(h){ var meta=mrMeta_(h), tech=market[meta[1]]||market.SPY, tol=mrTolerance_(h,meta,tech,weights), act=mrAction_(h,meta,tech,tol); return {ticker:h.ticker, qty:h.qty, value:mrMoney_(h.value), pnl:mrMoney_(h.pnl), pnlPct:mrPct_(h.pnlPct), tolerance:tol, thesis:meta[2], action:act.action, est:mrMoney_(act.est), trimPnl:mrMoney_(act.trimPnl), reason:act.reason+' Proxy '+meta[1]+': '+tech.read+' ('+tech.compositeScore.toFixed(2)+').', sleeve:meta[0], priority:mrPriority_(act.action), sortValue:h.value}; }).sort(function(a,b){ return a.priority-b.priority||b.sortValue-a.sortValue; }); }
function mrPortfolioCards_(sh,row,rows){ rows.forEach(function(r){ var bg=r.action.indexOf('Add-QTY')===0?'#dcfce7':r.action.indexOf('Trim-QTY')===0?'#fef3c7':'#f8fafc'; sh.getRange(row,1,1,6).merge().setValue(r.ticker+' | '+r.action+' | '+r.sleeve).setFontWeight('bold').setFontSize(11).setBackground(bg).setBorder(true,true,true,true,null,null,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(row,25); row++; row=mrTable_(sh,row,['Qty','Value','Total P&L','Tolerance','Action $','Trim P&L'],[[r.qty,r.value,r.pnl+' / '+r.pnlPct,r.tolerance,r.est,r.trimPnl]],[1,1,1,1,1,1]); row=mrMini_(sh,row,'Thesis: '+r.thesis,34); row=mrMini_(sh,row,'Reason: '+r.reason,40); }); return row; }
function mrMeta_(h){ if(MR_MAP[h.ticker]) return MR_MAP[h.ticker]; if(String(h.type).indexOf('etf')>=0) return ['ETF / Unmapped','SPY','ETF exposure; map sleeve manually if material.']; return ['Equity / Unmapped','SPY','Single-stock exposure; validate thesis manually.']; }
function mrTolerance_(h,meta,t,weights){ var sleeve=meta[0], sw=weights[sleeve]||0, safety=sleeve.indexOf('Safety')>=0||sleeve.indexOf('Income')>=0||sleeve.indexOf('Credit')>=0; if(safety&&sw>0.45) return 'Overweight safety'; if(safety&&sw>0.18) return 'Safety overlap'; if(safety) return 'Income tilt'; if(h.pnlPct>0.50) return 'Profit outlier'; if(h.pnlPct<-0.10&&t.compositeScore<0.45) return 'Weak'; if(h.weight<0.015) return 'Too small'; if(h.weight>0.08||sw>0.18) return 'Near limit'; if(t.compositeScore<0.35&&t.read!=='Needs data') return 'Out of tolerance'; return 'In tolerance'; }
function mrAction_(h,meta,t,tol){ var price=h.price||(h.value/h.qty)||0, qty=Math.max(0,Math.floor(h.qty)), safety=meta[0].indexOf('Safety')>=0||meta[0].indexOf('Income')>=0||meta[0].indexOf('Credit')>=0, pnlPer=h.qty?h.pnl/h.qty:0; function pack(label,q,why){ return {action:label, est:q*price, trimPnl:label.indexOf('Trim-QTY')===0?q*pnlPer:0, reason:why}; } if(qty<=0) return pack('Hold',0,'No share quantity available.'); if(t.read==='Needs data') return pack('Hold',0,'No market technical data yet.'); if(tol==='Profit outlier'&&t.compositeScore<0.65){var q=Math.max(1,Math.floor(qty*0.10)); return pack('Trim-QTY '+q,q,'Profit outlier and mapped trend is not leadership.');} if((tol==='Weak'||tol==='Out of tolerance')&&t.compositeScore<0.40&&!safety){var q2=Math.max(1,Math.floor(qty*0.15)); return pack('Trim-QTY '+q2,q2,'Weak mapped trend plus tolerance pressure.');} if((tol==='Overweight safety'||tol==='Safety overlap')&&safety){var q3=Math.max(1,Math.floor(qty*0.05)); return pack('Trim-QTY '+q3,q3,'Safety sleeve overlap; trim only if reallocating to confirmed leadership.');} if(tol==='Too small'&&t.compositeScore>=0.75&&!safety){var q4=price>=100?1:Math.max(1,Math.floor(500/Math.max(price,1))); return pack('Add-QTY '+q4,q4,'Small position with mapped technical leadership.');} return pack('Hold',0,'Hold-no-add until macro/news confirms.'); }
function mrPriority_(s){ return String(s).indexOf('Add-QTY')===0?1:String(s).indexOf('Trim-QTY')===0?2:9; }
function mrTotalAction_(rows,prefix){ return rows.reduce(function(sum,r){ return String(r.action).indexOf(prefix)===0?sum+mrNum_(r.est):sum; },0); }
function mrTotalTrimPnl_(rows){ return rows.reduce(function(sum,r){ return String(r.action).indexOf('Trim-QTY')===0?sum+mrNum_(r.trimPnl):sum; },0); }
function mrCountAction_(rows,prefix){ return rows.filter(function(r){ return String(r.action).indexOf(prefix)===0; }).length; }

function mrMacroCatalystRows_(){ var events=mrNormalizeEvents_(mrEventSeed_()); return events.slice(0,12).map(function(e){ return [e.date,e.time,e.name,e.impact,String(e.days),e.why]; }); }
function mrEventSeed_(){ var out=[], now=mrToday_(), months=6; for(var i=0;i<months;i++){ var d=new Date(now.getFullYear(), now.getMonth()+i, 1), y=d.getFullYear(), m=d.getMonth(); out.push(mrEv_('ISM Manufacturing PMI', mrFirstBiz_(y,m), '10:00 AM ET', 'MEDIUM','Early growth signal for cyclicals.')); out.push(mrEv_('ISM Services PMI', mrNthBiz_(y,m,3), '10:00 AM ET', 'MEDIUM','Services surprise can move yields and growth stocks.')); out.push(mrEv_('Retail Sales', mrNthWeekday_(y,m,4,2), '8:30 AM ET', 'MEDIUM','Consumer strength can support or weaken risk-on moves.')); out.push(mrEv_('PCE Inflation / Personal Income', mrLastBiz_(y,m), '8:30 AM ET', 'HIGH','Fed preferred inflation gauge; affects yields and growth.')); if([0,3,6,9].indexOf(m)!==-1) out.push(mrEv_('GDP Report', mrNthWeekday_(y,m,4,4), '8:30 AM ET', 'HIGH','Confirms growth or slowdown risk.')); } [[2026,6,29,'FOMC Rate Decision'],[2026,8,16,'FOMC + SEP'],[2026,9,28,'FOMC Rate Decision'],[2026,11,9,'FOMC + SEP']].forEach(function(x){ out.push(mrEv_(x[3], new Date(x[0],x[1],x[2]), '2:00 PM ET','HIGH','Fed decision can move yields, QQQ, and volatility.')); }); return out; }
function mrEv_(name,date,time,impact,why){ return {name:name,d:mrStrip_(date),time:time,impact:impact,why:why}; }
function mrNormalizeEvents_(events){ var t=mrToday_(), end=mrAddDays_(t,120), seen={}, out=[]; events.forEach(function(e){ if(e.d<t||e.d>end) return; var key=e.name+'|'+e.d; if(seen[key]) return; seen[key]=true; out.push({name:e.name,date:mrDate_(e.d),time:e.time,impact:e.impact,why:e.why,days:mrDays_(t,e.d),raw:e.d}); }); out.sort(function(a,b){ return a.raw-b.raw; }); return out; }

function mrReadHoldings_(ss){ var sh=ss.getSheetByName(MR.holdingsSheet); if(!sh) throw new Error('Missing Holdings tab.'); var vals=sh.getDataRange().getDisplayValues(); if(vals.length<2) return []; var h=vals[0].map(mrNorm_), ix={}; h.forEach(function(x,i){ix[x]=i;}); var fb={ticker_symbol:6,security_name:7,security_type:8,quantity:10,cost_basis:11,institution_price:12,institution_value:13,calculated_market_value:14,unrealized_gain_loss:15,unrealized_gain_loss_pct:16,portfolio_weight:17}; function get(row,k){var i=ix[k]!==undefined?ix[k]:fb[k]; return i===undefined||i>=row.length?'':row[i];} var out=[]; vals.slice(1).forEach(function(r){ var ticker=String(get(r,'ticker_symbol')||'').trim().toUpperCase(), name=String(get(r,'security_name')||''), type=String(get(r,'security_type')||'').toLowerCase(); if(!ticker) return; if(mrExclude_(ticker,name,type)) return; var qty=mrNum_(get(r,'quantity')), value=mrNum_(get(r,'institution_value'))||mrNum_(get(r,'calculated_market_value')), price=mrNum_(get(r,'institution_price')); if(!price&&value&&qty) price=value/qty; out.push({ticker:ticker,name:name,type:type,qty:qty,cost:mrNum_(get(r,'cost_basis')),price:price,value:value,pnl:mrNum_(get(r,'unrealized_gain_loss')),pnlPct:mrPctNum_(get(r,'unrealized_gain_loss_pct')),weight:mrPctNum_(get(r,'portfolio_weight'))}); }); return out; }
function mrExclude_(ticker,name,type){ var n=String(name||'').toLowerCase(); return type==='cash'||ticker.indexOf('CUR:')===0||['VMFXX','SWVXX','SPAXX','FDRXX'].indexOf(ticker)>=0||n.indexOf('money market')>=0||n.indexOf('sweep')>=0; }

function mrReset_(sh){ try{sh.showColumns(1,sh.getMaxColumns());}catch(e){} sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).breakApart(); sh.clear(); sh.setHiddenGridlines(true); if(sh.getMaxRows()<260) sh.insertRowsAfter(sh.getMaxRows(),260-sh.getMaxRows()); if(sh.getMaxColumns()<6) sh.insertColumnsAfter(sh.getMaxColumns(),6-sh.getMaxColumns()); [80,105,105,90,86,235].forEach(function(w,i){sh.setColumnWidth(i+1,w);}); if(sh.getMaxColumns()>6){try{sh.hideColumns(7,sh.getMaxColumns()-6);}catch(e2){}} sh.getRange(1,1,sh.getMaxRows(),6).setFontFamily(MR.font).setFontSize(10).setWrap(true).setVerticalAlignment('middle').setNumberFormat('@'); }
function mrTitle_(sh,row,title,sub){ sh.getRange(row,1,1,6).merge().setValue(title).setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#111827').setFontColor('#fff'); sh.setRowHeight(row,30); row++; sh.getRange(row,1,1,6).merge().setValue(sub).setFontSize(9).setBackground('#eef2ff').setFontColor('#374151'); sh.setRowHeight(row,30); return row+2; }
function mrSection_(sh,row,title){ sh.getRange(row,1,1,6).merge().setValue(title).setFontSize(12).setFontWeight('bold').setBackground('#dbeafe').setFontColor('#111827').setVerticalAlignment('middle'); sh.setRowHeight(row,24); return row+1; }
function mrPara_(sh,row,text,h){ sh.getRange(row,1,1,6).merge().setValue(text).setFontSize(10).setBackground('#fff').setBorder(true,true,true,true,null,null,'#e5e7eb',SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(row,h||58); return row+2; }
function mrMini_(sh,row,text,h){ sh.getRange(row,1,1,6).merge().setValue(text).setFontSize(9).setBackground('#fff').setBorder(true,true,true,true,null,null,'#e5e7eb',SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(row,h||36); return row+1; }
function mrTable_(sh,row,headers,rows,spans){ rows=rows||[]; spans=spans||headers.map(function(){return 1;}); var col=1; for(var h=0;h<headers.length;h++){ var span=spans[h]||1; sh.getRange(row,col,1,span).merge().setValue(headers[h]).setFontWeight('bold').setFontSize(8).setBackground('#1f2937').setFontColor('#fff').setHorizontalAlignment('center').setBorder(true,true,true,true,true,true,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID); col+=span; } sh.setRowHeight(row,22); row++; rows.forEach(function(r){ col=1; for(var c=0;c<headers.length;c++){ var sp=spans[c]||1; sh.getRange(row,col,1,sp).merge().setValue(r[c]===undefined?'':String(r[c])).setFontSize(9).setBackground('#fff').setFontColor('#111827').setWrap(true).setBorder(true,true,true,true,true,true,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID); col+=sp; } sh.setRowHeight(row,mrRowH_(r)); row++; }); return row+1; }
function mrRowH_(r){ var len=String((r||[]).join(' ')).length; return len>230?60:len>145?48:34; }
function mrFinal_(sh,lastRow){ sh.getRange(1,1,Math.max(1,lastRow),6).setFontFamily(MR.font).setNumberFormat('@').setWrap(true); for(var r=1;r<=lastRow;r++){ if(sh.getRowHeight(r)>82) sh.setRowHeight(r,82); } }

function mrEma_(arr,p){ var k=2/(p+1), out=[], ema=arr[0]; for(var i=0;i<arr.length;i++){ema=i===0?arr[i]:arr[i]*k+ema*(1-k);out.push(ema);} return out; }
function mrMacd_(arr){ var e12=mrEma_(arr,12), e26=mrEma_(arr,26), m=[]; for(var i=0;i<arr.length;i++) m.push(e12[i]-e26[i]); var s=mrEma_(m,9), macd=mrLast_(m), sig=mrLast_(s); return {macd:macd,signal:sig}; }
function mrRsi_(arr,p){ var g=[],l=[]; for(var i=1;i<arr.length;i++){var d=arr[i]-arr[i-1]; g.push(Math.max(0,d)); l.push(Math.max(0,-d));} var ag=mrSma_(g.slice(-p),p)||0, al=mrSma_(l.slice(-p),p)||0; if(al===0) return 100; return 100-(100/(1+ag/al)); }
function mrSma_(arr,p){ if(!arr.length) return 0; var n=Math.min(arr.length,p), s=0; for(var i=arr.length-n;i<arr.length;i++) s+=arr[i]; return s/n; }
function mrRet_(arr,n){ return arr.length>n&&arr[arr.length-1-n]?arr[arr.length-1]/arr[arr.length-1-n]-1:0; }
function mrVol_(arr,n){ var rets=[]; for(var i=Math.max(1,arr.length-n);i<arr.length;i++) rets.push(arr[i]/arr[i-1]-1); if(!rets.length) return 0; var mean=rets.reduce(function(a,b){return a+b;},0)/rets.length; var v=rets.reduce(function(a,b){return a+Math.pow(b-mean,2);},0)/rets.length; return Math.sqrt(v)*Math.sqrt(252); }
function mrDrawdown_(arr,n){ var slice=arr.slice(-n), peak=0, min=0; slice.forEach(function(x){peak=Math.max(peak,x); if(peak) min=Math.min(min,x/peak-1);}); return min; }
function mrLast_(arr){ return arr[arr.length-1]; }
function mrUnique_(arr){ var m={},out=[]; arr.forEach(function(x){ if(!m[x]){m[x]=true;out.push(x);} }); return out; }
function mrNorm_(x){ return String(x||'').trim().toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,''); }
function mrNum_(x){ if(x===null||x===undefined||x==='') return 0; if(typeof x==='number') return x; var s=String(x), neg=s.indexOf('(')>=0&&s.indexOf(')')>=0; s=s.replace(/[,$%\s()]/g,''); var n=Number(s); if(isNaN(n)) return 0; return neg?-n:n; }
function mrPctNum_(x){ if(typeof x==='number') return Math.abs(x)>1?x/100:x; var s=String(x||''), n=mrNum_(s); return s.indexOf('%')>=0?n/100:n; }
function mrMoney_(n){ return '$'+Number(n||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}); }
function mrPct_(n){ return n===undefined||n===null||isNaN(Number(n))?'n/a':(Number(n)*100).toFixed(2)+'%'; }
function mrNow_(){ return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'); }
function mrToday_(){ var n=new Date(); return new Date(n.getFullYear(),n.getMonth(),n.getDate()); }
function mrStrip_(d){ return new Date(d.getFullYear(),d.getMonth(),d.getDate()); }
function mrAddDays_(d,days){ var x=mrStrip_(d); x.setDate(x.getDate()+days); return x; }
function mrDays_(a,b){ return Math.round((mrStrip_(b).getTime()-mrStrip_(a).getTime())/86400000); }
function mrDate_(d){ return Utilities.formatDate(d,Session.getScriptTimeZone(),'MMM d, yyyy'); }
function mrWeekend_(d){ return d.getDay()===0||d.getDay()===6; }
function mrFirstBiz_(y,m){ var d=new Date(y,m,1); while(mrWeekend_(d)) d.setDate(d.getDate()+1); return mrStrip_(d); }
function mrNthBiz_(y,m,n){ var d=new Date(y,m,1), c=0; while(d.getMonth()===m){ if(!mrWeekend_(d)) c++; if(c===n) return mrStrip_(d); d.setDate(d.getDate()+1); } return mrLastBiz_(y,m); }
function mrNthWeekday_(y,m,w,n){ var d=new Date(y,m,1), c=0; while(d.getMonth()===m){ if(d.getDay()===w) c++; if(c===n) return mrStrip_(d); d.setDate(d.getDate()+1); } return mrLastBiz_(y,m); }
function mrLastBiz_(y,m){ var d=new Date(y,m+1,0); while(mrWeekend_(d)) d.setDate(d.getDate()-1); return mrStrip_(d); }
