var GEX_TPSL={holdingsSheet:'Holdings',reportSheet:'GEX Take Profit & Stop Limit',phoneSheet:'Phone-Gmail',font:'Times New Roman',maxTickers:40,maxDte:45,minOi:1,emailRecipient:'mohitsingh2031@gmail.com',cboeBase:'https://cdn.cboe.com/api/global/delayed_quotes/options/'};

function buildGexTakeProfitStopLimit(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var holdings=gexReadHoldings_(ss).slice(0,GEX_TPSL.maxTickers);
  if(!holdings.length) throw new Error('No usable holdings tickers found. Pull Holdings first.');
  var sh=ss.getSheetByName(GEX_TPSL.reportSheet)||ss.insertSheet(GEX_TPSL.reportSheet);
  gexPrepareSheet_(sh);
  var rows=[],errors=[];
  holdings.forEach(function(h){try{var r=gexAnalyzeTicker_(h.ticker,h.price);rows.push(gexBuildOutputRow_(h,r));Utilities.sleep(250);}catch(e){errors.push(h.ticker+': '+e.message);rows.push([h.ticker,h.qty,gexMoney_(h.cost),gexMoney_(h.costPerShare),h.price||0,'n/a','n/a','n/a','n/a','','n/a','','']);}});
  rows=rows.filter(gexRowHasUsableGex_);
  rows.sort(gexSortByTpPctDesc_);
  var row=1;
  row=gexTitle_(sh,row,'GEX Take Profit & Stop Limit','Built '+gexNow_()+'. Spot uses live GOOGLEFINANCE formulas. TP %, SL %, and EST. Loss use cost per share from Holdings.');
  row=gexSection_(sh,row,'How To Read This');
  row=gexParagraph_(sh,row,'Call Wall is the largest call-open-interest strike and is treated as upside resistance/magnet. Put Wall is the largest put-open-interest strike and is treated as downside support/risk. Gamma Flip is where cumulative estimated dealer GEX crosses zero. TP % = Take Profit / Cost Per Share - 1. SL % = Stop Limit / Cost Per Share - 1. EST. Loss = (Stop Limit - Cost Per Share) x Qty. Rows with no usable GEX levels are hidden.');
  row=gexSection_(sh,row,'Portfolio GEX TP / SL Map');
  row=gexTable_(sh,row,['Ticker','Qty','Cost Basis','Cost / Share','Live Spot','Call Wall','Put Wall','Gamma Flip','Take Profit','TP %','Stop Limit','SL %','EST. Loss'],rows);
  if(errors.length){row=gexSection_(sh,row,'Hidden / Data Issues');row=gexParagraph_(sh,row,errors.join('\n'));}
  gexFinalize_(sh,row);
  gexBuildPhoneGmailTab_(ss,rows);
  return 'GEX TP & SL complete. Main rows shown: '+rows.length+'. Phone-Gmail tab updated with row 4 totals and EST. Loss.'+(errors.length?' Hidden/issues: '+errors.length+'.':'');
}

function buildPhoneGmailGexReport(){return buildGexTakeProfitStopLimit();}
function buildGexTakeProfitStopLimitAndPhoneGmail(){return buildGexTakeProfitStopLimit();}

function sendPhoneGmailReportEmail(){
  buildGexTakeProfitStopLimit();
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(GEX_TPSL.phoneSheet);
  if(!sh)throw new Error('Missing Phone-Gmail tab.');
  var html=gexPhoneEmailHtml_(sh);
  GmailApp.sendEmail(GEX_TPSL.emailRecipient,'Daily GEX TP / SL Phone-Gmail Report - '+gexNow_(),'Your Phone-Gmail report is attached in HTML format.',{htmlBody:html});
  return 'Phone-Gmail report emailed to '+GEX_TPSL.emailRecipient+'.';
}

function createDailyPhoneGmailEmailTrigger(){
  deleteDailyPhoneGmailEmailTriggers();
  ScriptApp.newTrigger('sendPhoneGmailReportEmail').timeBased().everyDays(1).atHour(17).create();
  return 'Daily Phone-Gmail email trigger created for about 5 PM project time. Recipient: '+GEX_TPSL.emailRecipient+'.';
}

function deleteDailyPhoneGmailEmailTriggers(){
  ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='sendPhoneGmailReportEmail')ScriptApp.deleteTrigger(t);});
  return 'Deleted existing Phone-Gmail email triggers.';
}

function gexAnalyzeTicker_(ticker,fallbackSpot){var chain=gexFetchCboeChain_(ticker),spot=Number(chain.spot||fallbackSpot||0);if(!spot)throw new Error('missing spot price');var options=chain.options||[];if(!options.length)throw new Error('empty option chain');var byStrike={},callOi={},putOi={};options.forEach(function(o){var p=gexParseOption_(o,ticker);if(!p||p.dte<0||p.dte>GEX_TPSL.maxDte||p.oi<GEX_TPSL.minOi)return;var gamma=p.gamma||gexBsGamma_(spot,p.strike,p.dte,p.iv||0.35);var gx=gamma*p.oi*100*spot*spot*0.01;if(p.type==='P'){gx=-Math.abs(gx);putOi[String(p.strike)]=(putOi[String(p.strike)]||0)+p.oi;}else{gx=Math.abs(gx);callOi[String(p.strike)]=(callOi[String(p.strike)]||0)+p.oi;}byStrike[String(p.strike)]=(byStrike[String(p.strike)]||0)+gx;});var strikes=Object.keys(byStrike).map(Number).sort(function(a,b){return a-b;});if(!strikes.length)throw new Error('no option rows inside '+GEX_TPSL.maxDte+' DTE');return{spot:spot,callWall:gexMaxOiStrike_(callOi),putWall:gexMaxOiStrike_(putOi),gammaFlip:gexGammaFlip_(strikes,byStrike,spot),source:chain.source};}
function gexBuildOutputRow_(h,r){var spot=r.spot||h.price||0;var tp=r.callWall&&r.callWall>spot?r.callWall:gexNearestAbove_(spot,[r.callWall,r.gammaFlip]);var slc=[r.putWall,r.gammaFlip].filter(function(x){return x&&x<spot;});var sl=slc.length?Math.max.apply(null,slc):(r.putWall||r.gammaFlip||'');return[h.ticker,h.qty,gexMoney_(h.cost),gexMoney_(h.costPerShare),spot,gexPrice_(r.callWall),gexPrice_(r.putWall),gexPrice_(r.gammaFlip),gexPrice_(tp),'',gexPrice_(sl),'',''];}
function gexRowHasUsableGex_(r){return gexNum_(r[5])||gexNum_(r[6])||gexNum_(r[7])||gexNum_(r[8])||gexNum_(r[10]);}
function gexSortByTpPctDesc_(a,b){return gexTpPctSortValue_(b)-gexTpPctSortValue_(a);}
function gexTpPctSortValue_(row){var cost=gexNum_(row[3]),tp=gexNum_(row[8]);return(!cost||!tp)?-999999:(tp/cost)-1;}
function gexFetchCboeChain_(ticker){var sym=String(ticker||'').trim().toUpperCase().replace(/[^A-Z0-9.]/g,''),tries=[sym,sym.replace('.','-')],lastErr='';for(var i=0;i<tries.length;i++){try{var url=GEX_TPSL.cboeBase+encodeURIComponent(tries[i])+'.json';var res=UrlFetchApp.fetch(url,{muteHttpExceptions:true,followRedirects:true,headers:{'User-Agent':'Mozilla/5.0'}});if(res.getResponseCode()!==200){lastErr='Cboe HTTP '+res.getResponseCode();continue;}var data=JSON.parse(res.getContentText());var options=data.options||(data.data&&data.data.options)||[];var current=data.current_price||data.currentPrice||(data.data&&(data.data.current_price||data.data.currentPrice));if(!options.length){lastErr='no options in Cboe response';continue;}return{spot:Number(current||0),options:options,source:'Cboe delayed quotes'};}catch(e){lastErr=e.message;}}throw new Error(lastErr||'Cboe fetch failed');}
function gexParseOption_(o,ticker){var option=String(o.option||o.option_symbol||o.symbol||'').toUpperCase(),type=String(o.option_type||o.type||'').toUpperCase().charAt(0),strike=Number(o.strike||o.strike_price||0),exp=o.expiration_date||o.expiration||o.expiry||'';if((!type||!strike||!exp)&&option){var m=option.match(/(\d{6})([CP])(\d{8})$/);if(m){type=m[2];strike=Number(m[3])/1000;exp='20'+m[1].slice(0,2)+'-'+m[1].slice(2,4)+'-'+m[1].slice(4,6);}}if(type!=='C'&&type!=='P')return null;if(!strike)return null;var dte=gexDte_(exp),oi=Number(o.open_interest||o.openInterest||o.oi||0),gamma=Number(o.gamma||(o.greeks&&o.greeks.gamma)||0),iv=Number(o.iv||o.implied_volatility||(o.greeks&&o.greeks.iv)||0);if(iv>3)iv=iv/100;return{type:type,strike:strike,dte:dte,oi:oi,gamma:gamma,iv:iv};}
function gexBsGamma_(s,k,dte,iv){if(!s||!k||!dte)return 0;var t=Math.max(dte/365,1/365),v=Math.max(Number(iv||0.35),0.05),d1=(Math.log(s/k)+(0.5*v*v)*t)/(v*Math.sqrt(t));return gexNormPdf_(d1)/(s*v*Math.sqrt(t));}
function gexGammaFlip_(strikes,byStrike,spot){var cum=0,best=null,bestAbs=Infinity,prev=null;for(var i=0;i<strikes.length;i++){var k=strikes[i];cum+=byStrike[String(k)]||0;if(Math.abs(cum)<bestAbs){bestAbs=Math.abs(cum);best=k;}if(prev!==null&&((prev<=0&&cum>=0)||(prev>=0&&cum<=0)))return k;prev=cum;}return best;}
function gexMaxOiStrike_(map){var best='',bestOi=-1;Object.keys(map||{}).forEach(function(k){if(map[k]>bestOi){bestOi=map[k];best=Number(k);}});return best;}
function gexNearestAbove_(spot,arr){var a=arr.filter(function(x){return x&&x>spot;}).sort(function(a,b){return a-b;});return a.length?a[0]:'';}
function gexNormPdf_(x){return Math.exp(-0.5*x*x)/Math.sqrt(2*Math.PI);}
function gexDte_(exp){var d=new Date(exp);if(isNaN(d.getTime()))return 9999;var today=new Date();return Math.ceil((new Date(d.getFullYear(),d.getMonth(),d.getDate())-new Date(today.getFullYear(),today.getMonth(),today.getDate()))/86400000);}
function gexNow_(){return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');}
function gexMoney_(n){return n||n===0?'$'+Number(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}):'n/a';}
function gexPrice_(n){return n?Number(n).toFixed(2):'n/a';}
function gexReadHoldings_(ss){var sh=ss.getSheetByName(GEX_TPSL.holdingsSheet);if(!sh)throw new Error('Missing Holdings tab.');var values=sh.getDataRange().getDisplayValues();if(values.length<2)return[];var headers=values[0].map(gexNorm_),ix={};headers.forEach(function(h,i){ix[h]=i;});function col(names,fallback){for(var i=0;i<names.length;i++){if(ix[names[i]]!==undefined)return ix[names[i]];}return fallback;}var cTicker=col(['ticker_symbol','ticker','symbol'],6),cQty=col(['quantity','qty','shares'],10),cCost=col(['cost_basis','cost','basis'],11),cPrice=col(['institution_price','price','current_price'],12),out=[],seen={};values.slice(1).forEach(function(r){var ticker=String(r[cTicker]||'').trim().toUpperCase();if(!ticker||seen[ticker]||gexExcludeTicker_(ticker))return;var qty=gexNum_(r[cQty]),cost=gexNum_(r[cCost]);seen[ticker]=true;out.push({ticker:ticker,qty:qty,cost:cost,costPerShare:qty?cost/qty:0,price:gexNum_(r[cPrice])});});return out;}
function gexExcludeTicker_(t){return t.indexOf('CUR:')===0||['VMFXX','SWVXX','SPAXX','FDRXX'].indexOf(t)>=0;}
function gexNorm_(x){return String(x||'').trim().toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'');}
function gexNum_(x){if(typeof x==='number')return x;var s=String(x||''),neg=s.indexOf('(')>=0;s=s.replace(/[,$%\s()]/g,'');var n=Number(s);return isNaN(n)?0:(neg?-n:n);}
function gexGoogleFinanceFormula_(ticker,fallback){var t=String(ticker||'').replace(/"/g,''),fb=Number(fallback||0);return '=IFERROR(GOOGLEFINANCE("'+t+'","price"),'+fb+')';}

function gexPrepareSheet_(sh){try{sh.showColumns(1,sh.getMaxColumns());}catch(e){}sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).breakApart();sh.clear();sh.setHiddenGridlines(true);if(sh.getMaxColumns()<13)sh.insertColumnsAfter(sh.getMaxColumns(),13-sh.getMaxColumns());sh.getRange(1,1,sh.getMaxRows(),13).setFontFamily(GEX_TPSL.font).setFontSize(10).setWrap(true).setVerticalAlignment('middle');}
function gexTitle_(sh,row,title,sub){sh.getRange(row,1,1,13).merge().setValue(title).setFontSize(18).setFontWeight('bold').setBackground('#111827').setFontColor('#fff').setHorizontalAlignment('center');sh.setRowHeight(row,34);row++;sh.getRange(row,1,1,13).merge().setValue(sub).setBackground('#eef2ff').setFontColor('#374151');sh.setRowHeight(row,28);return row+2;}
function gexSection_(sh,row,title){sh.getRange(row,1,1,13).merge().setValue(title).setFontSize(13).setFontWeight('bold').setBackground('#dbeafe').setFontColor('#111827');sh.setRowHeight(row,24);return row+1;}
function gexParagraph_(sh,row,text){sh.getRange(row,1,1,13).merge().setValue(text).setWrap(true).setBorder(true,true,true,true,null,null,'#e5e7eb',SpreadsheetApp.BorderStyle.SOLID);sh.setRowHeight(row,58);return row+2;}
function gexTable_(sh,row,headers,rows){var data=[headers].concat(rows),range=sh.getRange(row,1,data.length,headers.length);range.setValues(data).setWrap(true).setVerticalAlignment('middle').setBorder(true,true,true,true,true,true,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID);sh.getRange(row,1,1,headers.length).setFontWeight('bold').setFontSize(9).setBackground('#1f2937').setFontColor('#fff').setHorizontalAlignment('center');for(var r=1;r<data.length;r++){var rr=row+r,ticker=String(rows[r-1][0]||''),fallbackSpot=gexNum_(rows[r-1][4]);sh.getRange(rr,5).setFormula(gexGoogleFinanceFormula_(ticker,fallbackSpot)).setNumberFormat('$#,##0.00');sh.getRange(rr,10).setFormula('=IFERROR(I'+rr+'/D'+rr+'-1,"")').setNumberFormat('0.00%');sh.getRange(rr,12).setFormula('=IFERROR(K'+rr+'/D'+rr+'-1,"")').setNumberFormat('0.00%');sh.getRange(rr,13).setFormula('=IFERROR((K'+rr+'-D'+rr+')*B'+rr+',"")').setNumberFormat('$#,##0.00;-$#,##0.00');sh.setRowHeight(rr,28);}gexHideNaColumns_(sh,rows);sh.setRowHeight(row,26);return row+data.length+2;}
function gexHideNaColumns_(sh,rows){[6,7,8,9,10,11,12,13].forEach(function(col){var sourceCol=col;if(col===10)sourceCol=9;if(col===12||col===13)sourceCol=11;var hasUsable=rows.some(function(r){var v=String(r[sourceCol-1]===undefined||r[sourceCol-1]===null?'':r[sourceCol-1]).trim().toLowerCase();return v!==''&&v!=='n/a'&&v!=='na'&&v!=='#n/a';});if(!hasUsable){try{sh.hideColumns(col);}catch(e){}}});}
function gexFinalize_(sh,lastRow){var widths=[62,52,86,86,84,78,78,78,84,64,84,64,86];for(var c=1;c<=13;c++)sh.setColumnWidth(c,widths[c-1]);if(sh.getMaxColumns()>13)sh.hideColumns(14,sh.getMaxColumns()-13);sh.setFrozenRows(0);}

function gexBuildPhoneGmailTab_(ss,rows){var sh=ss.getSheetByName(GEX_TPSL.phoneSheet)||ss.insertSheet(GEX_TPSL.phoneSheet);gexPhonePrepare_(sh);var r=1;r=gexPhoneTitle_(sh,r,'GEX TP / SL','Phone-Gmail printable one-page view • '+gexNow_());r=gexPhoneSummary_(sh,r,rows.length);r=gexPhoneTotals_(sh,r,rows);r=gexPhoneTable_(sh,r,rows);gexPhoneFinalize_(sh,r);}
function gexPhonePrepare_(sh){try{sh.showColumns(1,sh.getMaxColumns());}catch(e){}sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).breakApart();sh.clear();sh.setHiddenGridlines(true);if(sh.getMaxColumns()<9)sh.insertColumnsAfter(sh.getMaxColumns(),9-sh.getMaxColumns());sh.getRange(1,1,sh.getMaxRows(),9).setFontFamily(GEX_TPSL.font).setFontSize(10).setWrap(true).setVerticalAlignment('middle');}
function gexPhoneTitle_(sh,row,title,sub){sh.getRange(row,1,1,9).merge().setValue(title).setFontSize(19).setFontWeight('bold').setFontColor('#ffffff').setBackground('#111827').setHorizontalAlignment('center');sh.setRowHeight(row,34);row++;sh.getRange(row,1,1,9).merge().setValue(sub).setFontSize(9).setFontColor('#e5e7eb').setBackground('#374151').setHorizontalAlignment('center');sh.setRowHeight(row,22);return row+1;}
function gexPhoneSummary_(sh,row,count){sh.getRange(row,1,1,9).merge().setValue('iPhone/Gmail one-page table: TP and SL side-by-side, TP % and SL % side-by-side, plus EST. Loss. % and loss use cost/share. Rows: '+count).setFontSize(9).setFontColor('#111827').setBackground('#fef3c7').setHorizontalAlignment('center');sh.setRowHeight(row,34);return row+1;}
function gexPhoneTotals_(sh,row,rows){var pos=0,neg=0;rows.forEach(function(x){var n=gexEstLossNumber_(x[10],x[3],x[1]);if(n>0)pos+=n;if(n<0)neg+=n;});sh.getRange(row,1,1,4).merge().setValue('Positive Total EST. Loss\n'+gexMoney_(pos)).setBackground('#dcfce7').setFontWeight('bold').setHorizontalAlignment('center');sh.getRange(row,5,1,5).merge().setValue('Negative Total EST. Loss\n'+gexMoney_(neg)).setBackground('#fee2e2').setFontWeight('bold').setHorizontalAlignment('center');sh.setRowHeight(row,38);return row+1;}
function gexPhoneTable_(sh,row,rows){var out=[['Ticker','Qty','Cost/Shr','Spot','TP','TP %','SL','SL %','EST. Loss']];rows.forEach(function(x){out.push([x[0],x[1],x[3],x[4],x[8],gexPctFromCost_(x[8],x[3]),x[10],gexPctFromCost_(x[10],x[3]),gexEstLoss_(x[10],x[3],x[1])]);});var rg=sh.getRange(row,1,out.length,9);rg.setValues(out).setWrap(true).setVerticalAlignment('middle').setBorder(true,true,true,true,true,true,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID);sh.getRange(row,1,1,9).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff').setHorizontalAlignment('center').setFontSize(9);for(var i=1;i<out.length;i++){var rr=row+i;sh.getRange(rr,5,1,2).setBackground('#dcfce7').setFontWeight('bold');sh.getRange(rr,7,1,3).setBackground('#fee2e2').setFontWeight('bold');sh.setRowHeight(rr,25);}return row+out.length+1;}
function gexPctFromCost_(level,cost){var l=gexNum_(level),c=gexNum_(cost);return(!l||!c)?'n/a':((l/c)-1).toLocaleString('en-US',{style:'percent',minimumFractionDigits:2,maximumFractionDigits:2});}
function gexEstLossNumber_(sl,cost,qty){var s=gexNum_(sl),c=gexNum_(cost),q=gexNum_(qty);return(!s||!c||!q)?0:(s-c)*q;}
function gexEstLoss_(sl,cost,qty){var n=gexEstLossNumber_(sl,cost,qty);return n===0?'n/a':n.toLocaleString('en-US',{style:'currency',currency:'USD',minimumFractionDigits:2,maximumFractionDigits:2});}
function gexPhoneFinalize_(sh,lastRow){var widths=[58,42,62,58,58,52,58,52,72];for(var c=1;c<=9;c++)sh.setColumnWidth(c,widths[c-1]);if(sh.getMaxColumns()>9)sh.hideColumns(10,sh.getMaxColumns()-9);sh.setFrozenRows(0);sh.getRange(1,1,Math.max(lastRow,1),9).setVerticalAlignment('middle');}

function gexPhoneEmailHtml_(sh){var lastRow=sh.getLastRow(),cols=9,range=sh.getRange(1,1,lastRow,cols),v=range.getDisplayValues(),bg=range.getBackgrounds(),fc=range.getFontColors(),fw=range.getFontWeights(),ha=range.getHorizontalAlignments();var html='<div style="font-family:Arial,sans-serif"><h2 style="margin:0 0 8px">GEX TP / SL Phone-Gmail Report</h2><table style="border-collapse:collapse;font-size:13px">';for(var r=0;r<v.length;r++){html+='<tr>';for(var c=0;c<cols;c++){if(c>0&&v[r][c]===''&&v[r][c-1]!==''&&bg[r][c]===bg[r][c-1])continue;var span=1;while(c+span<cols&&v[r][c+span]===''&&bg[r][c+span]===bg[r][c])span++;var tag=r===4?'th':'td';html+='<'+tag+' colspan="'+span+'" style="border:1px solid #cbd5e1;padding:6px 7px;background:'+bg[r][c]+';color:'+fc[r][c]+';font-weight:'+fw[r][c]+';text-align:'+(ha[r][c]||'center')+';white-space:pre-line">'+gexHtmlEscape_(v[r][c])+'</'+tag+'>';c+=span-1;}html+='</tr>';}html+='</table></div>';return html;}
function gexHtmlEscape_(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
