/**
 * Portfolio Link - Holdings timestamp footer
 * Overrides pullHoldingsToSheet so every holdings refresh writes a timestamp footer
 * below the last holding row. If holdings grow/shrink, the footer moves automatically.
 */

function pullHoldingsToSheet(){
  setupDashboard();
  const items=getStoredItems_();
  if(!items.length)throw new Error('No linked brokerage items. Open Settings and connect first.');
  const rows=[],errors=[];
  items.forEach(item=>{
    try{
      getHoldingsRowsForItem_(item).forEach(r=>rows.push(r));
      item.last_successful_pull=now_();
      item.last_error='';
    }catch(e){
      item.last_error=e&&e.message?e.message:String(e);
      errors.push((item.institution_name||item.item_id)+': '+item.last_error);
    }
  });
  applyPortfolioWeights_(rows);
  rows.sort((a,b)=>safeNumber_(b[13])-safeNumber_(a[13]));
  const sh=getSpreadsheet_().getSheetByName(SHEETS.holdings);
  replaceSheetData_(sh,HEADERS.holdings,rows);
  addHoldingsUpdatedTimestamp_(sh,rows.length);
  setStoredItems_(items);
  return{status:errors.length?'partial':'success',rows:rows.length,message:errors.length?'Holdings pulled with errors: '+errors.join(' | '):'Holdings pulled successfully. Timestamp updated.'};
}

function addHoldingsUpdatedTimestamp_(sh,rowCount){
  if(!sh)return;
  clearHoldingsUpdatedTimestamp_(sh);
  const footerRow=Math.max(2,Number(rowCount||0)+3);
  const label=holdingsUpdatedTimestampText_();
  const lastCol=Math.min(Math.max(HEADERS.holdings.length,8),sh.getMaxColumns());
  try{sh.getRange(footerRow,1,1,lastCol).breakApart();}catch(e){}
  const rng=sh.getRange(footerRow,1,1,lastCol);
  rng.merge()
    .setValue(label)
    .setFontFamily('Times New Roman')
    .setFontSize(12)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#111827')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true,true,true,true,true,true,'#64748b',SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(footerRow,30);
}

function clearHoldingsUpdatedTimestamp_(sh){
  const maxRows=sh.getMaxRows();
  const maxCols=Math.min(Math.max(HEADERS.holdings.length,8),sh.getMaxColumns());
  const values=sh.getRange(1,1,maxRows,1).getDisplayValues();
  for(let i=0;i<values.length;i++){
    const v=String(values[i][0]||'').trim();
    if(v.indexOf('Updated ')===0){
      try{sh.getRange(i+1,1,1,maxCols).breakApart();}catch(e){}
      sh.getRange(i+1,1,1,maxCols).clearContent().clearFormat();
    }
  }
}

function holdingsUpdatedTimestampText_(){
  const tz=Session.getScriptTimeZone();
  const now=new Date();
  return 'Updated '+Utilities.formatDate(now,tz,'hh:mm a')+' on '+Utilities.formatDate(now,tz,'MMMM d, yyyy');
}
