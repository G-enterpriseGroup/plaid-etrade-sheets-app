/**
 * Portfolio Link - Phone-Gmail GEX TP/SL
 * Creates a phone-friendly Gmail/iPhone view from the GEX Take Profit & Stop Limit tab.
 * Tab name: Phone-Gmail
 * Layout is narrow, card-style, and easy to copy/screenshot/email from a phone.
 */

var PHONE_GMAIL_GEX = {
  sourceSheet: 'GEX Take Profit & Stop Limit',
  sheetName: 'Phone-Gmail',
  font: 'Times New Roman'
};

function buildPhoneGmailGexReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Build/refresh the normal GEX tab first when the function exists.
  if (typeof buildGexTakeProfitStopLimit === 'function') {
    buildGexTakeProfitStopLimit();
  }

  var source = ss.getSheetByName(PHONE_GMAIL_GEX.sourceSheet);
  if (!source) throw new Error('Missing GEX Take Profit & Stop Limit tab. Run GEX TP & SL first.');

  var rows = phoneGmailReadGexRows_(source);
  if (!rows.length) throw new Error('No usable GEX TP/SL rows found for Phone-Gmail.');

  var sh = ss.getSheetByName(PHONE_GMAIL_GEX.sheetName) || ss.insertSheet(PHONE_GMAIL_GEX.sheetName);
  phoneGmailPrepare_(sh);

  var r = 1;
  r = phoneGmailTitle_(sh, r, 'GEX TP / SL', 'Phone-Gmail view • ' + phoneGmailNow_());
  r = phoneGmailSummary_(sh, r, rows.length);

  rows.forEach(function(x, i) {
    r = phoneGmailCard_(sh, r, x, i + 1);
  });

  phoneGmailFinalize_(sh, r);
  return 'Phone-Gmail report complete. Cards: ' + rows.length + '. Check the Phone-Gmail tab.';
}

function phoneGmailReadGexRows_(source) {
  SpreadsheetApp.flush();
  Utilities.sleep(750);

  var values = source.getDataRange().getDisplayValues();
  var headerRow = -1;
  for (var i = 0; i < values.length; i++) {
    if (values[i].indexOf('Ticker') >= 0 && values[i].indexOf('Take Profit') >= 0 && values[i].indexOf('Stop Limit') >= 0) {
      headerRow = i;
      break;
    }
  }
  if (headerRow < 0) return [];

  var headers = values[headerRow];
  var idx = {};
  headers.forEach(function(h, i) { idx[String(h || '').trim()] = i; });

  var out = [];
  for (var r = headerRow + 1; r < values.length; r++) {
    var row = values[r];
    var ticker = String(row[idx['Ticker']] || '').trim();
    if (!ticker || ticker === 'Data Issues' || ticker === 'Hidden / Data Issues') break;

    var tp = phoneGmailClean_(row[idx['Take Profit']]);
    var sl = phoneGmailClean_(row[idx['Stop Limit']]);
    if (phoneGmailIsNa_(tp) && phoneGmailIsNa_(sl)) continue;

    out.push({
      ticker: ticker,
      qty: phoneGmailClean_(row[idx['Qty']]),
      costBasis: phoneGmailClean_(row[idx['Cost Basis']]),
      costPerShare: phoneGmailClean_(row[idx['Cost / Share']]),
      liveSpot: phoneGmailClean_(row[idx['Live Spot']]),
      callWall: phoneGmailClean_(row[idx['Call Wall']]),
      putWall: phoneGmailClean_(row[idx['Put Wall']]),
      gammaFlip: phoneGmailClean_(row[idx['Gamma Flip']]),
      takeProfit: tp,
      tpPct: phoneGmailClean_(row[idx['TP %']]),
      stopLimit: sl,
      slPct: phoneGmailClean_(row[idx['SL %']])
    });
  }

  out.sort(function(a, b) { return phoneGmailPctNum_(b.tpPct) - phoneGmailPctNum_(a.tpPct); });
  return out;
}

function phoneGmailPrepare_(sh) {
  try { sh.showColumns(1, sh.getMaxColumns()); } catch(e) {}
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.setHiddenGridlines(true);
  if (sh.getMaxColumns() < 6) sh.insertColumnsAfter(sh.getMaxColumns(), 6 - sh.getMaxColumns());
  sh.getRange(1, 1, sh.getMaxRows(), 6)
    .setFontFamily(PHONE_GMAIL_GEX.font)
    .setFontSize(11)
    .setWrap(true)
    .setVerticalAlignment('middle');
}

function phoneGmailTitle_(sh, row, title, subtitle) {
  sh.getRange(row, 1, 1, 6).merge()
    .setValue(title)
    .setFontSize(22)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#111827')
    .setHorizontalAlignment('center');
  sh.setRowHeight(row, 38);
  row++;
  sh.getRange(row, 1, 1, 6).merge()
    .setValue(subtitle)
    .setFontSize(10)
    .setFontColor('#e5e7eb')
    .setBackground('#374151')
    .setHorizontalAlignment('center');
  sh.setRowHeight(row, 24);
  return row + 2;
}

function phoneGmailSummary_(sh, row, count) {
  sh.getRange(row, 1, 1, 6).merge()
    .setValue('Built for iPhone/Gmail: TP and SL are side-by-side; TP % and SL % are side-by-side. Percentages are from cost/share, not spot.')
    .setFontSize(10)
    .setFontColor('#111827')
    .setBackground('#fef3c7')
    .setHorizontalAlignment('center');
  sh.setRowHeight(row, 42);
  row++;
  sh.getRange(row, 1, 1, 6).merge()
    .setValue('Cards shown: ' + count + ' • Sorted highest TP % to lowest')
    .setFontSize(10)
    .setFontColor('#111827')
    .setBackground('#dbeafe')
    .setHorizontalAlignment('center');
  sh.setRowHeight(row, 26);
  return row + 2;
}

function phoneGmailCard_(sh, row, x, rank) {
  var green = '#dcfce7';
  var red = '#fee2e2';
  var blue = '#dbeafe';
  var gray = '#f3f4f6';
  var dark = '#1f2937';

  sh.getRange(row, 1, 1, 6).merge()
    .setValue(rank + '. ' + x.ticker)
    .setFontSize(18)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground(dark)
    .setHorizontalAlignment('center');
  sh.setRowHeight(row, 30);
  row++;

  sh.getRange(row, 1, 1, 2).merge().setValue('Qty\n' + x.qty).setBackground(gray).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(row, 3, 1, 2).merge().setValue('Cost/Share\n' + x.costPerShare).setBackground(gray).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(row, 5, 1, 2).merge().setValue('Live Spot\n' + x.liveSpot).setBackground(gray).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(row, 48);
  row++;

  sh.getRange(row, 1, 1, 3).merge().setValue('TAKE PROFIT').setBackground(green).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(row, 4, 1, 3).merge().setValue('STOP LIMIT').setBackground(red).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(row, 24);
  row++;

  sh.getRange(row, 1, 1, 3).merge().setValue(x.takeProfit).setBackground('#bbf7d0').setFontSize(17).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(row, 4, 1, 3).merge().setValue(x.stopLimit).setBackground('#fecaca').setFontSize(17).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(row, 34);
  row++;

  sh.getRange(row, 1, 1, 3).merge().setValue(x.tpPct).setBackground(green).setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(row, 4, 1, 3).merge().setValue(x.slPct).setBackground(red).setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(row, 30);
  row++;

  sh.getRange(row, 1, 1, 2).merge().setValue('Call Wall\n' + x.callWall).setBackground(blue).setHorizontalAlignment('center');
  sh.getRange(row, 3, 1, 2).merge().setValue('Put Wall\n' + x.putWall).setBackground(blue).setHorizontalAlignment('center');
  sh.getRange(row, 5, 1, 2).merge().setValue('G-Flip\n' + x.gammaFlip).setBackground(blue).setHorizontalAlignment('center');
  sh.setRowHeight(row, 46);

  sh.getRange(row - 5, 1, 6, 6).setBorder(true, true, true, true, true, true, '#9ca3af', SpreadsheetApp.BorderStyle.SOLID);
  return row + 2;
}

function phoneGmailFinalize_(sh, lastRow) {
  var widths = [70, 70, 70, 70, 70, 70];
  for (var c = 1; c <= 6; c++) sh.setColumnWidth(c, widths[c - 1]);
  if (sh.getMaxColumns() > 6) sh.hideColumns(7, sh.getMaxColumns() - 6);
  sh.setFrozenRows(0);
  sh.getRange(1, 1, Math.max(lastRow, 1), 6).setVerticalAlignment('middle');
}

function phoneGmailClean_(v) {
  var s = String(v === null || v === undefined ? '' : v).trim();
  return s || 'n/a';
}

function phoneGmailIsNa_(v) {
  var s = String(v || '').trim().toLowerCase();
  return !s || s === 'n/a' || s === 'na' || s === '#n/a';
}

function phoneGmailPctNum_(v) {
  var s = String(v || '').replace('%', '').replace(',', '').trim();
  var n = Number(s);
  return isNaN(n) ? -999999 : n;
}

function phoneGmailNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
