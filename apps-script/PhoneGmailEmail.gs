/**
 * Portfolio Link - Phone-Gmail email sender
 * Separate Apps Script file for emailing the Phone-Gmail tab from visible A:I columns.
 * Column B / Qty is ALWAYS excluded from the email, even if Google reports it visible.
 * Preserves merged cells, colors, font weights, alignments, borders, row heights, column widths, and HYPERLINK URLs.
 */

var PHONE_GMAIL_EMAIL = {
  sheetName: 'Phone-Gmail',
  recipient: 'mohitsingh2031@gmail.com',
  subjectPrefix: 'Daily GEX TP / SL Phone-Gmail Report'
};

function sendPhoneGmailReportEmail() {
  if (typeof buildGexTakeProfitStopLimit === 'function') buildGexTakeProfitStopLimit();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(PHONE_GMAIL_EMAIL.sheetName);
  if (!sh) throw new Error('Missing Phone-Gmail tab. Run GEX TP & SL first.');
  try { sh.hideColumns(2); } catch(e) {}
  var html = phoneGmailBuildExactHtml_(sh);
  var subject = PHONE_GMAIL_EMAIL.subjectPrefix + ' - ' + phoneGmailNow_();
  GmailApp.sendEmail(PHONE_GMAIL_EMAIL.recipient, subject, 'Open this email in HTML view to see the formatted Phone-Gmail report.', {htmlBody: html});
  return 'Phone-Gmail report emailed to ' + PHONE_GMAIL_EMAIL.recipient + '. Qty column B force-excluded. Links preserved.';
}

function createDailyPhoneGmailEmailTrigger() {
  deleteDailyPhoneGmailEmailTriggers();
  ScriptApp.newTrigger('sendPhoneGmailReportEmail').timeBased().everyDays(1).atHour(17).create();
  return 'Daily Phone-Gmail email trigger created for about 5 PM project time. Recipient: ' + PHONE_GMAIL_EMAIL.recipient + '.';
}

function deleteDailyPhoneGmailEmailTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendPhoneGmailReportEmail') ScriptApp.deleteTrigger(t);
  });
  return 'Deleted existing Phone-Gmail email triggers.';
}

function phoneGmailBuildExactHtml_(sh) {
  SpreadsheetApp.flush();
  var lastRow = sh.getLastRow();
  var maxCol = 9;

  // Hard rule from Raj: never include Qty column B in email.
  var visibleCols = [];
  for (var c = 1; c <= maxCol; c++) {
    if (c === 2) continue;
    var hidden = false;
    try { hidden = sh.isColumnHiddenByUser(c); } catch(e) { hidden = false; }
    if (!hidden) visibleCols.push(c);
  }
  if (!visibleCols.length) visibleCols = [1,3,4,5,6,7,8,9];

  var range = sh.getRange(1, 1, lastRow, maxCol);
  var values = range.getDisplayValues();
  var formulas = range.getFormulas();
  var rich = range.getRichTextValues();
  var bgs = range.getBackgrounds();
  var colors = range.getFontColors();
  var weights = range.getFontWeights();
  var sizes = range.getFontSizes();
  var aligns = range.getHorizontalAlignments();
  var valigns = range.getVerticalAlignments();
  var wraps = range.getWraps();
  var skipped = {};
  var spans = {};

  range.getMergedRanges().forEach(function(m) {
    var r0 = m.getRow(), c0 = m.getColumn(), rs = m.getNumRows(), cs = m.getNumColumns();
    if (r0 < 1 || c0 < 1 || r0 > lastRow || c0 > maxCol) return;
    var visibleInMerge = [];
    for (var cc = c0; cc < c0 + cs && cc <= maxCol; cc++) {
      if (visibleCols.indexOf(cc) >= 0) visibleInMerge.push(cc);
    }
    if (!visibleInMerge.length) return;
    var anchorCol = visibleInMerge[0];
    spans[r0 + ':' + anchorCol] = {rowspan: rs, colspan: visibleInMerge.length};
    for (var r = r0; r < r0 + rs; r++) {
      visibleInMerge.forEach(function(cc) {
        if (!(r === r0 && cc === anchorCol)) skipped[r + ':' + cc] = true;
      });
    }
  });

  var html = '';
  html += '<div style="font-family:Arial,sans-serif;margin:0;padding:0;max-width:760px">';
  html += '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;table-layout:fixed;width:auto;margin:0;font-family:Arial,sans-serif;font-size:13px;line-height:1.18">';
  html += '<colgroup>';
  visibleCols.forEach(function(c) { html += '<col style="width:' + Math.max(42, sh.getColumnWidth(c)) + 'px">'; });
  html += '</colgroup>';

  for (var rr = 1; rr <= lastRow; rr++) {
    html += '<tr style="height:' + sh.getRowHeight(rr) + 'px">';
    visibleCols.forEach(function(cc) {
      if (skipped[rr + ':' + cc]) return;
      var sp = spans[rr + ':' + cc] || {rowspan: 1, colspan: 1};
      var v = values[rr - 1][cc - 1];
      var bg = bgs[rr - 1][cc - 1];
      var color = colors[rr - 1][cc - 1];
      var weight = weights[rr - 1][cc - 1];
      var size = sizes[rr - 1][cc - 1];
      var align = aligns[rr - 1][cc - 1] || 'center';
      var valign = valigns[rr - 1][cc - 1] || 'middle';
      var white = wraps[rr - 1][cc - 1] ? 'pre-line' : 'nowrap';
      var link = phoneGmailCellLink_(rich[rr - 1][cc - 1], formulas[rr - 1][cc - 1]);
      var content = phoneGmailHtmlEscape_(v);
      if (link) content = '<a href="' + phoneGmailHtmlEscape_(link) + '" target="_blank" rel="noopener noreferrer" style="color:#1155cc;text-decoration:underline">' + content + '</a>';
      var tag = rr === 5 ? 'th' : 'td';
      html += '<' + tag + ' rowspan="' + sp.rowspan + '" colspan="' + sp.colspan + '" style="' +
        'border:1px solid #cbd5e1;' +
        'padding:4px 6px;' +
        'background:' + bg + ';' +
        'color:' + color + ';' +
        'font-weight:' + weight + ';' +
        'font-size:' + size + 'px;' +
        'text-align:' + align + ';' +
        'vertical-align:' + valign + ';' +
        'white-space:' + white + ';' +
        'box-sizing:border-box;' +
        '">' + content + '</' + tag + '>';
    });
    html += '</tr>';
  }
  html += '</table></div>';
  return html;
}

function phoneGmailCellLink_(richCell, formula) {
  var link = '';
  try { link = richCell && richCell.getLinkUrl(); } catch(e) { link = ''; }
  if (link) return link;
  var f = String(formula || '');
  var m = f.match(/^=HYPERLINK\("([^"]+)"\s*,/i);
  if (m && m[1]) return m[1];
  return '';
}

function phoneGmailHtmlEscape_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function phoneGmailNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
