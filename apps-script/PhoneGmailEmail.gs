/**
 * Portfolio Link - Phone-Gmail email sender
 * Separate Apps Script file for emailing the Phone-Gmail tab exactly from cells A:I.
 * It preserves merged cells, colors, font weights, alignments, borders, and row/column sizing in HTML.
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
  var html = phoneGmailBuildExactHtml_(sh);
  var subject = PHONE_GMAIL_EMAIL.subjectPrefix + ' - ' + phoneGmailNow_();
  GmailApp.sendEmail(PHONE_GMAIL_EMAIL.recipient, subject, 'Open this email in HTML view to see the formatted Phone-Gmail report.', {htmlBody: html});
  return 'Phone-Gmail report emailed to ' + PHONE_GMAIL_EMAIL.recipient + '.';
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
  var cols = 9;
  var range = sh.getRange(1, 1, lastRow, cols);
  var values = range.getDisplayValues();
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
    var r0 = m.getRow();
    var c0 = m.getColumn();
    var rs = m.getNumRows();
    var cs = m.getNumColumns();
    if (r0 < 1 || c0 < 1 || r0 > lastRow || c0 > cols) return;
    spans[r0 + ':' + c0] = {rowspan: rs, colspan: Math.min(cs, cols - c0 + 1)};
    for (var r = r0; r < r0 + rs; r++) {
      for (var c = c0; c < c0 + cs && c <= cols; c++) {
        if (!(r === r0 && c === c0)) skipped[r + ':' + c] = true;
      }
    }
  });

  var html = '';
  html += '<div style="font-family:Arial,sans-serif;margin:0;padding:0;max-width:760px">';
  html += '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;table-layout:fixed;width:auto;margin:0;font-family:Arial,sans-serif;font-size:13px;line-height:1.18">';
  html += '<colgroup>';
  for (var c = 1; c <= cols; c++) html += '<col style="width:' + Math.max(42, sh.getColumnWidth(c)) + 'px">';
  html += '</colgroup>';

  for (var rr = 1; rr <= lastRow; rr++) {
    html += '<tr style="height:' + sh.getRowHeight(rr) + 'px">';
    for (var cc = 1; cc <= cols; cc++) {
      if (skipped[rr + ':' + cc]) continue;
      var sp = spans[rr + ':' + cc] || {rowspan: 1, colspan: 1};
      var v = values[rr - 1][cc - 1];
      var bg = bgs[rr - 1][cc - 1];
      var color = colors[rr - 1][cc - 1];
      var weight = weights[rr - 1][cc - 1];
      var size = sizes[rr - 1][cc - 1];
      var align = aligns[rr - 1][cc - 1] || 'center';
      var valign = valigns[rr - 1][cc - 1] || 'middle';
      var white = wraps[rr - 1][cc - 1] ? 'pre-line' : 'nowrap';
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
        '">' + phoneGmailHtmlEscape_(v) + '</' + tag + '>';
    }
    html += '</tr>';
  }
  html += '</table></div>';
  return html;
}

function phoneGmailHtmlEscape_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function phoneGmailNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
