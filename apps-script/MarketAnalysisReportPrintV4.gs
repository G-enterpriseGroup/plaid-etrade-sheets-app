/**
 * Market Analysis Report V4 post-processor.
 * Runs the existing local report engine, then cleans the printed sheet:
 * - keeps source list in sidebar only
 * - standardizes action labels to add / trim / hold
 * - adds estimated P&L for each trim card
 */

function buildMarketAnalysisReportV4() {
  var message = buildMarketAnalysisReport();
  mrV4CleanPrintedReport_();
  return message + ' Portrait print cleanup applied: source list removed from sheet, trim language standardized, estimated trim P&L added.';
}

function mrV4CleanPrintedReport_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Report Market Analysis');
  if (!sh) throw new Error('Report Market Analysis sheet not found.');
  var lastRow = Math.max(1, sh.getLastRow());
  var lastCol = Math.min(6, sh.getLastColumn());

  mrV4RemoveSourceList_(sh, lastRow);
  lastRow = Math.max(1, sh.getLastRow());
  mrV4StandardizeWords_(sh, lastRow, lastCol);
  mrV4AddTrimPnlRows_(sh);
  mrV4FinalPortraitPolish_(sh);
}

function mrV4RemoveSourceList_(sh, lastRow) {
  var vals = sh.getRange(1, 1, lastRow, 1).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    var v = String(vals[i][0] || '').toLowerCase();
    if (v.indexOf('source list') >= 0 || v.indexOf('chicago-style source list') >= 0) {
      sh.deleteRows(i + 1, Math.max(1, lastRow - i));
      return;
    }
  }
}

function mrV4StandardizeWords_(sh, lastRow, lastCol) {
  var rng = sh.getRange(1, 1, lastRow, lastCol);
  var vals = rng.getValues();
  var trimLabel = 'Tr' + 'im-QTY';
  var reduceLabel = 'tr' + 'im/reduce';
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      var s = String(vals[r][c] == null ? '' : vals[r][c]);
      if (!s) continue;
      s = s.replace(/Sell-QTY/g, trimLabel);
      s = s.replace(/sell\/reduce/gi, reduceLabel);
      s = s.replace(/sell\s*\/\s*reduce/gi, reduceLabel);
      s = s.replace(/sell candidates/gi, 'trim candidates');
      s = s.replace(/sells/gi, 'trims');
      s = s.replace(/selling/gi, 'trimming');
      vals[r][c] = s;
    }
  }
  rng.setValues(vals);
}

function mrV4AddTrimPnlRows_(sh) {
  var trimLabel = 'Tr' + 'im-QTY';
  var row = 1;
  while (row <= sh.getLastRow()) {
    var title = String(sh.getRange(row, 1).getDisplayValue() || '');
    if (title.indexOf(trimLabel) >= 0) {
      var qtyMatch = title.match(/Trim-QTY\s+(\d+)/i);
      var actionQty = qtyMatch ? Number(qtyMatch[1]) : 0;
      var dataRow = row + 2;
      var noteRow = row + 5;
      var totalQty = mrV4Num_(sh.getRange(dataRow, 1).getDisplayValue());
      var totalPnl = mrV4Num_(sh.getRange(dataRow, 3).getDisplayValue());
      var estTrimPnl = totalQty && actionQty ? (totalPnl / totalQty) * actionQty : 0;
      var nextText = String(sh.getRange(noteRow, 1).getDisplayValue() || '');
      if (nextText.indexOf('Estimated P&L on this trim') < 0) {
        sh.insertRowsAfter(noteRow - 1, 1);
        sh.getRange(noteRow, 1, 1, 6).merge()
          .setValue('Estimated P&L on this trim: ' + mrV4Money_(estTrimPnl) + '  |  Math: action quantity × average unrealized P&L per share. This is an estimate, not a realized-tax calculation.')
          .setFontFamily('Times New Roman')
          .setFontSize(9)
          .setWrap(true)
          .setVerticalAlignment('middle')
          .setBackground('#fff7ed')
          .setBorder(true, true, true, true, null, null, '#fed7aa', SpreadsheetApp.BorderStyle.SOLID);
        sh.setRowHeight(noteRow, 40);
        row = noteRow + 1;
      }
    }
    row++;
  }
}

function mrV4FinalPortraitPolish_(sh) {
  sh.setHiddenGridlines(true);
  var maxRows = Math.max(1, sh.getLastRow());
  sh.getRange(1, 1, maxRows, 6).setFontFamily('Times New Roman').setNumberFormat('@').setWrap(true).setVerticalAlignment('middle');
  [80, 105, 105, 90, 86, 235].forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  if (sh.getMaxColumns() > 6) {
    try { sh.hideColumns(7, sh.getMaxColumns() - 6); } catch (e) {}
  }
  for (var r = 1; r <= maxRows; r++) {
    if (sh.getRowHeight(r) > 82) sh.setRowHeight(r, 82);
  }
}

function mrV4Num_(x) {
  if (x === null || x === undefined || x === '') return 0;
  var s = String(x);
  var neg = s.indexOf('(') >= 0 && s.indexOf(')') >= 0;
  s = s.replace(/[,$%\s()]/g, '');
  var n = Number(s);
  if (isNaN(n)) return 0;
  return neg ? -n : n;
}

function mrV4Money_(n) {
  return '$' + Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
