/**
 * Portfolio Link Market Analysis Python Bridge
 * Reads the Holdings tab, sends it to the Python backend, and writes Report Market Analysis.
 * Required Script Properties:
 *   MARKET_BACKEND_URL=https://your-backend-domain.com
 *   MARKET_BACKEND_TOKEN=your-private-token
 * Run: buildMarketAnalysisReport()
 * Test: testMarketBackendConnection()
 * Version: python-backend-bridge-v1
 */

var MR_BRIDGE = {
  holdingsSheet: 'Holdings',
  reportSheet: 'Report Market Analysis',
  backendUrlProp: 'MARKET_BACKEND_URL',
  backendTokenProp: 'MARKET_BACKEND_TOKEN',
  font: 'Times New Roman'
};

function buildMarketAnalysisReport() {
  var payload = mrBuildBackendPayload_();
  var report = mrCallMarketBackend_(payload);
  mrWriteBackendReport_(report);
  return 'Python Market Analysis complete. Report written: ' + MR_BRIDGE.reportSheet;
}

function testMarketBackendConnection() {
  var baseUrl = mrGetBackendBaseUrl_();
  var res = UrlFetchApp.fetch(baseUrl + '/health', {
    method: 'get',
    muteHttpExceptions: true
  });
  return 'Status ' + res.getResponseCode() + ': ' + res.getContentText();
}

function setMarketBackendConfig(url, token) {
  if (!url) throw new Error('Missing backend URL.');
  var props = PropertiesService.getScriptProperties();
  props.setProperty(MR_BRIDGE.backendUrlProp, String(url).trim().replace(/\/+$/, ''));
  if (token) props.setProperty(MR_BRIDGE.backendTokenProp, String(token).trim());
  return 'Market backend settings saved.';
}

function mrBuildBackendPayload_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MR_BRIDGE.holdingsSheet);
  if (!sh) throw new Error('Missing Holdings tab. Pull holdings first.');

  var range = sh.getDataRange();
  var values = range.getDisplayValues();
  if (!values || values.length < 2) throw new Error('Holdings tab is empty. Pull holdings first.');

  var headers = values[0];
  var rows = values.slice(1).filter(function(row) {
    return row.join('').trim() !== '';
  });

  if (!rows.length) throw new Error('No Holdings rows found under the header. Pull holdings first.');

  return {
    headers: headers,
    rows: rows,
    source: 'Google Sheets Holdings tab',
    spreadsheet_id: ss.getId(),
    created_at: mrNow_()
  };
}

function mrCallMarketBackend_(payload) {
  var baseUrl = mrGetBackendBaseUrl_();
  var token = PropertiesService.getScriptProperties().getProperty(MR_BRIDGE.backendTokenProp) || '';
  var headers = { 'Content-Type': 'application/json' };
  if (token) headers['X-Backend-Token'] = token;

  var res = UrlFetchApp.fetch(baseUrl + '/analyze', {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var text = res.getContentText();
  var data;
  try {
    data = JSON.parse(text);
  } catch (err) {
    throw new Error('Python backend returned non-JSON response. Status ' + code + ': ' + text);
  }

  if (code < 200 || code >= 300) {
    throw new Error('Python backend error ' + code + ': ' + (data.detail || data.error || text));
  }

  return data;
}

function mrGetBackendBaseUrl_() {
  var url = PropertiesService.getScriptProperties().getProperty(MR_BRIDGE.backendUrlProp);
  if (!url) {
    throw new Error('Missing MARKET_BACKEND_URL in Apps Script Properties. Deploy the Python backend first, then set MARKET_BACKEND_URL and MARKET_BACKEND_TOKEN.');
  }
  return String(url).trim().replace(/\/+$/, '');
}

function mrWriteBackendReport_(report) {
  if (!report || !report.sections) throw new Error('Backend report is missing sections.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MR_BRIDGE.reportSheet) || ss.insertSheet(MR_BRIDGE.reportSheet);
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), Math.min(13, sh.getMaxColumns()))
    .setFontFamily(MR_BRIDGE.font)
    .setFontSize(9)
    .setWrap(true)
    .setVerticalAlignment('top');

  var row = 1;
  row = mrTitle_(sh, row, report.title || 'Raj Market Rotation Report', 'Built ' + (report.as_of || mrNow_()) + '. Powered by Python backend technical engine.');

  report.sections.forEach(function(section) {
    if (!section) return;
    if (section.type === 'paragraph') {
      if (section.title) row = mrSection_(sh, row, section.title);
      row = mrParagraph_(sh, row, section.text || '');
    } else if (section.type === 'table') {
      row = mrSection_(sh, row, section.title || 'Table');
      row = mrTable_(sh, row, section.headers || [], section.rows || []);
    }
  });

  mrFinalize_(sh, row);
}

function mrTitle_(sh, row, title, subtitle) {
  sh.getRange(row, 1, 1, 13)
    .merge()
    .setValue(title)
    .setFontSize(18)
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#ffffff');
  row++;
  sh.getRange(row, 1, 1, 13)
    .merge()
    .setValue(subtitle)
    .setFontSize(9)
    .setFontColor('#374151')
    .setBackground('#eef2ff');
  return row + 2;
}

function mrSection_(sh, row, title) {
  sh.getRange(row, 1, 1, 13)
    .merge()
    .setValue(title)
    .setFontSize(13)
    .setFontWeight('bold')
    .setBackground('#dbeafe')
    .setFontColor('#111827');
  return row + 1;
}

function mrParagraph_(sh, row, text) {
  sh.getRange(row, 1, 1, 13)
    .merge()
    .setValue(text)
    .setFontSize(9)
    .setBackground('#ffffff')
    .setWrap(true)
    .setVerticalAlignment('top');
  sh.setRowHeight(row, 42);
  return row + 2;
}

function mrTable_(sh, row, headers, rows) {
  headers = headers || [];
  rows = rows || [];
  var width = Math.max(headers.length, 1);
  var cleanRows = rows.map(function(r) {
    r = r || [];
    var out = [];
    for (var i = 0; i < width; i++) out.push(r[i] === undefined || r[i] === null ? '' : r[i]);
    return out;
  });
  var cleanHeaders = [];
  for (var h = 0; h < width; h++) cleanHeaders.push(headers[h] || '');
  var data = [cleanHeaders].concat(cleanRows);

  var range = sh.getRange(row, 1, data.length, width);
  range
    .setValues(data)
    .setWrap(true)
    .setVerticalAlignment('top')
    .setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(row, 1, 1, width)
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');

  if (data.length > 1) {
    sh.getRange(row + 1, 1, data.length - 1, width)
      .setBackground('#ffffff')
      .setFontColor('#111827');
  }

  return row + data.length + 2;
}

function mrFinalize_(sh, lastRow) {
  var maxCols = Math.max(13, sh.getMaxColumns());
  for (var c = 1; c <= Math.min(maxCols, 13); c++) {
    var w = 110;
    if (c === 8 || c === 12) w = 220;
    if (c === 13) w = 135;
    sh.setColumnWidth(c, w);
  }
  sh.getRange(1, 1, Math.max(1, lastRow), Math.min(13, sh.getMaxColumns())).setFontFamily(MR_BRIDGE.font);
  sh.autoResizeRows(1, Math.max(1, lastRow));
}

function mrNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
