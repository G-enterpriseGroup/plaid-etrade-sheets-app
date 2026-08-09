/**
 * Put Money Flow - Yahoo/yfinance bridge
 *
 * Architecture:
 *   Google Sheet / Apps Script -> GitHub Actions -> Python/yfinance -> GitHub artifact
 *   -> Apps Script -> Option Quote Cache / dropdowns
 *
 * No Google Cloud project, service account, or Google Sheets API credential is used.
 * This file is additive and does not modify the existing Portfolio Link / Plaid code.
 *
 * Existing Script Properties reused:
 *   GITHUB_MARKET_ACCESS
 *   GITHUB_REPO_FULL_NAME
 *   GITHUB_BRANCH
 *
 * Optional:
 *   GITHUB_PMF_WORKFLOW   (defaults to put-money-flow-yfinance.yml)
 */

var PMF_YF = {
  putSheet: 'Put Money Flow',
  cacheSheet: 'Option Quote Cache',
  pickerSheet: 'YF Option Picker',
  bridgeSheet: 'YF Automation',
  firstRow: 73,
  lastRow: 112,
  workflow: 'put-money-flow-yfinance.yml',
  githubApi: 'https://api.github.com',
  timezone: 'America/New_York'
};

function setupPutMoneyFlowYahooAutomation() {
  var cfg = pmfGitHubConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var put = ss.getSheetByName(PMF_YF.putSheet);
  if (!put) throw new Error('Missing Put Money Flow tab.');

  var cache = pmfEnsureSheet_(ss, PMF_YF.cacheSheet, 500, 7, true);
  var picker = pmfEnsureSheet_(ss, PMF_YF.pickerSheet, 600, 80, true);
  var bridge = pmfEnsureSheet_(ss, PMF_YF.bridgeSheet, 2000, 9, true);

  cache.getRange(1, 1, 1, 7).setValues([[
    'Yahoo Quote Key','Bid','Ask','Last','Mark','Source / Status','Updated At'
  ]]).setFontWeight('bold');
  if (cache.getMaxRows() > 1) cache.getRange(2, 1, cache.getMaxRows() - 1, 7).clearContent();

  bridge.getRange(1, 1, 1, 9).setValues([[
    'Request ID','Mode','Blotter Row','Ticker','Expiration','Created At','Status','Message','Artifact ID'
  ]]).setFontWeight('bold');

  put.getRange('A70').setValue(
    'AUTOMATED OPTION PICKER: choose option type → type ticker → Yahoo expiration dropdown → Yahoo strike dropdown. GitHub/yfinance refreshes option marks automatically.'
  );
  put.getRange('R72').setValue('Yahoo Quote Key');

  for (var row = PMF_YF.firstRow; row <= PMF_YF.lastRow; row++) {
    put.getRange(row, 10).setFormula(
      '=IF($C' + row + '=\"\",\"\",IF(OR($C' + row + '=\"STOCK\",$C' + row + '=\"ETF\"),IFERROR(GOOGLEFINANCE($D' + row + ',\"price\"),\"\"),IFERROR(INDEX(\'' + PMF_YF.cacheSheet + '\'!$E:$E,MATCH($R' + row + ',\'' + PMF_YF.cacheSheet + '\'!$A:$A,0)),\"\")))'
    );
    put.getRange(row, 14).setFormula(
      '=IF($C' + row + '=\"\",\"\",IF(OR($C' + row + '=\"STOCK\",$C' + row + '=\"ETF\"),\"GOOGLEFINANCE\",IFERROR(INDEX(\'' + PMF_YF.cacheSheet + '\'!$F:$F,MATCH($R' + row + ',\'' + PMF_YF.cacheSheet + '\'!$A:$A,0)),\"WAITING FOR YAHOO\")))'
    );
    put.getRange(row, 18).setFormula(
      '=IF(OR(NOT(REGEXMATCH($C' + row + ',\"CALL|PUT|LEAPS\")),$D' + row + '=\"\",$E' + row + '=\"\",$F' + row + '=\"\"),\"\",UPPER($D' + row + ')&\":\"&YEAR($E' + row + ')&\":\"&MONTH($E' + row + ')&\":\"&DAY($E' + row + ')&\":\"&IF(REGEXMATCH($C' + row + ',\"CALL\"),\"CALL\",\"PUT\")&\":\"&TEXT($F' + row + ',\"0.###\"))'
    );
    put.getRange(row, 19).setFormula(
      '=IF($C' + row + '=\"\",\"\",IF(OR($C' + row + '=\"STOCK\",$C' + row + '=\"ETF\"),\"Google Finance\",IF($N' + row + '=\"YFINANCE\",\"Yahoo/yfinance mark • automatic\",IF($N' + row + '=\"WAITING FOR YAHOO\",\"Waiting for Yahoo/yfinance refresh\",$N' + row + '))))'
    );
    put.getRange(row, 5).setNumberFormat('yyyy-mm-dd');

    var cols = pmfPickerCols_(row);
    picker.getRange(1, cols.exp).setValue('Row ' + row + ' expirations');
    picker.getRange(1, cols.strike).setValue('Row ' + row + ' strikes');
  }

  pmfDeleteTriggers_();
  ScriptApp.newTrigger('pmfYahooOnEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('pmfYahooPollResults').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('pmfYahooRefreshAllQuotes').timeBased().everyMinutes(5).create();

  SpreadsheetApp.flush();

  var optionRows = [];
  for (var r = PMF_YF.firstRow; r <= PMF_YF.lastRow; r++) {
    if (!pmfIsOptionType_(put.getRange(r, 3).getDisplayValue())) continue;
    var ticker = String(put.getRange(r, 4).getDisplayValue() || '').trim().toUpperCase();
    if (!ticker) continue;
    pmfRequestExpirations_(put, r);
    if (put.getRange(r, 5).getValue()) pmfRequestStrikes_(put, r);
    if (pmfBuildContract_(put, r)) optionRows.push(r);
  }
  if (optionRows.length) pmfDispatchQuoteRows_(put, optionRows);

  return 'Yahoo/yfinance automation installed. Repo: ' + cfg.repo + '. No Google Cloud or service account is used.';
}

function removePutMoneyFlowYahooAutomation() {
  pmfDeleteTriggers_();
  return 'Put Money Flow Yahoo automation triggers removed.';
}

function testPutMoneyFlowYahooGitHub() {
  var cfg = pmfGitHubConfig_();
  var url = PMF_YF.githubApi + '/repos/' + cfg.repo;
  var res = pmfGitHubFetch_(url, {method:'get'});
  var obj = JSON.parse(res.getContentText());
  return 'GitHub OK: ' + obj.full_name + ' • branch ' + cfg.branch + ' • workflow ' + cfg.workflow;
}

function pmfYahooOnEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== PMF_YF.putSheet) return;

  var r1 = Math.max(PMF_YF.firstRow, e.range.getRow());
  var r2 = Math.min(PMF_YF.lastRow, e.range.getLastRow());
  var c1 = e.range.getColumn();
  var c2 = e.range.getLastColumn();
  if (r1 > r2 || c2 < 3 || c1 > 6) return;

  for (var row = r1; row <= r2; row++) {
    var typ = sh.getRange(row, 3).getDisplayValue();
    if (!pmfIsOptionType_(typ)) continue;

    if (c1 <= 4 && c2 >= 3) {
      pmfClearPicker_(sh, row, true, true);
      var ticker = String(sh.getRange(row, 4).getDisplayValue() || '').trim();
      if (ticker) pmfRequestExpirations_(sh, row);
      continue;
    }

    if (c1 <= 5 && c2 >= 5) {
      pmfClearPicker_(sh, row, false, true);
      if (sh.getRange(row, 5).getValue()) pmfRequestStrikes_(sh, row);
      continue;
    }

    if (c1 <= 6 && c2 >= 6) {
      if (pmfBuildContract_(sh, row)) pmfDispatchQuoteRows_(sh, [row]);
    }
  }
}

function pmfYahooRefreshAllQuotes() {
  if (!pmfIsMarketWindow_()) return 'Outside regular U.S. market refresh window.';
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PMF_YF.putSheet);
  if (!sh) return 'Put Money Flow missing.';
  var rows = [];
  for (var r = PMF_YF.firstRow; r <= PMF_YF.lastRow; r++) {
    if (pmfBuildContract_(sh, r)) rows.push(r);
  }
  if (!rows.length) return 'No complete option rows to refresh.';
  return pmfDispatchQuoteRows_(sh, rows);
}

function pmfYahooPollResults() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bridge = ss.getSheetByName(PMF_YF.bridgeSheet);
  if (!bridge || bridge.getLastRow() < 2) return 'No pending Yahoo requests.';

  var values = bridge.getRange(2, 1, bridge.getLastRow() - 1, 9).getValues();
  var handled = 0;

  for (var i = 0; i < values.length; i++) {
    if (String(values[i][6] || '') !== 'PENDING') continue;
    var sheetRow = i + 2;
    var requestId = String(values[i][0] || '');
    var created = values[i][5] instanceof Date ? values[i][5] : new Date(values[i][5]);

    if (created && !isNaN(created.getTime()) && (Date.now() - created.getTime()) > 20 * 60 * 1000) {
      bridge.getRange(sheetRow, 7, 1, 2).setValues([['TIMEOUT','No GitHub artifact arrived within 20 minutes.']]);
      continue;
    }

    try {
      var artifact = pmfFindArtifact_(requestId);
      if (!artifact) continue;
      var result = pmfDownloadArtifactJson_(artifact);
      pmfApplyResult_(result);
      bridge.getRange(sheetRow, 7, 1, 3).setValues([['DONE','Applied ' + result.mode + ' result.',artifact.id || '']]);
      handled++;
    } catch (err) {
      bridge.getRange(sheetRow, 8).setValue(String(err && err.message ? err.message : err));
    }
  }

  return 'Applied ' + handled + ' Yahoo result(s).';
}

function pmfRequestExpirations_(sh, row) {
  var ticker = String(sh.getRange(row, 4).getDisplayValue() || '').trim().toUpperCase();
  var type = pmfOptionType_(sh.getRange(row, 3).getDisplayValue());
  if (!ticker || !type) return;
  pmfDispatch_('expirations', row, ticker, type, '', '[]');
}

function pmfRequestStrikes_(sh, row) {
  var ticker = String(sh.getRange(row, 4).getDisplayValue() || '').trim().toUpperCase();
  var type = pmfOptionType_(sh.getRange(row, 3).getDisplayValue());
  var expiration = pmfDateIso_(sh.getRange(row, 5).getValue());
  if (!ticker || !type || !expiration) return;
  pmfDispatch_('strikes', row, ticker, type, expiration, '[]');
}

function pmfDispatchQuoteRows_(sh, rows) {
  var contracts = [];
  rows.forEach(function(row){
    var c = pmfBuildContract_(sh, row);
    if (c) contracts.push(c);
  });
  if (!contracts.length) return 'No complete option rows.';
  return pmfDispatch_('quotes', 0, '', '', '', JSON.stringify(contracts));
}

function pmfBuildContract_(sh, row) {
  var typ = sh.getRange(row, 3).getDisplayValue();
  if (!pmfIsOptionType_(typ)) return null;
  var ticker = String(sh.getRange(row, 4).getDisplayValue() || '').trim().toUpperCase();
  var expiration = pmfDateIso_(sh.getRange(row, 5).getValue());
  var strike = Number(sh.getRange(row, 6).getValue());
  if (!ticker || !expiration || !isFinite(strike) || strike <= 0) return null;
  var type = pmfOptionType_(typ);
  var d = new Date(expiration + 'T12:00:00');
  var key = ticker + ':' + d.getUTCFullYear() + ':' + (d.getUTCMonth() + 1) + ':' + d.getUTCDate() + ':' + type + ':' + pmfTrimStrike_(strike);
  return {row:row, ticker:ticker, type:type, expiration:expiration, strike:strike, key:key};
}

function pmfDispatch_(mode, row, ticker, optionType, expiration, contractsJson) {
  var cfg = pmfGitHubConfig_();
  var requestId = Utilities.getUuid().replace(/-/g, '');
  var url = PMF_YF.githubApi + '/repos/' + cfg.repo + '/actions/workflows/' + encodeURIComponent(cfg.workflow) + '/dispatches';
  var payload = {
    ref: cfg.branch,
    inputs: {
      request_id: requestId,
      mode: String(mode || ''),
      row: row ? String(row) : '',
      ticker: String(ticker || ''),
      option_type: String(optionType || ''),
      expiration: String(expiration || ''),
      contracts_json: String(contractsJson || '[]')
    }
  };
  var res = pmfGitHubFetch_(url, {method:'post', contentType:'application/json', payload:JSON.stringify(payload)});
  if (res.getResponseCode() !== 204) throw new Error('GitHub workflow dispatch failed: HTTP ' + res.getResponseCode() + ' ' + res.getContentText());
  pmfRecordRequest_(requestId, mode, row, ticker, expiration);
  return requestId;
}

function pmfRecordRequest_(id, mode, row, ticker, expiration) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = pmfEnsureSheet_(ss, PMF_YF.bridgeSheet, 2000, 9, true);
  if (sh.getLastRow() < 1) sh.getRange(1,1,1,9).setValues([['Request ID','Mode','Blotter Row','Ticker','Expiration','Created At','Status','Message','Artifact ID']]);
  sh.appendRow([id, mode, row || '', ticker || '', expiration || '', new Date(), 'PENDING', '', '']);
}

function pmfFindArtifact_(requestId) {
  var cfg = pmfGitHubConfig_();
  var name = 'pmf-' + requestId;
  var url = PMF_YF.githubApi + '/repos/' + cfg.repo + '/actions/artifacts?name=' + encodeURIComponent(name) + '&per_page=5';
  var res = pmfGitHubFetch_(url, {method:'get'});
  var obj = JSON.parse(res.getContentText());
  var list = obj.artifacts || [];
  if (!list.length) return null;
  list.sort(function(a,b){ return String(b.created_at || '').localeCompare(String(a.created_at || '')); });
  return list[0];
}

function pmfDownloadArtifactJson_(artifact) {
  var first = pmfGitHubFetch_(artifact.archive_download_url, {method:'get', followRedirects:false});
  var code = first.getResponseCode();
  var blob;
  if (code >= 300 && code < 400) {
    var headers = first.getAllHeaders();
    var location = headers.Location || headers.location;
    if (!location) throw new Error('GitHub artifact redirect had no Location header.');
    var second = UrlFetchApp.fetch(location, {method:'get', muteHttpExceptions:true, followRedirects:true});
    if (second.getResponseCode() < 200 || second.getResponseCode() >= 300) throw new Error('Artifact download failed: HTTP ' + second.getResponseCode());
    blob = second.getBlob().setContentType('application/zip');
  } else if (code >= 200 && code < 300) {
    blob = first.getBlob().setContentType('application/zip');
  } else {
    throw new Error('Artifact download failed: HTTP ' + code + ' ' + first.getContentText());
  }

  var files = Utilities.unzip(blob);
  if (!files.length) throw new Error('GitHub artifact ZIP was empty.');
  var target = files.filter(function(b){ return /pmf_result\.json$/i.test(b.getName()); })[0] || files[0];
  return JSON.parse(target.getDataAsString('UTF-8'));
}

function pmfApplyResult_(result) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var put = ss.getSheetByName(PMF_YF.putSheet);
  var picker = pmfEnsureSheet_(ss, PMF_YF.pickerSheet, 600, 80, true);

  if (result.mode === 'expirations') {
    var row = Number(result.row);
    if (row < PMF_YF.firstRow || row > PMF_YF.lastRow) return;
    if (String(put.getRange(row,4).getDisplayValue() || '').trim().toUpperCase() !== String(result.ticker || '').trim().toUpperCase()) return;
    var cols = pmfPickerCols_(row);
    picker.getRange(2, cols.exp, picker.getMaxRows() - 1, 1).clearContent();
    var exps = (result.expirations || []).map(function(x){ return new Date(String(x) + 'T12:00:00'); });
    if (exps.length) {
      picker.getRange(2, cols.exp, exps.length, 1).setValues(exps.map(function(x){return [x];})).setNumberFormat('yyyy-mm-dd');
      var rule = SpreadsheetApp.newDataValidation().requireValueInRange(picker.getRange(2, cols.exp, exps.length, 1), true).setAllowInvalid(false).build();
      put.getRange(row,5).setDataValidation(rule).setNote('Yahoo/yfinance expirations loaded automatically.');
    } else {
      put.getRange(row,5).clearDataValidations().setNote('Yahoo returned no listed expirations for this ticker.');
    }
    return;
  }

  if (result.mode === 'strikes') {
    var r = Number(result.row);
    if (r < PMF_YF.firstRow || r > PMF_YF.lastRow) return;
    var currentTicker = String(put.getRange(r,4).getDisplayValue() || '').trim().toUpperCase();
    var currentExp = pmfDateIso_(put.getRange(r,5).getValue());
    if (currentTicker !== String(result.ticker || '').trim().toUpperCase() || currentExp !== String(result.expiration || '')) return;
    var c = pmfPickerCols_(r);
    picker.getRange(2, c.strike, picker.getMaxRows() - 1, 1).clearContent();
    var strikes = (result.strikes || []).map(Number).filter(function(x){return isFinite(x);});
    if (strikes.length) {
      picker.getRange(2, c.strike, strikes.length, 1).setValues(strikes.map(function(x){return [x];}));
      var srule = SpreadsheetApp.newDataValidation().requireValueInRange(picker.getRange(2, c.strike, strikes.length, 1), true).setAllowInvalid(false).build();
      put.getRange(r,6).setDataValidation(srule).setNote('Yahoo/yfinance strikes for the selected expiration.');
    } else {
      put.getRange(r,6).clearDataValidations().setNote('Yahoo returned no strikes for this expiration/type.');
    }
    return;
  }

  if (result.mode === 'quotes') {
    pmfUpsertQuoteCache_(result.quotes || [], result.updated_at || '');
  }
}

function pmfUpsertQuoteCache_(quotes, updatedAt) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = pmfEnsureSheet_(ss, PMF_YF.cacheSheet, 500, 7, true);
  var map = {};
  var last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2,1,last-1,7).getValues().forEach(function(r){ if (r[0]) map[String(r[0])] = r; });
  }
  quotes.forEach(function(q){
    if (!q || !q.key) return;
    map[String(q.key)] = [
      q.key,
      q.bid === null || q.bid === undefined ? '' : q.bid,
      q.ask === null || q.ask === undefined ? '' : q.ask,
      q.last === null || q.last === undefined ? '' : q.last,
      q.mark === null || q.mark === undefined ? '' : q.mark,
      q.status || 'YFINANCE ERROR',
      updatedAt || new Date()
    ];
  });
  var rows = Object.keys(map).sort().map(function(k){ return map[k]; });
  if (sh.getMaxRows() > 1) sh.getRange(2,1,sh.getMaxRows()-1,7).clearContent();
  if (rows.length) {
    if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
    sh.getRange(2,1,rows.length,7).setValues(rows);
    sh.getRange(2,2,rows.length,4).setNumberFormat('$0.00');
  }
  SpreadsheetApp.flush();
}

function pmfClearPicker_(put, row, clearExpiration, clearStrike) {
  var ss = put.getParent();
  var picker = pmfEnsureSheet_(ss, PMF_YF.pickerSheet, 600, 80, true);
  var cols = pmfPickerCols_(row);
  if (clearExpiration) {
    put.getRange(row,5).clearContent().clearDataValidations().clearNote();
    picker.getRange(2,cols.exp,picker.getMaxRows()-1,1).clearContent();
  }
  if (clearStrike) {
    put.getRange(row,6).clearContent().clearDataValidations().clearNote();
    picker.getRange(2,cols.strike,picker.getMaxRows()-1,1).clearContent();
  }
}

function pmfPickerCols_(row) {
  var i = row - PMF_YF.firstRow;
  return {exp:1 + i * 2, strike:2 + i * 2};
}

function pmfGitHubConfig_() {
  var p = PropertiesService.getScriptProperties();
  var token = String(p.getProperty('GITHUB_MARKET_ACCESS') || '').trim();
  var repo = String(p.getProperty('GITHUB_REPO_FULL_NAME') || 'G-enterpriseGroup/plaid-etrade-sheets-app').trim();
  var branch = String(p.getProperty('GITHUB_BRANCH') || 'main').trim();
  var workflow = String(p.getProperty('GITHUB_PMF_WORKFLOW') || PMF_YF.workflow).trim();
  if (!token) throw new Error('Missing Script Property GITHUB_MARKET_ACCESS.');
  if (!repo) throw new Error('Missing Script Property GITHUB_REPO_FULL_NAME.');
  return {token:token, repo:repo, branch:branch, workflow:workflow};
}

function pmfGitHubFetch_(url, options) {
  var cfg = pmfGitHubConfig_();
  options = options || {};
  var headers = options.headers || {};
  headers.Authorization = 'Bearer ' + cfg.token;
  headers.Accept = 'application/vnd.github+json';
  headers['X-GitHub-Api-Version'] = '2022-11-28';
  options.headers = headers;
  options.muteHttpExceptions = true;
  if (options.followRedirects === undefined) options.followRedirects = true;
  var res = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  if (code >= 400) throw new Error('GitHub API HTTP ' + code + ': ' + res.getContentText().slice(0,1000));
  return res;
}

function pmfEnsureSheet_(ss, name, minRows, minCols, hidden) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getMaxRows() < minRows) sh.insertRowsAfter(sh.getMaxRows(), minRows - sh.getMaxRows());
  if (sh.getMaxColumns() < minCols) sh.insertColumnsAfter(sh.getMaxColumns(), minCols - sh.getMaxColumns());
  if (hidden) try { sh.hideSheet(); } catch(e) {}
  return sh;
}

function pmfDeleteTriggers_() {
  var names = {pmfYahooOnEdit:true, pmfYahooPollResults:true, pmfYahooRefreshAllQuotes:true};
  ScriptApp.getProjectTriggers().forEach(function(t){ if (names[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t); });
}

function pmfIsOptionType_(value) {
  return /CALL|PUT|LEAPS/i.test(String(value || ''));
}

function pmfOptionType_(value) {
  var s = String(value || '').toUpperCase();
  if (s.indexOf('CALL') >= 0) return 'CALL';
  if (s.indexOf('PUT') >= 0) return 'PUT';
  return '';
}

function pmfDateIso_(value) {
  if (!value) return '';
  var d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, PMF_YF.timezone, 'yyyy-MM-dd');
}

function pmfTrimStrike_(value) {
  var n = Number(value);
  if (!isFinite(n)) return '';
  return String(Number(n.toFixed(3)));
}

function pmfIsMarketWindow_() {
  var now = new Date();
  var day = Number(Utilities.formatDate(now, PMF_YF.timezone, 'u'));
  if (day >= 6) return false;
  var h = Number(Utilities.formatDate(now, PMF_YF.timezone, 'H'));
  var m = Number(Utilities.formatDate(now, PMF_YF.timezone, 'm'));
  var minutes = h * 60 + m;
  return minutes >= (9 * 60 + 25) && minutes <= (16 * 60 + 10);
}
