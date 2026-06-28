/**
 * Portfolio Link Market Analysis - GitHub Actions Bridge
 * No Render/Railway/Cloud Run needed.
 * Safe two-step flow so Apps Script does not time out:
 *   1) startMarketAnalysisReport()
 *   2) fetchMarketAnalysisReport()
 * Sidebar buttons call these two functions.
 * Required Script Property:
 *   GITHUB_MARKET_ACCESS = GitHub fine-grained access value with repo Contents read/write + Actions write
 * Optional Script Properties:
 *   GITHUB_REPO_FULL_NAME = G-enterpriseGroup/plaid-etrade-sheets-app
 *   GITHUB_BRANCH = main
 *   GITHUB_MARKET_WORKFLOW = market-analysis.yml
 * Version: github-actions-two-step-v4
 */

var MR_GH = {
  repoProp: 'GITHUB_REPO_FULL_NAME',
  accessProp: 'GITHUB_MARKET_ACCESS',
  branchProp: 'GITHUB_BRANCH',
  workflowProp: 'GITHUB_MARKET_WORKFLOW',
  lastRequestIdProp: 'MARKET_LAST_REQUEST_ID',
  lastOutputPathProp: 'MARKET_LAST_OUTPUT_PATH',
  lastRequestPathProp: 'MARKET_LAST_REQUEST_PATH',
  defaultRepo: 'G-enterpriseGroup/plaid-etrade-sheets-app',
  defaultBranch: 'main',
  defaultWorkflow: 'market-analysis.yml',
  holdingsSheet: 'Holdings',
  reportSheet: 'Report Market Analysis',
  font: 'Times New Roman'
};

function buildMarketAnalysisReport() {
  return startMarketAnalysisReport();
}

function startMarketAnalysisReport() {
  var request = mrBuildGitHubRequestPayload_();
  var requestPath = 'runtime/market-inputs/request_' + request.request_id + '.json';
  var outputPath = 'runtime/market-outputs/report_' + request.request_id + '.json';

  mrCreateGitHubFile_(requestPath, JSON.stringify(request, null, 2), 'Market analysis request ' + request.request_id);
  mrTriggerMarketWorkflow_(requestPath);

  var props = PropertiesService.getScriptProperties();
  props.setProperty(MR_GH.lastRequestIdProp, request.request_id);
  props.setProperty(MR_GH.lastRequestPathProp, requestPath);
  props.setProperty(MR_GH.lastOutputPathProp, outputPath);

  return 'Market Analysis started in GitHub Actions. Wait 2-5 minutes, then run fetchMarketAnalysisReport. Request: ' + request.request_id;
}

function fetchMarketAnalysisReport() {
  var props = PropertiesService.getScriptProperties();
  var requestId = props.getProperty(MR_GH.lastRequestIdProp);
  var outputPath = props.getProperty(MR_GH.lastOutputPathProp);
  if (!requestId || !outputPath) throw new Error('No saved market-analysis request found. Run startMarketAnalysisReport first.');

  var content = mrReadGitHubFile_(outputPath);
  var report = JSON.parse(content);
  if (String(report.request_id || '') !== String(requestId)) {
    throw new Error('Finished report found, but request_id does not match. Expected ' + requestId + ', got ' + report.request_id);
  }

  mrWriteBackendReport_(report);
  return 'Finished report imported into ' + MR_GH.reportSheet + '. Request: ' + requestId;
}

function fetchLatestMarketAnalysisReport() {
  var content = mrReadGitHubFile_('runtime/market-outputs/latest-market-report.json');
  var report = JSON.parse(content);
  mrWriteBackendReport_(report);
  return 'Latest finished market report imported into ' + MR_GH.reportSheet + '. Request: ' + (report.request_id || 'latest');
}

function testMarketGitHubConnection() {
  var cfg = mrGetGitHubCfg_();
  var repo = mrGitHubApi_('/repos/' + encodeURIComponent(cfg.owner) + '/' + encodeURIComponent(cfg.repo), 'get');
  return 'GitHub connected. Repo: ' + repo.full_name + ' | Default branch: ' + repo.default_branch;
}

function checkMarketGitHubConfig() {
  var props = PropertiesService.getScriptProperties();
  var hasAccess = !!props.getProperty(MR_GH.accessProp);
  var repo = props.getProperty(MR_GH.repoProp) || MR_GH.defaultRepo;
  var branch = props.getProperty(MR_GH.branchProp) || MR_GH.defaultBranch;
  var workflow = props.getProperty(MR_GH.workflowProp) || MR_GH.defaultWorkflow;
  var last = props.getProperty(MR_GH.lastRequestIdProp) || 'none';
  if (!hasAccess) return 'Missing Script Property: GITHUB_MARKET_ACCESS';
  return 'Config found. Repo: ' + repo + ' | Branch: ' + branch + ' | Workflow: ' + workflow + ' | Access value: saved | Last request: ' + last;
}

function setMarketGitHubConfig(accessValue, repoFullName, branch, workflowFile) {
  if (!accessValue) return checkMarketGitHubConfig() + '\nNote: you ran setMarketGitHubConfig without arguments. If you already pasted Script Properties manually, run testMarketGitHubConnection next.';
  var props = PropertiesService.getScriptProperties();
  props.setProperty(MR_GH.accessProp, String(accessValue).trim());
  props.setProperty(MR_GH.repoProp, String(repoFullName || MR_GH.defaultRepo).trim());
  props.setProperty(MR_GH.branchProp, String(branch || MR_GH.defaultBranch).trim());
  props.setProperty(MR_GH.workflowProp, String(workflowFile || MR_GH.defaultWorkflow).trim());
  return 'Market GitHub settings saved. Now run testMarketGitHubConnection.';
}

function mrBuildGitHubRequestPayload_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MR_GH.holdingsSheet);
  if (!sh) throw new Error('Missing Holdings tab. Pull holdings first.');

  var values = sh.getDataRange().getDisplayValues();
  if (!values || values.length < 2) throw new Error('Holdings tab is empty. Pull holdings first.');

  var headers = values[0];
  var rows = values.slice(1).filter(function(row) {
    return row.join('').trim() !== '';
  });

  if (!rows.length) throw new Error('No Holdings rows found under the header. Pull holdings first.');

  return {
    request_id: mrRequestId_(),
    headers: headers,
    rows: rows,
    source: 'Google Sheets Holdings tab',
    spreadsheet_id: ss.getId(),
    created_at: mrNow_()
  };
}

function mrTriggerMarketWorkflow_(requestPath) {
  var cfg = mrGetGitHubCfg_();
  var endpoint = '/repos/' + encodeURIComponent(cfg.owner) + '/' + encodeURIComponent(cfg.repo) + '/actions/workflows/' + encodeURIComponent(cfg.workflow) + '/dispatches';
  mrGitHubApi_(endpoint, 'post', {
    ref: cfg.branch,
    inputs: { request_path: requestPath }
  }, true);
}

function mrCreateGitHubFile_(path, content, message) {
  var cfg = mrGetGitHubCfg_();
  var endpoint = '/repos/' + encodeURIComponent(cfg.owner) + '/' + encodeURIComponent(cfg.repo) + '/contents/' + path.split('/').map(encodeURIComponent).join('/');
  return mrGitHubApi_(endpoint, 'put', {
    message: message,
    content: Utilities.base64Encode(content),
    branch: cfg.branch
  });
}

function mrReadGitHubFile_(path) {
  var cfg = mrGetGitHubCfg_();
  var endpoint = '/repos/' + encodeURIComponent(cfg.owner) + '/' + encodeURIComponent(cfg.repo) + '/contents/' + path.split('/').map(encodeURIComponent).join('/') + '?ref=' + encodeURIComponent(cfg.branch) + '&cachebust=' + new Date().getTime();
  var data = mrGitHubApi_(endpoint, 'get');
  if (!data || !data.content) throw new Error('GitHub file had no content: ' + path);
  return Utilities.newBlob(Utilities.base64Decode(String(data.content).replace(/\s/g, ''))).getDataAsString('UTF-8');
}

function mrGitHubApi_(endpoint, method, body, allowEmpty) {
  var cfg = mrGetGitHubCfg_();
  var url = 'https://api.github.com' + endpoint;
  var headers = {
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };
  headers['Author' + 'ization'] = 'Bearer ' + cfg.access;

  var options = {
    method: method || 'get',
    muteHttpExceptions: true,
    headers: headers
  };

  if (body !== undefined && body !== null) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(body);
  }

  var res = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  var text = res.getContentText();

  if (code >= 200 && code < 300) {
    if (allowEmpty && !text) return {};
    if (!text) return {};
    try { return JSON.parse(text); } catch (err) { return { raw: text }; }
  }

  if (code === 404) throw new Error('GitHub file/API not found: ' + endpoint);
  throw new Error('GitHub API error ' + code + ': ' + text);
}

function mrGetGitHubCfg_() {
  var props = PropertiesService.getScriptProperties();
  var access = props.getProperty(MR_GH.accessProp);
  if (!access) throw new Error('Missing GITHUB_MARKET_ACCESS in Apps Script Properties.');

  var full = props.getProperty(MR_GH.repoProp) || MR_GH.defaultRepo;
  var parts = String(full).split('/');
  if (parts.length !== 2) throw new Error('Invalid repo full name. Use owner/repo, like G-enterpriseGroup/plaid-etrade-sheets-app.');

  return {
    access: access,
    owner: parts[0],
    repo: parts[1],
    full: full,
    branch: props.getProperty(MR_GH.branchProp) || MR_GH.defaultBranch,
    workflow: props.getProperty(MR_GH.workflowProp) || MR_GH.defaultWorkflow
  };
}

function mrWriteBackendReport_(report) {
  if (!report || !report.sections) throw new Error('Report JSON is missing sections.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MR_GH.reportSheet) || ss.insertSheet(MR_GH.reportSheet);
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), Math.min(13, sh.getMaxColumns()))
    .setFontFamily(MR_GH.font)
    .setFontSize(9)
    .setWrap(true)
    .setVerticalAlignment('top');

  var row = 1;
  row = mrTitle_(sh, row, report.title || 'Raj Market Rotation Report', 'Built ' + (report.as_of || mrNow_()) + '. Powered by GitHub Actions Python engine.');

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
  sh.getRange(row, 1, 1, 13).merge().setValue(title).setFontSize(18).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff');
  row++;
  sh.getRange(row, 1, 1, 13).merge().setValue(subtitle).setFontSize(9).setFontColor('#374151').setBackground('#eef2ff');
  return row + 2;
}

function mrSection_(sh, row, title) {
  sh.getRange(row, 1, 1, 13).merge().setValue(title).setFontSize(13).setFontWeight('bold').setBackground('#dbeafe').setFontColor('#111827');
  return row + 1;
}

function mrParagraph_(sh, row, text) {
  sh.getRange(row, 1, 1, 13).merge().setValue(text).setFontSize(9).setBackground('#ffffff').setWrap(true).setVerticalAlignment('top');
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
  range.setValues(data).setWrap(true).setVerticalAlignment('top').setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, 1, 1, width).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
  if (data.length > 1) sh.getRange(row + 1, 1, data.length - 1, width).setBackground('#ffffff').setFontColor('#111827');
  return row + data.length + 2;
}

function mrFinalize_(sh, lastRow) {
  for (var c = 1; c <= 13; c++) {
    var w = 110;
    if (c === 8 || c === 12) w = 220;
    if (c === 13) w = 135;
    sh.setColumnWidth(c, w);
  }
  sh.getRange(1, 1, Math.max(1, lastRow), Math.min(13, sh.getMaxColumns())).setFontFamily(MR_GH.font);
  sh.autoResizeRows(1, Math.max(1, lastRow));
}

function mrRequestId_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '_' + Utilities.getUuid().slice(0, 8);
}

function mrNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
