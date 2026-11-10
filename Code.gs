/**
 * Daily ATS jobs into a Google Sheet, via the fecklesspleb/ats-jobs-scraper
 * Actor on Apify (https://apify.com/fecklesspleb/ats-jobs-scraper).
 *
 * Script Properties (Apps Script > Project Settings):
 *   APIFY_TOKEN  Your Apify API token. Required. Never paste it into this file.
 *   BOARDS       Comma-separated board URLs or slugs, for example:
 *                https://boards.greenhouse.io/airbnb, lever:spotify, ashby:ramp
 *
 * Run runDaily() once by hand to grant permissions, then installTrigger()
 * once to schedule it every day at 08:00 UTC (change the project time zone
 * in appsscript.json if you want a different clock).
 *
 * Each run is capped at maxTotalChargeUsd 0.10. Jobs are de-duplicated by
 * uid (ats:board:id), not by URL. Errors are written to the "Run log" sheet.
 */

var ACTOR_ID = 'fecklesspleb~ats-jobs-scraper';
var ACTOR_SYNC_URL = 'https://api.apify.com/v2/acts/' + ACTOR_ID + '/run-sync-get-dataset-items';
var MAX_TOTAL_CHARGE_USD = 0.10;
var MAX_TOTAL_JOBS = 50; // Actor price is $0.002 per job; 50 jobs = $0.10
var STATE_STORE_NAME = 'ats-jobs-google-sheet';
var JOBS_SHEET_NAME = 'Jobs';
var LOG_SHEET_NAME = 'Run log';
var COLUMNS = [
  'uid', 'ats', 'board', 'company', 'title', 'location',
  'workplace', 'posted_at', 'salary', 'url', 'first_seen'
];
var LOG_COLUMNS = ['time', 'status', 'added', 'message'];

function runDaily() {
  try {
    var result = fetchAndAppend_();
    logRow_('ok', result.added, result.message);
  } catch (e) {
    var message = redact_(String(e && e.message ? e.message : e));
    logRow_('error', 0, message);
    throw e;
  }
}

function installTrigger() {
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'runDaily') {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  ScriptApp.newTrigger('runDaily').timeBased().everyDays(1).atHour(8).create();
}

function fetchAndAppend_() {
  var props = PropertiesService.getScriptProperties();
  var token = String(props.getProperty('APIFY_TOKEN') || '').trim();
  if (!token) {
    throw new Error(
      'Set Script Property APIFY_TOKEN (Apify Console > Settings > API & Integrations). ' +
      'Do not put the token in this file.'
    );
  }
  var boards = parseBoards_(props.getProperty('BOARDS'));
  if (!boards.length) {
    throw new Error('Set Script Property BOARDS to a comma-separated list of board URLs or slugs.');
  }

  var url = ACTOR_SYNC_URL +
    '?maxTotalChargeUsd=' + encodeURIComponent(String(MAX_TOTAL_CHARGE_USD)) +
    '&timeout=280&clean=true';
  var body = {
    companies: boards,
    includeDescription: false,
    onlyNewJobs: true,
    stateStoreName: STATE_STORE_NAME,
    maxTotalJobs: MAX_TOTAL_JOBS
  };

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  var text = response.getContentText() || '';
  if (code === 408) {
    throw new Error("The run did not finish within run-sync's 300 s limit. List fewer boards.");
  }
  if (code < 200 || code >= 300) {
    throw new Error('Apify API error ' + code + ': ' + text.substring(0, 300));
  }
  var items;
  try {
    items = JSON.parse(text);
  } catch (err) {
    throw new Error('Unexpected Apify response (not JSON): ' + text.substring(0, 300));
  }
  if (Object.prototype.toString.call(items) !== '[object Array]') {
    throw new Error('Unexpected Apify response: ' + String(items).substring(0, 300));
  }

  var sheet = ensureSheet_(JOBS_SHEET_NAME, COLUMNS);
  var seen = loadSeen_(sheet);
  var nowIso = Utilities.formatDate(new Date(), 'Etc/UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var added = [];
  var batch = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    if (!isJobItem_(item)) continue;
    var row = fromActorItem_(item, nowIso);
    if (!row.uid) continue;
    if (seen[row.uid] || batch[row.uid]) continue;
    batch[row.uid] = true;
    added.push(row);
  }
  if (added.length) {
    var values = [];
    for (var r = 0; r < added.length; r++) {
      var line = [];
      for (var c = 0; c < COLUMNS.length; c++) {
        var val = added[r][COLUMNS[c]];
        line.push(val == null ? '' : val);
      }
      values.push(line);
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, COLUMNS.length).setValues(values);
  }
  return {
    added: added.length,
    message: items.length + ' records from Apify, ' + added.length + ' new rows appended'
  };
}

function parseBoards_(raw) {
  var out = [];
  String(raw || '').split(',').forEach(function (part) {
    var s = String(part || '').replace(/^\s+|\s+$/g, '');
    if (s && s.charAt(0) !== '#') out.push(s);
  });
  return out;
}

function isJobItem_(item) {
  if (!item || item.type === 'status') return false;
  if (item.status && item.status !== 'active') return false;
  return !!(item.url && item.title && item.uid);
}

function fromActorItem_(item, seenAt) {
  return {
    uid: item.uid,
    ats: item.ats || '',
    board: item.companySlug || '',
    company: item.companyName || item.companySlug || '',
    title: String(item.title || '').replace(/^\s+|\s+$/g, ''),
    location: item.location || '',
    workplace: item.workArrangement || '',
    posted_at: item.postedAt || '',
    salary: item.salaryText || '',
    url: item.url || '',
    first_seen: seenAt
  };
}

function ensureSheet_(name, columns) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
  }
  return sheet;
}

function loadSeen_(sheet) {
  var values = sheet.getDataRange().getValues();
  var seen = {};
  if (!values.length) return seen;
  var header = values[0];
  var uidCol = 0;
  for (var i = 0; i < header.length; i++) {
    if (String(header[i]) === 'uid') uidCol = i;
  }
  for (var r = 1; r < values.length; r++) {
    var uid = String(values[r][uidCol] || '');
    if (uid) seen[uid] = true;
  }
  return seen;
}

function logRow_(status, added, message) {
  var sheet = ensureSheet_(LOG_SHEET_NAME, LOG_COLUMNS);
  var time = Utilities.formatDate(new Date(), 'Etc/UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  sheet.appendRow([time, status, added, redact_(message)]);
}

function redact_(message) {
  var text = String(message || '');
  var token = '';
  try {
    token = String((PropertiesService.getScriptProperties().getProperty('APIFY_TOKEN') || '')).trim();
  } catch (e) {
    token = '';
  }
  if (token) text = text.split(token).join('[redacted]');
  return text;
}
