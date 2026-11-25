/**
 * Log.js — Audit Trail, Run Log and Exceptions tabs.
 *
 * Audit Trail: one row per value the engine writes (old -> new, with source cell).
 * Run Log:     one row per run (prices, import, apply, checks) with a summary.
 * Exceptions:  current open items. A banker's "Reviewed" sign-off is kept across
 *              runs, keyed on (category, ticker, field, code).
 */

var AUDIT_HEADERS = ['Timestamp', 'Run ID', 'User', 'Action', 'Company', 'Ticker', 'Field', 'Old value', 'New value', 'Source file', 'Source tab', 'Source cell', 'Source label', 'Source raw value', 'Note'];
var RUN_HEADERS = ['Run ID', 'Started', 'Finished', 'User', 'Type', 'Trigger', 'Status', 'Summary'];
var EXC_HEADERS = ['Key', 'Severity', 'Category', 'Company', 'Ticker', 'Field', 'Issue', 'Source', 'First seen', 'Last seen', 'Run ID', 'Status', 'Reviewed by', 'Banker comment'];

function sheetWithHeaders(name, headers, opts) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader(sh.getRange(1, 1, 1, headers.length));
    sh.setFrozenRows(1);
    if (opts && opts.widths) opts.widths.forEach(function (w, i) { if (w) sh.setColumnWidth(i + 1, w); });
  }
  return sh;
}

function currentUser() {
  try { return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'scheduled'; } catch (e) { return 'scheduled'; }
}

function newRunId(type) {
  var tz = ss_().getSpreadsheetTimeZone();
  return type.toUpperCase().slice(0, 3) + '-' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmmss');
}

function startRun(type, trigger) {
  var run = { id: newRunId(type), type: type, trigger: trigger || 'manual', started: new Date(), user: currentUser() };
  return run;
}

function finishRun(run, status, summary) {
  var sh = sheetWithHeaders(SHEETS.RUNS, RUN_HEADERS, { widths: [170, 150, 150, 200, 90, 90, 90, 700] });
  sh.insertRowAfter(1);
  sh.getRange(2, 1, 1, RUN_HEADERS.length).setValues([[run.id, run.started, new Date(), run.user, run.type, run.trigger, status, summary]])
    .setBackground(null).setFontColor('#000000').setFontWeight('normal');
  sh.getRange(2, 2, 1, 2).setNumberFormat('dd-mmm-yyyy hh:mm');
}

/** Append audit rows: [{action, company, ticker, field, oldV, newV, source:{file,sheet,cell,label,raw}, note}] */
function writeAudit(run, rows) {
  if (!rows.length) return;
  var sh = sheetWithHeaders(SHEETS.AUDIT, AUDIT_HEADERS, { widths: [140, 160, 180, 90, 170, 130, 120, 110, 110, 260, 150, 150, 260, 120, 300] });
  var now = new Date();
  var data = rows.map(function (a) {
    var s = a.source || {};
    return [now, run.id, run.user, a.action, a.company, a.ticker, a.field, display(a.oldV), display(a.newV), s.file || '', s.sheet || '', s.cell || '', s.label || '', display(s.raw), a.note || s.note || ''];
  });
  sh.getRange(sh.getLastRow() + 1, 1, data.length, AUDIT_HEADERS.length).setValues(data);
}

function display(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v;
  return v;
}

/**
 * Replace the open exceptions of the given categories with `items`,
 * carrying over Reviewed status and comments by key. When scopeTickers is given,
 * only those companies' rows are replaced (others keep their open items).
 * items: [{severity, category, company, ticker, field, code, message, source}]
 */
function syncExceptions(run, categories, items, scopeTickers) {
  var sh = sheetWithHeaders(SHEETS.EXCEPTIONS, EXC_HEADERS, { widths: [10, 80, 90, 170, 130, 120, 620, 260, 110, 110, 160, 90, 170, 260] });
  var last = sh.getLastRow();
  var existing = last > 1 ? sh.getRange(2, 1, last - 1, EXC_HEADERS.length).getValues() : [];
  var prev = {};
  var keep = [];
  existing.forEach(function (row) {
    var inScope = categories.indexOf(row[2]) >= 0 && (!scopeTickers || row[2] === 'Import' || scopeTickers.indexOf(row[4]) >= 0);
    if (inScope) prev[row[0]] = row; else keep.push(row);
  });
  var now = new Date();
  var fresh = items.map(function (it) {
    var key = [it.category, it.ticker || '', it.field || '', it.code].join('|');
    var p = prev[key];
    var status = p && p[11] === 'Reviewed' ? 'Reviewed' : 'Open';
    return [key, it.severity, it.category, it.company || '', it.ticker || '', FIELD_LABELS[it.field] || it.field || '', it.message,
      it.source || '', p ? p[8] : now, now, run.id, status, p ? p[12] : '', p ? p[13] : ''];
  });
  var all = fresh.concat(keep);
  var rank = { ERROR: 0, REVIEW: 1, INFO: 2 };
  all.sort(function (a, b) {
    var s = (a[11] === 'Reviewed') - (b[11] === 'Reviewed');
    if (s) return s;
    return (rank[a[1]] - rank[b[1]]) || String(a[3]).localeCompare(String(b[3]));
  });
  if (last > 1) sh.getRange(2, 1, last - 1, EXC_HEADERS.length).clearContent().clearFormat();
  if (all.length) {
    sh.getRange(2, 1, all.length, EXC_HEADERS.length).setValues(all);
    sh.getRange(2, 9, all.length, 2).setNumberFormat('dd-mmm-yyyy');
    sh.getRange(2, 7, all.length, 1).setWrap(true);
    var colors = all.map(function (r) {
      var c = r[11] === 'Reviewed' ? '#E6F4EA' : r[1] === 'ERROR' ? '#FCE8E6' : r[1] === 'REVIEW' ? '#FFF4CE' : '#F1F3F4';
      return [c];
    });
    sh.getRange(2, 2, all.length, 1).setBackgrounds(colors).setFontWeight('bold');
    var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['Open', 'Reviewed'], true).build();
    sh.getRange(2, 12, all.length, 1).setDataValidation(statusRule);
    sh.getRange(2, 12, all.length, 3).setBackground('#FFF2CC');
  }
  sh.hideColumns(1);
  return all.length;
}

/** Read exceptions as objects (for the web app and report). */
function readExceptions() {
  var sh = ss_().getSheetByName(SHEETS.EXCEPTIONS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, EXC_HEADERS.length).getValues().map(function (r) {
    return {
      key: r[0], severity: r[1], category: r[2], company: r[3], ticker: r[4], field: r[5], message: r[6], source: r[7],
      firstSeen: r[8] instanceof Date ? r[8].toISOString() : String(r[8]), lastSeen: r[9] instanceof Date ? r[9].toISOString() : String(r[9]),
      runId: r[10], status: r[11], reviewedBy: r[12], comment: r[13]
    };
  });
}

/** Mark exceptions reviewed (from the web app). */
function markExceptionsReviewed(keys, comment) {
  var sh = ss_().getSheetByName(SHEETS.EXCEPTIONS);
  if (!sh || sh.getLastRow() < 2) return 0;
  var rng = sh.getRange(2, 1, sh.getLastRow() - 1, EXC_HEADERS.length);
  var vals = rng.getValues();
  var n = 0, user = currentUser();
  vals.forEach(function (r) {
    if (keys.indexOf(r[0]) >= 0) { r[11] = 'Reviewed'; r[12] = user; if (comment) r[13] = comment; n++; }
  });
  rng.setValues(vals);
  return n;
}

function readRuns(limit) {
  var sh = ss_().getSheetByName(SHEETS.RUNS);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = Math.min(limit || 15, sh.getLastRow() - 1);
  return sh.getRange(2, 1, n, RUN_HEADERS.length).getValues().map(function (r) {
    return { id: r[0], started: r[1] instanceof Date ? r[1].toISOString() : '', user: r[3], type: r[4], trigger: r[5], status: r[6], summary: r[7] };
  });
}
