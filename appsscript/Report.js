/**
 * Report.js — the Validation Report tab: what changed, what was flagged,
 * and what needs a banker's judgement. Rebuilt on demand.
 */

var REPORT_SHEET = 'Validation Report';

function buildValidationReport() {
  var ss = ss_();
  var sh = ss.getSheetByName(REPORT_SHEET) || ss.insertSheet(REPORT_SHEET);
  sh.clear();
  sh.setHiddenGridlines(true);
  var tz = ss.getSpreadsheetTimeZone();
  var model = loadModel();
  var snap = modelSnapshot();
  var ex = readExceptions();
  var runs = readRuns(30);
  var row = 1;

  function title(text) { sh.getRange(row, 2).setValue(text).setFontSize(18).setFontFamily('Georgia').setFontColor(BRAND.midnight); row += 1; }
  function h2(text) { row += 1; sh.getRange(row, 2).setValue(text).setFontSize(12).setFontWeight('bold').setFontColor(BRAND.primary); row += 1; }
  function line(text, opts) { var r = sh.getRange(row, 2).setValue(text).setWrap(true); if (opts && opts.muted) r.setFontColor('#5F6368'); row += 1; }
  function table(headers, rows, widths) {
    if (!rows.length) { line('None.', { muted: true }); return; }
    sh.getRange(row, 2, 1, headers.length).setValues([headers]);
    styleHeader(sh.getRange(row, 2, 1, headers.length));
    sh.getRange(row + 1, 2, rows.length, headers.length).setValues(rows).setWrap(true).setVerticalAlignment('top');
    row += rows.length + 1;
  }

  title('Areta Comps — Validation Report');
  line('Generated ' + Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy HH:mm') + ' by ' + currentUser() + ' · ' + ss.getName(), { muted: true });
  line('Price date: ' + snap.priceDate + ' · Prior price date: ' + snap.priorDate + ' · Financials: ' + snap.financialsSource, { muted: true });

  // 1. Summary
  var open = ex.filter(function (e) { return e.status !== 'Reviewed'; });
  var c = function (s) { return open.filter(function (e) { return e.severity === s; }).length; };
  h2('1. Summary');
  table(['Measure', 'Value'], [
    ['Companies in model', model.layout.companies.length],
    ['Peer groups', model.layout.groups.length],
    ['Open ERROR items (value not written / broken calc)', c('ERROR')],
    ['Open REVIEW items (banker judgement)', c('REVIEW')],
    ['INFO items', c('INFO')],
    ['Items signed off as Reviewed', ex.length - open.length],
    ['Peer-group medians: EV / Revenue, EV / EBITDA (all companies)', snap.allMedian ? snap.allMedian.evRevenue + ' · ' + snap.allMedian.evEbitda : '']
  ]);

  // 2. Items requiring banker judgement
  h2('2. Items requiring banker judgement');
  line('These values were written to the model but rest on a judgement the engine will not make. Mark each Reviewed on the Exceptions tab once decided.', { muted: true });
  table(['Company', 'Field', 'Issue', 'Source', 'Status'], open.filter(function (e) { return e.severity === 'REVIEW'; }).map(function (e) { return [e.company || '(model)', e.field, e.message, e.source, e.status]; }));

  // 3. Errors
  h2('3. Errors (not written, or calculation problems)');
  table(['Company', 'Field', 'Issue', 'Source'], open.filter(function (e) { return e.severity === 'ERROR'; }).map(function (e) { return [e.company || '(model)', e.field, e.message, e.source]; }));

  // 4. Changes made
  h2('4. Changes made (latest runs)');
  var audit = ss.getSheetByName(SHEETS.AUDIT);
  var changes = [];
  if (audit && audit.getLastRow() > 1) {
    var vals = audit.getRange(2, 1, audit.getLastRow() - 1, AUDIT_HEADERS.length).getValues();
    var latestRuns = {};
    runs.filter(function (r) { return r.type === 'apply' || r.type === 'prices'; }).slice(0, 2).forEach(function (r) { latestRuns[r.id] = true; });
    changes = vals.filter(function (v) { return latestRuns[v[1]]; }).map(function (v) {
      return [v[4], v[6], fmtCell(v[7]), fmtCell(v[8]), [v[9], v[10], v[11]].filter(String).join(' › ')];
    });
  }
  table(['Company', 'Field', 'Old', 'New', 'Source'], changes);

  // 5. Informational
  h2('5. For information');
  table(['Company', 'Issue'], open.filter(function (e) { return e.severity === 'INFO'; }).map(function (e) { return [e.company || '(model)', e.message]; }));

  // 6. Signed off
  h2('6. Signed off');
  table(['Company', 'Issue', 'Reviewed by', 'Comment'], ex.filter(function (e) { return e.status === 'Reviewed'; }).map(function (e) { return [e.company, e.message, e.reviewedBy, e.comment]; }));

  // 7. Run history
  h2('7. Run history');
  table(['Run', 'When', 'Type', 'Status', 'Summary'], runs.map(function (r) { return [r.id, r.started ? Utilities.formatDate(new Date(r.started), tz, 'dd-MMM HH:mm') : '', r.type, r.status, r.summary]; }));

  // 8. Method
  h2('8. Method and conventions');
  [
    'EV = market cap + net debt + NCI + preferred, as reported by CIQ (no Areta adjustments). Net debt = total debt − cash & short-term investments − long-term marketable securities, matching CIQ\'s TEV bridge.',
    'Every import rebuilds market cap (CIQ price × shares) and EV and reconciles them to CIQ within ' + pct(getSettings().reconTolerance, 1) + '. A failed reconciliation blocks the value.',
    'Revenue = CIQ standardised Total Revenue; EBITDA = CIQ standardised EBITDA. Where CIQ shows no EBITDA (financial-company template) the engine shows implied EBITDA = operating income + D&A and flags it.',
    'Prices: GOOGLEFINANCE last price, cross-checked against Yahoo Finance; never overwritten by a failed or invalid quote. Multiples show NM when the denominator is zero or negative; medians ignore NM.'
  ].forEach(function (t) { line('• ' + t); });

  sh.setColumnWidth(1, 24); sh.setColumnWidth(2, 220); sh.setColumnWidth(3, 160); sh.setColumnWidth(4, 520); sh.setColumnWidth(5, 280); sh.setColumnWidth(6, 120);
  sh.getRange(1, 1, row, 7).setFontFamily('Arial');
  sh.getRange(1, 2).setFontFamily('Georgia');
  return { sheet: REPORT_SHEET, gid: sh.getSheetId() };
}

function fmtCell(v) {
  if (typeof v === 'number') return Math.abs(v) >= 100 ? fmt(v, 0) : fmt(v, 2);
  if (v instanceof Date) return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'dd-MMM-yyyy');
  return v;
}
