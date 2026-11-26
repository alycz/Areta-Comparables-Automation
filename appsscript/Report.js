/**
 * Report.js — the Validation Report tab.
 *
 * One tab, rebuilt in place each time (always the current state). Each build
 * also saves a dated PDF copy to the "Reports" Drive folder, so past reports
 * are kept without cluttering the workbook.
 */

var REPORT_SHEET = 'Validation Report';

/* Layout: A gutter | B..G content | H gutter. Widths chosen so tables read without cramped wrapping. */
var R_COLS = { first: 2, count: 6 };
var R_WIDTHS = [24, 190, 150, 470, 300, 110, 110, 24];

function buildValidationReport() {
  var ss = ss_();
  var sh = ss.getSheetByName(REPORT_SHEET) || ss.insertSheet(REPORT_SHEET);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.setHiddenGridlines(true);
  if (sh.getMaxColumns() > R_WIDTHS.length) sh.deleteColumns(R_WIDTHS.length + 1, sh.getMaxColumns() - R_WIDTHS.length);
  R_WIDTHS.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  var tz = ss.getSpreadsheetTimeZone();
  var model = loadModel();
  var snap = modelSnapshot();
  var ex = readExceptions();
  var runs = readRuns(30);
  var open = ex.filter(function (e) { return e.status !== 'Reviewed'; });
  var cnt = function (s) { return open.filter(function (e) { return e.severity === s; }).length; };
  var row = 1;

  // Brand band + title.
  sh.getRange(row, 1, 1, 8).setBackground(BRAND.primary); sh.setRowHeight(row, 10); row += 2;
  full(row, 'Validation Report').setFontFamily(FONT_TITLE).setFontSize(24).setFontColor(BRAND.ink); sh.setRowHeight(row, 40); row++;
  full(row, 'Areta Comps · ' + ss.getName() + ' · generated ' + Utilities.formatDate(new Date(), tz, 'dd MMM yyyy, HH:mm') + ' (' + tz + ') by ' + currentUser())
    .setFontColor('#5F6368'); row += 2;

  // KPI cards.
  var kpis = [
    ['Companies', model.layout.companies.length, BRAND.ink],
    ['Open errors', cnt('ERROR'), cnt('ERROR') ? '#C5221F' : '#137333'],
    ['For review', cnt('REVIEW'), cnt('REVIEW') ? '#8A5A00' : '#137333'],
    ['Info', cnt('INFO'), '#5F6368'],
    ['Signed off', ex.length - open.length, '#137333'],
    ['Median EV / EBITDA', snap.allMedian ? snap.allMedian.evEbitda : '–', BRAND.ink]
  ];
  kpis.forEach(function (k, i) {
    var c = R_COLS.first + i;
    sh.getRange(row, c).setValue(k[0]).setFontSize(9).setFontColor('#5F6368').setBackground(BRAND.frost).setHorizontalAlignment('left');
    sh.getRange(row + 1, c).setValue(k[1]).setFontSize(20).setFontWeight('bold').setFontColor(k[2]).setBackground(BRAND.frost).setHorizontalAlignment('left');
  });
  sh.getRange(row, R_COLS.first, 2, 6).setBorder(null, null, null, null, true, null, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.setRowHeight(row, 24); sh.setRowHeight(row + 1, 40);
  row += 3;

  // Data as-of strip.
  kv([
    ['Price date', snap.priceDate || '–'],
    ['Prior price date', snap.priorDate || '–'],
    ['Financials', snap.financialsSource || '–'],
    ['Median EV / Revenue (all)', snap.allMedian ? snap.allMedian.evRevenue : '–']
  ]);

  section('1  Items requiring banker judgement',
    'Written to the model, but resting on a judgement the engine will not make. Decide, then mark Reviewed on the Exceptions tab.');
  table(['Company', 'Field', 'Issue', 'Source', 'Status'], [1, 1, 1, 1, 2],
    open.filter(function (e) { return e.severity === 'REVIEW'; }).map(function (e) { return [e.company || 'Model', e.field, e.message, e.source, e.status]; }),
    { pill: 4 });

  section('2  Errors', 'Values that were not written (the previous value is kept) or calculations that failed a check.');
  table(['Company', 'Field', 'Issue', 'Source', 'Status'], [1, 1, 1, 1, 2],
    open.filter(function (e) { return e.severity === 'ERROR'; }).map(function (e) { return [e.company || 'Model', e.field, e.message, e.source, e.status]; }),
    { empty: 'No errors. Every value passed validation and every formula recomputes.' });

  section('3  Changes in the latest runs', 'From the Audit Trail: latest financial apply and latest price refresh.');
  var audit = ss.getSheetByName(SHEETS.AUDIT), changes = [];
  if (audit && audit.getLastRow() > 1) {
    var vals = audit.getRange(2, 1, audit.getLastRow() - 1, AUDIT_HEADERS.length).getValues();
    var latest = {};
    var seen = {};
    runs.forEach(function (r) { if ((r.type === 'apply' || r.type === 'prices') && !seen[r.type] && r.status !== 'Failed') { seen[r.type] = true; latest[r.id] = true; } });
    changes = vals.filter(function (v) { return latest[v[1]]; }).map(function (v) {
      return [v[4], v[6], fmtCell(v[7]) + '  →  ' + fmtCell(v[8]), [v[9], v[10], v[11]].filter(String).join(' › '), v[3]];
    });
  }
  table(['Company', 'Field', 'Old  →  New', 'Source', 'Type'], [1, 1, 1, 1, 2], changes.slice(0, 400));
  if (changes.length > 400) note(changes.length - 400 + ' more on the Audit Trail tab.');

  section('4  For information', 'Recorded for completeness; no action needed.');
  table(['Company', 'Field', 'Note', '', ''], [1, 1, 3], open.filter(function (e) { return e.severity === 'INFO'; }).map(function (e) { return [e.company || 'Model', e.field, e.message, '', '']; }));

  section('5  Signed off', 'Items a banker has reviewed, with their comment.');
  table(['Company', 'Field', 'Issue', 'Comment', 'Reviewed by'], [1, 1, 1, 1, 2],
    ex.filter(function (e) { return e.status === 'Reviewed'; }).map(function (e) { return [e.company || 'Model', e.field, e.message, e.comment, e.reviewedBy]; }),
    { empty: 'Nothing signed off yet.' });

  section('6  Run history', 'Every scheduled and manual run.');
  table(['When', 'Type', 'Summary', '', 'Status'], [1, 1, 2, 0, 2], runs.map(function (r) {
    return [r.started ? Utilities.formatDate(new Date(r.started), tz, 'dd MMM HH:mm') : '', r.type + ' · ' + r.trigger, r.summary, '', r.status];
  }), { pill: 4 });

  section('7  Method and conventions', '');
  [
    'EV = market cap + net debt + NCI + preferred, as reported by CIQ. Net debt = total debt − cash & short-term investments − long-term marketable securities (CIQ\'s TEV bridge).',
    'Every import rebuilds market cap (CIQ price × shares) and EV and reconciles them to CIQ within ' + pct(getSettings().reconTolerance, 1) + '. A value that does not reconcile is blocked.',
    'Revenue = CIQ standardised Total Revenue. EBITDA = CIQ standardised EBITDA; where CIQ has none (financial template) the engine shows implied EBITDA = operating income + D&A and flags it.',
    'Prices: Google Finance last price, cross-checked against Yahoo Finance and saved to Price History. Prior price = the saved close at least ' + getSettings().priorPriceMinDays + ' days earlier. A failed quote never overwrites a good price.',
    'Multiples show NM when the denominator is zero or negative; medians ignore NM.'
  ].forEach(function (t) { full(row, '•  ' + t).setWrap(true).setFontColor('#3C4043'); sh.setRowHeight(row, 34); row++; });

  row++;
  full(row, 'This tab is rebuilt each time the report runs. Dated PDF copies are kept in Drive › Reports.').setFontSize(9).setFontColor('#9AA0A6');

  sh.getRange(1, 1, row, 8).setFontFamily(FONT_BODY).setVerticalAlignment('top');
  sh.getRange(3, R_COLS.first).setFontFamily(FONT_TITLE);
  if (sh.getMaxRows() > row + 2) sh.deleteRows(row + 3, sh.getMaxRows() - row - 2);
  SpreadsheetApp.flush();

  var pdf = savePdf_(sh, tz);
  return { sheet: REPORT_SHEET, gid: sh.getSheetId(), pdfUrl: pdf };

  /* ---------- helpers (closures over sh / row) ---------- */
  function full(r, text) {
    var rng = sh.getRange(r, R_COLS.first, 1, R_COLS.count).merge();
    rng.setValue(text);
    return rng;
  }
  function note(text) { full(row, text).setFontColor('#5F6368').setFontSize(9); row++; }
  function section(title, sub) {
    row++;
    full(row, title).setFontSize(13).setFontWeight('bold').setFontColor(BRAND.midnight)
      .setBorder(null, null, true, null, null, null, BRAND.primary, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.setRowHeight(row, 30); row++;
    if (sub) { full(row, sub).setFontColor('#5F6368').setFontSize(9).setWrap(true); sh.setRowHeight(row, 22); row++; }
  }
  function kv(pairs) {
    pairs.forEach(function (p) {
      sh.getRange(row, R_COLS.first).setValue(p[0]).setFontColor('#5F6368');
      sh.getRange(row, R_COLS.first + 1, 1, R_COLS.count - 1).merge().setValue(p[1]).setFontWeight('bold');
      row++;
    });
  }
  /**
   * spans: how many sheet columns each logical column takes (0 = skip). Sum must be 6.
   * opts.pill: index of a status column to colour.
   */
  function table(headers, spans, rows, opts) {
    opts = opts || {};
    if (!rows.length) { full(row, opts.empty || 'None.').setFontColor('#5F6368').setFontStyle('italic'); row++; return; }
    var startCols = [], c = 0;
    spans.forEach(function (sp) { startCols.push(c); c += sp || 0; });
    function grid(values) {
      var out = ['', '', '', '', '', ''];
      values.forEach(function (v, i) { if (spans[i]) out[startCols[i]] = v === null || v === undefined ? '' : v; });
      return out;
    }
    var n = rows.length + 1;
    var block = sh.getRange(row, R_COLS.first, n, R_COLS.count);
    block.setValues([grid(headers)].concat(rows.map(grid))).setWrap(true);
    // Merge multi-column cells in one call per logical column.
    spans.forEach(function (sp, i) { if (sp > 1) sh.getRange(row, R_COLS.first + startCols[i], n, sp).mergeAcross(); });
    sh.getRange(row, R_COLS.first, 1, R_COLS.count).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(BRAND.midnight);
    sh.setRowHeight(row, 26);
    var bodyRows = rows.length, first = row + 1;
    var bg = rows.map(function (rv, i) { var col = i % 2 ? '#F7F9FD' : '#FFFFFF'; return [col, col, col, col, col, col]; });
    var fc = rows.map(function () { return ['#202124', '#5F6368', '#202124', '#5F6368', '#202124', '#202124']; });
    if (opts.pill !== undefined && spans[opts.pill]) {
      var pc = R_COLS.first + startCols[opts.pill] - R_COLS.first;
      rows.forEach(function (rv, i) {
        var v = String(rv[opts.pill]);
        var col = /Open|REVIEW|Partial|Exceptions/.test(v) ? ['#FFF4CE', '#8A5A00'] : /Failed|ERROR/.test(v) ? ['#FCE8E6', '#C5221F'] : ['#E6F4EA', '#137333'];
        for (var k = 0; k < (spans[opts.pill] || 1); k++) { bg[i][pc + k] = col[0]; fc[i][pc + k] = col[1]; }
      });
    }
    var body = sh.getRange(first, R_COLS.first, bodyRows, R_COLS.count);
    body.setBackgrounds(bg).setFontColors(fc);
    body.setBorder(null, null, true, null, null, true, '#E4E8F2', SpreadsheetApp.BorderStyle.SOLID);
    row = first + bodyRows + 1;
  }
}

/** Save the report tab as a dated PDF in Drive › Reports. Returns its URL ('' on failure). */
function savePdf_(sh, tz) {
  try {
    var ss = ss_();
    var folders = getFolders();
    var it = folders.root.getFoldersByName('Reports');
    var reports = it.hasNext() ? it.next() : folders.root.createFolder('Reports');
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf&gid=' + sh.getSheetId() +
      '&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=CENTER&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4';
    var res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return '';
    var name = 'Validation Report ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HHmm') + '.pdf';
    return reports.createFile(res.getBlob().setName(name)).getUrl();
  } catch (e) {
    return '';
  }
}

function fmtCell(v) {
  if (v === '' || v === null || v === undefined) return '–';
  if (typeof v === 'number') return Math.abs(v) >= 100 ? fmt(v, 0) : Math.abs(v) < 1 && v !== 0 ? fmt(v, 4) : fmt(v, 2);
  if (v instanceof Date) return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'dd-MMM-yyyy');
  return v;
}
