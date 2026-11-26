/**
 * Style.js — presentation only: Comps tab layout polish and the Start Here landing tab.
 * Values, formulas and the template's colour conventions are untouched.
 */

var START_SHEET = 'Start Here';
var FONT_BODY = 'Arial';
var FONT_TITLE = 'Georgia';

/** Column widths (px) that fit each Comps header on one line. Keys = model fields. */
var COMPS_WIDTHS = {
  company: 210, ticker: 150, price: 95, priorPrice: 125, change: 95, sharesOut: 120, marketCap: 150,
  netDebt: 140, nci: 100, preferred: 140, ev: 110, ltmRevenue: 165, ltmEbitda: 165, revenueGrowth: 145,
  evRevenue: 120, evEbitda: 120, revenueBasis: 330, ebitdaBasis: 330, financialPeriod: 250, notes: 560
};

/** Layout polish for the Comps tab. Idempotent. */
function styleComps_(model) {
  model = model || loadModel();
  var sh = model.sheet;
  Object.keys(COMPS_WIDTHS).forEach(function (k) {
    if (model.map.cols[k] !== undefined) sh.setColumnWidth(model.map.cols[k] + 1, COMPS_WIDTHS[k]);
  });
  // Header row: one line, vertically centred.
  var hdr = sh.getRange(model.map.headerRow + 1, 1, 1, sh.getLastColumn());
  hdr.setWrap(false).setVerticalAlignment('middle');
  sh.setRowHeight(model.map.headerRow + 1, 34);

  // No vertical freeze line (it cut through the title block); freeze only down to the header row,
  // so the line sits directly under the dark header bar and the title area stays clean.
  sh.setFrozenColumns(0);
  sh.setFrozenRows(model.map.headerRow + 1);

  // Title block: one line each, spanning the table instead of squeezing into Company/Ticker.
  var cName = model.map.cols.company + 1;
  var lastCol = sh.getLastColumn();
  var spanTo = Math.min(lastCol, (model.map.cols.ev !== undefined ? model.map.cols.ev + 1 : cName + 7));
  [2, 3].forEach(function (r) {
    sh.getRange(r, 1, 1, lastCol).breakApart();
    var rng = sh.getRange(r, cName, 1, spanTo - cName + 1);
    try { rng.merge(); } catch (e) { /* layout differs: leave it */ }
    rng.setWrap(false).setVerticalAlignment('middle').setHorizontalAlignment('left');
  });
  sh.getRange(2, cName).setFontSize(16).setFontWeight('bold').setFontColor('#0B0F1A');
  sh.getRange(3, cName).setFontSize(9).setFontColor('#5F6368');
  sh.setRowHeight(1, 12);
  sh.setRowHeight(2, 34);
  sh.setRowHeight(3, 20);
  sh.setRowHeight(4, 26);
  if (model.map.headerRow + 1 > 5) sh.setRowHeight(5, 10);

  // Header value cells: let "Financials source" fill its box.
  var fs = model.header.financialsSource;
  if (fs) {
    var span = Math.max(1, (model.map.cols.ev !== undefined ? model.map.cols.ev : fs.col + 3) - fs.col + 1);
    var box = sh.getRange(fs.row + 1, fs.col + 1, 1, span);
    if (!box.isPartOfMerge()) { try { box.merge(); } catch (e) { /* ignore */ } }
    box.setBackground('#DDEBF7').setHorizontalAlignment('left');
  }
  // Header values (price date, prior price date, financials source) read left to right.
  ['priceDate', 'priorDate', 'financialsSource'].forEach(function (k) {
    var h = model.header[k];
    if (h) sh.getRange(h.row + 1, h.col + 1).setHorizontalAlignment('left').setVerticalAlignment('middle');
  });
  var pr = model.header.priorDate;
  if (pr) sh.getRange(pr.row + 1, pr.col + 1).setNumberFormat('dd-mmm-yyyy');
  // Text columns left, numbers stay right (template convention).
  ['company', 'ticker', 'revenueBasis', 'ebitdaBasis', 'financialPeriod', 'notes'].forEach(function (k) {
    var c = model.map.cols[k];
    if (c !== undefined) sh.getRange(model.map.headerRow + 2, c + 1, Math.max(1, sh.getLastRow() - model.map.headerRow - 1), 1).setHorizontalAlignment('left');
  });

  // Notes and text columns: wrap so nothing spills into the next cell.
  ['notes', 'revenueBasis', 'ebitdaBasis', 'financialPeriod'].forEach(function (k) {
    var c = model.map.cols[k];
    if (c === undefined) return;
    sh.getRange(model.map.headerRow + 2, c + 1, Math.max(1, sh.getLastRow() - model.map.headerRow - 1), 1).setWrap(true).setVerticalAlignment('middle');
  });
  var notesCol = model.map.cols.notes;
  if (notesCol !== undefined) sh.getRange(model.map.headerRow + 2, notesCol + 1, Math.max(1, sh.getLastRow() - model.map.headerRow - 1), 1).setFontColor('#8A5A00');
}

/* ------------------------- Shared page kit ------------------------- */
/* Branded tabs (Start Here, Read Me, Config) share one layout: column A is a
 * gutter, content runs B..F, brand band on row 1, logo row 3, title row 4. */

var PAGE_WIDTHS = [28, 230, 210, 210, 210, 210, 28];
var RUN_ALL_LABEL = '▶ Run all steps (1 → 5)';

/** Reset a tab and draw the brand header. Returns the first free row. */
function brandPage_(sh, title, subtitle, widths) {
  widths = widths || PAGE_WIDTHS;
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.getImages().forEach(function (im) { im.remove(); });
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(0); sh.setFrozenColumns(0);
  var n = widths.length;
  if (sh.getMaxColumns() > n) sh.deleteColumns(n + 1, sh.getMaxColumns() - n);
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange(1, 1, sh.getMaxRows(), n).setFontFamily(FONT_BODY).setFontSize(10).setVerticalAlignment('middle').setFontColor('#3C4043');
  var span = n - 2;

  sh.setRowHeight(1, 14);
  sh.getRange(1, 1, 1, n).setBackground(BRAND.primary);
  sh.setRowHeight(2, 22);
  sh.setRowHeight(3, 56);
  try {
    var img = sh.insertImage(Utilities.newBlob(Utilities.base64Decode(LOGO_PNG_B64), 'image/png', 'areta.png'), 2, 3, 0, 4);
    img.setWidth(206).setHeight(45);
  } catch (e) { sh.getRange(3, 2).setValue('Areta').setFontSize(26).setFontWeight('bold').setFontColor(BRAND.ink); }
  sh.getRange(3, n - 2, 1, 2).merge().setValue('COMPS ENGINE').setFontWeight('bold').setFontSize(9).setFontColor(BRAND.midnight).setHorizontalAlignment('right');
  sh.getRange(4, 2, 1, span).merge().setValue(title).setFontFamily(FONT_TITLE).setFontSize(24).setFontColor(BRAND.ink);
  sh.setRowHeight(4, 46);
  if (subtitle) {
    sh.getRange(5, 2, 1, span).merge().setValue(subtitle).setWrap(true).setFontColor('#5F6368').setFontSize(11);
    sh.setRowHeight(5, textHeight_(subtitle, sumWidths_(widths, 1, span), 11));
  }
  return 7;
}

function sumWidths_(widths, from, count) { var t = 0; for (var i = from; i < from + count; i++) t += widths[i] || 0; return t; }

/** Row height for wrapped text in merged cells (Sheets cannot auto-fit merges). */
function textHeight_(text, px, size) {
  size = size || 10;
  var perLine = Math.max(20, Math.floor(px / (size * 0.56)));
  var lines = String(text).split('\n').reduce(function (n, l) { return n + Math.max(1, Math.ceil(l.length / perLine)); }, 0);
  return Math.max(24, Math.round(lines * size * 1.6 + 10));
}

function sectionTitle_(sh, r, text, span) {
  span = span || 5;
  sh.getRange(r, 2, 1, span).merge().setValue(text).setFontWeight('bold').setFontSize(12).setFontColor(BRAND.midnight)
    .setBorder(null, null, true, null, null, null, BRAND.primary, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(r, 32);
}

function tableHeader_(range) {
  range.setBackground(BRAND.midnight).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('left');
}

/** A label (column B) and wrapped text (C..F), like the Start Here steps. */
function labelRow_(sh, r, label, text, opts) {
  opts = opts || {};
  sh.getRange(r, 2).setValue(label).setFontWeight('bold').setFontColor(opts.labelColor || BRAND.ink).setWrap(true);
  sh.getRange(r, 3, 1, 4).merge().setValue(text).setWrap(true).setFontColor('#3C4043');
  sh.setRowHeight(r, Math.max(opts.min || 30, textHeight_(text, 840), textHeight_(label, 220)));
  sh.getRange(r, 2, 1, 5).setBorder(null, null, true, null, null, null, '#E4E8F2', SpreadsheetApp.BorderStyle.SOLID);
}

/** Full-width paragraph across B..F. */
function paraRow_(sh, r, text, opts) {
  opts = opts || {};
  var rng = sh.getRange(r, 2, 1, 5).merge().setValue(text).setWrap(true).setFontColor(opts.color || '#3C4043');
  if (opts.size) rng.setFontSize(opts.size);
  sh.setRowHeight(r, textHeight_(text, 1060, opts.size || 10));
  return rng;
}

function footer_(sh, r, text) {
  sh.getRange(r, 2, 1, 5).merge().setValue(text).setFontSize(9).setFontColor('#9AA0A6').setWrap(true);
  sh.setRowHeight(r, 26);
}

/** Left-align everything on a branded page (except the brand label) and trim spare rows. */
function finishPage_(sh, lastRow) {
  var n = sh.getMaxColumns();
  sh.getRange(4, 1, Math.max(1, sh.getMaxRows() - 3), n).setHorizontalAlignment('left');
  if (sh.getMaxRows() > lastRow + 3) sh.deleteRows(lastRow + 4, sh.getMaxRows() - lastRow - 3);
}

/** Log-style tabs: every column left-aligned, header styled, dates readable. */
function alignLeft_(sh) {
  if (!sh) return;
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setHorizontalAlignment('left');
}

function polishDataTabs_() {
  var ss = ss_();
  [SHEETS.EXCEPTIONS, SHEETS.REVIEW, SHEETS.HISTORY, SHEETS.AUDIT, SHEETS.RUNS].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    alignLeft_(sh);
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setVerticalAlignment('middle');
    var lc = sh.getLastColumn();
    if (lc) { styleHeader(sh.getRange(1, 1, 1, lc)); sh.setRowHeight(1, 30); }
  });
  var audit = ss.getSheetByName(SHEETS.AUDIT);
  if (audit && audit.getLastRow() > 1) audit.getRange(2, 1, audit.getLastRow() - 1, 1).setNumberFormat('dd-mmm-yyyy hh:mm');
  var hist = ss.getSheetByName(SHEETS.HISTORY);
  if (hist && hist.getLastRow() > 1) hist.getRange(2, 9, hist.getLastRow() - 1, 1).setNumberFormat('dd-mmm-yyyy hh:mm');
}

/* ----------------------------- Start Here ----------------------------- */

/** Areta-branded landing tab with live status and the steps in order. */
function buildStartHere_() {
  var ss = ss_();
  var sh = ss.getSheetByName(START_SHEET) || ss.insertSheet(START_SHEET, 0);
  var r = brandPage_(sh, 'Trading comparables, maintained automatically',
    'Share prices refresh every week. New Capital IQ exports are read, checked against CIQ and applied with a full audit trail. Anything that needs a banker\'s judgement is flagged, never assumed.');

  // Live status cards (formulas, so they are always current).
  sectionTitle_(sh, r, 'Status'); r++;
  var cards = [
    ['Price date', '=IFERROR(TEXT(Comps!C4,"dd-mmm-yyyy"),"–")'],
    ['Prior price date', '=IFERROR(TEXT(Comps!F4,"dd-mmm-yyyy"),"–")'],
    ['Open errors', '=COUNTIFS(Exceptions!B:B,"ERROR",Exceptions!L:L,"Open")'],
    ['For banker review', '=COUNTIFS(Exceptions!B:B,"REVIEW",Exceptions!L:L,"Open")'],
    ['Financials', '=IF(Comps!I4="","–",SUBSTITUTE(Comps!I4,"CIQ export, ",""))']
  ];
  cards.forEach(function (c, i) {
    var col = 2 + i;
    sh.getRange(r, col).setValue(c[0]).setFontSize(9).setFontColor('#5F6368').setBackground(BRAND.frost);
    sh.getRange(r + 1, col).setFormula(c[1]).setFontSize(i === 4 ? 10 : 18).setFontWeight('bold').setFontColor(i === 2 ? '#C5221F' : i === 3 ? '#8A5A00' : BRAND.ink).setBackground(BRAND.frost).setWrap(true);
  });
  sh.setRowHeight(r, 24); sh.setRowHeight(r + 1, 46);
  sh.getRange(r, 2, 2, 5).setBorder(null, null, null, null, true, null, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK);
  r += 3;

  sectionTitle_(sh, r, 'How to use it: menu Areta Comps, in this order'); r++;
  [
    ['1. Import CIQ exports', 'When companies report: drop their Capital IQ "Financials" exports in the CIQ inbox (link below), then run step 1. Each file is matched by ticker, read, and reconciled to CIQ. Results land on the Review Queue tab.'],
    ['2. Apply financial changes', 'Look over the Review Queue: Ready = passed every check, Review = flagged for judgement, Blocked = never written. Untick anything you disagree with, then run step 2. Only the blue input cells change, and the queue empties once applied.'],
    ['3. Refresh share prices', 'Runs automatically every Monday at 07:00 London time. Run it by hand for fresh prices before a meeting. Failed or suspicious quotes never overwrite a good price.'],
    ['4. Run checks', 'Recomputes every formula independently and looks for hardcodes, #errors, missing inputs and outliers. Results go to the Exceptions tab. Tick Reviewed there to sign items off.'],
    ['5. Build validation report', 'Rebuilds the Validation Report tab (one tab, always current) and saves a dated PDF copy to the Reports folder in Drive.'],
    [RUN_ALL_LABEL, 'Runs 1 to 5 in one go. Ready and Review items are applied (Review items stay open on Exceptions for sign-off). Blocked items are never written.']
  ].forEach(function (s, i) { labelRow_(sh, r, s[0], s[1], { labelColor: i === 5 ? BRAND.primary : BRAND.ink, min: 42 }); r++; });
  r++;

  sectionTitle_(sh, r, 'Links'); r++;
  var links = [];
  try { links.push(['CIQ inbox', 'Drop new Capital IQ exports here', getFolders().inbox.getUrl()]); } catch (e) { /* not set up */ }
  try { var w = webAppUrl_(); if (w) links.push(['Web dashboard', 'Same steps in a browser page. Sign in with the web app password (Tools › Set web app password).', w]); } catch (e) { /* none */ }
  links.forEach(function (l) {
    sh.getRange(r, 2).setValue(l[0]).setFontWeight('bold').setFontColor(BRAND.ink);
    sh.getRange(r, 3, 1, 3).merge().setValue(l[1]).setFontColor('#3C4043').setWrap(true);
    sh.getRange(r, 6).setFormula('=HYPERLINK("' + l[2] + '","Open ↗")').setFontColor(BRAND.primary).setFontWeight('bold');
    sh.setRowHeight(r, 32);
    sh.getRange(r, 2, 1, 5).setBorder(null, null, true, null, null, null, '#E4E8F2', SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });
  r++;

  sectionTitle_(sh, r, 'What each tab is for'); r++;
  [
    ['Comps', 'The model. Yellow = weekly inputs, blue = quarterly inputs, black = formulas.'],
    ['Exceptions', 'Open items: ERROR (value not written), REVIEW (judgement), INFO. Sign-offs persist between runs.'],
    ['Review Queue', 'Proposed financial changes with old value, new value and the CIQ source cell. Empty when nothing is waiting.'],
    ['Validation Report', 'Summary of updates, exceptions and items for manual review.'],
    ['Read Me · Config', 'How the model works; schedule, thresholds, access list and ticker map (no code needed).'],
    ['Price History · Audit Trail · Run Log', 'Every price saved, every value written (with source), every run.']
  ].forEach(function (t) { labelRow_(sh, r, t[0], t[1], { min: 30 }); r++; });
  r++;
  sectionTitle_(sh, r, 'Access'); r++;
  paraRow_(sh, r, 'Anyone you share this Sheet with (as an editor) can use the menu and sidebar with their own Google account; every run and change is logged under their email. The web dashboard works in any browser with the web app password, and logs the name entered at sign-in. Scheduled refreshes run under the account that set them up.');
  r += 2;
  footer_(sh, r, 'Built for Areta · Google Sheets + Apps Script · source and docs: github.com/alycz/Areta-Comparables-Automation');
  finishPage_(sh, r);
  return sh;
}

/* ------------------------------- Read Me ------------------------------- */

/**
 * Re-lay the template's Read Me in the branded style. The original text is
 * captured once (bold lines = headings) and kept in Script Properties, so the
 * tab can be rebuilt any number of times without losing Areta's wording.
 */
function buildReadMe_() {
  var sh = ss_().getSheetByName(SHEETS.README);
  if (!sh) return;
  var props = PropertiesService.getScriptProperties();
  var src = null;
  try { src = JSON.parse(props.getProperty('README_SOURCE') || 'null'); } catch (e) { src = null; }
  if (!src || !src.length) {
    src = [];
    var lastRow = Math.max(1, sh.getLastRow()), lastCol = Math.max(1, Math.min(4, sh.getLastColumn()));
    var rng = sh.getRange(1, 1, lastRow, lastCol);
    var vals = rng.getDisplayValues(), weights = rng.getFontWeights();
    for (var i = 0; i < vals.length; i++) {
      for (var c = 0; c < vals[i].length; c++) {
        var t = String(vals[i][c]).trim();
        if (!t) continue;
        if (/^Automation \(Areta Comps engine\)/.test(t) || t === 'COMPS ENGINE') { i = vals.length; break; }
        src.push({ h: weights[i][c] === 'bold', t: t });
        break;
      }
    }
    if (!src.length) return;
    props.setProperty('README_SOURCE', JSON.stringify(src));
  }

  var r = brandPage_(sh, 'Read Me', 'How the comps model is built, what changes when, and the conventions behind every number.');
  var first = true;
  src.forEach(function (line) {
    if (line.h) { if (!first) r++; sectionTitle_(sh, r, line.t); r++; first = false; return; }
    first = false;
    var m = line.t.match(/^([^:]{3,40}):\s+(.+)$/);
    if (m) labelRow_(sh, r, m[1], m[2], { min: 30 });
    else paraRow_(sh, r, line.t);
    r++;
  });
  r++;
  sectionTitle_(sh, r, 'Automation'); r++;
  labelRow_(sh, r, 'Areta Comps engine', 'Prices, CIQ imports, checks and the validation report are automated. The Start Here tab lists the steps in order; run them from the Areta Comps menu or the web dashboard.', { min: 30 }); r++;
  labelRow_(sh, r, 'Colour conventions', 'Unchanged. The engine only writes yellow (weekly) and blue (quarterly) input cells, and keeps the black formula cells as formulas.', { min: 30 }); r++;
  r++;
  footer_(sh, r, 'Text above is from the Areta comps template; layout by the Areta Comps engine.');
  finishPage_(sh, r);
}
