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

  // Title and subtitle sit inside the frozen columns, so the freeze line never cuts through text.
  var cName = model.map.cols.company + 1, cTick = model.map.cols.ticker + 1;
  var frozenCols = sh.getFrozenColumns();
  if (frozenCols >= cTick) {
    [2, 3].forEach(function (r) {
      var rng = sh.getRange(r, cName, 1, cTick - cName + 1);
      if (!rng.isPartOfMerge()) { try { rng.merge(); } catch (e) { /* layout differs: leave it */ } }
      rng.setWrap(true).setVerticalAlignment('top');
    });
    sh.getRange(2, cName).setFontSize(14);
    sh.setRowHeight(2, 46);
    sh.getRange(3, cName).setFontSize(9).setFontColor('#5F6368');
    sh.setRowHeight(3, 70);
  }

  // Header value cells: let "Financials source" fill its box.
  var fs = model.header.financialsSource;
  if (fs) {
    var span = Math.max(1, (model.map.cols.ev !== undefined ? model.map.cols.ev : fs.col + 3) - fs.col + 1);
    var box = sh.getRange(fs.row + 1, fs.col + 1, 1, span);
    if (!box.isPartOfMerge()) { try { box.merge(); } catch (e) { /* ignore */ } }
    box.setBackground('#DDEBF7').setHorizontalAlignment('left');
  }
  var pr = model.header.priorDate;
  if (pr) sh.getRange(pr.row + 1, pr.col + 1).setNumberFormat('dd-mmm-yyyy');

  // Notes and text columns: wrap so nothing spills into the next cell.
  ['notes', 'revenueBasis', 'ebitdaBasis', 'financialPeriod'].forEach(function (k) {
    var c = model.map.cols[k];
    if (c === undefined) return;
    sh.getRange(model.map.headerRow + 2, c + 1, Math.max(1, sh.getLastRow() - model.map.headerRow - 1), 1).setWrap(true).setVerticalAlignment('middle');
  });
  var notesCol = model.map.cols.notes;
  if (notesCol !== undefined) sh.getRange(model.map.headerRow + 2, notesCol + 1, Math.max(1, sh.getLastRow() - model.map.headerRow - 1), 1).setFontColor('#8A5A00');
}

/** Areta-branded landing tab with live status and the steps in order. */
function buildStartHere_() {
  var ss = ss_();
  var sh = ss.getSheetByName(START_SHEET) || ss.insertSheet(START_SHEET, 0);
  sh.clear();
  sh.getImages().forEach(function (im) { im.remove(); });
  sh.setHiddenGridlines(true);
  var maxCols = sh.getMaxColumns();
  if (maxCols > 8) sh.deleteColumns(9, maxCols - 8);
  [28, 230, 210, 210, 210, 210, 28, 28].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('A:H').setFontFamily(FONT_BODY).setFontSize(10).setVerticalAlignment('middle');

  // Brand band.
  sh.setRowHeight(1, 18);
  sh.getRange('A1:H1').setBackground(BRAND.primary);
  sh.setRowHeight(2, 24);
  sh.setRowHeight(3, 60);
  try {
    var img = sh.insertImage(Utilities.newBlob(Utilities.base64Decode(LOGO_PNG_B64), 'image/png', 'areta.png'), 2, 3, 0, 6);
    img.setWidth(220).setHeight(48);
  } catch (e) { sh.getRange('B3').setValue('Areta').setFontSize(28).setFontWeight('bold'); }
  sh.getRange('D3:F3').merge().setValue('COMPS ENGINE').setFontWeight('bold').setFontSize(10).setFontColor(BRAND.ink).setHorizontalAlignment('right');
  sh.getRange('B4:F4').merge().setValue('Trading comparables, maintained automatically').setFontFamily(FONT_TITLE).setFontSize(24).setFontColor(BRAND.ink);
  sh.setRowHeight(4, 44);
  sh.getRange('B5:F5').merge().setValue('Share prices refresh every week. New Capital IQ exports are read, checked against CIQ and applied with a full audit trail. Anything that needs a banker\'s judgement is flagged, never assumed.')
    .setWrap(true).setFontColor('#5F6368').setFontSize(11);
  sh.setRowHeight(5, 40);

  // Live status cards (formulas, so they are always current).
  var r = 7;
  sectionTitle_(sh, r, 'Status'); r++;
  var cards = [
    ['Price date', '=IFERROR(TEXT(Comps!C4,"dd-mmm-yyyy"),"–")'],
    ['Prior price date', '=IFERROR(TEXT(Comps!F4,"dd-mmm-yyyy"),"–")'],
    ['Open errors', '=COUNTIFS(Exceptions!B:B,"ERROR",Exceptions!L:L,"Open")'],
    ['For banker review', '=COUNTIFS(Exceptions!B:B,"REVIEW",Exceptions!L:L,"Open")'],
    ['Financials', '=IF(Comps!I4="","–",SUBSTITUTE(Comps!I4,"CIQ export, ",""))']
  ];
  cards.forEach(function (c, i) {
    var col = 2 + (i % 5);
    sh.getRange(r, col).setValue(c[0]).setFontSize(9).setFontColor('#5F6368').setBackground(BRAND.frost);
    sh.getRange(r + 1, col).setFormula(c[1]).setFontSize(i === 4 ? 10 : 18).setFontWeight('bold').setFontColor(i === 2 ? '#C5221F' : i === 3 ? '#8A5A00' : BRAND.ink).setBackground(BRAND.frost).setWrap(true);
  });
  sh.setRowHeight(r, 22); sh.setRowHeight(r + 1, 44);
  sh.getRange(r, 2, 2, 5).setBorder(null, null, null, null, true, null, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK);
  r += 3;

  // Steps.
  sectionTitle_(sh, r, 'How to use it: menu Areta Comps, in this order'); r++;
  var steps = [
    ['1. Import CIQ exports', 'When companies report: drop their Capital IQ "Financials" exports in the CIQ inbox (link below), then run step 1. Each file is matched by ticker, read, and reconciled to CIQ. Results land on the Review Queue tab.'],
    ['2. Apply financial changes', 'Look over the Review Queue: Ready = passed every check, Review = flagged for judgement, Blocked = never written. Untick anything you disagree with, then run step 2. Only the blue input cells change.'],
    ['3. Refresh share prices', 'Runs automatically every Monday at 07:00 London time. Run it by hand for fresh prices before a meeting. Failed or suspicious quotes never overwrite a good price.'],
    ['4. Run checks', 'Recomputes every formula independently and looks for hardcodes, #errors, missing inputs and outliers. Results go to the Exceptions tab. Tick Reviewed there to sign items off.'],
    ['5. Build validation report', 'Rebuilds the Validation Report tab (one tab, always current) and saves a dated PDF copy to the Reports folder in Drive.'],
    ['▶ Run all steps', 'Runs 1 to 5 in one go. Ready and Review items are applied (Review items stay open on Exceptions for sign-off). Blocked items are never written.']
  ];
  steps.forEach(function (s, i) {
    sh.getRange(r, 2).setValue(s[0]).setFontWeight('bold').setFontColor(i === 5 ? BRAND.primary : BRAND.ink);
    sh.getRange(r, 3, 1, 4).merge().setValue(s[1]).setWrap(true).setFontColor('#3C4043');
    sh.setRowHeight(r, 42);
    sh.getRange(r, 2, 1, 5).setBorder(null, null, true, null, null, null, '#E4E8F2', SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });
  r++;

  // Links.
  sectionTitle_(sh, r, 'Links'); r++;
  var links = [];
  try { links.push(['CIQ inbox (drop exports here)', getFolders().inbox.getUrl()]); } catch (e) { /* not set up */ }
  try { var w = webAppUrl_(); if (w) links.push(['Web dashboard', w]); } catch (e) { /* none */ }
  links.forEach(function (l) {
    sh.getRange(r, 2).setValue(l[0]);
    sh.getRange(r, 3, 1, 4).merge().setFormula('=HYPERLINK("' + l[1] + '","Open ↗")').setFontColor(BRAND.primary);
    r++;
  });
  r++;

  // Tabs.
  sectionTitle_(sh, r, 'What each tab is for'); r++;
  [
    ['Comps', 'The model. Yellow = weekly inputs, blue = quarterly inputs, black = formulas.'],
    ['Exceptions', 'Open items: ERROR (value not written), REVIEW (judgement), INFO. Sign-offs persist between runs.'],
    ['Review Queue', 'Proposed financial changes with old value, new value and the CIQ source cell.'],
    ['Validation Report', 'Summary of updates, exceptions and items for manual review.'],
    ['Config', 'Schedule, thresholds, access list and ticker map. No code needed.'],
    ['Price History · Audit Trail · Run Log', 'Every price saved, every value written (with source), every run.']
  ].forEach(function (t) {
    sh.getRange(r, 2).setValue(t[0]).setFontWeight('bold');
    sh.getRange(r, 3, 1, 4).merge().setValue(t[1]).setWrap(true).setFontColor('#3C4043');
    sh.setRowHeight(r, 26);
    r++;
  });
  r++;
  sectionTitle_(sh, r, 'Access'); r++;
  sh.getRange(r, 2, 1, 5).merge().setValue('Anyone you share this Sheet with (as an editor) can use the menu, sidebar and web dashboard with their own Google account. Every run and change is logged under their email. Scheduled refreshes run under the account that set them up. To limit the web dashboard to named people, list their emails on the Config tab (allowedUsers).')
    .setWrap(true).setFontColor('#3C4043');
  sh.setRowHeight(r, 48);
  r += 2;
  sh.getRange(r, 2, 1, 5).merge().setValue('Built for Areta · Google Sheets + Apps Script · source and docs: github.com/alycz/Areta-Comparables-Automation').setFontSize(9).setFontColor('#9AA0A6');
  sh.setFrozenRows(0);
  if (sh.getMaxRows() > r + 2) sh.deleteRows(r + 3, sh.getMaxRows() - r - 2);
  return sh;
}

function sectionTitle_(sh, r, text) {
  sh.getRange(r, 2, 1, 5).merge().setValue(text).setFontWeight('bold').setFontSize(12).setFontColor(BRAND.midnight)
    .setBorder(null, null, true, null, null, null, BRAND.primary, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(r, 30);
}
