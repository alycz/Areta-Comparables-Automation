/**
 * Config.js — sheet names, settings, and the editable Config tab.
 *
 * Bankers change behaviour from the Config tab (tickers, thresholds, schedule).
 * Nothing here needs editing for normal use.
 */

var APP_NAME = 'Areta Comps';

var SHEETS = {
  COMPS: 'Comps',
  README: 'Read Me',
  CONFIG: 'Config',
  REVIEW: 'Review Queue',
  EXCEPTIONS: 'Exceptions',
  HISTORY: 'Price History',
  AUDIT: 'Audit Trail',
  RUNS: 'Run Log',
  FEED: '_PriceFeed'
};

/** Editable settings shown on the Config tab: key -> [default, description]. */
var SETTING_DEFS = {
  scheduleDay: ['MONDAY', 'Weekly price refresh day (MONDAY … SUNDAY)'],
  scheduleHour: [7, 'Hour of the weekly refresh, in the spreadsheet time zone (0-23)'],
  dailySnapshot: ['YES', 'Also save a price snapshot every weekday (YES/NO). Builds history; Comps tab still rolls weekly.'],
  crossCheck: ['YES', 'Cross-check Google Finance against a second public source (YES/NO)'],
  crossCheckTolerance: [0.02, 'Flag when the two price sources differ by more than this'],
  weeklyMoveAlert: [0.20, 'Flag weekly price moves larger than this'],
  staleDays: [5, 'Flag a quote whose last trade is older than this many days'],
  priorPriceMinDays: [6, 'Prior price = latest saved close at least this many days before the price date'],
  reconTolerance: [0.005, 'Market cap / EV must rebuild to CIQ within this tolerance'],
  impliedSharesTolerance: [0.02, 'Flag when CIQ implied shares differ from listed shares by more than this'],
  growthOutlier: [1.0, 'Flag revenue growth beyond ± this'],
  ebitdaMarginOutlier: [1.0, 'Flag LTM EBITDA whose size exceeds this multiple of LTM revenue (e.g. crypto fair-value swings)'],
  timeZone: ['Europe/London', 'Time zone for the schedule and dates (Areta bankers: Europe/London)'],
  revenueBasisTolerance: [0.02, 'Flag when CIQ revenue differs from as-reported revenue by more than this'],
  changeShares: [0.10, 'Flag a change in shares vs the current model larger than this'],
  changeFinancials: [0.30, 'Flag a change in LTM revenue / EBITDA vs the current model larger than this'],
  reportingLagDays: [45, 'Days after quarter end by which results are expected (for stale-period flags)'],
  multipleOutlierEvEbitda: [100, 'Flag EV / EBITDA above this'],
  multipleOutlierEvRevenue: [50, 'Flag EV / Revenue above this'],
  autoApplyClean: ['NO', 'Apply imported financials without the review step when there are no ERROR items (YES/NO)'],
  allowedUsers: ['', 'Comma-separated emails allowed to use the web app. Blank = anyone the Sheet is shared with.']
};

/** Read settings from the Config tab, falling back to defaults. */
function getSettings() {
  var s = {};
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) { s[k] = DEFAULT_SETTINGS[k]; });
  Object.keys(SETTING_DEFS).forEach(function (k) { s[k] = SETTING_DEFS[k][0]; });
  var sh = ss_().getSheetByName(SHEETS.CONFIG);
  if (!sh) return s;
  var vals = sh.getDataRange().getValues();
  // Keys may sit in column A (older layout) or B (branded layout); the value is the next cell.
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < Math.min(2, vals[r].length - 1); c++) {
      var k = String(vals[r][c] || '').trim();
      if (k && SETTING_DEFS.hasOwnProperty(k)) { if (vals[r][c + 1] !== '') s[k] = vals[r][c + 1]; break; }
    }
  }
  return s;
}

/**
 * Ticker map from the Config tab:
 * Model ticker -> { google, crossCheck, active }
 */
function getTickerMap() {
  var out = {};
  var sh = ss_().getSheetByName(SHEETS.CONFIG);
  if (!sh) return out;
  var vals = sh.getDataRange().getValues();
  var start = -1, c0 = 0;
  for (var r = 0; r < vals.length && start < 0; r++) {
    for (var c = 0; c < 2; c++) if (String(vals[r][c]).trim() === 'Model ticker') { start = r + 1; c0 = c; break; }
  }
  if (start < 0) return out;
  for (var i = start; i < vals.length; i++) {
    var t = String(vals[i][c0] || '').trim();
    if (!t) continue;
    out[t.toUpperCase()] = {
      google: String(vals[i][c0 + 1] || '').trim() || toGoogleTicker(t),
      crossCheck: String(vals[i][c0 + 2] || '').trim() || tickerSymbol(t),
      active: String(vals[i][c0 + 3] || 'YES').trim().toUpperCase() !== 'NO'
    };
  }
  return out;
}

/** Create / refresh the Config tab (branded layout). Keeps any values bankers have changed. */
function ensureConfigSheet(companies) {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEETS.CONFIG) || ss.insertSheet(SHEETS.CONFIG);
  var current = getSettings();
  var existingTickers = getTickerMap();
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  var top = brandPage_(sh, 'Configuration', 'Everything the engine does is driven from this tab. Change the yellow cells; no code is needed. Schedule changes apply after Areta Comps › Tools › Set up; ticker changes apply on the next price refresh.', [28, 230, 190, 170, 150, 330, 28]);
  var r = top;

  sectionTitle_(sh, r, 'Settings', 5); r++;
  var rows = [['Setting', 'Value', 'What it does', '', '']];
  Object.keys(SETTING_DEFS).forEach(function (k) { rows.push([k, current[k], SETTING_DEFS[k][1], '', '']); });
  sh.getRange(r, 2, rows.length, 5).setValues(rows);
  sh.getRange(r, 4, rows.length, 3).mergeAcross();
  tableHeader_(sh.getRange(r, 2, 1, 5));
  sh.getRange(r + 1, 3, rows.length - 1, 1).setBackground('#FFF2CC').setFontColor('#0000FF');
  sh.getRange(r + 1, 4, rows.length - 1, 3).setFontColor('#5F6368').setWrap(true);
  sh.getRange(r + 1, 2, rows.length - 1, 5).setBorder(null, null, true, null, null, true, '#E4E8F2', SpreadsheetApp.BorderStyle.SOLID);
  for (var i = 0; i < rows.length; i++) sh.setRowHeight(r + i, 26);
  r += rows.length + 1;

  sectionTitle_(sh, r, 'Ticker map', 5); r++;
  var trows = [['Model ticker', 'Google Finance ticker', 'Cross-check symbol', 'Active (YES/NO)', 'Company']];
  (companies || []).forEach(function (c) {
    var ex = existingTickers[c.ticker.toUpperCase()] || {};
    trows.push([c.ticker, ex.google || toGoogleTicker(c.ticker), ex.crossCheck || tickerSymbol(c.ticker), ex.active === false ? 'NO' : 'YES', c.name]);
  });
  sh.getRange(r, 2, trows.length, 5).setValues(trows);
  tableHeader_(sh.getRange(r, 2, 1, 5));
  if (trows.length > 1) {
    sh.getRange(r + 1, 3, trows.length - 1, 3).setBackground('#FFF2CC').setFontColor('#0000FF');
    sh.getRange(r + 1, 2, trows.length - 1, 5).setBorder(null, null, true, null, null, true, '#E4E8F2', SpreadsheetApp.BorderStyle.SOLID);
  }
  for (var j = 0; j < trows.length; j++) sh.setRowHeight(r + j, 24);
  r += trows.length + 1;
  footer_(sh, r, 'Yellow cells are inputs. Values are read by name, so rows can be reordered but setting names must not be renamed.');
  finishPage_(sh, r);
  return sh;
}

/** Areta brand tokens (Brand Guidelines 2026, p.9). */
var BRAND = {
  primary: '#0E4FF5',
  midnight: '#002B99',
  midday: '#789DFF',
  frost: '#DDE7FF',
  ink: '#0B0F1A'
};

function styleHeader(range) {
  range.setBackground(BRAND.midnight).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle');
}
