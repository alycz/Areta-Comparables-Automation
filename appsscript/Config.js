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
  revenueBasisTolerance: [0.02, 'Flag when CIQ revenue differs from as-reported revenue by more than this'],
  changeShares: [0.10, 'Flag a change in shares vs the current model larger than this'],
  changeFinancials: [0.30, 'Flag a change in LTM revenue / EBITDA vs the current model larger than this'],
  reportingLagDays: [45, 'Days after quarter end by which results are expected (for stale-period flags)'],
  multipleOutlierEvEbitda: [100, 'Flag EV / EBITDA above this'],
  multipleOutlierEvRevenue: [50, 'Flag EV / Revenue above this'],
  autoApplyClean: ['NO', 'Apply imported financials without the review step when there are no ERROR items (YES/NO)'],
  allowedUsers: ['', 'Comma-separated emails allowed to use the web app. Blank = anyone the Sheet is shared with.']
};

var CONFIG_LAYOUT = {
  settingsHeaderRow: 3,   // 'Setting | Value | Description'
  tickerTitleRow: 0       // computed: below the settings
};

/** Read settings from the Config tab, falling back to defaults. */
function getSettings() {
  var s = {};
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) { s[k] = DEFAULT_SETTINGS[k]; });
  Object.keys(SETTING_DEFS).forEach(function (k) { s[k] = SETTING_DEFS[k][0]; });
  var sh = ss_().getSheetByName(SHEETS.CONFIG);
  if (!sh) return s;
  var vals = sh.getDataRange().getValues();
  for (var r = 0; r < vals.length; r++) {
    var k = String(vals[r][0] || '').trim();
    if (k && SETTING_DEFS.hasOwnProperty(k) && vals[r][1] !== '') s[k] = vals[r][1];
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
  var start = -1;
  for (var r = 0; r < vals.length; r++) if (String(vals[r][0]).trim() === 'Model ticker') { start = r + 1; break; }
  if (start < 0) return out;
  for (var i = start; i < vals.length; i++) {
    var t = String(vals[i][0] || '').trim();
    if (!t) continue;
    out[t.toUpperCase()] = {
      google: String(vals[i][1] || '').trim() || toGoogleTicker(t),
      crossCheck: String(vals[i][2] || '').trim() || tickerSymbol(t),
      active: String(vals[i][3] || 'YES').trim().toUpperCase() !== 'NO'
    };
  }
  return out;
}

/** Create / refresh the Config tab. Keeps any values bankers have changed. */
function ensureConfigSheet(companies) {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEETS.CONFIG) || ss.insertSheet(SHEETS.CONFIG);
  var current = getSettings();
  var existingTickers = getTickerMap();
  sh.clear();
  sh.getRange('A1').setValue('Areta Comps — configuration').setFontSize(14).setFontWeight('bold').setFontColor(BRAND.midnight);
  sh.getRange('A2').setValue('Change values in column B. No code changes are needed. Ticker overrides apply from the next price refresh.').setFontColor('#555555');
  var rows = [['Setting', 'Value', 'Description']];
  Object.keys(SETTING_DEFS).forEach(function (k) { rows.push([k, current[k], SETTING_DEFS[k][1]]); });
  sh.getRange(CONFIG_LAYOUT.settingsHeaderRow + 1, 1, rows.length, 3).setValues(rows);
  styleHeader(sh.getRange(CONFIG_LAYOUT.settingsHeaderRow + 1, 1, 1, 3));
  sh.getRange(CONFIG_LAYOUT.settingsHeaderRow + 2, 2, rows.length - 1, 1).setBackground('#FFF2CC').setFontColor('#0000FF');

  var tStart = CONFIG_LAYOUT.settingsHeaderRow + rows.length + 3;
  sh.getRange(tStart, 1).setValue('Ticker map').setFontWeight('bold').setFontColor(BRAND.midnight);
  var trows = [['Model ticker', 'Google Finance ticker', 'Cross-check symbol', 'Active (YES/NO)', 'Company']];
  (companies || []).forEach(function (c) {
    var ex = existingTickers[c.ticker.toUpperCase()] || {};
    trows.push([c.ticker, ex.google || toGoogleTicker(c.ticker), ex.crossCheck || tickerSymbol(c.ticker), ex.active === false ? 'NO' : 'YES', c.name]);
  });
  sh.getRange(tStart + 1, 1, trows.length, 5).setValues(trows);
  styleHeader(sh.getRange(tStart + 1, 1, 1, 5));
  if (trows.length > 1) sh.getRange(tStart + 2, 2, trows.length - 1, 3).setBackground('#FFF2CC').setFontColor('#0000FF');
  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 160); sh.setColumnWidth(4, 120); sh.setColumnWidth(5, 220);
  sh.getRange('C:C').setWrap(true);
  sh.setColumnWidth(3, 520);
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
  range.setBackground(BRAND.midnight).setFontColor('#FFFFFF').setFontWeight('bold');
}
