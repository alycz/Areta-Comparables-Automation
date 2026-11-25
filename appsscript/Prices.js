/**
 * Prices.js — automated share-price updates.
 *
 * Source:      GOOGLEFINANCE (live quote), via a hidden _PriceFeed tab.
 * Cross-check: Yahoo Finance public chart endpoint (no key). Optional.
 * History:     every snapshot is saved to the Price History tab.
 * Weekly roll: Price = latest close; Prior Price = latest saved close at least
 *              `priorPriceMinDays` before it. Re-running the same week is safe:
 *              the prior price is looked up, never shifted blindly.
 * Failures:    a missing / invalid quote never overwrites the existing price.
 */

var HISTORY_HEADERS = ['Price date', 'Ticker', 'Google ticker', 'Price (US$)', 'Currency', 'Source', 'Cross-check price', 'Run ID', 'Saved at'];

/**
 * Refresh prices.
 * opts.mode: 'weekly' (default: update Comps) or 'snapshot' (history only)
 * opts.trigger: 'manual' | 'schedule'
 */
function refreshPrices(opts) {
  opts = opts || {};
  var mode = opts.mode || 'weekly';
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another update is running. Try again in a minute.');
  var run = startRun('prices', opts.trigger);
  try {
    var settings = getSettings();
    var model = loadModel();
    var tmap = getTickerMap();
    var tz = ss_().getSpreadsheetTimeZone();
    var todayIso = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var cos = model.layout.companies.filter(function (c) { var t = tmap[c.ticker.toUpperCase()]; return !t || t.active; });

    var quotes = fetchGoogleQuotes(cos.map(function (c) { var t = tmap[c.ticker.toUpperCase()]; return t ? t.google : toGoogleTicker(c.ticker); }));
    var useCross = String(settings.crossCheck).toUpperCase() === 'YES';
    var history = readHistory();
    var issues = [], audit = [], histRows = [];
    var priceDates = [];
    var updated = 0, failed = 0;

    cos.forEach(function (co, i) {
      var q = quotes[i];
      var t = tmap[co.ticker.toUpperCase()] || { google: toGoogleTicker(co.ticker), crossCheck: tickerSymbol(co.ticker) };
      var cur = rowValues(model, co.row);
      var secondary = null, yahooHist = null;
      if (useCross || !historyFor(history, co.ticker).length) {
        var y = fetchYahoo(t.crossCheck);
        if (y) { secondary = y.price; yahooHist = y.history; }
      }
      var priceDate = q.tradeDate || todayIso;
      // Prior price: history first, then Yahoo daily closes (first run), else keep current prior.
      var prior = null, priorSrc = '';
      var h = historyFor(history, co.ticker);
      var pick = pickPriorPrice(h, priceDate, Number(settings.priorPriceMinDays));
      if (pick) { prior = pick; priorSrc = 'Price History'; }
      else if (yahooHist) { pick = pickPriorPrice(yahooHist, priceDate, Number(settings.priorPriceMinDays)); if (pick) { prior = pick; priorSrc = 'Yahoo Finance daily close (seed)'; } }

      var pIssues = validatePrice({ ticker: co.ticker, price: q.price, currency: q.currency, tradeDate: q.tradeDate, todayIso: todayIso, prior: prior ? prior.price : num(cur.priorPrice), secondary: secondary, error: q.error }, settings);
      if (useCross && secondary === null && q.price) pIssues.push({ severity: SEVERITY.INFO, code: 'CROSS_CHECK_UNAVAILABLE', field: 'price', message: 'Second price source unavailable; Google Finance price used without cross-check.' });
      if (!prior && mode === 'weekly') pIssues.push({ severity: SEVERITY.REVIEW, code: 'NO_PRIOR_PRICE', field: 'priorPrice', message: 'No saved close at least ' + settings.priorPriceMinDays + ' days before ' + priceDate + '. Prior price left unchanged; weekly move will be right after next week\'s run.' });
      var blocking = pIssues.some(function (x) { return x.severity === SEVERITY.ERROR; });

      pIssues.forEach(function (x) {
        issues.push({ severity: x.severity, category: 'Price', company: co.name, ticker: co.ticker, field: x.field, code: x.code, message: x.message, source: 'GOOGLEFINANCE("' + t.google + '")' });
      });
      if (blocking) { failed++; return; }

      histRows.push([priceDate, co.ticker, t.google, q.price, q.currency || '', 'GOOGLEFINANCE', secondary === null ? '' : secondary, run.id, new Date()]);
      priceDates.push(priceDate);
      if (mode !== 'weekly') return;

      if (cur.price !== q.price) {
        setCell(model.sheet, co.row, model.map.cols.price, q.price);
        audit.push({ action: 'Price', company: co.name, ticker: co.ticker, field: 'Price', oldV: cur.price, newV: q.price, source: { file: 'GOOGLEFINANCE', sheet: '', cell: t.google, label: 'Last price, trade date ' + priceDate, raw: q.price } });
      }
      if (prior && cur.priorPrice !== prior.price) {
        setCell(model.sheet, co.row, model.map.cols.priorPrice, prior.price);
        audit.push({ action: 'Price', company: co.name, ticker: co.ticker, field: 'Prior Price', oldV: cur.priorPrice, newV: prior.price, source: { file: priorSrc, sheet: '', cell: t.google, label: 'Close on ' + prior.date, raw: prior.price } });
      }
      updated++;
    });

    // Header dates.
    if (mode === 'weekly' && priceDates.length) {
      var pd = mostCommon(priceDates);
      if (model.header.priceDate) setCell(model.sheet, model.header.priceDate.row, model.header.priceDate.col, isoToDate(pd));
      var priorDates = [];
      model.layout.companies.forEach(function (co) { var p = pickPriorPrice(historyFor(history.concat(histToObjs(histRows)), co.ticker), pd, Number(settings.priorPriceMinDays)); if (p) priorDates.push(p.date); });
      if (model.header.priorDate && priorDates.length) {
        setCell(model.sheet, model.header.priorDate.row, model.header.priorDate.col, isoToDate(mostCommon(priorDates)));
        model.sheet.getRange(model.header.priorDate.row + 1, model.header.priorDate.col + 1).setNumberFormat('dd-mmm-yyyy');
      }
      var mixed = priceDates.filter(function (d) { return d !== pd; });
      if (mixed.length) issues.push({ severity: SEVERITY.REVIEW, category: 'Price', company: '', ticker: '', field: 'price', code: 'MIXED_PRICE_DATES', message: mixed.length + ' quote(s) have a trade date other than ' + pd + '. Check for halted or illiquid names.', source: '' });
    }

    appendHistory(histRows);
    writeAudit(run, audit);
    syncExceptions(run, ['Price'], issues);
    SpreadsheetApp.flush();
    var summary = (mode === 'weekly' ? 'Updated ' + updated + ' of ' + cos.length + ' prices' : 'Saved ' + histRows.length + ' price snapshots') +
      (failed ? ', ' + failed + ' failed (previous price kept)' : '') + '. ' + countBy(issues) + '.';
    finishRun(run, failed ? 'Partial' : 'OK', summary);
    if (mode === 'weekly') runModelChecks({ trigger: opts.trigger, quiet: true });
    return { runId: run.id, summary: summary, updated: updated, failed: failed, issues: issues.length };
  } catch (e) {
    finishRun(run, 'Failed', String(e && e.message || e));
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/** Scheduled entry points (installed by installTriggers). */
function scheduledWeeklyPrices() { refreshPrices({ mode: 'weekly', trigger: 'schedule' }); }
function scheduledDailySnapshot() {
  var d = new Date().getDay();
  if (d === 0 || d === 6) return;
  refreshPrices({ mode: 'snapshot', trigger: 'schedule' });
}

/**
 * Read GOOGLEFINANCE quotes through the hidden feed tab.
 * Returns [{price, currency, tradeDate, error}] in the order of `googleTickers`.
 */
function fetchGoogleQuotes(googleTickers) {
  var ss = ss_();
  var feed = ss.getSheetByName(SHEETS.FEED) || ss.insertSheet(SHEETS.FEED);
  feed.hideSheet();
  feed.clear();
  var rows = googleTickers.map(function (t, i) {
    var r = i + 2;
    return [t, '=IFERROR(GOOGLEFINANCE($A' + r + ',"price"),"ERR")', '=IFERROR(GOOGLEFINANCE($A' + r + ',"currency"),"")', '=IFERROR(TO_TEXT(TEXT(GOOGLEFINANCE($A' + r + ',"tradetime"),"yyyy-mm-dd")),"")'];
  });
  feed.getRange(1, 1, 1, 4).setValues([['Google ticker', 'Price', 'Currency', 'Trade date']]);
  if (!rows.length) return [];
  feed.getRange(2, 1, rows.length, 4).setValues(rows);
  var vals;
  for (var attempt = 0; attempt < 8; attempt++) {
    SpreadsheetApp.flush();
    Utilities.sleep(attempt === 0 ? 1500 : 2500);
    vals = feed.getRange(2, 1, rows.length, 4).getValues();
    var loading = vals.some(function (v) { return v[1] === '' || /loading/i.test(String(v[1])); });
    if (!loading) break;
  }
  return vals.map(function (v) {
    var p = typeof v[1] === 'number' ? v[1] : null;
    return { price: p, currency: String(v[2] || ''), tradeDate: /^\d{4}-\d{2}-\d{2}$/.test(String(v[3])) ? String(v[3]) : null, error: p === null ? 'GOOGLEFINANCE returned "' + v[1] + '" for ' + v[0] : null };
  });
}

/** Yahoo Finance chart endpoint: latest price + ~1 month of daily closes. null on any failure. */
function fetchYahoo(symbol) {
  try {
    var url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol) + '?range=1mo&interval=1d';
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (res.getResponseCode() !== 200) return null;
    var j = JSON.parse(res.getContentText());
    var r = j.chart && j.chart.result && j.chart.result[0];
    if (!r) return null;
    var closes = (r.indicators.quote[0].close || []);
    var hist = [];
    (r.timestamp || []).forEach(function (ts, i) {
      if (typeof closes[i] === 'number') hist.push({ date: Utilities.formatDate(new Date(ts * 1000), 'America/New_York', 'yyyy-MM-dd'), price: closes[i] });
    });
    var price = r.meta && typeof r.meta.regularMarketPrice === 'number' ? r.meta.regularMarketPrice : (hist.length ? hist[hist.length - 1].price : null);
    if (r.meta && r.meta.currency && r.meta.currency !== 'USD') return null;
    return { price: price, history: hist };
  } catch (e) {
    return null;
  }
}

function readHistory() {
  var sh = ss_().getSheetByName(SHEETS.HISTORY);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, HISTORY_HEADERS.length).getValues().map(function (r) {
    return { date: r[0] instanceof Date ? Utilities.formatDate(r[0], ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(r[0]), ticker: String(r[1]), price: r[3] };
  });
}

function histToObjs(rows) { return rows.map(function (r) { return { date: r[0], ticker: r[1], price: r[3] }; }); }

function historyFor(history, ticker) {
  var t = String(ticker).toUpperCase();
  return history.filter(function (h) { return String(h.ticker).toUpperCase() === t; });
}

/** Append snapshots; one row per (date, ticker): a re-run on the same day replaces it. */
function appendHistory(rows) {
  if (!rows.length) return;
  var sh = sheetWithHeaders(SHEETS.HISTORY, HISTORY_HEADERS, { widths: [100, 140, 160, 100, 80, 120, 130, 170, 150] });
  var tz = ss_().getSpreadsheetTimeZone();
  var last = sh.getLastRow();
  var existing = last > 1 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
  var idx = {};
  existing.forEach(function (r, i) {
    var d = r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]);
    idx[d + '|' + String(r[1]).toUpperCase()] = i + 2;
  });
  var appends = [];
  rows.forEach(function (r) {
    var out = [isoToDate(r[0])].concat(r.slice(1));
    var at = idx[r[0] + '|' + String(r[1]).toUpperCase()];
    if (at) sh.getRange(at, 1, 1, out.length).setValues([out]); else appends.push(out);
  });
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, HISTORY_HEADERS.length).setValues(appends);
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).setNumberFormat('dd-mmm-yyyy');
  sh.getRange(2, 9, sh.getLastRow() - 1, 1).setNumberFormat('dd-mmm-yyyy hh:mm');
}

function isoToDate(iso) {
  if (!iso) return '';
  var p = String(iso).split('-');
  return new Date(+p[0], +p[1] - 1, +p[2]);
}

function mostCommon(arr) { return mode(arr); }

function countBy(issues) {
  var c = { ERROR: 0, REVIEW: 0, INFO: 0 };
  issues.forEach(function (i) { c[i.severity] = (c[i.severity] || 0) + 1; });
  return c.ERROR + ' errors, ' + c.REVIEW + ' for review, ' + c.INFO + ' info';
}
