/**
 * Checks.js — model-level validation, run after every price or financial update
 * and on demand.
 *
 *  - Formula integrity: black cells must hold the engine's formula, not a hardcode.
 *  - Recalculation: every formula result is recomputed in JS and compared.
 *  - Errors: any #REF!/#N/A/#DIV/0! etc. on the Comps tab.
 *  - Missing inputs: blank yellow / blue cells.
 *  - Sanity: EV <= 0, multiples outliers, stale price date, mixed periods in a peer group.
 */

function runModelChecks(opts) {
  opts = opts || {};
  var run = startRun('checks', opts.trigger);
  try {
    var settings = getSettings();
    var model = loadModel();
    var sh = model.sheet;
    var formulas = sh.getDataRange().getFormulas();
    var issues = [];
    var tz = ss_().getSpreadsheetTimeZone();
    function add(sev, co, field, code, msg, cell) {
      issues.push({ severity: sev, category: 'Model', company: co ? co.name : '', ticker: co ? co.ticker : '', field: field, code: code, message: msg, source: cell ? SHEETS.COMPS + '!' + cell : '' });
    }

    // Errors anywhere on the tab.
    model.values.forEach(function (row, r) {
      row.forEach(function (v, c) {
        if (typeof v === 'string' && /^#(REF!|N\/A|DIV\/0!|VALUE!|NAME\?|NUM!|ERROR!|NULL!)/.test(v)) add(SEVERITY.ERROR, null, '', 'CELL_ERROR', 'Cell shows ' + v + '.', a1(r, c));
      });
    });

    var formulaFields = ['change', 'marketCap', 'ev', 'evRevenue', 'evEbitda'];
    var inputFields = ['price', 'priorPrice', 'sharesOut', 'netDebt', 'nci', 'preferred', 'ltmRevenue', 'ltmEbitda', 'revenueGrowth', 'revenueBasis', 'ebitdaBasis', 'financialPeriod'];
    var repaired = 0;
    model.layout.companies.forEach(function (co) {
      var v = rowValues(model, co.row);
      var expect = rowFormulas(model, co.row);
      formulaFields.forEach(function (k) {
        var col = model.map.cols[k]; if (col === undefined) return;
        var f = formulas[co.row][col];
        if (!f) add(SEVERITY.ERROR, co, k, 'HARDCODED_FORMULA', MODEL_HEADERS[k] + ' is a hardcoded value, not a formula. Use "Repair formulas" to restore it.', a1(co.row, col));
        else if (f.replace(/\s/g, '') !== expect[k].replace(/\s/g, '')) add(SEVERITY.REVIEW, co, k, 'FORMULA_CHANGED', MODEL_HEADERS[k] + ' formula differs from the standard (' + f + '). Check it is intentional.', a1(co.row, col));
      });
      inputFields.forEach(function (k) {
        var col = model.map.cols[k]; if (col === undefined) return;
        if (v[k] === '' || v[k] === null) add(k === 'priorPrice' ? SEVERITY.REVIEW : SEVERITY.ERROR, co, k, 'MISSING_INPUT', MODEL_HEADERS[k] + ' is blank.', a1(co.row, col));
        else if (formulas[co.row][col] && k !== 'price') add(SEVERITY.REVIEW, co, k, 'INPUT_IS_FORMULA', MODEL_HEADERS[k] + ' is an input cell but holds a formula.', a1(co.row, col));
      });
      // Independent recalculation.
      var calc = computeRow(v);
      formulaFields.forEach(function (k) {
        var col = model.map.cols[k]; if (col === undefined) return;
        var got = v[k], want = calc[k];
        var ok = (want === null && (got === '' || got === null)) || (want === 'NM' && got === 'NM') ||
          (typeof want === 'number' && typeof got === 'number' && Math.abs(got - want) <= 1e-6 * Math.max(1, Math.abs(want)));
        if (!ok && formulas[co.row][col]) add(SEVERITY.ERROR, co, k, 'RECALC_MISMATCH', MODEL_HEADERS[k] + ' shows ' + got + ' but recomputes to ' + want + '.', a1(co.row, col));
      });
      if (typeof v.ev === 'number' && v.ev <= 0) add(SEVERITY.REVIEW, co, 'ev', 'NON_POSITIVE_EV', 'EV is ' + fmt(v.ev) + ': net cash exceeds market cap. Multiples are not meaningful.', a1(co.row, model.map.cols.ev));
      if (typeof v.evEbitda === 'number' && v.evEbitda > Number(settings.multipleOutlierEvEbitda)) add(SEVERITY.REVIEW, co, 'evEbitda', 'MULTIPLE_OUTLIER', 'EV / EBITDA of ' + v.evEbitda.toFixed(1) + 'x is above ' + settings.multipleOutlierEvEbitda + 'x. Consider excluding from the median.', a1(co.row, model.map.cols.evEbitda));
      if (typeof v.evRevenue === 'number' && v.evRevenue > Number(settings.multipleOutlierEvRevenue)) add(SEVERITY.REVIEW, co, 'evRevenue', 'MULTIPLE_OUTLIER', 'EV / Revenue of ' + v.evRevenue.toFixed(1) + 'x is above ' + settings.multipleOutlierEvRevenue + 'x.', a1(co.row, model.map.cols.evRevenue));
    });

    // Median formulas present.
    model.layout.groups.forEach(function (g) {
      if (g.medianRow < 0) { add(SEVERITY.REVIEW, { name: g.name + ' (peer group)', ticker: '' }, '', 'NO_MEDIAN_ROW', 'Peer group "' + g.name + '" has no Median row.', ''); return; }
      MEDIAN_FIELDS.forEach(function (k) {
        var col = model.map.cols[k]; if (col === undefined) return;
        if (!formulas[g.medianRow][col]) add(SEVERITY.ERROR, { name: g.name + ' (peer group)', ticker: '' }, k, 'MEDIAN_HARDCODED', g.name + ' median of ' + MODEL_HEADERS[k] + ' is not a formula.', a1(g.medianRow, col));
      });
      // Mixed periods within a group.
      var pcol = model.map.cols.financialPeriod;
      if (pcol !== undefined) {
        var ps = g.rows.map(function (r) { return String(model.values[r][pcol] || '').replace(/\s*\(.*\)$/, ''); }).filter(String);
        var uniq = ps.filter(function (p, i) { return ps.indexOf(p) === i; });
        if (uniq.length > 1) add(SEVERITY.REVIEW, { name: g.name + ' (peer group)', ticker: '' }, 'financialPeriod', 'GROUP_MIXED_PERIODS', g.name + ': companies are on different LTM periods (' + uniq.join(', ') + '). Median compares unlike periods.', '');
      }
    });

    // Price date staleness.
    var pdCell = model.header.priceDate;
    if (pdCell) {
      var pd = model.values[pdCell.row][pdCell.col];
      if (pd instanceof Date) {
        var age = Math.round((new Date() - pd) / 86400000);
        if (age > 8) add(SEVERITY.REVIEW, null, 'price', 'PRICES_STALE', 'Prices are ' + age + ' days old (price date ' + Utilities.formatDate(pd, tz, 'dd-MMM-yyyy') + '). Run a price refresh.', a1(pdCell.row, pdCell.col));
      } else add(SEVERITY.REVIEW, null, 'price', 'NO_PRICE_DATE', 'Price date is blank. Run a price refresh.', a1(pdCell.row, pdCell.col));
    }

    syncExceptions(run, ['Model'], issues);
    var summary = 'Checked ' + model.layout.companies.length + ' companies, ' + model.layout.groups.length + ' peer groups. ' + countBy(issues) + '.';
    finishRun(run, issues.some(function (i) { return i.severity === 'ERROR'; }) ? 'Exceptions' : 'OK', summary);
    return { runId: run.id, summary: summary, issues: issues.length };
  } catch (e) {
    finishRun(run, 'Failed', String(e && e.message || e));
    throw e;
  }
}

/** Menu / web-app action: restore every standard formula. */
function repairFormulas() {
  var run = startRun('repair', 'manual');
  var n = installFormulas();
  finishRun(run, 'OK', 'Re-installed ' + n + ' formulas on the Comps tab.');
  SpreadsheetApp.flush();
  return runModelChecks({ quiet: true });
}
