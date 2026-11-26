/**
 * Model.js — reads and writes the Comps tab without changing its structure.
 *
 * Columns and companies are located by header text and ticker, so the engine
 * works on the supplied template and on an existing Areta model whose column
 * order differs. Only input cells (yellow / blue) and formula cells are touched;
 * formats, colours and layout are left alone.
 */

function compsSheet() {
  var sh = ss_().getSheetByName(SHEETS.COMPS);
  if (!sh) throw new Error('No "' + SHEETS.COMPS + '" tab found. Run Comparables Automation › Set up first.');
  return sh;
}

/** Full model context: values, header map, layout. */
function loadModel() {
  var sh = compsSheet();
  var values = sh.getDataRange().getValues();
  var map = mapModelHeaders(values);
  if (!map) throw new Error('Could not find the header row (a row containing "Company" and "Ticker") on the Comps tab.');
  var required = ['company', 'ticker', 'price', 'priorPrice', 'sharesOut', 'netDebt', 'nci', 'preferred', 'ltmRevenue', 'ltmEbitda', 'marketCap', 'ev', 'evRevenue', 'evEbitda'];
  var missing = required.filter(function (k) { return map.cols[k] === undefined; });
  if (missing.length) throw new Error('Comps tab is missing columns: ' + missing.map(function (k) { return MODEL_HEADERS[k]; }).join(', '));
  var layout = readModelLayout(values, map);
  return { sheet: sh, values: values, map: map, layout: layout, header: locateHeaderCells(values) };
}

/** Find the 'Price date' / 'Prior price' / 'Financials source' label cells (value goes one cell right). */
function locateHeaderCells(values) {
  var out = {};
  var want = { priceDate: 'price date', priorDate: 'prior price', financialsSource: 'financials source' };
  for (var r = 0; r < Math.min(values.length, 8); r++) {
    for (var c = 0; c < values[r].length; c++) {
      var v = normLabel(values[r][c]);
      Object.keys(want).forEach(function (k) { if (v === want[k] && !out[k]) out[k] = { row: r, col: c + 1 }; });
    }
  }
  return out;
}

/** Current values for one company row as {field: value}. */
function rowValues(model, row) {
  var o = {};
  Object.keys(model.map.cols).forEach(function (k) { o[k] = model.values[row][model.map.cols[k]]; });
  return o;
}

/** Write a single cell (0-based row/col) keeping the cell's format. */
function setCell(sh, row, col, value) {
  sh.getRange(row + 1, col + 1).setValue(value);
}

function colA1(model, field, row) {
  return colLetter(model.map.cols[field]) + (row + 1);
}

/**
 * The formulas behind the black cells. "NM" when the denominator is zero or
 * negative (Read Me convention); blank while inputs are missing.
 */
function rowFormulas(model, r) {
  var c = function (f) { return colA1(model, f, r); };
  var f = {};
  f.change = '=IF(AND(ISNUMBER(' + c('price') + '),ISNUMBER(' + c('priorPrice') + '),N(' + c('priorPrice') + ')<>0),' + c('price') + '/' + c('priorPrice') + '-1,"")';
  f.marketCap = '=IF(AND(ISNUMBER(' + c('price') + '),ISNUMBER(' + c('sharesOut') + ')),' + c('price') + '*' + c('sharesOut') + ',"")';
  f.ev = '=IF(AND(ISNUMBER(' + c('marketCap') + '),ISNUMBER(' + c('netDebt') + ')),' + c('marketCap') + '+' + c('netDebt') + '+N(' + c('nci') + ')+N(' + c('preferred') + '),"")';
  f.evRevenue = '=IF(AND(ISNUMBER(' + c('ev') + '),ISNUMBER(' + c('ltmRevenue') + ')),IF(' + c('ltmRevenue') + '>0,' + c('ev') + '/' + c('ltmRevenue') + ',"NM"),"")';
  f.evEbitda = '=IF(AND(ISNUMBER(' + c('ev') + '),ISNUMBER(' + c('ltmEbitda') + ')),IF(' + c('ltmEbitda') + '>0,' + c('ev') + '/' + c('ltmEbitda') + ',"NM"),"")';
  return f;
}

/** Columns that get a median row (MEDIAN ignores "NM" text and blanks). */
var MEDIAN_FIELDS = ['change', 'revenueGrowth', 'evRevenue', 'evEbitda'];

function medianFormula(model, field, rows) {
  if (!rows.length) return '';
  var col = colLetter(model.map.cols[field]);
  // Contiguous groups -> one range each.
  var parts = [], start = rows[0], prev = rows[0];
  for (var i = 1; i <= rows.length; i++) {
    if (i < rows.length && rows[i] === prev + 1) { prev = rows[i]; continue; }
    parts.push(start === prev ? col + (start + 1) : col + (start + 1) + ':' + col + (prev + 1));
    if (i < rows.length) { start = rows[i]; prev = rows[i]; }
  }
  return '=IFERROR(MEDIAN(' + parts.join(',') + '),"NM")';
}

/**
 * Install (or repair) every formula on the Comps tab. Idempotent.
 * Returns the number of cells written.
 */
function installFormulas(model) {
  model = model || loadModel();
  var sh = model.sheet, n = 0;
  model.layout.companies.forEach(function (co) {
    var f = rowFormulas(model, co.row);
    Object.keys(f).forEach(function (k) {
      if (model.map.cols[k] === undefined) return;
      sh.getRange(co.row + 1, model.map.cols[k] + 1).setFormula(f[k]); n++;
    });
  });
  model.layout.groups.forEach(function (g) {
    if (g.medianRow < 0) return;
    MEDIAN_FIELDS.forEach(function (k) {
      if (model.map.cols[k] === undefined) return;
      sh.getRange(g.medianRow + 1, model.map.cols[k] + 1).setFormula(medianFormula(model, k, g.rows)); n++;
    });
  });
  if (model.layout.globalMedianRow >= 0) {
    var all = model.layout.companies.map(function (c) { return c.row; });
    MEDIAN_FIELDS.forEach(function (k) {
      if (model.map.cols[k] === undefined) return;
      sh.getRange(model.layout.globalMedianRow + 1, model.map.cols[k] + 1).setFormula(medianFormula(model, k, all)); n++;
    });
  }
  return n;
}

/** Comps snapshot for the web app. */
function modelSnapshot() {
  var m = loadModel();
  var disp = m.sheet.getDataRange().getDisplayValues();
  var groups = m.layout.groups.map(function (g) {
    return {
      name: g.name,
      companies: g.rows.map(function (r) {
        var v = rowValues(m, r);
        return {
          row: r + 1, name: v.company, ticker: v.ticker,
          price: num(v.price), priorPrice: num(v.priorPrice), change: num(v.change), marketCap: num(v.marketCap), ev: num(v.ev),
          ltmRevenue: num(v.ltmRevenue), ltmEbitda: num(v.ltmEbitda), revenueGrowth: num(v.revenueGrowth),
          evRevenue: typeof v.evRevenue === 'number' ? v.evRevenue : String(v.evRevenue || ''),
          evEbitda: typeof v.evEbitda === 'number' ? v.evEbitda : String(v.evEbitda || ''),
          period: String(v.financialPeriod || ''), ebitdaBasis: String(v.ebitdaBasis || ''), notes: String(v.notes || '')
        };
      }),
      median: g.medianRow >= 0 ? {
        change: num(m.values[g.medianRow][m.map.cols.change]),
        revenueGrowth: num(m.values[g.medianRow][m.map.cols.revenueGrowth]),
        evRevenue: disp[g.medianRow][m.map.cols.evRevenue], evEbitda: disp[g.medianRow][m.map.cols.evEbitda]
      } : null
    };
  });
  var gm = m.layout.globalMedianRow;
  var h = m.header;
  function hv(k) { return h[k] ? m.values[h[k].row][h[k].col] : ''; }
  // Title as Areta wrote it on the Comps tab ("Exchanges & Trading Infrastructure | Trading Comparables").
  var rawTitle = '';
  for (var tr = 0; tr < Math.min(4, disp.length) && !rawTitle; tr++) for (var tc = 0; tc < Math.min(4, disp[tr].length); tc++) if (String(disp[tr][tc]).indexOf('|') > 0 || /comparables/i.test(disp[tr][tc])) { rawTitle = String(disp[tr][tc]); break; }
  var parts = rawTitle.split('|').map(function (x) { return x.trim(); });
  var nCos = groups.reduce(function (a, g) { return a + g.companies.length; }, 0);
  return {
    title: parts[0] || 'Trading comparables', titleSub: parts[1] || 'Trading Comparables', companyCount: nCos, groupCount: groups.length,
    groups: groups,
    allMedian: gm >= 0 ? { evRevenue: disp[gm][m.map.cols.evRevenue], evEbitda: disp[gm][m.map.cols.evEbitda], revenueGrowth: disp[gm][m.map.cols.revenueGrowth], change: disp[gm][m.map.cols.change] } : null,
    priceDate: dateOut(hv('priceDate')), priorDate: dateOut(hv('priorDate')), financialsSource: String(hv('financialsSource') || '')
  };
}

function num(v) { return typeof v === 'number' && isFinite(v) ? v : null; }
function dateOut(v) { return v instanceof Date ? Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'dd-MMM-yyyy') : String(v || ''); }
