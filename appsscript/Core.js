/**
 * Core.js — pure logic shared by Apps Script and the Node test suite.
 *
 * Nothing in this file touches SpreadsheetApp, DriveApp or any Google service.
 * It works on plain 2-D arrays (exactly what Range.getValues() returns), which
 * means every rule here is unit-tested in Node against the real CIQ exports.
 *
 * Sections
 *   1. Small helpers (label normalisation, number parsing, A1 addresses, dates)
 *   2. CIQ workbook parser  -> parseCiqWorkbook()
 *   3. Financial validation -> validateFinancials()
 *   4. Price validation     -> validatePrice()
 *   5. Model helpers        -> header mapping, ticker matching, calculations
 */

/* ============================== 1. Helpers ============================== */

var CORE_VERSION = '1.0.0';

/** Severity levels used everywhere (Exceptions tab, report, web app). */
var SEVERITY = {
  ERROR: 'ERROR',   // value NOT written; existing value kept
  REVIEW: 'REVIEW', // value written, banker judgement required
  INFO: 'INFO'      // for the record only
};

/** Lower-case, strip footnote markers / punctuation noise, collapse spaces. */
function normLabel(s) {
  return String(s == null ? '' : s)
    .replace(/[¹²³⁴⁵⁶⁷⁸⁹⁰†‡*]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * Parse a CIQ cell into a number.
 *   '-'            -> 0     (CIQ shows nil balances as a dash)
 *   'NA','NM',''   -> null  (not available / not meaningful)
 *   numbers        -> number
 */
function ciqNumber(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return null;
  var s = String(v).trim();
  if (s === '-' || s === '–') return 0;
  if (s === '' || /^(na|nm|n\/a)$/i.test(s)) return null;
  var cleaned = s.replace(/[,$\s]/g, '');
  var neg = /^\(.*\)$/.test(cleaned);
  cleaned = cleaned.replace(/[()]/g, '');
  var pct = /%$/.test(cleaned);
  cleaned = cleaned.replace(/%$/, '');
  var n = Number(cleaned);
  if (!isFinite(n) || cleaned === '') return null;
  if (neg) n = -n;
  if (pct) n = n / 100;
  return n;
}

/** True when the raw CIQ cell is the nil dash. */
function isCiqDash(v) {
  return typeof v === 'string' && /^\s*[-–]\s*$/.test(v);
}

/** 0-based column index -> letters (0 -> A). */
function colLetter(c) {
  var s = '';
  c = c + 1;
  while (c > 0) {
    var m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

/** 0-based (row, col) -> A1. */
function a1(r, c) {
  return colLetter(c) + (r + 1);
}

var MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };

/** 'Jun-30-2026' -> 'YYYY-MM-DD' (or null). */
function parseCiqDateText(s) {
  var m = /([A-Za-z]{3})-(\d{1,2})-(\d{4})/.exec(String(s || ''));
  if (!m) return null;
  var mon = MONTHS[m[1].toLowerCase()];
  if (mon === undefined) return null;
  return isoDate(new Date(Date.UTC(+m[3], mon, +m[2])));
}

/** Date -> 'YYYY-MM-DD' using UTC parts. */
function isoDate(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return null;
  var mm = String(d.getUTCMonth() + 1);
  var dd = String(d.getUTCDate());
  return d.getUTCFullYear() + '-' + (mm.length < 2 ? '0' + mm : mm) + '-' + (dd.length < 2 ? '0' + dd : dd);
}

/** Excel serial / Date / ISO string -> 'YYYY-MM-DD'. */
function anyToIsoDate(v) {
  if (v instanceof Date) {
    // Sheets returns local-midnight dates; shift by the offset so UTC parts match.
    return isoDate(new Date(v.getTime() - v.getTimezoneOffset() * 60000));
  }
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    return isoDate(new Date(Date.UTC(1899, 11, 30) + v * 86400000));
  }
  if (typeof v === 'string') {
    if (/^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
    return parseCiqDateText(v);
  }
  return null;
}

/** 'YYYY-MM-DD' -> 'Jun-30-2026' (CIQ style, used in labels). */
function prettyDate(iso) {
  if (!iso) return '';
  var p = iso.split('-');
  var names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return names[+p[1] - 1] + '-' + p[2] + '-' + p[0];
}

function daysBetween(isoA, isoB) {
  return Math.round((Date.parse(isoB) - Date.parse(isoA)) / 86400000);
}

/**
 * Parse a CIQ period header such as
 *   'LTM²\n12 months\nJun-30-2026A'
 *   'LTM² Press Release\n12 months\nSep-30-2026A'
 *   '12 months\nDec-31-2027E'
 */
function parsePeriodHeader(h) {
  var s = String(h || '');
  var flat = s.replace(/\s+/g, ' ');
  var months = /(\d+)\s*months/i.exec(flat);
  return {
    raw: flat.trim(),
    isLtm: /\bLTM\b/i.test(flat),
    isPressRelease: /press release/i.test(flat),
    isEstimate: /\d{4}E\b/.test(flat),
    months: months ? +months[1] : null,
    endDate: parseCiqDateText(flat)
  };
}

/* ========================= 2. CIQ workbook parser ========================= */

/**
 * Find the first row in `rows` whose column-A label matches.
 * opts.exact: whole-label match (after normalisation); default is startsWith.
 * opts.from / opts.to: 0-based row bounds (to is exclusive).
 */
function findRow(rows, label, opts) {
  opts = opts || {};
  var want = normLabel(label);
  var from = opts.from || 0;
  var to = Math.min(opts.to == null ? rows.length : opts.to, rows.length);
  for (var r = from; r < to; r++) {
    var got = normLabel(rows[r] && rows[r][0]);
    if (opts.exact ? got === want : got.indexOf(want) === 0) return r;
  }
  return -1;
}

/** Scan the sheet's top rows for 'Key:' 'Value' pairs (CIQ export settings). */
function readSettings(rows) {
  var out = {};
  var lim = Math.min(rows.length, 13);
  for (var r = 0; r < lim; r++) {
    var row = rows[r] || [];
    for (var c = 0; c < row.length; c++) {
      var k = String(row[c] || '').trim();
      if (/:$/.test(k) && c + 1 < row.length) out[normLabel(k.slice(0, -1))] = String(row[c + 1] || '').trim();
      if (/^in (millions|thousands|billions|units)/i.test(k)) out.scale = k;
    }
  }
  return out;
}

/** 'Coinbase Global, Inc. (NasdaqGS:COIN) > Financials > Key Stats' */
function parseIdentity(title, fileName) {
  var m = /^(.*?)\s*\(([A-Za-z]+):([A-Za-z0-9.\-]+)\)/.exec(String(title || ''));
  if (m) return { name: m[1].trim(), exchange: m[2], symbol: m[3], ticker: m[2] + ':' + m[3] };
  // Fallback: '<Name> <EXCH> <SYM> Financials.xls'
  var f = /^(.*)\s+([A-Za-z]+)\s+([A-Z0-9.]+)\s+Financials/.exec(String(fileName || ''));
  if (f) return { name: f[1].trim(), exchange: f[2], symbol: f[3], ticker: f[2] + ':' + f[3] };
  return { name: null, exchange: null, symbol: null, ticker: null };
}

/** Build a source reference used for the audit trail. */
function srcRef(fileName, sheet, r, c, label, raw) {
  return { file: fileName, sheet: sheet, cell: a1(r, c), label: String(label || '').replace(/\s+/g, ' ').trim(), raw: raw };
}

/**
 * Parse one CIQ "Financials" export.
 * @param {Object<string, Array<Array<*>>>} sheets  tab name -> 2-D values
 * @param {string} fileName
 * @return {Object} record with identity, fields, sources and issues
 */
function parseCiqWorkbook(sheets, fileName) {
  var rec = {
    fileName: fileName,
    identity: null,
    fields: {},     // model field -> value (in $mm / decimal)
    sources: {},    // model field -> srcRef
    ciq: {},        // CIQ reference values used for reconciliation only
    meta: {},
    issues: []
  };
  function issue(sev, code, field, msg) {
    rec.issues.push({ severity: sev, code: code, field: field || '', message: msg });
  }

  var ks = sheets['Key Stats'];
  if (!ks) {
    issue(SEVERITY.ERROR, 'NOT_CIQ_EXPORT', '', 'No "Key Stats" tab found. This does not look like a Capital IQ Financials export.');
    rec.identity = parseIdentity('', fileName);
    return rec;
  }

  // ---- Identity -------------------------------------------------------
  var title = '';
  for (var r0 = 0; r0 < Math.min(ks.length, 8) && !title; r0++) {
    var v0 = ks[r0] && ks[r0][0];
    if (v0 && /\(.+:.+\)/.test(String(v0))) title = String(v0);
  }
  rec.identity = parseIdentity(title, fileName);
  if (!rec.identity.ticker) issue(SEVERITY.ERROR, 'NO_TICKER', '', 'Could not read the company ticker from the export title or file name.');

  // ---- Export settings (units, currency, conversion) --------------------
  var ksSet = readSettings(ks);
  rec.meta.scale = ksSet.scale || '';
  rec.meta.currencySetting = ksSet.currency || '';
  rec.meta.conversion = ksSet.conversion || '';
  rec.meta.dilution = ksSet.dilution || '';
  if (!/millions/i.test(rec.meta.scale)) {
    issue(SEVERITY.ERROR, 'UNITS', '', 'Export is not "In Millions" (found "' + rec.meta.scale + '"). Re-export from CIQ in millions.');
  }

  // ---- Key Financials block (LTM revenue, EBITDA, growth) ---------------
  var hdr = findRow(ks, 'For the Fiscal Period Ending');
  var ltmCol = -1, period = null;
  if (hdr >= 0) {
    for (var c = 1; c < ks[hdr].length; c++) {
      var p = parsePeriodHeader(ks[hdr][c]);
      if (p.isLtm) { ltmCol = c; period = p; break; }
    }
  }
  if (ltmCol < 0) {
    issue(SEVERITY.ERROR, 'NO_LTM', 'ltmRevenue', 'No LTM column in Key Stats > Key Financials.');
  } else {
    rec.meta.period = period;
    var curRow = findRow(ks, 'Currency', { from: hdr + 1, to: hdr + 3, exact: true });
    rec.meta.ksCurrency = curRow >= 0 ? String(ks[curRow][ltmCol]).trim() : '';

    var blockEnd = findCapStart(ks, hdr);
    if (blockEnd < 0) blockEnd = hdr + 30;

    var revRow = findRow(ks, 'Total Revenue', { from: hdr, to: blockEnd, exact: true });
    if (revRow >= 0) {
      setField('ltmRevenue', ciqNumber(ks[revRow][ltmCol]), 'Key Stats', revRow, ltmCol, ks[revRow][0], ks[revRow][ltmCol]);
      var gRow = findRow(ks, 'Growth Over Prior Year', { from: revRow + 1, to: revRow + 2 });
      if (gRow >= 0) setField('revenueGrowth', ciqNumber(ks[gRow][ltmCol]), 'Key Stats', gRow, ltmCol, 'Total Revenue growth over prior year', ks[gRow][ltmCol]);
    }
    var ebRow = findRow(ks, 'EBITDA', { from: hdr, to: blockEnd, exact: true });
    if (ebRow >= 0) setField('ltmEbitda', ciqNumber(ks[ebRow][ltmCol]), 'Key Stats', ebRow, ltmCol, ks[ebRow][0], ks[ebRow][ltmCol]);
  }

  // ---- Latest Capitalization (EV bridge) --------------------------------
  var capStart = findCapStart(ks, 0);
  var capEnd = capStart >= 0 ? findRow(ks, 'Valuation Multiples', { from: capStart }) : -1;
  if (capEnd < 0) capEnd = capStart + 40;
  function capRow(label, exact) { return capStart < 0 ? -1 : findRow(ks, label, { from: capStart, to: capEnd, exact: exact }); }
  var bridgeEnd = capRow('= Total Enterprise Value');
  function bridgeRow(label) { return capStart < 0 || bridgeEnd < 0 ? -1 : findRow(ks, label, { from: capStart, to: bridgeEnd }); }

  var rPx = capRow('Share Price', true);
  var rSh = capRow('Shares Out.', true);
  var rMc = capRow('Market Capitalization', true);
  if (rPx >= 0) rec.ciq.price = ciqNumber(ks[rPx][1]);
  if (rMc >= 0) rec.ciq.marketCap = ciqNumber(ks[rMc][1]);
  if (rSh >= 0) setField('sharesOut', ciqNumber(ks[rSh][1]), 'Key Stats', rSh, 1, ks[rSh][0], ks[rSh][1]);
  var rImp = capRow('Implied Shares Outstanding');
  if (rImp >= 0) rec.ciq.impliedShares = ciqNumber(ks[rImp][1]);
  var rCapCur = capRow('Currency', true);
  rec.meta.capCurrency = rCapCur >= 0 ? String(ks[rCapCur][1]).trim() : '';

  var cash = null, debt = null, ltms = 0, nci = null, pref = null;
  if (bridgeEnd >= 0) {
    rec.meta.evSource = 'Key Stats > Latest Capitalization';
    var rCash = bridgeRow('- Cash & Short Term Investments');
    var rDebt = bridgeRow('+ Total Debt');
    var rPref = bridgeRow('+ Pref. Equity');
    var rNci = bridgeRow('+ Total Minority Interest');
    var rLtms = bridgeRow('- Long Term Marketable Securities');
    rec.ciq.tev = ciqNumber(ks[bridgeEnd][1]);
    cash = rCash >= 0 ? ciqNumber(ks[rCash][1]) : null;
    debt = rDebt >= 0 ? ciqNumber(ks[rDebt][1]) : null;
    ltms = rLtms >= 0 ? (ciqNumber(ks[rLtms][1]) || 0) : 0;
    if (cash !== null && debt !== null) {
      rec.fields.netDebt = debt - cash - ltms;
      rec.sources.netDebt = {
        file: fileName, sheet: 'Key Stats',
        cell: a1(rDebt, 1) + ' − ' + a1(rCash, 1) + (rLtms >= 0 && ltms ? ' − ' + a1(rLtms, 1) : ''),
        label: 'Total Debt − Cash & ST Investments' + (ltms ? ' − LT Marketable Securities' : ''),
        raw: debt + ' − ' + cash + (ltms ? ' − ' + ltms : '')
      };
    }
    if (rPref >= 0) setField('preferred', ciqNumber(ks[rPref][1]), 'Key Stats', rPref, 1, ks[rPref][0], ks[rPref][1]);
    if (rNci >= 0) setField('nci', ciqNumber(ks[rNci][1]), 'Key Stats', rNci, 1, ks[rNci][0], ks[rNci][1]);
    rec.ciq.cash = cash; rec.ciq.debt = debt; rec.ciq.ltMarketable = ltms;
  } else {
    // Financial-company template: CIQ shows no EV bridge in Key Stats.
    // Fall back to Capital Structure Summary (latest balance sheet column).
    rec.meta.evSource = 'Capital Structure Summary (no EV bridge in Key Stats)';
    issue(SEVERITY.REVIEW, 'NO_EV_BRIDGE', 'netDebt',
      'CIQ shows no enterprise value bridge for this company (financial-company template). Net debt and NCI taken from Capital Structure Summary, latest quarter. Confirm debt is corporate debt, not operating/client funding.');
    var cs = sheets['Capital Structure Summary'];
    if (!cs) {
      issue(SEVERITY.ERROR, 'MISSING_FIELD', 'netDebt', 'No Capital Structure Summary tab to source net debt from.');
    } else {
      var csHdr = findRow(cs, 'For the Fiscal Period Ending');
      var csCol = -1;
      if (csHdr >= 0) for (var cc = cs[csHdr].length - 1; cc >= 1; cc--) { if (String(cs[csHdr][cc] || '').trim()) { csCol = cc; break; } }
      var csSet = readSettings(cs);
      var csCurRow = findRow(cs, 'Currency', { from: csHdr + 1, to: csHdr + 3, exact: true });
      var csCur = csCurRow >= 0 ? String(cs[csCurRow][csCol]).trim() : '';
      rec.meta.csPeriod = csCol >= 0 ? parsePeriodHeader(cs[csHdr][csCol]) : null;
      if (csCur && csCur !== 'USD') issue(SEVERITY.ERROR, 'CURRENCY', 'netDebt', 'Capital Structure Summary is in ' + csCur + ', not USD.');
      if (!/millions/i.test(csSet.scale || '')) issue(SEVERITY.ERROR, 'UNITS', 'netDebt', 'Capital Structure Summary is not in millions.');
      var rTd = findRow(cs, 'Total Debt', { exact: true });
      var rTc = findRow(cs, 'Total Cash & ST Investments');
      var rMi = findRow(cs, 'Total Minority Interest');
      var rPe = findRow(cs, 'Total Preferred Equity');
      var rStb = findRow(cs, 'Total Short-Term Borrowings');
      if (csCol >= 0 && rTd >= 0 && rTc >= 0) {
        debt = ciqNumber(cs[rTd][csCol]);
        cash = ciqNumber(cs[rTc][csCol]);
        rec.fields.netDebt = debt - cash;
        rec.sources.netDebt = {
          file: fileName, sheet: 'Capital Structure Summary',
          cell: a1(rTd, csCol) + ' − ' + a1(rTc, csCol),
          label: 'Total Debt − Total Cash & ST Investments (' + (rec.meta.csPeriod ? rec.meta.csPeriod.raw : '') + ')',
          raw: debt + ' − ' + cash
        };
        rec.ciq.cash = cash; rec.ciq.debt = debt;
        if (rStb >= 0) {
          var stb = ciqNumber(cs[rStb][csCol]) || 0;
          rec.ciq.shortTermBorrowings = stb;
          if (debt > 0 && stb / debt > 0.5) {
            issue(SEVERITY.REVIEW, 'BROKER_FUNDING_IN_DEBT', 'netDebt',
              'Short-term borrowings are ' + Math.round(stb / debt * 100) + '% of total debt ($' + fmt(stb) + 'mm of $' + fmt(debt) + 'mm). Likely broker / securities-lending funding. The Read Me excludes balance-sheet-intensive brokers for this reason; banker to decide treatment.');
          }
        }
      } else {
        issue(SEVERITY.ERROR, 'MISSING_FIELD', 'netDebt', 'Could not find Total Debt / Total Cash in Capital Structure Summary.');
      }
      if (csCol >= 0 && rMi >= 0) setField('nci', ciqNumber(cs[rMi][csCol]), 'Capital Structure Summary', rMi, csCol, cs[rMi][0], cs[rMi][csCol]);
      if (csCol >= 0 && rPe >= 0) setField('preferred', ciqNumber(cs[rPe][csCol]), 'Capital Structure Summary', rPe, csCol, cs[rPe][0], cs[rPe][csCol]);
      else { rec.fields.preferred = 0; rec.sources.preferred = { file: fileName, sheet: 'Key Stats', cell: '', label: 'No preferred equity line in export: set to 0', raw: '' }; }
    }
  }

  // Implied shares (Up-C structures, e.g. Tradeweb) may live in Historical Capitalization.
  var hc = sheets['Historical Capitalization'];
  if (hc && rec.ciq.impliedShares == null) {
    var rIs = findRow(hc, 'Implied Shares Outstanding');
    if (rIs >= 0) {
      for (var hcC = hc[rIs].length - 1; hcC >= 1; hcC--) {
        var iv = ciqNumber(hc[rIs][hcC]);
        if (iv !== null) { rec.ciq.impliedShares = iv; break; }
      }
    }
  }

  // ---- Income statement: reporting currency, as-reported revenue, filing date
  var is = sheets['Income Statement'];
  if (is) {
    var isSet = readSettings(is);
    rec.meta.isCurrencySetting = isSet.currency || '';
    var isHdr = findRow(is, 'For the Fiscal Period Ending');
    var isCol = -1;
    if (isHdr >= 0) for (var ic = 1; ic < is[isHdr].length; ic++) { if (parsePeriodHeader(is[isHdr][ic]).isLtm) { isCol = ic; break; } }
    if (isCol >= 0) {
      var isCur = findRow(is, 'Currency', { from: isHdr + 1, to: isHdr + 3, exact: true });
      rec.meta.reportedCurrency = isCur >= 0 ? String(is[isCur][isCol]).trim() : '';
      var rAr = findRow(is, 'As Reported Total Revenue');
      if (rAr >= 0) rec.ciq.asReportedRevenue = ciqNumber(is[rAr][isCol]);
      var rFd = findRow(is, 'Filing Date', { exact: true });
      if (rFd >= 0) rec.meta.filingDate = anyToIsoDate(is[rFd][isCol]);
    }
  }

  // ---- Implied EBITDA for financial-company templates (no EBITDA line) ---
  var ebitdaImplied = false;
  if (rec.fields.ltmEbitda == null && is) {
    var imp = impliedEbitda(sheets, fileName);
    if (imp) {
      ebitdaImplied = true;
      rec.fields.ltmEbitda = imp.value;
      rec.sources.ltmEbitda = imp.source;
      issue(SEVERITY.REVIEW, 'IMPLIED_EBITDA', 'ltmEbitda',
        'CIQ does not report EBITDA for this company (financial template). Implied LTM EBITDA = operating income $' + fmt(imp.ebit, 1) + 'mm + D&A $' + fmt(imp.da, 1) + 'mm (' + imp.daLabel + ') = $' + fmt(imp.value, 1) + 'mm. Banker to confirm this is a meaningful metric for the business.');
    }
  }

  // ---- Labels written to the model (basis / period columns) -------------
  if (period) {
    rec.fields.financialPeriod = 'LTM ' + prettyDate(period.endDate) + (period.isPressRelease ? ' (press release)' : '');
    rec.sources.financialPeriod = { file: fileName, sheet: 'Key Stats', cell: a1(hdr, ltmCol), label: 'Key Financials period header', raw: period.raw };
  }
  var fx = rec.meta.reportedCurrency && rec.meta.reportedCurrency !== 'USD'
    ? ' · reported in ' + rec.meta.reportedCurrency + ', converted at CIQ spot' : '';
  rec.fields.revenueBasis = 'CIQ Total Revenue (standardised)' + fx;
  rec.fields.ebitdaBasis = (ebitdaImplied ? 'Implied: CIQ operating income + D&A' : 'CIQ EBITDA (standardised, unadjusted)') + fx;
  rec.sources.revenueBasis = { file: fileName, sheet: 'Key Stats', cell: '', label: 'Derived from CIQ export settings', raw: rec.meta.currencySetting };
  rec.sources.ebitdaBasis = rec.sources.revenueBasis;

  // ---- Missing-field checks --------------------------------------------
  ['sharesOut', 'netDebt', 'nci', 'preferred', 'ltmRevenue', 'ltmEbitda', 'revenueGrowth'].forEach(function (f) {
    if (rec.fields[f] === undefined || rec.fields[f] === null) {
      var already = rec.issues.some(function (i) { return i.field === f && i.severity === SEVERITY.ERROR; });
      if (!already) issue(SEVERITY.ERROR, 'MISSING_FIELD', f, FIELD_LABELS[f] + ' not found or not available in the export. Existing model value kept.');
      delete rec.fields[f];
    }
  });

  return rec;

  function setField(field, value, sheet, r, c, label, raw) {
    rec.fields[field] = value;
    rec.sources[field] = srcRef(fileName, sheet, r, c, label, raw);
    if (isCiqDash(raw)) rec.sources[field].note = 'CIQ shows "-" (nil): treated as 0';
  }
}

/** Row where the Key Stats capitalization block starts ('Latest' or 'Current Capitalization'). */
function findCapStart(ks, from) {
  for (var r = from || 0; r < ks.length; r++) {
    if (/^(latest|current) capitalization/.test(normLabel(ks[r] && ks[r][0]))) return r;
  }
  return -1;
}

/** Column index of the LTM period in a statement tab, or -1. */
function ltmColumn(rows) {
  var h = findRow(rows, 'For the Fiscal Period Ending');
  if (h < 0) return { hdr: -1, col: -1 };
  for (var c = 1; c < rows[h].length; c++) if (parsePeriodHeader(rows[h][c]).isLtm) return { hdr: h, col: c };
  return { hdr: h, col: -1 };
}

/**
 * EBITDA proxy when CIQ shows none: Operating Income + D&A.
 * D&A from the income statement if it is a separate opex line, otherwise
 * cash-flow 'Depreciation & Amort., Total' + 'Other Amortization'.
 */
function impliedEbitda(sheets, fileName) {
  var is = sheets['Income Statement'], cf = sheets['Cash Flow'];
  if (!is) return null;
  var L = ltmColumn(is);
  if (L.col < 0) return null;
  var rOi = findRow(is, 'Operating Income', { exact: true });
  if (rOi < 0) return null;
  var ebit = ciqNumber(is[rOi][L.col]);
  if (ebit == null) return null;
  var da = null, daLabel = '', cells = [];
  var rDa = findRow(is, 'Depreciation & Amort.', { exact: true });
  if (rDa >= 0 && ciqNumber(is[rDa][L.col])) {
    da = ciqNumber(is[rDa][L.col]); daLabel = 'income statement'; cells.push('Income Statement!' + a1(rDa, L.col));
  } else if (cf) {
    var C = ltmColumn(cf);
    var rT = findRow(cf, 'Depreciation & Amort., Total', { exact: true });
    var rO = findRow(cf, 'Other Amortization', { exact: true });
    if (C.col >= 0 && rT >= 0) {
      da = (ciqNumber(cf[rT][C.col]) || 0) + (rO >= 0 ? (ciqNumber(cf[rO][C.col]) || 0) : 0);
      daLabel = 'cash flow D&A total' + (rO >= 0 ? ' + other amortization' : '');
      cells.push('Cash Flow!' + a1(rT, C.col)); if (rO >= 0) cells.push('Cash Flow!' + a1(rO, C.col));
    }
  }
  if (da == null) return null;
  return {
    value: ebit + da, ebit: ebit, da: da, daLabel: daLabel,
    source: { file: fileName, sheet: 'Income Statement', cell: a1(rOi, L.col) + ' + ' + cells.join(' + '), label: 'Implied EBITDA = Operating Income + D&A (' + daLabel + ')', raw: ebit + ' + ' + da }
  };
}

var FIELD_LABELS = {
  price: 'Price', priorPrice: 'Prior Price', sharesOut: 'Shares outstanding', netDebt: 'Net debt', nci: 'NCI',
  preferred: 'Preferred equity', ltmRevenue: 'LTM revenue', ltmEbitda: 'LTM EBITDA', revenueGrowth: 'Revenue growth',
  revenueBasis: 'Revenue basis', ebitdaBasis: 'EBITDA basis', financialPeriod: 'Financial period'
};

function fmt(n, dp) {
  if (n === null || n === undefined || !isFinite(n)) return 'n/a';
  dp = dp == null ? 0 : dp;
  var s = Math.abs(n).toFixed(dp).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '-' : '') + s;
}

function pct(n, dp) {
  if (n === null || n === undefined || !isFinite(n)) return 'n/a';
  return (n * 100).toFixed(dp == null ? 1 : dp) + '%';
}

/* ======================== 3. Financial validation ======================== */

/** Most recent quarter end that companies should have reported by `todayIso`. */
function expectedLatestQuarter(todayIso, lagDays) {
  lagDays = lagDays == null ? 45 : lagDays;
  var t = Date.parse(todayIso);
  var d = new Date(t - lagDays * 86400000);
  var y = d.getUTCFullYear(), m = d.getUTCMonth();
  var qEndMonth = m - ((m + 1) % 3);           // 2,5,8,11 are quarter-end months
  if ((m + 1) % 3 === 0) qEndMonth = m;
  var end = new Date(Date.UTC(y, qEndMonth + 1, 0));
  if (end.getTime() > d.getTime()) end = new Date(Date.UTC(y, qEndMonth - 2, 0));
  return isoDate(end);
}

/** Mode of an array of strings. */
function mode(arr) {
  var counts = {}, best = null, n = 0;
  arr.forEach(function (x) { if (!x) return; counts[x] = (counts[x] || 0) + 1; if (counts[x] > n) { n = counts[x]; best = x; } });
  return best;
}

/**
 * Validate one parsed record. Adds issues to rec.issues and returns them.
 * ctx: { todayIso, peerPeriod, existing: {field: value}, settings }
 */
function validateFinancials(rec, ctx) {
  ctx = ctx || {};
  var s = ctx.settings || DEFAULT_SETTINGS;
  var f = rec.fields, ciq = rec.ciq, meta = rec.meta;
  function issue(sev, code, field, msg) { rec.issues.push({ severity: sev, code: code, field: field || '', message: msg }); }

  // Currency of the figures that land in the model.
  if (meta.ksCurrency && meta.ksCurrency !== 'USD') {
    issue(SEVERITY.ERROR, 'CURRENCY', 'ltmRevenue', 'Key Financials are in ' + meta.ksCurrency + ', model is US$. Re-export with Currency = US Dollar.');
    delete f.ltmRevenue; delete f.ltmEbitda;
  }
  if (meta.capCurrency && meta.capCurrency !== 'USD') {
    issue(SEVERITY.ERROR, 'CURRENCY', 'sharesOut', 'Capitalization is in ' + meta.capCurrency + ', model is US$.');
    delete f.netDebt; delete f.nci; delete f.preferred;
  }
  if (meta.reportedCurrency && meta.reportedCurrency !== 'USD') {
    issue(SEVERITY.REVIEW, 'FX_TRANSLATION', 'ltmRevenue',
      'Company reports in ' + meta.reportedCurrency + '. CIQ translated LTM figures to US$ at today\'s spot rate (setting: "' + (meta.conversion || 'n/a') + '"), not period-average rates. Confirm the FX convention.');
  }

  // Reporting period.
  var per = meta.period;
  if (per) {
    if (per.isPressRelease) {
      issue(SEVERITY.REVIEW, 'PRESS_RELEASE', 'financialPeriod',
        'LTM to ' + prettyDate(per.endDate) + ' is built from a press release (not yet in a 10-Q/10-K). Figures may change on filing, and EBITDA may be less complete.');
    }
    if (ctx.peerPeriod && per.endDate && per.endDate !== ctx.peerPeriod) {
      issue(SEVERITY.REVIEW, 'PERIOD_MISMATCH', 'financialPeriod',
        'LTM ends ' + prettyDate(per.endDate) + ' while most peers end ' + prettyDate(ctx.peerPeriod) + '. Multiples are not on a like-for-like period.');
    }
    if (ctx.todayIso && per.endDate) {
      var exp = expectedLatestQuarter(ctx.todayIso, s.reportingLagDays);
      if (per.endDate < exp) {
        issue(SEVERITY.INFO, 'STALE_PERIOD', 'financialPeriod',
          'Latest LTM is ' + prettyDate(per.endDate) + '; results to ' + prettyDate(exp) + ' should be published by now. Request a refreshed CIQ export.');
      }
    }
    if (per.months && per.months !== 12) issue(SEVERITY.ERROR, 'PERIOD_LENGTH', 'financialPeriod', 'LTM column covers ' + per.months + ' months, not 12.');
  }

  // Reconciliations — prove the mapping by rebuilding CIQ's own numbers.
  if (ciq.price != null && f.sharesOut != null && ciq.marketCap != null && ciq.marketCap !== 0) {
    var mc = ciq.price * f.sharesOut;
    var dMc = mc / ciq.marketCap - 1;
    if (Math.abs(dMc) > s.reconTolerance) {
      issue(SEVERITY.ERROR, 'RECON_MARKET_CAP', 'sharesOut',
        'Price × shares ($' + fmt(mc) + 'mm) does not match CIQ market cap ($' + fmt(ciq.marketCap) + 'mm), diff ' + pct(dMc, 2) + '. Check the share count mapping.');
    }
  }
  if (ciq.tev != null && ciq.marketCap != null && f.netDebt != null) {
    var ev = ciq.marketCap + f.netDebt + (f.nci || 0) + (f.preferred || 0);
    var dEv = ev / ciq.tev - 1;
    if (Math.abs(dEv) > s.reconTolerance) {
      issue(SEVERITY.ERROR, 'RECON_EV', 'netDebt',
        'Rebuilt EV ($' + fmt(ev) + 'mm) does not match CIQ TEV ($' + fmt(ciq.tev) + 'mm), diff ' + pct(dEv, 2) + '. Check net debt / NCI / preferred mapping.');
    } else {
      rec.meta.evReconciled = true;
    }
  }

  // Capital structure judgement items.
  if (ciq.impliedShares != null && f.sharesOut) {
    var dImp = ciq.impliedShares / f.sharesOut - 1;
    if (Math.abs(dImp) > s.impliedSharesTolerance) {
      issue(SEVERITY.REVIEW, 'IMPLIED_SHARES', 'sharesOut',
        'CIQ also reports ' + fmt(ciq.impliedShares, 1) + 'mm implied shares (incl. exchangeable units) vs ' + fmt(f.sharesOut, 1) + 'mm listed shares (' + pct(dImp) + '). Model uses listed shares per CIQ convention; banker to confirm.');
    }
  }
  if (f.nci != null && f.nci < 0) issue(SEVERITY.REVIEW, 'NEGATIVE_NCI', 'nci', 'NCI is negative ($' + fmt(f.nci, 1) + 'mm), which reduces EV. Confirm the treatment.');
  if (f.preferred) issue(SEVERITY.INFO, 'PREFERRED', 'preferred', 'Preferred equity of $' + fmt(f.preferred) + 'mm included in EV.');
  if (f.nci && f.nci > 0 && rec.fields.sharesOut && ciq.marketCap && f.nci / ciq.marketCap > 0.05) {
    issue(SEVERITY.INFO, 'MATERIAL_NCI', 'nci', 'NCI of $' + fmt(f.nci) + 'mm (' + pct(f.nci / ciq.marketCap) + ' of market cap) included in EV.');
  }

  // Earnings quality.
  if (f.ltmEbitda != null && f.ltmEbitda <= 0) issue(SEVERITY.INFO, 'NEGATIVE_EBITDA', 'ltmEbitda', 'LTM EBITDA is ' + (f.ltmEbitda < 0 ? 'negative' : 'zero') + ' ($' + fmt(f.ltmEbitda, 1) + 'mm): EV / EBITDA shows NM.');
  if (f.ltmRevenue != null && f.ltmRevenue <= 0) issue(SEVERITY.REVIEW, 'NON_POSITIVE_REVENUE', 'ltmRevenue', 'LTM revenue is not positive: EV / Revenue shows NM.');
  if (f.revenueGrowth != null && Math.abs(f.revenueGrowth) > s.growthOutlier) {
    issue(SEVERITY.REVIEW, 'GROWTH_OUTLIER', 'revenueGrowth', 'Revenue growth of ' + pct(f.revenueGrowth) + ' is outside ±' + pct(s.growthOutlier, 0) + '. Check for M&A, IPO or a change of basis.');
  }
  if (ciq.asReportedRevenue != null && f.ltmRevenue && (!meta.reportedCurrency || meta.reportedCurrency === 'USD')) {
    var dRev = ciq.asReportedRevenue / f.ltmRevenue - 1;
    if (Math.abs(dRev) > s.revenueBasisTolerance) {
      issue(SEVERITY.INFO, 'REVENUE_DEFINITION', 'ltmRevenue',
        'CIQ standardised revenue ($' + fmt(f.ltmRevenue) + 'mm) differs from as-reported revenue ($' + fmt(ciq.asReportedRevenue) + 'mm) by ' + pct(dRev) + '. Model uses CIQ standardised for comparability.');
    }
  }

  // Movement vs the values already in the model.
  var ex = ctx.existing || {};
  [['sharesOut', s.changeShares], ['ltmRevenue', s.changeFinancials], ['ltmEbitda', s.changeFinancials]].forEach(function (pair) {
    var k = pair[0], tol = pair[1];
    if (typeof ex[k] === 'number' && ex[k] !== 0 && typeof f[k] === 'number') {
      var d = f[k] / ex[k] - 1;
      if (Math.abs(d) > tol) issue(SEVERITY.REVIEW, 'LARGE_CHANGE', k, FIELD_LABELS[k] + ' moves ' + pct(d) + ' vs the current model (' + fmt(ex[k], 1) + ' → ' + fmt(f[k], 1) + ').');
    }
  });
  if (typeof ex.netDebt === 'number' && typeof f.netDebt === 'number' && ex.netDebt * f.netDebt < 0 && Math.abs(ex.netDebt - f.netDebt) > 100) {
    issue(SEVERITY.REVIEW, 'NET_DEBT_SIGN', 'netDebt', 'Net debt flips between net debt and net cash (' + fmt(ex.netDebt) + ' → ' + fmt(f.netDebt) + ').');
  }

  return rec.issues;
}

/* ========================== 4. Price validation ========================== */

/**
 * p: { ticker, price, prior, currency, tradeDate, secondary, todayIso }
 * returns issues[]
 */
function validatePrice(p, settings) {
  var s = settings || DEFAULT_SETTINGS;
  var out = [];
  function issue(sev, code, msg) { out.push({ severity: sev, code: code, field: 'price', message: msg }); }
  if (p.price === null || p.price === undefined || !isFinite(p.price)) {
    issue(SEVERITY.ERROR, 'PRICE_MISSING', 'No price returned' + (p.error ? ' (' + p.error + ')' : '') + '. Previous price kept.');
    return out;
  }
  if (p.price <= 0) { issue(SEVERITY.ERROR, 'PRICE_INVALID', 'Price ' + p.price + ' is not positive. Previous price kept.'); return out; }
  if (p.currency && p.currency !== 'USD') issue(SEVERITY.ERROR, 'PRICE_CURRENCY', 'Quote currency is ' + p.currency + ', model is US$.');
  if (p.tradeDate && p.todayIso) {
    var age = daysBetween(p.tradeDate, p.todayIso);
    if (age > s.staleDays) issue(SEVERITY.REVIEW, 'PRICE_STALE', 'Last trade ' + p.tradeDate + ' (' + age + ' days old). Possible halt or delisting.');
  }
  if (typeof p.prior === 'number' && p.prior > 0) {
    var mv = p.price / p.prior - 1;
    if (Math.abs(mv) > s.weeklyMoveAlert) issue(SEVERITY.REVIEW, 'PRICE_JUMP', 'Moved ' + pct(mv) + ' week on week. Check for corporate action (split, special dividend) or bad tick.');
  }
  if (typeof p.secondary === 'number' && p.secondary > 0) {
    var dv = p.price / p.secondary - 1;
    if (Math.abs(dv) > s.crossCheckTolerance) issue(SEVERITY.REVIEW, 'PRICE_DIVERGENCE', 'Google Finance $' + p.price.toFixed(2) + ' vs cross-check $' + p.secondary.toFixed(2) + ' (' + pct(dv) + ').');
  }
  return out;
}

/* ============================ 5. Model helpers ============================ */

var DEFAULT_SETTINGS = {
  reconTolerance: 0.005,        // 0.5% for market cap / EV reconciliation to CIQ
  impliedSharesTolerance: 0.02, // 2%
  growthOutlier: 1.0,           // ±100%
  revenueBasisTolerance: 0.02,  // 2%
  changeShares: 0.10,           // 10% move in shares vs model
  changeFinancials: 0.30,       // 30% move in LTM revenue / EBITDA vs model
  weeklyMoveAlert: 0.20,        // 20% weekly price move
  crossCheckTolerance: 0.02,    // 2% Google vs secondary source
  staleDays: 5,                 // last trade older than 5 days
  reportingLagDays: 45,         // results expected 45 days after quarter end
  priorPriceMinDays: 6          // prior price = latest close at least 6 days before price date
};

/**
 * Model columns, located by header text so the engine also works on an existing
 * Areta model with a different column order. Keys are internal field names.
 */
var MODEL_HEADERS = {
  company: 'Company', ticker: 'Ticker', price: 'Price ($)', priorPrice: 'Prior Price ($)', change: 'Δ % w/w',
  sharesOut: 'Shares (mm)', marketCap: 'Market Cap ($mm)', netDebt: 'Net Debt ($mm)', nci: 'NCI ($mm)',
  preferred: 'Preferred ($mm)', ev: 'EV ($mm)', ltmRevenue: 'LTM Revenue ($mm)', ltmEbitda: 'LTM EBITDA ($mm)',
  revenueGrowth: 'Revenue Growth', evRevenue: 'EV / Revenue', evEbitda: 'EV / EBITDA', revenueBasis: 'Revenue Basis',
  ebitdaBasis: 'EBITDA Basis', financialPeriod: 'Financial Period', notes: 'Notes'
};

/** Quarterly (blue) inputs written by the financial import, in model order. */
var FINANCIAL_FIELDS = ['sharesOut', 'netDebt', 'nci', 'preferred', 'ltmRevenue', 'ltmEbitda', 'revenueGrowth', 'revenueBasis', 'ebitdaBasis', 'financialPeriod'];

/**
 * Locate the header row and column of each field.
 * @return {{headerRow:number, cols:Object<string,number>, missing:string[]}} 0-based
 */
function mapModelHeaders(values) {
  for (var r = 0; r < Math.min(values.length, 30); r++) {
    var row = values[r].map(normLabel);
    if (row.indexOf('company') >= 0 && row.indexOf('ticker') >= 0) {
      var cols = {}, missing = [];
      Object.keys(MODEL_HEADERS).forEach(function (k) {
        var idx = row.indexOf(normLabel(MODEL_HEADERS[k]));
        if (idx >= 0) cols[k] = idx; else missing.push(MODEL_HEADERS[k]);
      });
      return { headerRow: r, cols: cols, missing: missing };
    }
  }
  return null;
}

/**
 * Read company rows, peer groups and median rows from the model grid.
 * A company row has a ticker; a group header has a label and no ticker;
 * a median row's label starts with "Median"; "All companies median" is global.
 */
function readModelLayout(values, map) {
  var cName = map.cols.company, cTick = map.cols.ticker;
  var groups = [], companies = [], medians = [], globalMedian = -1, current = null;
  for (var r = map.headerRow + 1; r < values.length; r++) {
    var name = String(values[r][cName] || '').trim();
    var tick = String(values[r][cTick] || '').trim();
    if (!name && !tick) continue;
    if (/^all companies median/i.test(name)) { globalMedian = r; continue; }
    if (/^median/i.test(name)) { medians.push({ row: r, group: current ? current.name : '' }); if (current) current.medianRow = r; continue; }
    if (tick) {
      var co = { row: r, name: name, ticker: tick, group: current ? current.name : '' };
      companies.push(co);
      if (current) current.rows.push(r);
      continue;
    }
    current = { name: name, row: r, rows: [], medianRow: -1 };
    groups.push(current);
  }
  return { groups: groups, companies: companies, medians: medians, globalMedianRow: globalMedian };
}

/** Normalise 'NasdaqGS:COIN' / 'NASDAQ:COIN' / 'COIN' to the bare symbol. */
function tickerSymbol(t) {
  var s = String(t || '').trim().toUpperCase();
  var i = s.lastIndexOf(':');
  return i >= 0 ? s.slice(i + 1) : s;
}

/** CIQ exchange prefix -> Google Finance exchange prefix. */
var GOOGLE_EXCHANGE = {
  NASDAQGS: 'NASDAQ', NASDAQGM: 'NASDAQ', NASDAQCM: 'NASDAQ', NASDAQ: 'NASDAQ',
  NYSE: 'NYSE', NYSEAM: 'NYSEAMERICAN', NYSEAMERICAN: 'NYSEAMERICAN', NYSEARCA: 'NYSEARCA',
  BATS: 'BATS', CBOE: 'BATS'
};

function toGoogleTicker(ciqTicker) {
  var s = String(ciqTicker || '').trim();
  var i = s.indexOf(':');
  if (i < 0) return s.toUpperCase();
  var ex = s.slice(0, i).toUpperCase(), sym = s.slice(i + 1).toUpperCase();
  return (GOOGLE_EXCHANGE[ex] || ex) + ':' + sym;
}

/** Match a CIQ record to a model company: exact ticker first, then symbol. */
function matchCompany(companies, ciqTicker) {
  var t = String(ciqTicker || '').toUpperCase();
  var exact = companies.filter(function (c) { return c.ticker.toUpperCase() === t; });
  if (exact.length === 1) return { company: exact[0], how: 'ticker' };
  var sym = tickerSymbol(t);
  var bySym = companies.filter(function (c) { return tickerSymbol(c.ticker) === sym; });
  if (bySym.length === 1) return { company: bySym[0], how: 'symbol' };
  if (bySym.length > 1) return { company: null, how: 'ambiguous' };
  return { company: null, how: 'none' };
}

/** JS mirror of the sheet formulas: used to cross-check the sheet's own results. */
function computeRow(v) {
  var out = {};
  out.marketCap = isNum(v.price) && isNum(v.sharesOut) ? v.price * v.sharesOut : null;
  out.ev = out.marketCap != null && isNum(v.netDebt) ? out.marketCap + v.netDebt + (v.nci || 0) + (v.preferred || 0) : null;
  out.change = isNum(v.price) && isNum(v.priorPrice) && v.priorPrice !== 0 ? v.price / v.priorPrice - 1 : null;
  out.evRevenue = out.ev != null && isNum(v.ltmRevenue) && v.ltmRevenue > 0 ? out.ev / v.ltmRevenue : 'NM';
  out.evEbitda = out.ev != null && isNum(v.ltmEbitda) && v.ltmEbitda > 0 ? out.ev / v.ltmEbitda : 'NM';
  return out;
}

function isNum(x) { return typeof x === 'number' && isFinite(x); }

/** Median ignoring non-numbers ("NM"); 'NM' when nothing is numeric. */
function medianIgnoringNM(arr) {
  var a = arr.filter(isNum).sort(function (x, y) { return x - y; });
  if (!a.length) return 'NM';
  var m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
}

/** Pick prior price from history: latest close at least minDays before priceDate. */
function pickPriorPrice(history, priceDateIso, minDays) {
  var best = null;
  history.forEach(function (h) {
    if (!h.date || !isNum(h.price)) return;
    if (daysBetween(h.date, priceDateIso) >= minDays && (!best || h.date > best.date)) best = h;
  });
  return best;
}

/* Node export (ignored by Apps Script, where `module` is undefined). */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    CORE_VERSION: CORE_VERSION, SEVERITY: SEVERITY, DEFAULT_SETTINGS: DEFAULT_SETTINGS, MODEL_HEADERS: MODEL_HEADERS,
    FINANCIAL_FIELDS: FINANCIAL_FIELDS, FIELD_LABELS: FIELD_LABELS,
    normLabel: normLabel, ciqNumber: ciqNumber, colLetter: colLetter, a1: a1, parsePeriodHeader: parsePeriodHeader,
    parseCiqDateText: parseCiqDateText, anyToIsoDate: anyToIsoDate, prettyDate: prettyDate, expectedLatestQuarter: expectedLatestQuarter,
    parseCiqWorkbook: parseCiqWorkbook, validateFinancials: validateFinancials, validatePrice: validatePrice, mode: mode,
    mapModelHeaders: mapModelHeaders, readModelLayout: readModelLayout, toGoogleTicker: toGoogleTicker, tickerSymbol: tickerSymbol,
    matchCompany: matchCompany, computeRow: computeRow, medianIgnoringNM: medianIgnoringNM, pickPriorPrice: pickPriorPrice, fmt: fmt, pct: pct
  };
}
