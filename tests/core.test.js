// Run with: npm test   (Node 18+, no dependencies)
const test = require('node:test');
const assert = require('node:assert');
const core = require('../appsscript/Core.js');
const { loadAll } = require('./load.js');

const files = loadAll();
const recs = files.map(f => core.parseCiqWorkbook(f.sheets, f.fileName));
const peer = core.mode(recs.map(r => r.meta.period && r.meta.period.endDate));
recs.forEach(r => core.validateFinancials(r, { todayIso: '2026-11-25', peerPeriod: peer }));
const by = t => recs.find(r => r.identity.symbol === t);
const codes = (r, sev) => r.issues.filter(i => !sev || i.severity === sev).map(i => i.code);
const close = (a, b, tol = 1e-6) => assert.ok(Math.abs(a - b) <= tol * Math.max(1, Math.abs(b)), `${a} != ${b}`);

test('all 20 exports parse and map to a ticker', () => {
  assert.strictEqual(recs.length, 20);
  recs.forEach(r => assert.ok(r.identity.ticker, r.fileName));
});

test('every export yields every model field (no ERROR items)', () => {
  recs.forEach(r => assert.deepStrictEqual(codes(r, 'ERROR'), [], r.identity.ticker));
  recs.forEach(r => core.FINANCIAL_FIELDS.forEach(f => assert.ok(r.fields[f] !== undefined, r.identity.ticker + ' ' + f)));
});

test('EV rebuilt from extracted fields reconciles to CIQ TEV for every company with a CIQ bridge', () => {
  const withBridge = recs.filter(r => r.ciq.tev != null);
  assert.strictEqual(withBridge.length, 18);
  withBridge.forEach(r => assert.ok(r.meta.evReconciled, r.identity.ticker));
});

test('Coinbase values match the export', () => {
  const r = by('COIN');
  close(r.fields.sharesOut, 263.836923);
  close(r.fields.netDebt, 6668.29 - 8788.843);
  close(r.fields.ltmRevenue, 6043.753);
  close(r.fields.ltmEbitda, 699.549);
  close(r.fields.revenueGrowth, -0.092388);
  assert.strictEqual(r.fields.financialPeriod, 'LTM Jun-30-2026');
  assert.strictEqual(r.sources.ltmRevenue.sheet, 'Key Stats');
});

test('Invesco preferred and NCI flow into EV', () => {
  const r = by('IVZ');
  close(r.fields.preferred, 2510.5);
  close(r.fields.nci, 998.3);
});

test('BlackRock: press-release LTM, period mismatch, Current Capitalization layout', () => {
  const r = by('BLK');
  assert.ok(codes(r).includes('PRESS_RELEASE'));
  assert.ok(codes(r).includes('PERIOD_MISMATCH'));
  close(r.fields.sharesOut, 154.6);
  assert.strictEqual(r.fields.financialPeriod, 'LTM Sep-30-2026 (press release)');
});

test('Coincheck: JPY reporter flagged, figures in USD', () => {
  const r = by('CNCK');
  assert.ok(codes(r, 'REVIEW').includes('FX_TRANSLATION'));
  assert.strictEqual(r.meta.ksCurrency, 'USD');
  assert.match(r.fields.revenueBasis, /JPY/);
});

test('Robinhood / Tradeweb: no CIQ EV bridge -> capital structure fallback + implied EBITDA', () => {
  const h = by('HOOD'), t = by('TW');
  for (const r of [h, t]) {
    assert.ok(codes(r).includes('NO_EV_BRIDGE'));
    assert.ok(codes(r).includes('IMPLIED_EBITDA'));
  }
  close(h.fields.netDebt, 22925 - 23905);
  close(h.fields.ltmEbitda, 2270 + 42 + 49);
  assert.ok(codes(h).includes('BROKER_FUNDING_IN_DEBT'));
  close(t.fields.ltmEbitda, 961.845 + 246.638);
  assert.ok(codes(t).includes('IMPLIED_SHARES'));
});

test('negative EBITDA and negative NCI are flagged', () => {
  assert.ok(codes(by('GEMI')).includes('NEGATIVE_EBITDA'));
  assert.ok(codes(by('PENN')).includes('NEGATIVE_NCI'));
});

test('CIQ dash means nil; NA means missing', () => {
  assert.strictEqual(core.ciqNumber('-'), 0);
  assert.strictEqual(core.ciqNumber('NA'), null);
  assert.strictEqual(core.ciqNumber('NM'), null);
  assert.strictEqual(core.ciqNumber('(1,234.5)'), -1234.5);
});

test('a non-millions export is blocked', () => {
  const f = JSON.parse(JSON.stringify(files[0], (k, v) => v));
  const sheets = files[0].sheets;
  const copy = {}; for (const k in sheets) copy[k] = sheets[k].map(r => r.slice());
  copy['Key Stats'][6][0] = 'In Thousands of the trading currency';
  const r = core.parseCiqWorkbook(copy, 'x.xls');
  assert.ok(codes(r, 'ERROR').includes('UNITS'));
});

test('large change vs existing model is flagged', () => {
  const r = core.parseCiqWorkbook(files.find(f => f.fileName.startsWith('Coinbase')).sheets, 'c.xls');
  core.validateFinancials(r, { todayIso: '2026-11-25', peerPeriod: '2026-06-30', existing: { sharesOut: 200, ltmRevenue: 6000 } });
  assert.ok(r.issues.some(i => i.code === 'LARGE_CHANGE' && i.field === 'sharesOut'));
  assert.ok(!r.issues.some(i => i.code === 'LARGE_CHANGE' && i.field === 'ltmRevenue'));
});

test('price validation', () => {
  const s = core.DEFAULT_SETTINGS;
  assert.deepStrictEqual(core.validatePrice({ price: null }, s).map(i => i.code), ['PRICE_MISSING']);
  assert.deepStrictEqual(core.validatePrice({ price: -1 }, s).map(i => i.code), ['PRICE_INVALID']);
  assert.ok(core.validatePrice({ price: 130, prior: 100 }, s).some(i => i.code === 'PRICE_JUMP'));
  assert.ok(core.validatePrice({ price: 103, secondary: 100 }, s).some(i => i.code === 'PRICE_DIVERGENCE'));
  assert.ok(core.validatePrice({ price: 100, tradeDate: '2026-11-10', todayIso: '2026-11-25' }, s).some(i => i.code === 'PRICE_STALE'));
  assert.deepStrictEqual(core.validatePrice({ price: 101, prior: 100, secondary: 101, currency: 'USD' }, s), []);
});

test('prior price picks the latest close at least 6 days earlier', () => {
  const h = [{ date: '2026-11-13', price: 1 }, { date: '2026-11-16', price: 2 }, { date: '2026-11-20', price: 3 }];
  assert.strictEqual(core.pickPriorPrice(h, '2026-11-23', 6).price, 2);
  assert.strictEqual(core.pickPriorPrice(h, '2026-11-14', 6), null);
});

test('model helpers: header map, layout, tickers, NM and medians', () => {
  const XLSX_LIKE = [[], [], [], [], [], ['', 'Company', 'Ticker', 'Price ($)', 'Prior Price ($)', 'Δ % w/w', 'Shares (mm)', 'Market Cap ($mm)', 'Net Debt ($mm)', 'NCI ($mm)', 'Preferred ($mm)', 'EV ($mm)', 'LTM Revenue ($mm)', 'LTM EBITDA ($mm)', 'Revenue Growth', 'EV / Revenue', 'EV / EBITDA', 'Revenue Basis', 'EBITDA Basis', 'Financial Period', 'Notes'],
    ['', 'Crypto'], ['', 'Coinbase', 'NasdaqGS:COIN'], ['', 'Median'], [], ['', 'All companies median']];
  const map = core.mapModelHeaders(XLSX_LIKE);
  assert.strictEqual(map.headerRow, 5);
  assert.deepStrictEqual(map.missing, []);
  const lay = core.readModelLayout(XLSX_LIKE, map);
  assert.strictEqual(lay.companies.length, 1);
  assert.strictEqual(lay.groups[0].medianRow, 8);
  assert.strictEqual(lay.globalMedianRow, 10);
  assert.strictEqual(core.toGoogleTicker('NasdaqGS:COIN'), 'NASDAQ:COIN');
  assert.strictEqual(core.toGoogleTicker('NYSEAM:EXOD'), 'NYSEAMERICAN:EXOD');
  assert.strictEqual(core.matchCompany(lay.companies, 'NASDAQ:COIN').how, 'symbol');
  const c = core.computeRow({ price: 10, priorPrice: 8, sharesOut: 100, netDebt: -50, nci: 5, preferred: 0, ltmRevenue: 100, ltmEbitda: -1 });
  assert.strictEqual(c.ev, 955); assert.strictEqual(c.evEbitda, 'NM'); close(c.change, 0.25);
  assert.strictEqual(core.medianIgnoringNM([1, 'NM', 3, 2]), 2);
  assert.strictEqual(core.medianIgnoringNM(['NM']), 'NM');
});

test('expected latest quarter honours the reporting lag', () => {
  assert.strictEqual(core.expectedLatestQuarter('2026-11-25', 45), '2026-09-30');
  assert.strictEqual(core.expectedLatestQuarter('2026-11-10', 45), '2026-06-30');
  assert.strictEqual(core.expectedLatestQuarter('2026-02-20', 45), '2025-12-31');
});

test('crypto fair-value swings in CIQ EBITDA are flagged (CoinShares, Bullish)', () => {
  for (const t of ['CSHR', 'BLSH']) {
    const i = by(t).issues.find(x => x.code === 'EBITDA_NOT_REPRESENTATIVE');
    assert.ok(i, t);
    assert.match(i.message, /fair-value/);
  }
  assert.ok(!codes(by('COIN')).includes('EBITDA_NOT_REPRESENTATIVE'));
});
