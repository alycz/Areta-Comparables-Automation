// Runs the real Core.js pipeline over every fixture and prints what the engine would write.
const core = require('../appsscript/Core.js');
const { loadAll } = require('../tests/load.js');
const today = process.argv[2] || new Date().toISOString().slice(0, 10);
const recs = loadAll().map(f => core.parseCiqWorkbook(f.sheets, f.fileName));
const peer = core.mode(recs.map(r => r.meta.period && r.meta.period.endDate));
for (const r of recs) {
  core.validateFinancials(r, { todayIso: today, peerPeriod: peer });
  const f = r.fields;
  console.log(`${(r.identity.ticker||'?').padEnd(16)} sh=${core.fmt(f.sharesOut,1).padStart(8)} nd=${core.fmt(f.netDebt).padStart(8)} nci=${core.fmt(f.nci).padStart(6)} pref=${core.fmt(f.preferred).padStart(6)} rev=${core.fmt(f.ltmRevenue).padStart(7)} ebitda=${core.fmt(f.ltmEbitda).padStart(7)} g=${core.pct(f.revenueGrowth).padStart(7)} ${f.financialPeriod} evRecon=${!!r.meta.evReconciled}`);
  for (const i of r.issues) console.log(`     ${i.severity.padEnd(6)} ${i.code}: ${i.message}`);
}
