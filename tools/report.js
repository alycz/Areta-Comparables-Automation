// Builds docs/VALIDATION_REPORT.md from the supplied CIQ exports using the same Core.js the engine runs.
// Usage: node tools/report.js 2026-11-25
const fs = require('fs');
const core = require('../appsscript/Core.js');
const { loadAll } = require('../tests/load.js');
const today = process.argv[2] || new Date().toISOString().slice(0, 10);
const recs = loadAll().map(f => core.parseCiqWorkbook(f.sheets, f.fileName));
const peer = core.mode(recs.map(r => r.meta.period && r.meta.period.endDate));
recs.forEach(r => core.validateFinancials(r, { todayIso: today, peerPeriod: peer }));
recs.sort((a, b) => a.identity.name.localeCompare(b.identity.name));
const all = recs.flatMap(r => r.issues.map(i => ({ ...i, co: r.identity.name, t: r.identity.ticker })));
const cnt = s => all.filter(i => i.severity === s).length;
const f = core.fmt, p = core.pct;
let md = `# Validation report: supplied CIQ exports\n\nGenerated ${today} by \`tools/report.js\`, which runs the same \`Core.js\` as the Apps Script engine. The live Sheet also builds this report on its **Validation Report** tab, with prices included.\n\n`;
md += `## Summary\n\n| | |\n|---|---|\n| Exports read | ${recs.length} of 20 |\n| Companies matched to the model | ${recs.length} |\n| Fields extracted | ${recs.reduce((n, r) => n + Object.keys(r.fields).length, 0)} |\n| Market cap and EV rebuilt and reconciled to CIQ (±0.5%) | ${recs.filter(r => r.meta.evReconciled).length} of ${recs.filter(r => r.ciq.tev != null).length} with a CIQ EV bridge |\n| ERROR items (value blocked) | ${cnt('ERROR')} |\n| REVIEW items (banker judgement) | ${cnt('REVIEW')} |\n| INFO items | ${cnt('INFO')} |\n| Peer LTM period | ${core.prettyDate(peer)} |\n\n`;
md += `## Items requiring banker judgement\n\n| Company | Issue | Detail |\n|---|---|---|\n`;
all.filter(i => i.severity === 'REVIEW').forEach(i => { md += `| ${i.co} | ${i.code} | ${i.message.replace(/\|/g, '/')} |\n`; });
md += `\n## Extracted values (US$ mm)\n\n| Company | Ticker | Shares (mm) | Net debt | NCI | Pref. | LTM revenue | LTM EBITDA | Growth | Period | EV source | Reconciles to CIQ |\n|---|---|--:|--:|--:|--:|--:|--:|--:|---|---|---|\n`;
recs.forEach(r => { const x = r.fields; md += `| ${r.identity.name} | ${r.identity.ticker} | ${f(x.sharesOut, 1)} | ${f(x.netDebt)} | ${f(x.nci)} | ${f(x.preferred)} | ${f(x.ltmRevenue)} | ${f(x.ltmEbitda)} | ${p(x.revenueGrowth)} | ${x.financialPeriod} | ${r.meta.evSource.startsWith('Key') ? 'CIQ TEV bridge' : 'Capital Structure Summary'} | ${r.ciq.tev == null ? 'n/a (no CIQ TEV)' : r.meta.evReconciled ? 'Yes' : 'NO'} |\n`; });
md += `\n## For information\n\n| Company | Issue | Detail |\n|---|---|---|\n`;
all.filter(i => i.severity === 'INFO').forEach(i => { md += `| ${i.co} | ${i.code} | ${i.message.replace(/\|/g, '/')} |\n`; });
md += `\n## Source traceability (example: Coinbase)\n\n| Field | Tab | Cell | Label | Raw |\n|---|---|---|---|---|\n`;
const cb = recs.find(r => r.identity.symbol === 'COIN');
Object.entries(cb.sources).forEach(([k, s]) => { md += `| ${core.FIELD_LABELS[k] || k} | ${s.sheet} | ${s.cell} | ${s.label} | ${s.raw} |\n`; });
md += `\nEvery value the engine writes is logged like this on the Sheet's **Audit Trail** tab.\n`;
fs.writeFileSync('docs/VALIDATION_REPORT.md', md);
console.log('wrote docs/VALIDATION_REPORT.md', cnt('ERROR'), cnt('REVIEW'), cnt('INFO'));
