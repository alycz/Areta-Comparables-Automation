// Loads JSON fixtures produced by tools/xls_to_fixture.py, reviving dates.
const fs = require('fs');
const path = require('path');
const DIR = path.join(__dirname, 'fixtures');
function revive(v) {
  if (v && typeof v === 'object' && v.$date) return new Date(v.$date + 'T00:00:00Z');
  return v;
}
function loadAll() {
  return fs.readdirSync(DIR).filter(f => f.endsWith('.json')).sort().map(f => {
    const d = JSON.parse(fs.readFileSync(path.join(DIR, f), 'utf8'));
    const sheets = {};
    for (const [k, rows] of Object.entries(d.sheets)) sheets[k] = rows.map(r => r.map(revive));
    return { fileName: d.fileName, sheets };
  });
}
module.exports = { loadAll };
