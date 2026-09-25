// Reads every workbook in data/raw and writes the website to _site/
// (index.html + data/dashboard-data.json). GitHub Actions runs this on every push.
//
//   node scripts/build-data.js
//
const fs = require('fs');
const path = require('path');
const P = require('./parsers');

const ROOT = path.join(__dirname, '..');
const RAW = path.join(ROOT, 'data', 'raw');
const OUT = path.join(ROOT, '_site');

function listFiles(dir) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap(e => {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) return listFiles(p);
    return /\.(xlsx|xlsm|xls|csv)$/i.test(e.name) && !e.name.startsWith('~$') ? [p] : [];
  }).sort();
}

const months = {};      // "2026-09" -> { rows: { "2026-09-01": {...} }, targets: {...} }
const report = [];
let problems = 0;

function mergeRows(rows) {
  for (const ds of Object.keys(rows)) {
    const id = ds.slice(0, 7);
    const m = months[id] || (months[id] = { rows: {}, targets: {} });
    m.rows[ds] = Object.assign(m.rows[ds] || {}, rows[ds]);
  }
}
function mergeTargets(id, t) {
  if (!id || !t || !Object.keys(t).length) return;
  const m = months[id] || (months[id] = { rows: {}, targets: {} });
  Object.assign(m.targets, t);
}

// Gross Efficiency first (whole year), then the monthly files, then anything else (e.g. PM)
const order = f => /gross/i.test(f) ? 0 : /utilit/i.test(f) ? 1 : /process/i.test(f) ? 2 : 3;
const files = listFiles(RAW).sort((a, b) => order(path.basename(a)) - order(path.basename(b)) || a.localeCompare(b));

for (const file of files) {
  const name = path.relative(RAW, file);
  try {
    const wb = P.XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: true });
    const res = P.parseKnown(wb, path.basename(file));
    if (res) {
      mergeRows(res.rows);
      for (const [id, v] of Object.entries(res.monthly || {})) {
        const m = months[id] || (months[id] = { rows: {}, targets: {} });
        if (v.vol) m.vol = v.vol;
        if (v.kpi) m.kpi = Object.assign(m.kpi || {}, v.kpi);
        if (v.targets) Object.assign(m.targets, v.targets);
      }
      const days = Object.keys(res.rows).sort();
      if (res.type === 'Daily Process Report') {
        for (const ds of days) mergeTargets(ds.slice(0, 7), Object.fromEntries(Object.entries({ loss: res.rows[ds].plTgt ?? res.targets.loss, brew: res.targets.brew }).filter(([, v]) => v != null)));
      } else {
        mergeTargets(res.month || (days.length ? days[days.length - 1].slice(0, 7) : null), res.targets);
      }
      const mids = Object.keys(res.monthly || {}).sort();
      report.push(`OK    ${name}: ${res.type}, ` + (mids.length && !days.length
        ? `${mids.length} month${mids.length > 1 ? 's' : ''} (${mids[0]} to ${mids[mids.length - 1]})`
        : `${days.length} days${days.length ? ` (${days[0]} to ${days[days.length - 1]})` : ''}`));
      continue;
    }
    // Other files: only PM / maintenance trackers are matched by column name
    if (!/pm|maint/i.test(path.basename(file))) {
      report.push(`SKIP  ${name}: not a Gross Efficiency, Utilities Tracking, Daily Process Report or PM file`);
      continue;
    }
    let best = null;
    for (const sn of wb.SheetNames) {
      const sh = P.analyzeSheet(sn, wb.Sheets[sn]);
      if (sh && sh.map.date != null && (!best || sh.score > best.score)) best = sh;
    }
    const pmCols = best ? ['pmPlanned', 'pmDone', 'pmPct'].filter(k => best.map[k] != null) : [];
    if (!best || !pmCols.length) {
      report.push(`WARN  ${name}: could not find a Date column plus PM planned / completed / compliance columns`);
      problems++;
      continue;
    }
    const rows = P.buildGeneric(best);
    mergeRows(rows);
    const days = Object.keys(rows).sort();
    report.push(`OK    ${name}: PM data from sheet "${best.name}" (${pmCols.map(k => best.headers[best.map[k]]).join(', ')}), ${days.length} days`);
  } catch (e) {
    report.push(`ERROR ${name}: ${e.message}`);
    problems++;
  }
}

// Optional target overrides kept in the repo (blank = use the targets found in the files)
let targets = {};
const tfile = path.join(ROOT, 'data', 'targets.json');
if (fs.existsSync(tfile)) {
  try { targets = JSON.parse(fs.readFileSync(tfile, 'utf8')); }
  catch (e) { report.push(`ERROR data/targets.json: ${e.message}`); problems++; }
}

// Monthly KPI tables (OEE, FTR) often carry placeholder values for months that haven't happened yet:
// keep only months up to the latest month that has daily data in that year.
const lastMonth = {};
for (const id of Object.keys(months)) if (Object.keys(months[id].rows).length) { const y = id.slice(0, 4); if (!lastMonth[y] || id > lastMonth[y]) lastMonth[y] = id; }
for (const id of Object.keys(months)) {
  const m = months[id], y = id.slice(0, 4);
  if (lastMonth[y] && id > lastMonth[y]) { delete m.kpi; delete m.vol; }
  if (!Object.keys(m.rows).length && !m.kpi && !m.vol) delete months[id];
}

const dayCount = Object.values(months).reduce((n, m) => n + Object.keys(m.rows).length, 0);
const out = {
  generatedAt: new Date().toISOString(),
  files: files.map(f => path.relative(RAW, f)),
  targets,
  months: Object.fromEntries(Object.keys(months).sort().map(k => [k, months[k]])),
};

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(path.join(OUT, 'data'), { recursive: true });
fs.copyFileSync(path.join(ROOT, 'index.html'), path.join(OUT, 'index.html'));
fs.writeFileSync(path.join(OUT, 'data', 'dashboard-data.json'), JSON.stringify(out));
fs.writeFileSync(path.join(OUT, 'data', 'build-report.txt'), report.join('\n') + '\n');
fs.writeFileSync(path.join(OUT, '.nojekyll'), '');

// Put a copy of the data inside index.html and in data/dashboard-data.js, so the page also works
// when it is opened straight from a computer (browsers block reading .json files from disk).
const json = JSON.stringify(out).replace(/</g, '\\u003c');
const js = `window.CBG_DATA=${json};\n`;
const embed = html => html.replace(/\/\*CBG_DATA_START\*\/[\s\S]*?\/\*CBG_DATA_END\*\//, () => `/*CBG_DATA_START*/window.CBG_DATA_EMBED=${json};/*CBG_DATA_END*/`);
const page = embed(fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8'));
fs.writeFileSync(path.join(OUT, 'index.html'), page);
fs.writeFileSync(path.join(OUT, 'data', 'dashboard-data.js'), js);
// Also keep a copy in the repository itself, so the site works whether GitHub Pages
// is set to "GitHub Actions" or to "Deploy from a branch".
fs.writeFileSync(path.join(ROOT, 'data', 'dashboard-data.json'), JSON.stringify(out));
fs.writeFileSync(path.join(ROOT, 'data', 'build-report.txt'), report.join('\n') + '\n');
fs.writeFileSync(path.join(ROOT, 'data', 'dashboard-data.js'), js);
fs.writeFileSync(path.join(ROOT, 'index.html'), page);

console.log(report.join('\n'));
console.log(`\n${files.length} files read, ${dayCount} days of data, ${Object.keys(months).length} months.`);
if (problems) console.log(`${problems} file(s) had problems - see the lines above.`);
if (!dayCount && files.length) { console.error('No usable data found.'); process.exit(1); }
