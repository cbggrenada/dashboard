// Reads every workbook in data/raw and writes the website to _site/
// (index.html + data/dashboard-data.json). GitHub Actions runs this on every push.
//
//   node scripts/build-data.js
//
const fs = require('fs');
const path = require('path');
const P = require('./parsers');
const { execFileSync } = require('child_process');

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

// When each file was last uploaded (its last commit), so the newest upload wins. Falls back to the file's date.
function uploadedAt(file) {
  try {
    const t = execFileSync('git', ['log', '-1', '--format=%ct', '--', path.relative(ROOT, file)], { cwd: ROOT, encoding: 'utf8' }).trim();
    if (t) return +t * 1000;
  } catch (e) { /* not a git checkout */ }
  return fs.statSync(file).mtimeMs;
}

// KPI names in the Ops KPIs MD report -> the dashboard's keys
const MD_KEYS = [[/^water/, 'water'], [/^fuel/, 'fuel'], [/^electric/, 'elec'], [/^c[o0]2/, 'co2'], [/^extract recovery/, 'er'], [/^brews per day/, 'bpd'],
  [/^process loss/, 'loss'], [/^oee ?%?$/, 'oee'], [/^oee util/, 'oeeUtil'], [/^total bottling loss/, 'bottLoss'], [/^bottling eff/, 'bottEff'], [/^bottling util/, 'bottUtil'],
  [/^(case per hr|production cases per h)/, 'cph'], [/^production cases$/, 'cases'], [/^plant avail/, 'avail'], [/^maintenance compl/, 'pm'], [/^ftr/, 'ftr']];
const mdKey = l => { const t = String(l).toLowerCase().replace(/\s+/g, ' ').trim(); const f = MD_KEYS.find(([re]) => re.test(t)); return f ? f[1] : null; };

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
// 1) read every file
const files = listFiles(RAW);
const parsed = [];
for (const file of files) {
  const name = path.relative(RAW, file);
  try {
    const wb = P.XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: true });
    const res = P.parseKnown(wb, path.basename(file));
    if (res) {
      const days = Object.keys(res.rows).sort(), mids = Object.keys(res.monthly || {}).sort();
      parsed.push({ name, res, days, mids, time: uploadedAt(file), last: days[days.length - 1] || (mids.length ? mids[mids.length - 1] + '-31' : '') });
      continue;
    }
    // Other files: only PM / maintenance trackers laid out one row per date are matched by column name
    if (!/pm|maint/i.test(path.basename(file))) {
      report.push(`SKIP  ${name}: not a file type the dashboard reads`);
      continue;
    }
    let best = null;
    for (const sn of wb.SheetNames) {
      const sh = P.analyzeSheet(sn, wb.Sheets[sn]);
      if (sh && sh.map.date != null && (!best || sh.score > best.score)) best = sh;
    }
    const pmCols = best ? ['pmPlanned', 'pmDone', 'pmPct'].filter(k => best.map[k] != null) : [];
    if (!best || !pmCols.length) {
      report.push(`WARN  ${name}: could not find PM compliance figures in this file`);
      problems++;
      continue;
    }
    const rows = P.buildGeneric(best), days = Object.keys(rows).sort();
    parsed.push({ name, res: { type: 'PM data', rows, targets: {} }, days, mids: [], time: uploadedAt(file), last: days[days.length - 1] || '' });
  } catch (e) {
    report.push(`ERROR ${name}: ${e.message}`);
    problems++;
  }
}

// 2) merge. Files of the same type are applied oldest data first, so when two files cover the
//    same month (for example a renamed copy of a report) the one with the newest data wins.
const rank = t => ['Gross Efficiency', 'Utilities Tracking', 'Daily Process Report', 'Monthly process loss (volume)'].indexOf(t);
parsed.sort((x, y) => rank(x.res.type) - rank(y.res.type) || x.last.localeCompare(y.last) || x.time - y.time || x.name.localeCompare(y.name));
// Month-by-month KPI workbooks (FTR, OEE) are kept up to date in one file for the whole year,
// so only the most recently uploaded copy for each year is read. An older copy (for example the same
// workbook saved under a slightly different name) is skipped, because its later months are often
// unfinished or placeholder figures.
const KPI_TYPES = ['FTR summary', 'OEE summary'];
// PM compliance files hold one sheet per month, so several can be uploaded side by side (for example the
// maintenance team's workbook and a file of earlier months); each month is taken from the latest file covering it.
const yearsOf = x => [...new Set(x.mids.map(m => m.slice(0, 4)))];
for (const x of parsed.filter(x => KPI_TYPES.includes(x.res.type) && x.mids.length && !x.days.length)) {
  const newer = parsed.find(y => y !== x && y.res.type === x.res.type && y.mids.length && !y.days.length && y.time > x.time && yearsOf(y).some(v => yearsOf(x).includes(v)));
  if (newer) { x.skip = true; report.push(`SKIP  ${x.name}: older copy of ${newer.name} (uploaded ${new Date(x.time).toISOString().slice(0, 10)}); only the newest upload is read. Delete it from data/raw to tidy up.`); }
}
for (const { name, res, days, mids, skip } of parsed) {
  if (skip) continue;
  mergeRows(res.rows);
  for (const [id, v] of Object.entries(res.monthly || {})) {
    const m = months[id] || (months[id] = { rows: {}, targets: {} });
    if (v.vol) m.vol = v.vol;
    if (v.kpi) m.kpi = Object.assign(m.kpi || {}, v.kpi);
    if (v.targets) Object.assign(m.targets, v.targets);
  }
  if (res.md) {   // Ops KPIs MD report: the official month-to-date and year-to-date figures, a whole month at a time
    const byM = {};
    for (const r of res.md) { const k = mdKey(r.label); if (!k) continue; const id = `${r.y}-${String(r.m).padStart(2, '0')}`; (byM[id] = byM[id] || {})[k] = { a: r.mtd.a, b: r.mtd.b, ly: r.mtd.ly, ya: r.ytd.a, yb: r.ytd.b, yly: r.ytd.ly }; }
    for (const [id, md] of Object.entries(byM)) { const m = months[id] || (months[id] = { rows: {}, targets: {} }); m.md = md; }
    res.mdMonths = Object.keys(byM).sort();
  }
  if (res.type === 'Daily Process Report') {
    for (const ds of days) mergeTargets(ds.slice(0, 7), Object.fromEntries(Object.entries({
      loss: res.rows[ds].plTgt ?? res.targets.loss, brew: res.targets.brew, bpd: res.rows[ds].bpdTgt ?? res.targets.bpd,
    }).filter(([, v]) => v != null)));
  } else {
    mergeTargets(res.month || (days.length ? days[days.length - 1].slice(0, 7) : null), res.targets);
  }
  if (res.mdMonths) { report.push(`OK    ${name}: ${res.type}, ${res.mdMonths.length} months (${res.mdMonths.join(', ')}); its figures replace the calculated ones for those months`); continue; }
  report.push(`OK    ${name}: ${res.type}, ` + (mids.length && !days.length
    ? `${mids.length} month${mids.length > 1 ? 's' : ''} (${mids[0]} to ${mids[mids.length - 1]})`
    : `${days.length} days${days.length ? ` (${days[0]} to ${days[days.length - 1]})` : ''}`));
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
  if (!Object.keys(m.rows).length && !m.kpi && !m.vol && !m.md) delete months[id];
}

// PM compliance year to date: the average of the monthly figures from January, as the SCTCM reports calculate it,
// worked out across all PM files (a single file may only hold the last few months)
for (const id of Object.keys(months)) { const md = months[id].md; if (md && md.pm && md.pm.a != null) (months[id].kpi = months[id].kpi || {}).pmMtd = md.pm.a; }
{
  const byYear = {};
  for (const id of Object.keys(months).sort()) {
    const k = months[id].kpi; if (!k || k.pmMtd == null) continue;
    const s = byYear[id.slice(0, 4)] || (byYear[id.slice(0, 4)] = { t: 0, n: 0 });
    s.t += k.pmMtd; s.n++; k.pmYtd = +(s.t / s.n).toFixed(3);
  }
}

// Ops KPIs MD report: its OEE, FTR and PM compliance and its budgets replace those from the other files for the months it covers
for (const id of Object.keys(months)) {
  const md = months[id].md; if (!md) continue;
  const k = months[id].kpi = months[id].kpi || {}, t = months[id].targets;
  for (const x of ['oee', 'ftr', 'pm']) if (md[x]) { if (md[x].a != null) k[x + 'Mtd'] = md[x].a; if (md[x].ya != null) k[x + 'Ytd'] = md[x].ya; }
  for (const x of ['cases', 'water', 'fuel', 'elec', 'co2', 'loss', 'bpd', 'oee', 'ftr', 'pm', 'cph']) if (md[x] && md[x].b != null) t[x] = md[x].b;
  if (md.er && md.er.b != null) t.brew = +(100 - md.er.b).toFixed(2);
}

// Say so when the latest month with daily data has no FTR / OEE figure yet (the cell is blank or shows an error such as #DIV/0!)
for (const [key, label, type] of [['ftrMtd', 'FTR', 'FTR summary'], ['oeeMtd', 'OEE', 'OEE summary']]) {
  for (const id of Object.values(lastMonth)) {
    const src = parsed.filter(x => x.res.type === type && !x.skip && x.mids.some(m => m.slice(0, 4) === id.slice(0, 4)));
    if (!src.length) continue;
    const k = months[id] && months[id].kpi;
    if (!k || k[key] == null) {
      const prev = Object.keys(months).filter(m => m < id && months[m].kpi && months[m].kpi[key] != null).sort().pop();
      report.push(`NOTE  ${src.map(x => x.name).join(', ')}: no ${label} figure for ${id} yet (the month's overall ${label} cell is empty or shows an error), so the dashboard shows ${label} up to ${prev || 'the last month that has one'}.`);
    }
  }
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
// Also keep a copy of the data in the repository (data/dashboard-data.*), which the page loads when hosted.
fs.writeFileSync(path.join(ROOT, 'data', 'dashboard-data.json'), JSON.stringify(out));
fs.writeFileSync(path.join(ROOT, 'data', 'build-report.txt'), report.join('\n') + '\n');
fs.writeFileSync(path.join(ROOT, 'data', 'dashboard-data.js'), js);
// index.html in the repository is left alone (only the published copy in _site gets the data built in),
// so uploading a new index.html never clashes with the build.

console.log(report.join('\n'));
console.log(`\n${files.length} files read, ${dayCount} days of data, ${Object.keys(months).length} months.`);
if (problems) console.log(`${problems} file(s) had problems - see the lines above.`);
if (!dayCount && files.length) { console.error('No usable data found.'); process.exit(1); }
