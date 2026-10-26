/**
 * Print the engine-only brief for a workbook, as plain text.
 *
 *   node tools/print-brief.mjs <workbook.xlsx> [--previous <last-week.xlsx>] [--sites A-01,B-02]
 *
 * No AI and no network. It runs the same buildBrief() the app uses, so it is a
 * way to check the figures against your own workbook before there is a browser
 * page, and a way to see exactly what the AI would be given.
 */
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
global.window = { XLSX: createRequire(import.meta.url)(path.join(root, 'assets/vendor/xlsx.full.min.js')) };
const { readWorkbook, parseWorkbook } = await import('../js/core/parser.js');
const { buildBrief, engineSummary } = await import('../js/core/brief.js');
const { dropMilestoneColumns, unregisteredSheets } = await import('../js/core/normalise.js');

const args = process.argv.slice(2);
const flag = (name) => { const i = args.indexOf(name); return i >= 0 ? args[i + 1] : null; };
const file = args.find((a, i) => !a.startsWith('--') && !(i > 0 && args[i - 1].startsWith('--')));
if (!file) {
  console.error('Usage: node tools/print-brief.mjs <workbook.xlsx> [--previous <file.xlsx>] [--sites A-01,B-02] [--json]');
  process.exit(2);
}

async function load(p) {
  const b = fs.readFileSync(p);
  return dropMilestoneColumns(parseWorkbook(await readWorkbook({
    name: path.basename(p), arrayBuffer: async () => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength),
  })));
}

const model = await load(file);
if (!model.sites?.length) {
  console.error('No sites were read from this workbook.');
  for (const p of model.problems || []) console.error(' - ' + (p.message || JSON.stringify(p)));
  process.exit(1);
}
const previousModel = flag('--previous') ? await load(flag('--previous')) : null;
const codes = flag('--sites') ? flag('--sites').split(',').map((s) => s.trim()).filter(Boolean) : null;
const b = buildBrief(model, { codes: codes || undefined, previousModel, workbookName: path.basename(file) });

if (args.includes('--json')) { console.log(JSON.stringify(b, null, 2)); process.exit(0); }

const pad = (v, n, right) => { const s = String(v ?? '-'); return right ? s.padStart(n) : s.padEnd(n); };
const line = (n = 78) => '-'.repeat(n);
const p = b.portfolio;

console.log('BIM TRACKER LITE - status brief   [calculated by the engine; no AI]');
console.log(`Run: ${b.title.runAt.replace('T', ' ').slice(0, 16)} UTC   Workbook: ${b.title.workbook}`);
console.log(`Sites tracked (${b.title.siteCount}): ${b.title.sites.join(', ')}`);
console.log(`Overall: ${p.pace.word.toUpperCase()}${p.pace.ratio != null ? ` (pace ${p.pace.ratio})` : ''}   Completion across sites: ${p.overallPctByWeight ?? '-'}%`);
console.log(p.pace.basis);
if (b.privacy.flagged) console.log(`\n!! PRIVACY: possible personal or place names found in: ${b.privacy.sites.join(', ')}. Check the workbook before sending anything to an AI.`);
for (const w of b.warnings) console.log('!! ' + w);
const unreg = unregisteredSheets(model);
if (unreg.length) console.log(`!! NOT READ: ${unreg.map((u) => u.name).join(', ')} hold tasks but no row in the Sites sheet lists them.`);
for (const w of b.dataWarnings) console.log(`!! ${w.severity === 'RED' ? '[WRONG SITE?] ' : ''}${w.site}: ${w.text}`);
const pc = b.portfolio.pace;
if (pc.excluded.length) console.log(`   Pace counts ${pc.sitesCounted} of ${pc.sitesTotal} sites. Not counted: ${pc.excluded.map((e) => `${e.code} (${e.reason})`).join(', ')}.`);

console.log('\nSITES\n' + line());
console.log(pad('Site', 8) + pad('Status', 7) + pad('Groups', 8) + pad('Done', 5, 1) + pad('WIP', 5, 1) + pad('Blk', 5, 1) + pad('Wait', 5, 1) + pad('NotSt', 6, 1) + pad('%', 6, 1) + '  Forecast vs target');
for (const s of b.sites) {
  if (s.noData) { console.log(pad(s.code, 8) + 'no data: ' + s.reason); continue; }
  const sc = s.schedule;
  const fc = sc.forecastDate ? `${sc.forecastDate} (${sc.varianceDays > 0 ? '+' : ''}${sc.varianceDays ?? '?'} d vs ${sc.target})` : 'not forecast';
  console.log(pad(s.code, 8) + pad(s.overall, 7) + pad(`${s.groups.complete}/${s.groups.total}`, 8)
    + pad(s.counts.finished, 5, 1) + pad(s.counts.wip, 5, 1) + pad(s.counts.blocked, 5, 1) + pad(s.counts.waiting, 5, 1)
    + pad(s.counts.notStarted, 6, 1) + pad(s.pctByWeight, 6, 1) + '  ' + fc);
}

for (const s of b.sites.filter((x) => !x.noData)) {
  console.log(`\nTASK GROUPS - ${s.code} (status as at ${s.reportingWeek}, ${s.statusDate})\n` + line());
  console.log(pad('Group', 32) + pad('Done', 6, 1) + pad('WIP', 5, 1) + pad('Pending', 9, 1) + '   (of which blocked/waiting)');
  for (const g of s.groups.rows) {
    console.log(pad(`${g.complete ? '*' : ' '} ${g.name}`.slice(0, 31), 32) + pad(g.done, 6, 1) + pad(g.wip, 5, 1) + pad(g.pending, 9, 1) + `   ${g.blocked + g.waiting || ''}`);
  }
  const pre = s.prerequisites;
  if (pre.pending.length || pre.notTracked.length) {
    console.log(`PREREQUISITES PENDING - ${s.code}`);
    for (const r of pre.pending) {
      console.log(`  ${r.id} ${r.name} - from ${r.provider || '?'}, needed by week ${r.byWeek ?? '?'}, for ${r.neededFor.join(', ') || '-'} [${r.status}]`);
    }
    if (pre.notTracked.length) {
      console.log(`  (${pre.notTracked.length} template prerequisites have no entry in this site's table, so their status is unknown: ${pre.notTracked.map((n) => n.id).join(', ')})`);
    }
  }
  if (s.blocked.length) {
    console.log(`BLOCKED / WAITING - ${s.code}`);
    for (const t of s.blocked) {
      console.log(`  ${t.taskId} ${t.name} [${t.status}] ${t.weeksStuck} wk since ${t.sinceWeek}; ${t.resource || 'no owner'}`
        + (t.reason ? ` - ${t.reason}` : ' - no log entry') + (t.waitingOn ? ` (waiting on ${t.waitingOn})` : ''));
    }
  }
  if (s.movement?.text) console.log(`SINCE LAST WEEK - ${s.code}: ${s.movement.text}`);
}

console.log('\nENGINE SUMMARY (used when the AI is not run)\n' + line());
for (const l of engineSummary(b)) console.log('  ' + l);
