/**
 * Brief verification.
 *
 * Every figure on the one-page brief is checked against a fixture whose answer
 * is known by hand (tools/fixture.mjs, site F-01 at W06):
 *
 *   Finished 4   WIP 2   Blocked 1   Waiting on 1   Not started 2   N/A 1
 *   live = 10, total = 11.  Weight done 5 of 11 = 45.5%.
 *
 *   Group C010 "Initial setup": 2 live, 2 done            -> complete
 *   Group C020 "Modelling":     8 live, 2 done, 2 WIP,
 *                               1 blocked, 1 waiting, 2 not started
 *                               pending = 8 - 2 - 2 = 4
 *
 *   Time: start 2026-01-05, target 2026-03-01 (55 days); status date
 *   2026-02-15 (41 days in) -> 74.5% elapsed.
 *   Expected weight by now = 11 x 0.7455 = 8.2; done = 5; pace 5 / 8.2 = 0.61
 *   -> below 0.85 -> "Poor".
 *
 * Run: node tools/verify-brief.mjs
 */

import { createRequire } from 'node:module';
import path from 'node:path';
import fsSync from 'node:fs';
import { fileURLToPath } from 'node:url';
import { fixtureWorkbook } from './fixture.mjs';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const require = createRequire(import.meta.url);
const XLSX = require(path.join(root, 'assets/vendor/xlsx.full.min.js'));
global.window = { XLSX };

const { readWorkbook, parseWorkbook, STATUS } = await import('../js/core/parser.js');
const { buildBrief, overallPace, engineSummary, PACE_BANDS } = await import('../js/core/brief.js');

let pass = 0, fail = 0;
const failures = [];
const check = (name, cond, detail = '') => {
  if (cond) pass++;
  else { fail++; failures.push(`${name}${detail ? ` — ${detail}` : ''}`); }
};
const near = (a, b, tol = 0.01) => a != null && Math.abs(a - b) <= tol;

const load = async (wb, name = 'x.xlsx') => {
  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  return parseWorkbook(await readWorkbook({ name, arrayBuffer: async () => buf }));
};
const fixture = () => load(fixtureWorkbook(XLSX, STATUS), 'fixture.xlsx');
const NOW = new Date(2026, 1, 16, 9, 30);   // Mon 16 Feb 2026, local

/* ---------------------------- 1. site counts ---------------------------- */
{
  const model = await fixture();
  const b = buildBrief(model, { codes: ['F-01'], now: NOW, workbookName: 'fixture.xlsx' });
  const s = b.sites[0];
  const c = s.counts;
  check('1 site in the brief', b.sites.length === 1);
  check('total 11 tasks', c.total === 11, String(c.total));
  check('N/A 1 excluded', c.na === 1 && c.live === 10, JSON.stringify(c));
  check('finished 4', c.finished === 4);
  check('WIP 2', c.wip === 2);
  check('blocked 1', c.blocked === 1);
  check('waiting on 1 (kept separate from blocked)', c.waiting === 1);
  check('not started 2', c.notStarted === 2);
  check('statuses add up to live', c.finished + c.wip + c.blocked + c.waiting + c.notStarted === c.live);
  check('completion by weight 45.5', s.pctByWeight === 45.5, String(s.pctByWeight));
  check('completion by count 40', s.pctByCount === 40, String(s.pctByCount));
  check('count and weight differ (weights are not all 1)', s.pctByCount !== s.pctByWeight);
  check('reporting week is W06', s.reportingWeek === 'W06');
  check('status date is the week end', s.statusDate === '2026-02-15', s.statusDate);
  check('time elapsed 74.5', s.timeElapsed === 74.5, String(s.timeElapsed));
  check('schedule is RED at SPI 0.61', s.scheduleRag === 'RED', s.scheduleRag);
  check('title block carries the run time', b.title.runAt === NOW.toISOString());
  check('title block carries the run date', b.title.reportDate === '2026-02-16', b.title.reportDate);
  check('title block names the workbook', b.title.workbook === 'fixture.xlsx');
  check('title block lists the sites tracked', JSON.stringify(b.title.sites) === '["F-01"]');
}

/* ---------------------------- 2. task groups ---------------------------- */
{
  const model = await fixture();
  const s = buildBrief(model, { codes: ['F-01'], now: NOW }).sites[0];
  const g = s.groups;
  const c010 = g.rows.find((r) => r.id === 'C010');
  const c020 = g.rows.find((r) => r.id === 'C020');
  check('2 task groups', g.total === 2, String(g.total));
  check('1 task group completely finished', g.complete === 1, String(g.complete));
  check('C010 is complete', c010?.complete === true && c010.done === 2 && c010.live === 2);
  check('C020 live 8', c020?.live === 8, String(c020?.live));
  check('C020 done 2, WIP 2', c020?.done === 2 && c020?.wip === 2);
  check('C020 blocked 1, waiting 1, not started 2', c020?.blocked === 1 && c020?.waiting === 1 && c020?.notStarted === 2);
  check('C020 pending 4 (not started + blocked + waiting)', c020?.pending === 4, String(c020?.pending));
  check('C020 is not complete', c020?.complete === false);
  check('every group: done + WIP + pending = live', g.rows.every((r) => r.done + r.wip + r.pending === r.live));
  const sumOf = (k) => g.rows.reduce((t, r) => t + r[k], 0);
  check('group rows add up to the site totals',
    sumOf('live') === s.counts.live && sumOf('done') === s.counts.finished && sumOf('wip') === s.counts.wip
    && sumOf('blocked') === s.counts.blocked && sumOf('waiting') === s.counts.waiting && sumOf('notStarted') === s.counts.notStarted);

  // A task whose group is unrecognised must still be counted somewhere.
  const m2 = await fixture();
  m2.sites[0].detail.tasks.find((t) => t.id === 'T020-04').categoryId = 'C999';
  const s2 = buildBrief(m2, { codes: ['F-01'], now: NOW }).sites[0];
  const orphan = s2.groups.rows.find((r) => r.id === '(none)');
  check('a task with an unknown group appears under "(no task group)"', !!orphan && orphan.live === 1 && orphan.wip === 1);
  check('orphaned task still counted in the group totals',
    s2.groups.rows.reduce((t, r) => t + r.live, 0) === s2.counts.live);
}

/* ------------------------ 3. prerequisites and blocked ------------------------ */
{
  const model = await fixture();
  const s = buildBrief(model, { codes: ['F-01'], now: NOW }).sites[0];
  check('one prerequisite is pending (P02)', s.prerequisites.pending.length === 1 && s.prerequisites.pending[0].id === 'P02');
  check('received prerequisite P01 is not listed', !s.prerequisites.pending.some((p) => p.id === 'P01') && !s.prerequisites.notTracked.some((p) => p.id === 'P01'));
  check('pending prerequisite says who and for which tasks',
    s.prerequisites.pending[0].provider === 'R3' && s.prerequisites.pending[0].neededFor.join() === 'T020-05,T020-06');
  check('two stuck tasks listed', s.blocked.length === 2, String(s.blocked.length));
  const wait = s.blocked.find((x) => x.taskId === 'T020-06');
  const blk = s.blocked.find((x) => x.taskId === 'T020-05');
  check('waiting task carries its logged reason', wait?.reason === 'Ceiling void zones not confirmed', wait?.reason);
  check('waiting task carries who it waits on', wait?.waitingOn === 'R3');
  check('waiting task is flagged as logged', wait?.logged === true);
  check('blocked task with no log entry is flagged as unlogged', blk?.logged === false && blk?.reason == null);
  check('stuck status is kept (Blocked vs Waiting on)', blk?.status === STATUS.BLOCKED && wait?.status === STATUS.WAITING);
  check('weeks stuck is measured', wait?.weeksStuck >= 1 && blk?.weeksStuck >= 1);
}

/* ------------------------------- 4. overall pace ------------------------------- */
{
  const model = await fixture();
  const b = buildBrief(model, { now: NOW });                 // both register sites
  const p = b.portfolio.pace;
  check('overall word is Poor at 0.61', p.word === 'Poor', p.word);
  check('pace ratio 0.61', near(p.ratio, 0.61), String(p.ratio));
  check('expected weight 8.2, done 5', near(p.expectedWeight, 8.2, 0.05) && p.doneWeight === 5, JSON.stringify(p));
  check('pace is RED', p.rag === 'RED');
  check('only the measurable site counts', p.measurableSites === 1);
  check('the basis says a site was left out', /1 site is not included/.test(p.basis), p.basis);
  check('portfolio overall % by weight 45.5', b.portfolio.overallPctByWeight === 45.5);
  check('chart has one bar per site with data', b.portfolio.chart.length === 1 && b.portfolio.chart[0].code === 'F-01');
  check('site without a detail sheet is reported, not hidden',
    b.sites.some((s) => s.code === 'F-02' && s.noData === true && /no detail sheet/.test(s.reason)));
  check('site count includes the no-data site', b.title.siteCount === 2 && b.portfolio.sitesWithoutData === 1);

  // Band edges, using stand-in sites so the cut-offs themselves are tested.
  const mk = (done, tElapsed) => ({ noData: false, submitted: false, totalWeight: 100, doneWeight: done, timeElapsed: tElapsed });
  check('exactly 0.95 is Good', overallPace([mk(47.5, 50)]).word === 'Good');
  check('just under 0.95 is Decent', overallPace([mk(47.4, 50)]).word === 'Decent');
  check('exactly 0.85 is Decent', overallPace([mk(42.5, 50)]).word === 'Decent');
  check('just under 0.85 is Poor', overallPace([mk(42.4, 50)]).word === 'Poor');
  check('bands are the full app\'s', PACE_BANDS.good === 0.95 && PACE_BANDS.decent === 0.85);
  // Pooled across sites: a big late site outweighs a small early one.
  const big = { noData: false, submitted: false, totalWeight: 900, doneWeight: 300, timeElapsed: 50 };   // expects 450
  const small = { noData: false, submitted: false, totalWeight: 100, doneWeight: 100, timeElapsed: 50 }; // expects 50
  const pooled = overallPace([big, small]);
  check('pace is pooled by weight, not averaged per site', near(pooled.ratio, 0.8) && pooled.word === 'Poor', JSON.stringify(pooled));
  check('a submitted site is left out of pace', overallPace([{ ...small, submitted: true }, big]).measurableSites === 1);
  check('all sites submitted reads Complete', overallPace([{ ...small, submitted: true }]).word === 'Complete');
  check('nothing measurable is stated, not guessed', overallPace([{ noData: true }]).word === 'Not measurable');
}

/* ------------------------------ 5. previous week ------------------------------ */
{
  // (a) same workbook: W05 -> W06 in the fixture has no change, by construction.
  const model = await fixture();
  const s = buildBrief(model, { codes: ['F-01'], now: NOW }).sites[0];
  const m = s.movement;
  check('same-workbook comparison is used when no file is uploaded', m?.source === 'same-workbook' && m.comparable === true);
  check('compares W05 with W06', m?.previousWeek === 'W05' && m?.currentWeek === 'W06');
  check('no change fixture shows no change', m?.finishedChange === 0 && m?.percentChange === 0 && m?.stuckChange === 0);
  check('movement sentence is written by the engine', /no change in completion/.test(m?.text || ''), m?.text);

  // (b) an uploaded previous file where one task was not yet finished.
  const prev = await fixture();
  for (const t of prev.sites[0].detail.tasks) t.weekly[5] = null;                // last week = W05
  for (const c of prev.sites[0].detail.categories) c.weekly[5] = null;
  prev.sites[0].detail.tasks.find((t) => t.id === 'T020-02').weekly[4] = STATUS.NOT_STARTED;
  const cur = await fixture();
  const s2 = buildBrief(cur, { codes: ['F-01'], now: NOW, previousModel: prev }).sites[0];
  const m2 = s2.movement;
  check('uploaded file is used when given', m2?.source === 'uploaded' && m2.comparable === true, JSON.stringify(m2));
  check('one more task finished since the uploaded file', m2?.finishedChange === 1 && m2.finishedThen === 3 && m2.finishedNow === 4, JSON.stringify(m2));
  check('completion rose 36.4 -> 45.5', m2?.pctThen === 36.4 && m2?.pctNow === 45.5 && m2.percentChange === 9.1, JSON.stringify(m2));
  check('sentence names the weeks and the source', /uploaded previous file \(W05 to W06\)/.test(m2?.text || ''), m2?.text);
  check('sentence says one more task finished', /1 more task finished/.test(m2?.text || ''), m2?.text);
  check('brief records where the comparison came from', buildBrief(cur, { codes: ['F-01'], now: NOW, previousModel: prev }).previousSource === 'uploaded');

  // (c) uploading the same week again must not be presented as a comparison.
  const same = await fixture();
  const s3 = buildBrief(cur, { codes: ['F-01'], now: NOW, previousModel: same }).sites[0].movement;
  check('same-week upload is refused as a comparison', s3?.comparable === false && s3.sameWeek === true, JSON.stringify(s3));
  check('and says so', /same week/.test(s3?.text || ''), s3?.text);

  // (d) previous file lacks the site.
  const other = await fixture();
  other.sites[0].code = 'Z-99';
  const s4 = buildBrief(cur, { codes: ['F-01'], now: NOW, previousModel: other }).sites[0].movement;
  check('site missing from the previous file is said, not guessed', s4?.comparable === false && /not in the previous file/.test(s4.text || ''), s4?.text);
}

/* ------------------------------ 6. site selection ------------------------------ */
{
  const model = await fixture();
  const b = buildBrief(model, { codes: ['F-01', 'Z-99'], now: NOW });
  check('unknown site code is skipped with a warning', b.sites.length === 1 && b.warnings.length === 1 && /Z-99/.test(b.warnings[0]), JSON.stringify(b.warnings));
  const one = buildBrief(model, { codes: ['F-01'], now: NOW });
  check('selecting one site leaves the other out of every part of the brief', !JSON.stringify(one).includes('F-02'));
  const two = buildBrief(model, { codes: ['F-02', 'F-01'], now: NOW });
  check('sites come out in register order regardless of selection order', two.title.sites.join() === 'F-01,F-02', two.title.sites.join());
  const none = buildBrief(model, { codes: [], now: NOW });
  check('an empty selection gives an empty brief, not every site', none.sites.length === 0 && none.title.siteCount === 0);
}

/* ------------------------------ 7. privacy flag ------------------------------ */
{
  const clean = buildBrief(await fixture(), { codes: ['F-01'], now: NOW });
  check('clean workbook is not flagged', clean.privacy.flagged === false && clean.privacy.sites.length === 0);

  const dirty = await fixture();
  dirty.sites[0].detail.tasks.find((t) => t.id === 'T020-04').name = 'Walls for John Smith';
  const d = buildBrief(dirty, { codes: ['F-01'], now: NOW });
  check('a personal-name-like phrase in a task name is flagged', d.privacy.flagged === true && d.privacy.sites[0] === 'F-01', JSON.stringify(d.privacy));
  check('the flag says where', d.sites[0].privacy.locations.length > 0);
  check('the flag never repeats the text it found', !JSON.stringify(d.sites[0].privacy).includes('John Smith'));
}

/* ------------------------------ 8. fallback summary ------------------------------ */
{
  const model = await fixture();
  const b = buildBrief(model, { codes: ['F-01'], now: NOW });
  const lines = engineSummary(b);
  const text = lines.join(' ');
  check('fallback states the pace', /poor/.test(lines[0]) && /0\.61/.test(lines[0]), lines[0]);
  check('fallback names the site and its forecast gap', /F-01/.test(text) && /days after the 2026-03-01 target/.test(text), text);
  check('fallback counts stuck tasks', /2 tasks blocked or waiting/.test(text), text);
  check('fallback counts pending prerequisites', /1 prerequisite pending/.test(text), text);
  check('fallback does not advise', !/should|recommend|reallocate/i.test(text), text);
  check('fallback of a one-site brief does not mention other sites', !/F-02/.test(text));
}

/* -------------------- 9. the shipped template must produce a brief -------------------- */
{
  const buf = fsSync.readFileSync(path.join(root, 'templates/BIM-Multi-Site-Tracker-TEMPLATE.xlsx'));
  const model = parseWorkbook(await readWorkbook({
    name: 'template.xlsx', arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
  }));
  const b = buildBrief(model, { now: NOW });
  const withData = b.sites.filter((s) => !s.noData);
  check('template brief has at least one site with data', withData.length >= 1, String(withData.length));
  for (const s of withData) {
    const sumOf = (k) => s.groups.rows.reduce((t, r) => t + r[k], 0);
    check(`template ${s.code}: group rows add up to the site totals`,
      sumOf('live') === s.counts.live && sumOf('done') === s.counts.finished && sumOf('wip') === s.counts.wip
      && sumOf('blocked') === s.counts.blocked && sumOf('waiting') === s.counts.waiting && sumOf('notStarted') === s.counts.notStarted,
      JSON.stringify(s.counts));
    check(`template ${s.code}: statuses add up to live`,
      s.counts.finished + s.counts.wip + s.counts.blocked + s.counts.waiting + s.counts.notStarted === s.counts.live);
    check(`template ${s.code}: groups complete never exceeds groups`, s.groups.complete <= s.groups.total);
    check(`template ${s.code}: blocked list length equals blocked + waiting`, s.blocked.length === s.counts.blocked + s.counts.waiting);
    check(`template ${s.code}: movement uses the same workbook`, s.movement?.source === 'same-workbook');
  }
  for (const s of withData) {
    check(`template ${s.code}: pending prerequisites are only ones the site recorded as outstanding`,
      s.prerequisites.pending.every((p) => p.status !== 'Not tracked' && /outstanding|overdue|pending|requested/i.test(p.status || 'outstanding')),
      JSON.stringify(s.prerequisites.pending.map((p) => p.status)));
  }
  check('template A-01: untracked template prerequisites are counted separately, not listed as pending',
    withData[0].prerequisites.notTracked.length >= 1
    && withData[0].prerequisites.pending.length === 2
    && !withData[0].prerequisites.pending.some((p) => withData[0].prerequisites.notTracked.some((n) => n.id === p.id)),
    `pending ${withData[0].prerequisites.pending.length}, untracked ${withData[0].prerequisites.notTracked.length}`);
  const s0 = withData[0];
  check('template worked example: 16.4% -> 21.3% complete since last week',
    s0.movement?.comparable && s0.movement.pctThen === 16.4 && s0.movement.pctNow === 21.3, JSON.stringify(s0.movement));
  check('template worked example: 1 more task finished', s0.movement?.finishedChange === 1, String(s0.movement?.finishedChange));
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) {
  console.log('\nFAILURES:');
  for (const f of failures) console.log('  ✗ ' + f);
  process.exit(1);
}
console.log('\nBrief verified.');
