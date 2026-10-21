/**
 * The one-page status brief — every figure on the page is calculated here.
 *
 * ENGINE ONLY. Nothing in this file calls, or knows about, an AI model. The AI
 * (added later) receives the finished result of buildBrief() and writes the
 * conclusion; it never recalculates a figure.
 *
 * It reuses parser.js, engine.js and report.js unchanged, so Lite and the full
 * app cannot disagree about a count, a forecast or a rating.
 */

import { STATUS } from './parser.js';
import { analyseSite, compareSites, lastReportedWeek, statusAt } from './engine.js';
import { computeSiteReport, RED, AMBER, GREEN } from './report.js';

/** Pace bands for the overall status word. Same cut-offs as the full app's SPI proxy. */
export const PACE_BANDS = { good: 0.95, decent: 0.85 };

const r1 = (n) => (n == null || Number.isNaN(n) ? null : Math.round(n * 10) / 10);
const r2 = (n) => (n == null || Number.isNaN(n) ? null : Math.round(n * 100) / 100);

function isoDate(d) {
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

/* ------------------------------------------------------------------ */
/* Task groups                                                         */
/* ------------------------------------------------------------------ */

/**
 * Count one set of tasks at a given week. "Pending" is everything that is not
 * finished and not in progress: not started + blocked + waiting on. Blocked and
 * waiting are also kept separately because they need different action.
 */
function tally(id, name, tasks, weekIndex) {
  const live = tasks.filter((t) => statusAt(t, weekIndex) !== STATUS.NA);
  const at = (t) => statusAt(t, weekIndex);
  const done = live.filter((t) => at(t) === STATUS.FINISHED).length;
  const wip = live.filter((t) => at(t) === STATUS.WIP).length;
  const blocked = live.filter((t) => at(t) === STATUS.BLOCKED).length;
  const waiting = live.filter((t) => at(t) === STATUS.WAITING).length;
  const notStarted = live.filter((t) => at(t) === STATUS.NOT_STARTED || at(t) == null).length;
  return {
    id, name,
    live: live.length, done, wip, blocked, waiting, notStarted,
    pending: live.length - done - wip,
    complete: live.length > 0 && done === live.length,
  };
}

function groupRows(site, a, weekIndex) {
  const tasks = site.detail.tasks;
  const known = new Set(a.categories.map((c) => c.id));
  const rows = a.categories.map((c) =>
    tally(c.id, c.name, tasks.filter((t) => t.categoryId === c.id), weekIndex));
  // A task whose group is missing or unrecognised must still be counted, or the
  // group rows would silently disagree with the site totals.
  const orphans = tasks.filter((t) => !known.has(t.categoryId));
  if (orphans.length) rows.push(tally('(none)', '(no task group)', orphans, weekIndex));
  return rows.filter((r) => r.live > 0);
}

/* ------------------------------------------------------------------ */
/* Movement since the previous week                                    */
/* ------------------------------------------------------------------ */

function movementSentence(m) {
  if (!m) return null;
  if (!m.comparable) return `No comparison with the previous week: ${m.reason}.`;
  const where = m.source === 'uploaded' ? 'the uploaded previous file' : 'the previous week in this workbook';
  const ch = m.percentChange;
  const pts = ch === 0 ? 'no change in completion' : `${ch > 0 ? '+' : ''}${ch} points complete`;
  const fin = m.finishedChange === 0
    ? 'no change in finished tasks'
    : `${Math.abs(m.finishedChange)} ${m.finishedChange > 0 ? 'more' : 'fewer'} task${Math.abs(m.finishedChange) === 1 ? '' : 's'} finished`;
  const stuck = m.stuckChange === 0
    ? `blocked or waiting unchanged at ${m.stuckNow}`
    : `${Math.abs(m.stuckChange)} ${m.stuckChange > 0 ? 'more' : 'fewer'} blocked or waiting (now ${m.stuckNow})`;
  return `Against ${where} (${m.previousWeek} to ${m.currentWeek}): ${pts} (${m.pctThen}% to ${m.pctNow}%), ${fin}; ${stuck}.`;
}

function movementFor(site, template, curA, previousModel) {
  let prevA = null;
  let source = null;
  if (previousModel) {
    const prevSite = (previousModel.sites || []).find((s) => s.code === site.code);
    if (!prevSite) return { source: 'uploaded', comparable: false, reason: 'this site is not in the previous file' };
    prevA = analyseSite(prevSite, previousModel.template);
    source = 'uploaded';
  } else {
    const idx = lastReportedWeek(site.detail);
    if (idx < 1) return { source: 'same-workbook', comparable: false, reason: 'the workbook has only one week of status' };
    prevA = analyseSite(site, template, { weekIndex: idx - 1 });
    source = 'same-workbook';
  }
  const c = compareSites(prevA, curA);
  if (!c) return null;
  if (!c.comparable) return { source, comparable: false, reason: c.reason };
  if (source === 'uploaded' && c.previousReportingWeek === c.currentReportingWeek) {
    return {
      source, comparable: false, sameWeek: true,
      reason: `the previous file reports the same week (${c.currentReportingWeek}) as this one`,
    };
  }
  return {
    source, comparable: true,
    previousWeek: c.previousReportingWeek, currentWeek: c.currentReportingWeek,
    pctThen: c.percentByWeightThen, pctNow: c.percentByWeightNow, percentChange: c.percentChange,
    finishedThen: c.finishedThen, finishedNow: c.finishedNow, finishedChange: c.finishedChange,
    stuckThen: c.stuckThen, stuckNow: c.stuckNow, stuckChange: c.stuckChange,
    newlyStuck: c.newlyStuck, cleared: c.cleared,
    forecastSlipChangeWeeks: c.slipChange,
  };
}

/* ------------------------------------------------------------------ */
/* One site                                                            */
/* ------------------------------------------------------------------ */

function siteBrief(site, template, { previousModel, reportDate }) {
  const a = analyseSite(site, template);
  const base = {
    code: site.code, description: site.description || '', priority: site.priority || '',
    wave: site.wave || '', target: site.target || null,
  };
  if (a.noData) {
    return {
      ...base, noData: true,
      reason: !site.detail
        ? 'In the register but no detail sheet was found, so nothing can be measured.'
        : 'The detail sheet has no weekly status filled in yet.',
    };
  }

  const rep = computeSiteReport(site, a, template, { reportDate });
  const weekIndex = lastReportedWeek(site.detail);
  const groups = groupRows(site, a, weekIndex);
  const groupsComplete = groups.filter((g) => g.complete).length;
  const m = rep.metrics;

  const movement = movementFor(site, template, a, previousModel);
  if (movement) movement.text = movementSentence(movement);

  return {
    ...base,
    noData: false,
    reportingWeek: a.reportingWeek?.label || null,
    statusDate: m.statusDate,
    submitted: rep.submitted,
    overall: rep.overall,
    scheduleRag: rep.dimensions.schedule.rag,
    counts: {
      total: a.taskCount, na: a.naCount, live: a.liveCount,
      finished: a.finishedCount, wip: a.wipCount, blocked: a.blockedCount,
      waiting: a.waitingCount, notStarted: a.notStartedCount,
    },
    groups: { total: groups.length, complete: groupsComplete, rows: groups },
    pctByCount: r1(m.completionCount),
    pctByWeight: r1(m.completionWeighted),
    doneWeight: a.doneWeight,
    totalWeight: a.totalWeight,
    timeElapsed: r1(m.timeElapsed),
    spiProxy: r2(m.spiProxy),
    schedule: {
      target: site.target || null,
      forecastDate: m.forecastDate || null,
      forecastMethod: m.forecastMethod || null,
      varianceDays: m.scheduleVariance ?? null,
    },
    // "Pending" means the site's own prerequisite table says it is outstanding.
    // Template defaults the site never recorded have an UNKNOWN status; they are
    // counted separately so they neither crowd the page nor pass as confirmed gaps.
    prerequisites: {
      pending: rep.constraints.open.filter((p) => !p.notTracked).map((p) => ({
        id: p.id, name: p.name, provider: p.provider, byWeek: p.byWeek,
        status: p.status, neededFor: p.neededFor,
      })),
      notTracked: rep.constraints.open.filter((p) => p.notTracked).map((p) => ({
        id: p.id, name: p.name, provider: p.provider,
      })),
    },
    blocked: a.stuckDetail.map((s) => ({
      taskId: s.taskId, name: s.name, group: s.category, status: s.status,
      resource: s.resource, weeksStuck: s.weeksStuck, sinceWeek: s.sinceWeek,
      reason: s.reason, waitingOn: s.waitingOn, expected: s.expected, logged: s.logged,
    })),
    resources: {
      maxWip: rep.resources.maxWip,
      rows: rep.resources.rows.map((r) => ({ resource: r.resource, open: r.open, wip: r.wip, stuck: r.stuck })),
    },
    movement,
    // Where the privacy scan found something — never the text itself.
    privacy: { flagged: !!rep.quality.privacy, locations: rep.quality.privacyLocations || [] },
    dataFindings: rep.quality.findings.length,
  };
}

/* ------------------------------------------------------------------ */
/* Overall pace                                                        */
/* ------------------------------------------------------------------ */

/**
 * Overall status word.
 *
 *   pace = weight finished so far / weight that should be finished by now
 *
 * "Should be finished by now" assumes straight-line progress between each
 * site's start and target dates, so for each site it is (total weight) x
 * (share of time elapsed). Summed over sites, so a large site counts for more
 * than a small one. It is the full app's SPI proxy, pooled across sites.
 *
 * Only sites that are still active and measurable count. Submitted sites have
 * no time left to be behind against, and sites without a start/target date or
 * weekly status cannot be measured. The result says how many were left out.
 */
export function overallPace(sites) {
  const withData = sites.filter((s) => !s.noData);
  const measurable = withData.filter((s) => !s.submitted && s.timeElapsed > 0 && s.totalWeight > 0);
  const expected = measurable.reduce((t, s) => t + s.totalWeight * (s.timeElapsed / 100), 0);
  const done = measurable.reduce((t, s) => t + s.doneWeight, 0);
  const skipped = sites.length - measurable.length;

  let word, rag, ratio = null;
  if (!measurable.length) {
    const allSubmitted = withData.length > 0 && withData.every((s) => s.submitted);
    word = allSubmitted ? 'Complete' : 'Not measurable';
    rag = allSubmitted ? GREEN : null;
  } else {
    ratio = expected > 0 ? done / expected : null;
    if (ratio == null) { word = 'Not measurable'; rag = null; }
    else if (ratio >= PACE_BANDS.good) { word = 'Good'; rag = GREEN; }
    else if (ratio >= PACE_BANDS.decent) { word = 'Decent'; rag = AMBER; }
    else { word = 'Poor'; rag = RED; }
  }

  const basis = ratio == null
    ? (word === 'Complete' ? 'Every tracked site has been submitted.' : 'No site has both dates and weekly status, so pace cannot be measured.')
    : `Finished weight is ${r2(ratio)} of what straight-line progress to each target date would expect by now `
      + `(Good at ${PACE_BANDS.good} or above, Decent at ${PACE_BANDS.decent} to just under ${PACE_BANDS.good}, Poor below ${PACE_BANDS.decent}).`
      + (skipped ? ` ${skipped} site${skipped === 1 ? ' is' : 's are'} not included (submitted, or not measurable).` : '');

  return { word, rag, ratio: r2(ratio), doneWeight: r1(done), expectedWeight: r1(expected), measurableSites: measurable.length, basis };
}

/* ------------------------------------------------------------------ */
/* The whole brief                                                     */
/* ------------------------------------------------------------------ */

/**
 * @param {object} model        parseWorkbook() output
 * @param {object} opts
 * @param {string[]} opts.codes  site codes to include, in any order
 * @param {object}  [opts.previousModel]  parseWorkbook() output of last week's file
 * @param {Date}    [opts.now]  run time (injectable so tests are deterministic)
 * @param {string}  [opts.workbookName]
 */
export function buildBrief(model, opts = {}) {
  const now = opts.now || new Date();
  const wanted = new Set(opts.codes || (model.sites || []).map((s) => s.code));
  const known = new Set((model.sites || []).map((s) => s.code));
  const warnings = [...wanted].filter((c) => !known.has(c)).map((c) => `Site ${c} is not in this workbook and was skipped.`);

  // Register order, not click order, so two runs of the same selection match.
  const chosen = (model.sites || []).filter((s) => wanted.has(s.code));
  const reportDate = opts.reportDate || isoDate(now);
  const sites = chosen.map((s) => siteBrief(s, model.template, { previousModel: opts.previousModel || null, reportDate }));

  const live = sites.filter((s) => !s.noData);
  const sum = (f) => live.reduce((t, s) => t + f(s), 0);
  const totalWeight = sum((s) => s.totalWeight);
  const doneWeight = sum((s) => s.doneWeight);

  const flagged = sites.filter((s) => s.privacy?.flagged).map((s) => s.code);

  return {
    kind: 'bim-tracker-lite.brief',
    version: 1,
    title: {
      runAt: now.toISOString(),
      reportDate,
      workbook: opts.workbookName || null,
      sites: sites.map((s) => s.code),
      siteCount: sites.length,
    },
    sites,
    portfolio: {
      sitesWithData: live.length,
      sitesWithoutData: sites.length - live.length,
      totals: {
        tasks: sum((s) => s.counts.live),
        finished: sum((s) => s.counts.finished),
        wip: sum((s) => s.counts.wip),
        blocked: sum((s) => s.counts.blocked),
        waiting: sum((s) => s.counts.waiting),
        notStarted: sum((s) => s.counts.notStarted),
        groupsComplete: sum((s) => s.groups.complete),
        groupsTotal: sum((s) => s.groups.total),
      },
      overallPctByWeight: totalWeight > 0 ? r1((doneWeight / totalWeight) * 100) : null,
      pace: overallPace(sites),
      chart: live.map((s) => ({
        code: s.code, pctByWeight: s.pctByWeight, timeElapsed: s.timeElapsed, rag: s.overall, submitted: s.submitted,
      })),
    },
    previousSource: opts.previousModel ? 'uploaded' : 'same-workbook',
    privacy: { flagged: flagged.length > 0, sites: flagged },
    warnings,
  };
}

/* ------------------------------------------------------------------ */
/* Fallback conclusion (no AI)                                         */
/* ------------------------------------------------------------------ */

/**
 * A plain, rule-based summary used when the AI is unavailable or has not been
 * run. Labelled as the engine's in the UI. It states facts; it does not advise.
 */
export function engineSummary(brief) {
  const lines = [];
  const p = brief.portfolio.pace;
  lines.push(p.ratio != null
    ? `Overall pace is ${p.word.toLowerCase()} (${p.ratio} of the expected progress to date).`
    : `Overall pace: ${p.word.toLowerCase()}.`);

  for (const s of brief.sites) {
    if (s.noData) { lines.push(`${s.code}: no progress data (${s.reason})`); continue; }
    const bits = [];
    if (s.overall !== GREEN) bits.push(`rated ${s.overall.toLowerCase()}`);
    if (s.schedule.varianceDays > 0 && s.schedule.forecastDate) {
      bits.push(`forecast ${s.schedule.forecastDate} is ${s.schedule.varianceDays} days after the ${s.schedule.target} target`);
    }
    const stuck = s.counts.blocked + s.counts.waiting;
    if (stuck) bits.push(`${stuck} task${stuck === 1 ? '' : 's'} blocked or waiting`);
    if (s.prerequisites.pending.length) bits.push(`${s.prerequisites.pending.length} prerequisite${s.prerequisites.pending.length === 1 ? '' : 's'} pending`);
    if (bits.length) lines.push(`${s.code}: ${bits.join('; ')}.`);
  }
  return lines;
}
