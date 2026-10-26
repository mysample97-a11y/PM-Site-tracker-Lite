/**
 * Exports: PDF, Word and Excel of the one-page brief.
 *
 * Format choices, so the app carries no extra dependencies (same as the full app):
 *  - Excel  .xlsx through the vendored SheetJS.
 *  - Word   .rtf, plain text that Word opens natively with tables and formatting.
 *  - PDF    the browser's own print-to-PDF, driven by a hidden frame. It uses the
 *           exact HTML shown on screen, so what you reviewed is what prints.
 *
 * Word and Excel cannot hold the SVG chart, so they carry the same figures as a
 * "Progress by site" table. The PDF has the chart.
 *
 * Exports are built from the brief, which already holds only the sites that were
 * selected. They never walk the workbook, so an unselected site cannot appear.
 */

import { downloadBlob } from './util.js';
import { buildPrintHTML, fmtD, fmtRun } from './briefview.js';

const stamp = (brief) => brief.title.reportDate || new Date().toISOString().slice(0, 10);
const nv = (v) => (v == null ? '' : v);

/* ============================ shared content ============================ */

/** The conclusion as plain labelled lines, for Word and Excel. */
export function conclusionLines(brief, conclusion, engineLines) {
  const out = [];
  for (const s of brief.sites) {
    if (!s.noData && s.movement?.text) out.push({ label: 'Calculated', site: s.code, text: s.movement.text });
  }
  if (!conclusion) {
    for (const l of engineLines || []) out.push({ label: 'Calculated, no AI', site: '', text: l });
    out.push({ label: 'Note', site: '', text: 'An AI-written conclusion has not been generated for this brief.' });
    return out;
  }
  const r = conclusion.result;
  const tag = `AI-written (${conclusion.meta?.model || 'model'})`;
  if (r.headline) out.push({ label: tag, site: '', text: r.headline });
  for (const s of r.sites) {
    out.push({ label: tag, site: s.site, text: `${s.verdict}. ${s.reason}${s.action ? ` Action: ${s.action}` : ''}` });
  }
  if (r.reallocation) out.push({ label: tag, site: '', text: `Resources: ${r.reallocation}` });
  for (const i of r.issues || []) out.push({ label: 'Checked on receipt', site: '', text: i });
  out.push({ label: 'Note', site: '', text: 'Written by an AI from the calculated figures. A judgement, to be checked by a person.' });
  return out;
}

const sitesRows = (brief) => brief.sites.map((s) => {
  if (s.noData) return [s.code, 'No data', '', '', '', '', '', '', '', s.reason, ''];
  const sc = s.schedule;
  return [
    s.code, s.overall, `${s.groups.complete} of ${s.groups.total}`,
    s.counts.finished, s.counts.wip, s.counts.blocked, s.counts.waiting, s.counts.notStarted,
    s.pctByWeight,
    `Target ${fmtD(sc.target)}; ${sc.forecastDate ? `forecast ${fmtD(sc.forecastDate)} (${sc.varianceDays > 0 ? '+' : ''}${nv(sc.varianceDays)} d)` : 'no forecast'}`,
    s.movement?.comparable ? `${s.movement.percentChange > 0 ? '+' : ''}${s.movement.percentChange} pts` : '',
  ];
});
const SITES_HEAD = ['Site', 'Rating', 'Task groups finished', 'Done', 'WIP', 'Blocked', 'Waiting on', 'Not started', '% done', 'Target / forecast', 'Since last week'];

const progressRows = (brief) => brief.portfolio.chart.map((c) => [c.code, c.pctByWeight, c.submitted ? 'submitted' : nv(c.timeElapsed), c.rag || '']);
const PROGRESS_HEAD = ['Site', '% of work complete', '% of time used', 'Rating'];

const groupRows = (brief) => brief.sites.filter((s) => !s.noData).flatMap((s) =>
  s.groups.rows.map((g) => [s.code, g.name, g.done, g.wip, g.blocked, g.waiting, g.notStarted, g.pending, g.complete ? 'Yes' : '']));
const GROUP_HEAD = ['Site', 'Task group', 'Done', 'WIP', 'Blocked', 'Waiting on', 'Not started', 'Pending', 'Group finished'];

const preqRows = (brief) => brief.sites.filter((s) => !s.noData).flatMap((s) => [
  ...s.prerequisites.pending.map((p) => [s.code, p.id, p.name, nv(p.provider), nv(p.byWeek), nv(p.status), (p.neededFor || []).join(', ')]),
  ...s.prerequisites.notTracked.map((p) => [s.code, p.id, p.name, nv(p.provider), '', 'No record for this site (status unknown)', '']),
]);
const PREQ_HEAD = ['Site', 'ID', 'Prerequisite', 'From', 'Needed by week', 'Status', 'Needed for tasks'];

const blockedRows = (brief) => brief.sites.filter((s) => !s.noData).flatMap((s) => s.blocked.map((t) => [
  s.code, t.taskId, t.name, nv(t.group), t.status, nv(t.weeksStuck), nv(t.resource), t.reason || 'No log entry', nv(t.waitingOn)]));
const BLOCKED_HEAD = ['Site', 'Task', 'Name', 'Group', 'Status', 'Weeks', 'Owner', 'Reason', 'Waiting on'];

const titleRows = (brief) => [
  ['Site progress brief'],
  ['Run', fmtRun(brief.title.runAt)],
  ['Workbook', nv(brief.title.workbook)],
  [`Sites tracked (${brief.title.siteCount})`, brief.title.sites.join(', ')],
  ['Overall', `${brief.portfolio.pace.word}${brief.portfolio.pace.ratio != null ? ` (pace ${brief.portfolio.pace.ratio} of plan)` : ''}`],
  ['Pace counts', `${brief.portfolio.pace.sitesCounted} of ${brief.portfolio.pace.sitesTotal} sites${brief.portfolio.pace.excluded.length ? `; not counted: ${brief.portfolio.pace.excluded.map((e) => `${e.code} (${e.reason})`).join(', ')}` : ''}`],
  ['Work complete across sites', `${nv(brief.portfolio.overallPctByWeight)}%`],
];

const warningRows = (brief) => (brief.dataWarnings || []).map((w) => [w.site, w.severity === 'RED' ? 'Check: may be the wrong site' : 'Check', w.text]);
const WARN_HEAD = ['Site', 'Severity', 'Warning about the source data'];

/* ================================ Excel ================================ */

export function buildBriefWorkbook(brief, conclusion, engineLines = [], XLSX = window.XLSX) {
  const wb = XLSX.utils.book_new();
  const add = (name, aoa, widths) => {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (widths) ws['!cols'] = widths.map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, name);
  };
  add('Brief', [...titleRows(brief), [], SITES_HEAD, ...sitesRows(brief), [], ['Progress by site'], PROGRESS_HEAD, ...progressRows(brief)],
    [26, 16, 18, 8, 8, 9, 11, 11, 9, 52, 16]);
  add('Data warnings', [WARN_HEAD, ...(warningRows(brief).length ? warningRows(brief) : [['', '', 'None found']])], [8, 28, 120]);
  add('Task groups', [GROUP_HEAD, ...groupRows(brief)], [8, 34, 8, 8, 9, 11, 12, 9, 14]);
  add('Prerequisites', [PREQ_HEAD, ...preqRows(brief)], [8, 8, 36, 12, 15, 34, 24]);
  add('Blocked and waiting', [BLOCKED_HEAD, ...blockedRows(brief)], [8, 10, 36, 24, 10, 8, 10, 46, 12]);
  add('Conclusion', [['Source', 'Site', 'Text'], ...conclusionLines(brief, conclusion, engineLines).map((l) => [l.label, l.site, l.text])], [26, 8, 120]);
  return wb;
}

export function exportExcel(brief, conclusion, engineLines = []) {
  const wb = buildBriefWorkbook(brief, conclusion, engineLines);
  window.XLSX.writeFile(wb, `site-brief-${stamp(brief)}.xlsx`, { bookType: 'xlsx' });
}

/* ================================= Word ================================= */

function rtfEscape(s) {
  return String(s ?? '')
    .replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}')
    .replace(/[\u0080-\uFFFF]/g, (c) => `\\u${c.charCodeAt(0) > 32767 ? c.charCodeAt(0) - 65536 : c.charCodeAt(0)}?`)
    .replace(/\r?\n/g, '\\par ');
}

const BORDER = '\\clbrdrt\\brdrs\\brdrw5\\clbrdrl\\brdrs\\brdrw5\\clbrdrb\\brdrs\\brdrw5\\clbrdrr\\brdrs\\brdrw5';
function rtfRow(cells, widths, header) {
  let x = 0;
  let out = '\\trowd\\trgaph60\\trleft0';
  widths.forEach((w) => { x += w; out += `${BORDER}${header ? '\\clcbpat1' : ''}\\cellx${x}`; });
  cells.forEach((c) => {
    out += `\\pard\\intbl\\fs16 ${header ? '\\b ' : ''}${rtfEscape(c)}${header ? '\\b0 ' : ''}\\cell`;
  });
  return `${out}\\row `;
}
const rtfTable = (head, rows, widths) => `${rtfRow(head, widths, true)}${rows.map((r) => rtfRow(r, widths, false)).join('')}\\pard\\sa100 \\par `;
const h = (t) => `\\pard\\sb200\\sa80\\b\\fs24 ${rtfEscape(t)}\\b0\\fs18\\par `;
const p = (t, o = {}) => `\\pard\\sa80${o.i ? '\\i' : ''}${o.b ? '\\b' : ''}\\fs${o.fs || 18} ${rtfEscape(t)}${o.b ? '\\b0' : ''}${o.i ? '\\i0' : ''}\\fs18\\par `;

/** The brief as one landscape Word-openable RTF string. */
export function buildBriefRTF(brief, conclusion, engineLines = []) {
  const head = '{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}{\\colortbl;\\red241\\green239\\blue250;}'
    + '\\landscape\\paperw16838\\paperh11906\\margl720\\margr720\\margt720\\margb720\\fs18 ';
  const t = brief.portfolio.totals;
  let body = `\\pard\\sa60\\b\\fs36 Site progress brief\\b0\\fs18\\par `
    + p(`Run ${fmtRun(brief.title.runAt)}  |  Workbook: ${brief.title.workbook || '-'}  |  Sites tracked (${brief.title.siteCount}): ${brief.title.sites.join(', ')}`, { i: true });
  body += p(`Overall: ${brief.portfolio.pace.word}${brief.portfolio.pace.ratio != null ? ` (pace ${brief.portfolio.pace.ratio} of plan)` : ''}. `
    + `${nv(brief.portfolio.overallPctByWeight)}% of work complete. ${t.groupsComplete} of ${t.groupsTotal} task groups finished; `
    + `${t.finished} of ${t.tasks} tasks done; ${t.wip} WIP; ${t.blocked + t.waiting} blocked or waiting.`, { b: true });
  if (brief.portfolio.pace.excluded.length) {
    body += p(`Pace counts ${brief.portfolio.pace.sitesCounted} of ${brief.portfolio.pace.sitesTotal} sites. Not counted: ${brief.portfolio.pace.excluded.map((e) => `${e.code} (${e.reason})`).join(', ')}.`, { i: true });
  }
  if (brief.dataWarnings?.length) {
    body += h('Check before relying on this brief');
    for (const w of brief.dataWarnings) body += p(`${w.severity === 'RED' ? '[Check: may be the wrong site] ' : ''}${w.site}: ${w.text}`);
  }
  if (brief.privacy.flagged) body += p(`Check before sharing: possible personal or place names were found in ${brief.privacy.sites.join(', ')}.`, { i: true });

  body += h('Sites at a glance') + rtfTable(SITES_HEAD, sitesRows(brief), [1300, 900, 1300, 750, 750, 850, 950, 1000, 800, 3900, 1300]);
  body += h('Progress by site') + rtfTable(PROGRESS_HEAD, progressRows(brief), [1500, 2500, 2500, 1500]);
  const groups = groupRows(brief);
  if (groups.length) body += h('Task groups') + rtfTable(GROUP_HEAD, groups, [900, 3800, 900, 900, 1000, 1200, 1200, 1000, 1500]);
  body += h('Pending prerequisites');
  const preq = preqRows(brief);
  body += preq.length ? rtfTable(PREQ_HEAD, preq, [900, 800, 3600, 1200, 1500, 3600, 2800]) : p('No prerequisites are pending.');
  body += h('Blocked and waiting tasks');
  const blocked = blockedRows(brief);
  body += blocked.length
    ? rtfTable(BLOCKED_HEAD, blocked, [800, 1000, 3000, 2000, 1100, 700, 900, 3600, 1100])
    : p('No tasks are blocked or waiting.');
  body += h('Conclusion');
  for (const l of conclusionLines(brief, conclusion, engineLines)) body += p(`[${l.label}]${l.site ? ` ${l.site}:` : ''} ${l.text}`);
  body += p('All figures are calculated by the app from the workbook. Only text marked AI-written was produced by an AI.', { i: true });
  return `${head}${body}}`;
}

export function exportWord(brief, conclusion, engineLines = []) {
  const rtf = buildBriefRTF(brief, conclusion, engineLines);
  downloadBlob(new Blob([rtf], { type: 'application/rtf' }), `site-brief-${stamp(brief)}.rtf`);
}

/* ================================= PDF ================================= */

/**
 * Prints the brief through a hidden frame, so there is no popup to be blocked.
 * Returns false if the frame could not be created.
 */
export function exportPDF(brief, conclusion, { engineLines = [], orientation = 'portrait' } = {}) {
  const html = buildPrintHTML(brief, conclusion, { engineLines, orientation });
  const frame = document.createElement('iframe');
  frame.setAttribute('aria-hidden', 'true');
  frame.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
  frame.srcdoc = html;
  const cleanup = () => setTimeout(() => frame.remove(), 500);
  frame.addEventListener('load', () => {
    try {
      const w = frame.contentWindow;
      w.addEventListener('afterprint', cleanup);
      w.focus();
      w.print();
    } catch { cleanup(); }
  });
  document.body.appendChild(frame);
  setTimeout(() => frame.remove(), 10 * 60 * 1000);   // never leave it behind
  return true;
}
