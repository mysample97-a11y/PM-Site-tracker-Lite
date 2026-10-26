/**
 * Draws the one-page brief as HTML. Used for the on-screen page and, unchanged,
 * for the PDF (the browser's print-to-PDF) — so what is reviewed is what prints.
 *
 * Two safety rules, both tested:
 *  1. Every piece of workbook text goes through esc(). A task name typed into a
 *     shared spreadsheet can never become markup.
 *  2. Colours are literal hex values, never CSS variables. In a separate print
 *     window an unresolved variable inside SVG silently falls back to black.
 */

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const RAGC = { GREEN: '#0E9F6E', AMBER: '#C2740A', RED: '#DC2626' };
const RAGW = { GREEN: 'Green', AMBER: 'Amber', RED: 'Red' };
const RAGS = { GREEN: '●', AMBER: '▲', RED: '■' };   // shape as well as colour

export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => (
  { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/** "2026-04-03" -> "3 Apr 2026". Parsed by hand so the time zone can never shift the day. */
export function fmtD(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  return m ? `${Number(m[3])} ${MONTHS[Number(m[2]) - 1]} ${m[1]}` : (iso ? String(iso) : '—');
}

/** Run time, in the viewer's own time zone: "16 Feb 2026, 09:30". */
export function fmtRun(iso) {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '—';
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}, ${p(d.getHours())}:${p(d.getMinutes())}`;
}

const num = (v, dp = 1) => (v == null ? '—' : (Math.round(v * 10 ** dp) / 10 ** dp).toString());

/* ================================== CSS ================================== */

export const BRIEF_CSS = `
.bf{font-family:Arial,Helvetica,sans-serif;font-size:12px;line-height:1.35;color:#1E1B36;background:#fff}
.bf *{box-sizing:border-box}
.bf-head{border-bottom:2px solid #241E58;padding-bottom:6px;margin-bottom:8px}
.bf h1{font-size:20px;margin:0 0 2px;color:#241E58}
.bf h2{font-size:13px;margin:12px 0 4px;color:#241E58;text-transform:uppercase;letter-spacing:.04em}
.bf h3{font-size:12px;margin:0 0 3px}
.bf-meta{color:#4A4570;font-size:11px}
.bf-sub{color:#7C769D;font-size:10px}
.bf-overall{display:flex;gap:14px;align-items:center;flex-wrap:wrap;margin:6px 0}
.bf-overall>div{min-width:0}
.bf-paceword{font-size:16px;font-weight:700;display:inline-block;padding:3px 10px;border-radius:3px;color:#fff}
.bf-pace-green{background:#0E9F6E}.bf-pace-amber{background:#C2740A}.bf-pace-red{background:#DC2626}.bf-pace-none{background:#7C769D}
.bf-chart{flex:1 1 320px;max-width:560px}
.bf-chart svg{width:100%;height:auto;display:block}
.bf-table{border-collapse:collapse;width:100%;margin:2px 0 6px}
.bf-table th{background:#F1EFFA;text-align:left;font-weight:700;font-size:10.5px;padding:3px 5px;border:1px solid #D9D4F3}
.bf-table td{padding:3px 5px;border:1px solid #D9D4F3;vertical-align:top}
.bf-table .num{text-align:right;font-variant-numeric:tabular-nums}
.bf-table tr{break-inside:avoid}
.bf-chip{display:inline-block;padding:1px 7px;border-radius:3px;font-size:10.5px;font-weight:700;color:#fff;white-space:nowrap}
.bf-green{background:#0E9F6E}.bf-amber{background:#C2740A}.bf-red{background:#DC2626}.bf-grey{background:#7C769D}
.bf-late{color:#DC2626;font-weight:700}.bf-ok{color:#0E9F6E;font-weight:700}
.bf-groups{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:8px}
.bf-groups .bf-table td:first-child{max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bf-done td:first-child{font-weight:700}
.bf-total td{background:#F7F5FE;font-weight:700}
.bf-note{font-size:11px;color:#4A4570;margin:2px 0}
.bf-concl{border:1px solid #B5ACE9;border-radius:4px;padding:8px 10px;margin-top:6px;break-inside:avoid}
.bf-badge{display:inline-block;font-size:10px;font-weight:700;padding:1px 6px;border-radius:3px;margin-left:6px;vertical-align:middle}
.bf-ai{background:#EDEBFD;color:#4A3FD6;border:1px solid #B5ACE9}
.bf-eng{background:#ECEAF7;color:#4A4570;border:1px solid #C3BEDF}
.bf-head-line{font-size:13px;font-weight:700;margin:2px 0 6px}
.bf-verdicts{list-style:none;margin:0;padding:0}
.bf-verdicts li{margin:3px 0}
.bf-act{color:#4A4570}
.bf-issues{font-size:10.5px;color:#C2740A;margin-top:4px}
.bf-foot{margin-top:8px;font-size:10px;color:#7C769D;border-top:1px solid #D9D4F3;padding-top:4px}
.bf-attn{border:1px solid #DC2626;border-left:5px solid #DC2626;background:#FDF0F0;padding:6px 10px;margin:6px 0;border-radius:3px;font-size:11.5px;break-inside:avoid}
.bf-attn.amber{border-color:#C2740A;background:#FCF2DF}
.bf-attn ul{margin:3px 0 0 16px;padding:0}.bf-attn li{margin:1px 0}
.bf-warn{background:#FCF2DF;border:1px solid #C2740A;color:#7A4A06;padding:6px 8px;border-radius:3px;margin:6px 0;font-size:11.5px}
`;

export const PRINT_CSS = (orientation = 'portrait') => `
@page{size:A4 ${orientation === 'landscape' ? 'landscape' : 'portrait'};margin:10mm}
html,body{margin:0;background:#fff}
*{-webkit-print-color-adjust:exact;print-color-adjust:exact}
.bf{font-size:8pt}
.bf h1{font-size:15pt}.bf h2{font-size:9pt}.bf-head-line{font-size:9.5pt}
.bf-meta,.bf-note{font-size:7.5pt}.bf-sub{font-size:6.5pt}
.bf-table th,.bf-chip{font-size:7pt}
.bf-paceword{font-size:11pt}
`;

/* ================================== chart ================================== */

/** One bar per site: work complete (bar) against time used (line). */
export function chartSVG(brief) {
  const rows = brief.portfolio.chart;
  if (!rows.length) return '';
  const W = 560, labelW = 70, trackW = 360, rowH = 26, top = 6;
  const H = top + rows.length * rowH + 24;
  const clampPct = (v) => Math.max(0, Math.min(100, Number(v) || 0));
  const parts = [];
  rows.forEach((r, i) => {
    const y = top + i * rowH;
    const w = (trackW * clampPct(r.pctByWeight)) / 100;
    const colour = r.submitted ? '#5B4FE9' : (RAGC[r.rag] || '#7C769D');
    parts.push(
      `<text x="0" y="${y + 16}" font-size="12" font-weight="700" fill="#1E1B36">${esc(r.code)}</text>`,
      `<rect x="${labelW}" y="${y + 3}" width="${trackW}" height="16" rx="2" fill="#E6E3F2"/>`,
      `<rect x="${labelW}" y="${y + 3}" width="${w.toFixed(1)}" height="16" rx="2" fill="${colour}"/>`,
    );
    if (r.timeElapsed != null && !r.submitted) {
      const x = labelW + (trackW * clampPct(r.timeElapsed)) / 100;
      parts.push(`<line x1="${x.toFixed(1)}" y1="${y}" x2="${x.toFixed(1)}" y2="${y + 22}" stroke="#1E1B36" stroke-width="2.5"/>`);
    }
    parts.push(`<text x="${labelW + trackW + 8}" y="${y + 16}" font-size="12" fill="#1E1B36">${esc(num(r.pctByWeight))}% done${r.timeElapsed != null && !r.submitted ? ` · ${esc(num(r.timeElapsed))}% of time used` : ''}</text>`);
  });
  const ly = top + rows.length * rowH + 14;
  parts.push(`<text x="${labelW}" y="${ly}" font-size="10" fill="#4A4570">Bar: work complete (by weight). Black line: share of time used, start to target.</text>`);
  const label = `Progress of ${rows.length} site${rows.length === 1 ? '' : 's'}: ${rows.map((r) => `${r.code} ${num(r.pctByWeight)} percent`).join(', ')}`;
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(label)}" xmlns="http://www.w3.org/2000/svg" font-family="Arial,Helvetica,sans-serif">${parts.join('')}</svg>`;
}

/* ================================== the page ================================== */

const chip = (rag) => (RAGW[rag]
  ? `<span class="bf-chip bf-${rag.toLowerCase()}">${RAGS[rag]} ${RAGW[rag]}</span>`
  : '<span class="bf-chip bf-grey">—</span>');

function sitesTable(brief) {
  const head = '<tr><th>Site</th><th>Rating</th><th class="num">Task groups finished</th><th class="num">Done</th><th class="num">WIP</th>'
    + '<th class="num">Blocked</th><th class="num">Waiting on</th><th class="num">Not started</th><th class="num">% done</th><th>Target / forecast</th><th>Since last week</th></tr>';
  const body = brief.sites.map((s) => {
    const name = `<td><b>${esc(s.code)}</b>${s.description ? `<div class="bf-sub">${esc(s.description)}</div>` : ''}</td>`;
    if (s.noData) return `<tr>${name}<td colspan="10">${esc(s.reason)}</td></tr>`;
    const sc = s.schedule;
    let fc = '<div>No forecast</div>';
    if (sc.forecastDate) {
      const v = sc.varianceDays;
      const tag = v == null ? '' : ` <span class="${v > 0 ? 'bf-late' : 'bf-ok'}">(${v > 0 ? '+' : ''}${v} d)</span>`;
      fc = `<div>Forecast ${esc(fmtD(sc.forecastDate))}${tag}</div>`;
    }
    const m = s.movement;
    const last = m?.comparable
      ? `<span title="${esc(m.text)}">${m.percentChange > 0 ? '+' : ''}${esc(num(m.percentChange))} pts</span>`
      : `<span title="${esc(m?.text || 'No comparison available')}">—</span>`;
    const c = s.counts;
    return `<tr>${name}<td>${chip(s.overall)}<div class="bf-sub">${esc(s.reportingWeek || '')} · ${esc(fmtD(s.statusDate))}</div></td>`
      + `<td class="num">${s.groups.complete} of ${s.groups.total}</td><td class="num">${c.finished}</td><td class="num">${c.wip}</td>`
      + `<td class="num">${c.blocked}</td><td class="num">${c.waiting}</td><td class="num">${c.notStarted}</td>`
      + `<td class="num">${esc(num(s.pctByWeight))}</td><td><div>Target ${esc(fmtD(sc.target))}</div>${fc}</td><td>${last}</td></tr>`;
  }).join('');
  return `<table class="bf-table">${head}${body}</table>`;
}

function groupsBlocks(brief) {
  const blocks = brief.sites.filter((s) => !s.noData).map((s) => {
    const rows = s.groups.rows.map((g) =>
      `<tr class="${g.complete ? 'bf-done' : ''}"><td title="${esc(g.name)}">${g.complete ? '✓ ' : ''}${esc(g.name)}</td>`
      + `<td class="num">${g.done}</td><td class="num">${g.wip}</td><td class="num">${g.pending}</td></tr>`).join('');
    const t = s.groups.rows.reduce((a, g) => ({ d: a.d + g.done, w: a.w + g.wip, p: a.p + g.pending }), { d: 0, w: 0, p: 0 });
    return `<div><h3>${esc(s.code)} <span class="bf-sub">as at ${esc(s.reportingWeek || '')}</span></h3>`
      + `<table class="bf-table"><tr><th>Task group</th><th class="num">Done</th><th class="num">WIP</th><th class="num">Pending</th></tr>${rows}`
      + `<tr class="bf-total"><td>All groups</td><td class="num">${t.d}</td><td class="num">${t.w}</td><td class="num">${t.p}</td></tr></table></div>`;
  });
  if (!blocks.length) return '';
  return `<h2>Task groups</h2><div class="bf-groups">${blocks.join('')}</div>`
    + '<div class="bf-note">Pending = not started, blocked or waiting on. A ✓ marks a group with every task finished.</div>';
}

function prerequisitesBlock(brief) {
  const sites = brief.sites.filter((s) => !s.noData);
  const rows = sites.flatMap((s) => s.prerequisites.pending.map((p) =>
    `<tr><td><b>${esc(s.code)}</b></td><td>${esc(p.id)} ${esc(p.name)}</td><td>${esc(p.provider || '—')}</td>`
    + `<td>${esc(p.byWeek ?? '—')}</td><td>${esc(p.status || '—')}</td><td>${esc((p.neededFor || []).join(', ') || '—')}</td></tr>`));
  const unknown = sites.filter((s) => s.prerequisites.notTracked.length).map((s) =>
    `<div class="bf-note">${esc(s.code)}: ${s.prerequisites.notTracked.length} standard prerequisite${s.prerequisites.notTracked.length === 1 ? ' has' : 's have'} no entry in the site's own table, so their status is unknown.</div>`);
  if (!rows.length && !unknown.length) return '<h2>Pending prerequisites</h2><div class="bf-note">No prerequisites are pending.</div>';
  const table = rows.length
    ? `<table class="bf-table"><tr><th>Site</th><th>Prerequisite</th><th>From</th><th>Needed by week</th><th>Status</th><th>Needed for tasks</th></tr>${rows.join('')}</table>`
    : '<div class="bf-note">No prerequisites are recorded as pending.</div>';
  return `<h2>Pending prerequisites</h2>${table}${unknown.join('')}`;
}

function blockedBlock(brief) {
  const rows = brief.sites.filter((s) => !s.noData).flatMap((s) => s.blocked.map((t) =>
    `<tr><td><b>${esc(s.code)}</b></td><td>${esc(t.taskId)} ${esc(t.name)}${t.group ? `<div class="bf-sub">${esc(t.group)}</div>` : ''}</td>`
    + `<td>${esc(t.status)}</td><td class="num">${esc(t.weeksStuck ?? '—')}</td><td>${esc(t.resource || '—')}</td>`
    + `<td>${t.reason ? esc(t.reason) : '<i>No log entry</i>'}${t.waitingOn ? ` <span class="bf-sub">(waiting on ${esc(t.waitingOn)})</span>` : ''}</td></tr>`));
  if (!rows.length) return '<h2>Blocked and waiting tasks</h2><div class="bf-note">No tasks are blocked or waiting.</div>';
  return `<h2>Blocked and waiting tasks</h2><table class="bf-table"><tr><th>Site</th><th>Task</th><th>Status</th><th class="num">Weeks</th><th>Owner</th><th>Reason</th></tr>${rows.join('')}</table>`
    + '<div class="bf-note">Blocked: the team can clear it. Waiting on: someone outside the team must act first.</div>';
}

function conclusionBlock(brief, conclusion, engineLines) {
  const moves = brief.sites.filter((s) => !s.noData && s.movement?.text).map((s) =>
    `<div class="bf-note"><b>${esc(s.code)}</b> — ${esc(s.movement.text)}</div>`).join('');
  const lastWeek = moves ? `<h3>Since last week <span class="bf-badge bf-eng">Calculated</span></h3>${moves}` : '';

  if (!conclusion) {
    const lines = engineLines.map((l) => `<li>${esc(l)}</li>`).join('');
    return `<h2>Conclusion</h2><div class="bf-concl">${lastWeek}`
      + `<h3 style="margin-top:6px">Summary <span class="bf-badge bf-eng">Calculated, no AI</span></h3><ul class="bf-verdicts">${lines}</ul>`
      + '<div class="bf-note"><i>An AI-written conclusion (on time or not, and where to move resources) has not been generated for this brief.</i></div></div>';
  }
  const r = conclusion.result;
  const meta = conclusion.meta || {};
  const verdictColour = { 'On track': 'green', 'At risk': 'amber', 'Will miss target': 'red', 'Cannot tell': 'grey' };
  const items = r.sites.map((s) =>
    `<li><span class="bf-chip bf-${verdictColour[s.verdict] || 'grey'}">${esc(s.verdict)}</span> <b>${esc(s.site)}</b> — ${esc(s.reason)}`
    + `${s.action ? `<div class="bf-act">Action: ${esc(s.action)}</div>` : ''}</li>`).join('');
  return `<h2>Conclusion</h2><div class="bf-concl">${lastWeek}`
    + `<h3 style="margin-top:6px">Assessment <span class="bf-badge bf-ai">AI-written · ${esc(meta.model || 'model')}${meta.at ? ` · ${esc(fmtRun(meta.at))}` : ''}</span></h3>`
    + `${r.headline ? `<div class="bf-head-line">${esc(r.headline)}</div>` : ''}<ul class="bf-verdicts">${items}</ul>`
    + `${r.reallocation ? `<div><b>Resources:</b> ${esc(r.reallocation)}</div>` : ''}`
    + `${r.issues?.length ? `<div class="bf-issues">Checked on receipt: ${r.issues.map(esc).join(' ')}</div>` : ''}`
    + '<div class="bf-note"><i>Written by an AI from the figures above. The figures are calculated; this reading of them is a judgement and should be checked by a person.</i></div></div>';
}

/** Problems with the source data, so nobody relies on a figure that may belong to another site. */
function attentionBlock(brief) {
  const w = brief.dataWarnings || [];
  if (!w.length) return '';
  const bySite = new Map();
  for (const x of w) {
    if (!bySite.has(x.site)) bySite.set(x.site, []);
    bySite.get(x.site).push(x);
  }
  const items = [...bySite.entries()].map(([site, list]) =>
    `<li><b>${esc(site)}</b> — ${list.map((x) => `${x.severity === 'RED' ? '■ ' : ''}${esc(x.text)}`).join(' ')}</li>`).join('');
  const red = w.some((x) => x.severity === 'RED');
  return `<div class="bf-attn${red ? '' : ' amber'}"><b>Check before relying on this brief</b>${red ? ' — a ■ item means a figure may belong to the wrong site' : ''}<ul>${items}</ul></div>`;
}

/**
 * @param {object} brief        buildBrief() result
 * @param {object|null} conclusion  { result, meta } or null
 * @param {object} opts
 * @param {string[]} opts.engineLines  engineSummary(brief), shown when there is no AI conclusion
 * @param {boolean} opts.screen  screen-only extras (privacy banner)
 */
export function renderBriefHTML(brief, conclusion, { engineLines = [], screen = true } = {}) {
  if (!brief.sites.length) return '<div class="bf"><p>No sites are selected.</p></div>';
  const p = brief.portfolio.pace;
  const paceCls = { GREEN: 'green', AMBER: 'amber', RED: 'red' }[p.rag] || 'none';
  const t = brief.portfolio.totals;
  const warn = screen && brief.privacy.flagged
    ? `<div class="bf-warn"><b>Check before sharing or sending to an AI:</b> possible personal or place names were found in ${esc(brief.privacy.sites.join(', '))}. The check is a guess and only warns.</div>` : '';
  const warnings = brief.warnings.map((w) => `<div class="bf-warn">${esc(w)}</div>`).join('');
  return `<div class="bf">
<div class="bf-head"><h1>Site progress brief</h1>
<div class="bf-meta">Run ${esc(fmtRun(brief.title.runAt))} · Workbook: ${esc(brief.title.workbook || '—')} · Sites tracked (${brief.title.siteCount}): <b>${esc(brief.title.sites.join(', '))}</b></div></div>
${warn}${warnings}${attentionBlock(brief)}
<div class="bf-overall"><div><span class="bf-paceword bf-pace-${paceCls}">Overall: ${esc(p.word)}</span>
<div class="bf-note">${p.ratio != null ? `Pace ${esc(p.ratio)} of plan, counting ${p.sitesCounted} of ${p.sitesTotal} site${p.sitesTotal === 1 ? '' : 's'} · ` : ''}${esc(num(brief.portfolio.overallPctByWeight))}% of work complete across ${brief.portfolio.sitesWithData} site${brief.portfolio.sitesWithData === 1 ? '' : 's'} with data</div>
${p.excluded?.length ? `<div class="bf-note">Not counted in pace: ${p.excluded.map((e) => `${esc(e.code)} (${esc(e.reason)})`).join(', ')}</div>` : ''}
<div class="bf-note">${t.groupsComplete} of ${t.groupsTotal} task groups finished · ${t.finished} of ${t.tasks} tasks done · ${t.wip} WIP · ${t.blocked + t.waiting} blocked or waiting</div></div>
<div class="bf-chart">${chartSVG(brief)}</div></div>
<h2>Sites at a glance</h2>${sitesTable(brief)}
${groupsBlocks(brief)}
${prerequisitesBlock(brief)}
${blockedBlock(brief)}
${conclusionBlock(brief, conclusion, engineLines)}
<div class="bf-foot">All figures are calculated by the app from the workbook. ${esc(p.basis)} Rating uses the tracker's red/amber/green rules.</div>
</div>`;
}

/** A complete standalone document for printing to PDF. */
export function buildPrintHTML(brief, conclusion, { engineLines = [], orientation = 'portrait' } = {}) {
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Site progress brief ${esc(brief.title.reportDate)}</title>`
    + `<style>${BRIEF_CSS}${PRINT_CSS(orientation)}</style></head><body>${renderBriefHTML(brief, conclusion, { engineLines, screen: false })}</body></html>`;
}
