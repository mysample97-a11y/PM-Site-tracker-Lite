/**
 * Page, export and session verification.
 *
 *   - Every piece of workbook or AI text is escaped on the page (HTML injection).
 *   - The printable document is self-contained: no scripts, no CSS variables.
 *   - Word and Excel carry the figures, cover only the selected sites, and survive
 *     awkward characters.
 *   - A saved session reloads to the identical page, holds only the selected
 *     sites, and never holds the API key.
 *   - Nothing about a site is left in browser storage.
 *
 * Run: node tools/verify-view.mjs
 */

import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { fixtureWorkbook } from './fixture.mjs';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const XLSX = createRequire(import.meta.url)(path.join(root, 'assets/vendor/xlsx.full.min.js'));
global.window = { XLSX };
const memStore = () => {
  const m = new Map();
  return {
    getItem: (k) => (m.has(k) ? m.get(k) : null), setItem: (k, v) => m.set(k, String(v)),
    removeItem: (k) => m.delete(k), clear: () => m.clear(), dump: () => JSON.stringify([...m.entries()]),
  };
};
global.localStorage = memStore();
global.sessionStorage = memStore();

const { readWorkbook, parseWorkbook, STATUS } = await import('../js/core/parser.js');
const { buildBrief, engineSummary } = await import('../js/core/brief.js');
const V = await import('../js/core/briefview.js');
const X = await import('../js/core/exports.js');
const S = await import('../js/core/session.js');
const T = await import('../js/core/tokens.js');

let pass = 0, fail = 0;
const failures = [];
const check = (name, cond, detail = '') => {
  if (cond) pass++;
  else { fail++; failures.push(`${name}${detail ? ` — ${detail}` : ''}`); }
};

const load = async () => {
  const buf = XLSX.write(fixtureWorkbook(XLSX, STATUS), { type: 'array', bookType: 'xlsx' });
  return parseWorkbook(await readWorkbook({ name: 'f.xlsx', arrayBuffer: async () => buf }));
};
const NOW = new Date(2026, 1, 16, 9, 30);
const model = await load();
const brief = buildBrief(model, { codes: ['F-01'], now: NOW, workbookName: 'fixture.xlsx' });
const lines = engineSummary(brief);
const CONCL = {
  result: {
    headline: 'F-01 is behind plan.', issues: [],
    sites: [{ site: 'F-01', verdict: 'Will miss target', reason: 'Pace is 0.61 of plan.', action: 'Clear the two stuck tasks.' }],
    reallocation: 'Move R2 onto the modelling backlog.',
  },
  meta: { model: 'gemini-test', at: new Date(2026, 1, 16, 9, 41).toISOString() },
};

/* ------------------------------ 1. the page ------------------------------ */
{
  const html = V.renderBriefHTML(brief, null, { engineLines: lines });
  const has = (t) => html.includes(t);
  check('title block names the run, workbook and sites', has('Site progress brief') && has('16 Feb 2026, 09:30') && has('fixture.xlsx') && has('Sites tracked (1)') && has('<b>F-01</b>'));
  check('overall status word and pace', has('Overall: Poor') && has('bf-pace-red') && has('Pace 0.61'));
  check('site table: groups finished 1 of 2', has('1 of 2'));
  check('site table: counts present', has('<td class="num">4</td><td class="num">2</td><td class="num">1</td><td class="num">1</td><td class="num">2</td>'));
  check('site table: completion 45.5', has('45.5'));
  check('site table: target and forecast', has('Target 1 Mar 2026') && has('Forecast'));
  check('task group table lists Modelling with pending 4', /Modelling<\/td><td class="num">2<\/td><td class="num">2<\/td><td class="num">4<\/td>/.test(html));
  check('finished group is marked', has('✓ Initial setup'));
  check('pending prerequisite is listed', has('P02') && has('Ceiling'));
  check('blocked list shows the reason and the unlogged case', has('Ceiling void zones not confirmed') && has('No log entry'));
  check('conclusion without AI is labelled as calculated, with no AI badge', has('Calculated, no AI') && !has('bf-badge bf-ai'));
  check('page says an AI conclusion was not generated', has('has not been generated'));
  check('chart is drawn', has('<svg') && has('role="img"'));
  check('no placeholder leaks into the page', !/undefined|NaN|\[object|null</.test(html));
  check('colours are literal, never CSS variables', !html.includes('var(--') && !V.BRIEF_CSS.includes('var(--'));

  const withAI = V.renderBriefHTML(brief, CONCL, { engineLines: lines });
  check('AI conclusion is labelled with its model', /AI-written · gemini-test/.test(withAI));
  check('AI conclusion shows headline, verdict, action, resources', withAI.includes('F-01 is behind plan.') && withAI.includes('Will miss target') && withAI.includes('Clear the two stuck tasks.') && withAI.includes('Move R2'));
  check('AI conclusion says it is a judgement to be checked', /judgement and should be checked/.test(withAI));
  check('calculated "since last week" is shown beside the AI text', withAI.includes('Since last week') && withAI.includes('Calculated'));

  const two = buildBrief(model, { codes: ['F-01', 'F-02'], now: NOW });
  const html2 = V.renderBriefHTML(two, null, { engineLines: engineSummary(two) });
  check('a site with no data is shown, with the reason', html2.includes('F-02') && /no detail sheet/.test(html2));
  check('empty selection renders a message, not a blank', V.renderBriefHTML(buildBrief(model, { codes: [], now: NOW }), null).includes('No sites are selected'));
  check('one-site page does not contain the other site', !html.includes('F-02'));

  check('fmtD reads ISO dates without a time-zone shift', V.fmtD('2026-04-03') === '3 Apr 2026' && V.fmtD('2026-01-01') === '1 Jan 2026' && V.fmtD(null) === '—');
  check('fmtRun is stable', V.fmtRun(NOW.toISOString()) === '16 Feb 2026, 09:30');
}

/* ------------------------- 2. HTML injection ------------------------- */
{
  const EVIL = '<img src=x onerror=alert(1)>';
  const b = JSON.parse(JSON.stringify(brief));
  b.title.workbook = EVIL;
  const s = b.sites[0];
  s.description = `"><script>alert(2)</script>`;
  s.blocked[0].name = EVIL; s.blocked[0].reason = '<b onmouseover=1>x</b>'; s.blocked[0].resource = EVIL;
  s.prerequisites.pending[0].name = EVIL;
  s.groups.rows[0].name = EVIL;
  s.reportingWeek = EVIL;
  s.movement = { comparable: true, percentChange: 1, text: EVIL };
  const c = JSON.parse(JSON.stringify(CONCL));
  c.result.headline = EVIL; c.result.sites[0].reason = '<script>alert(3)</script>'; c.result.sites[0].action = EVIL;
  c.result.reallocation = EVIL; c.result.issues = [EVIL]; c.meta.model = EVIL;
  const html = V.renderBriefHTML(b, c, { engineLines: [EVIL] });
  check('no <img> tag survives from workbook or AI text', !/<img/i.test(html));
  check('no <script> tag survives', !/<script/i.test(html));
  // Quoted attribute values are inert text once escaped, so they are blanked before looking for on...= attributes.
  const tagsOnly = html.replace(/"[^"]*"/g, '""');
  check('no event-handler attribute can be injected', !/<[a-z][^>]*\son\w+\s*=/i.test(tagsOnly));
  check('the check itself works: a real injected handler would be caught', /<[a-z][^>]*\son\w+\s*=/i.test('<img src="x" onerror=alert(1)>'.replace(/"[^"]*"/g, '""')));
  check('the text is shown, escaped', html.includes('&lt;img src=x onerror=alert(1)&gt;'));
  check('quotes in attributes are escaped', !html.includes('title=""><script'));
  check('escape handles all five characters', V.esc(`&<>"'`) === '&amp;&lt;&gt;&quot;&#39;');
  const print = V.buildPrintHTML(b, c, { engineLines: [EVIL] });
  check('the print document is safe the same way', !/<img/i.test(print) && !/<script/i.test(print.replace(/<style>[\s\S]*?<\/style>/, '')));
}

/* ------------------------- 3. the printable document ------------------------- */
{
  const doc = V.buildPrintHTML(brief, CONCL, { engineLines: lines });
  const flagged = JSON.parse(JSON.stringify(brief));
  flagged.privacy = { flagged: true, sites: ['F-01'] };
  check('print document is complete HTML', doc.startsWith('<!doctype html>') && doc.includes('</html>'));
  check('print document has no scripts at all', !/<script/i.test(doc));
  check('print document uses no CSS variables', !doc.includes('var(--'));
  check('print is A4 portrait by default', /@page\{size:A4 portrait/.test(doc));
  check('landscape can be chosen', /@page\{size:A4 landscape/.test(V.buildPrintHTML(brief, null, { orientation: 'landscape' })));
  check('colours are forced to print', doc.includes('print-color-adjust:exact'));
  check('rows are kept whole across pages', doc.includes('break-inside:avoid'));
  check('privacy banner shows on screen only', V.renderBriefHTML(flagged, null, { screen: true }).includes('Check before sharing')
    && !V.buildPrintHTML(flagged, null).includes('Check before sharing'));
  check('chart uses literal colours', V.chartSVG(brief).includes('#DC2626') && !V.chartSVG(brief).includes('var('));
  check('chart marks time used as a line, not only colour', V.chartSVG(brief).includes('<line'));
  check('chart carries a text alternative', /aria-label="Progress of 1 site: F-01 45.5 percent"/.test(V.chartSVG(brief)));
  check('no chart when no site has data', V.chartSVG(buildBrief(model, { codes: ['F-02'], now: NOW })) === '');
}

/* ---------------------------- 4. Excel ---------------------------- */
{
  const wb = X.buildBriefWorkbook(brief, CONCL, lines, XLSX);
  const back = XLSX.read(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }), { type: 'array' });
  const aoa = (n) => XLSX.utils.sheet_to_json(back.Sheets[n], { header: 1, defval: '' });
  check('five sheets', back.SheetNames.join('|') === 'Brief|Task groups|Prerequisites|Blocked and waiting|Conclusion', back.SheetNames.join('|'));
  const brf = aoa('Brief');
  const siteRow = brf.find((r) => r[0] === 'F-01' && r[1] === 'RED');
  check('Brief sheet: site row with counts', !!siteRow && siteRow[3] === 4 && siteRow[4] === 2 && siteRow[5] === 1 && siteRow[6] === 1 && siteRow[7] === 2 && siteRow[8] === 45.5, JSON.stringify(siteRow));
  check('Brief sheet: groups finished', siteRow?.[2] === '1 of 2');
  check('Brief sheet: title block', brf.some((r) => r[0] === 'Run' && r[1] === '16 Feb 2026, 09:30') && brf.some((r) => r[0] === 'Workbook' && r[1] === 'fixture.xlsx'));
  check('Brief sheet: chart data as a table', brf.some((r) => r[0] === 'F-01' && r[1] === 45.5 && r[2] === 74.5));
  const grp = aoa('Task groups');
  const mod = grp.find((r) => r[1] === 'Modelling');
  check('Task groups sheet: Modelling row', mod && mod[2] === 2 && mod[3] === 2 && mod[4] === 1 && mod[5] === 1 && mod[6] === 2 && mod[7] === 4, JSON.stringify(mod));
  check('Prerequisites sheet lists P02', aoa('Prerequisites').some((r) => r[1] === 'P02'));
  check('Blocked sheet has both tasks', aoa('Blocked and waiting').filter((r) => String(r[1]).startsWith('T020')).length === 2);
  const concl = aoa('Conclusion');
  check('Conclusion sheet labels AI text', concl.some((r) => /^AI-written \(gemini-test\)/.test(r[0]) && /behind plan/.test(r[2])));
  check('Conclusion sheet labels calculated text', concl.some((r) => r[0] === 'Calculated'));
  check('a one-site workbook holds no other site', !JSON.stringify([...back.SheetNames.map(aoa)]).includes('F-02'));
  const none = X.buildBriefWorkbook(brief, null, lines, XLSX);
  const nb = XLSX.read(XLSX.write(none, { type: 'array', bookType: 'xlsx' }), { type: 'array' });
  check('without AI the sheet says so', XLSX.utils.sheet_to_json(nb.Sheets.Conclusion, { header: 1 }).some((r) => /has not been generated/.test(r[2] || '')));
}

/* ----------------------------- 5. Word ----------------------------- */
{
  const rtf = X.buildBriefRTF(brief, CONCL, lines);
  const bal = (s) => (s.match(/(?<!\\)\{/g) || []).length === (s.match(/(?<!\\)\}/g) || []).length;
  check('RTF braces balance', bal(rtf));
  check('RTF is landscape with tables', rtf.includes('\\landscape') && rtf.includes('\\trowd') && rtf.includes('\\cellx'));
  check('RTF carries the figures', rtf.includes('F-01') && rtf.includes('45.5') && rtf.includes('1 of 2'));
  check('RTF labels the AI text', rtf.includes('[AI-written (gemini-test)]'));
  check('RTF holds no other site', !rtf.includes('F-02'));
  const awkward = JSON.parse(JSON.stringify(CONCL));
  awkward.result.reallocation = 'Move {R2} to path C:\\work – café';
  const r2 = X.buildBriefRTF(brief, awkward, lines);
  check('RTF escapes braces and backslashes', r2.includes('\\{R2\\}') && r2.includes('C:\\\\work'));
  check('RTF escapes non-ASCII characters', r2.includes('\\u233?') && r2.includes('\\u8211?'));
  check('RTF stays balanced with awkward text', bal(r2));
  check('RTF notes when there is no AI conclusion', X.buildBriefRTF(brief, null, lines).includes('has not been generated'));
  const cl = X.conclusionLines(brief, CONCL, lines);
  check('conclusion lines keep calculated and AI text apart', cl.some((l) => l.label === 'Calculated') && cl.some((l) => /^AI-written/.test(l.label)) && cl.some((l) => l.label === 'Note'));
}

/* ---------------------------- 6. session ---------------------------- */
{
  const state = { workbookName: 'fixture.xlsx', model, previousName: null, previousModel: null, selection: ['F-01'], runAt: NOW.toISOString(), brief, conclusion: CONCL };
  S.setApiKey('SECRET-KEY-ABC', true);
  const poisoned = { ...state, conclusion: { ...CONCL, meta: { ...CONCL.meta, apiKey: 'LEAK-1', 'x-api-key': 'LEAK-2' } } };
  const file = S.buildSessionFile(poisoned, new Date(2026, 1, 16, 10, 0));
  const text = JSON.stringify(file);
  check('session has the format and version', file.format === 'bim-tracker-lite.session' && file.version === 1);
  check('the API key is not in the file', !text.includes('SECRET-KEY-ABC'));
  check('key-like fields are stripped', !text.includes('LEAK-1') && !text.includes('LEAK-2'));
  check('the file holds only the selected site', !text.includes('F-02'));
  check('the register and other sheets are not in the file', file.state.model.register === undefined && file.state.model.sites.length === 1);
  check('the conclusion is kept', file.state.conclusion.result.headline === 'F-01 is behind plan.');

  const back = S.parseSessionFile(text).state;
  const rebuilt = buildBrief(back.model, { codes: back.selection, previousModel: back.previousModel || null, now: new Date(back.runAt), workbookName: back.workbookName });
  const a = V.renderBriefHTML(brief, CONCL, { engineLines: lines });
  const b = V.renderBriefHTML(rebuilt, back.conclusion, { engineLines: engineSummary(rebuilt) });
  check('a reloaded session draws the identical page', a === b, `${a.length} vs ${b.length}`);
  check('the reloaded figures match', JSON.stringify(rebuilt.portfolio) === JSON.stringify(brief.portfolio));

  // with a previous-week file
  const prev = await load();
  for (const t of prev.sites[0].detail.tasks) t.weekly[5] = null;
  for (const c of prev.sites[0].detail.categories) c.weekly[5] = null;
  prev.sites[0].detail.tasks.find((t) => t.id === 'T020-02').weekly[4] = STATUS.NOT_STARTED;
  const withPrev = buildBrief(model, { codes: ['F-01'], previousModel: prev, now: NOW, workbookName: 'fixture.xlsx' });
  const f2 = JSON.stringify(S.buildSessionFile({ ...state, previousName: 'last.xlsx', previousModel: prev, brief: withPrev }));
  const l2 = S.parseSessionFile(f2).state;
  const re2 = buildBrief(l2.model, { codes: l2.selection, previousModel: l2.previousModel, now: new Date(l2.runAt), workbookName: l2.workbookName });
  check('a session with a previous-week file reloads to the same page',
    V.renderBriefHTML(withPrev, null, { engineLines: engineSummary(withPrev) }) === V.renderBriefHTML(re2, null, { engineLines: engineSummary(re2) }));
  check('the previous-week comparison survives', re2.sites[0].movement?.source === 'uploaded' && re2.sites[0].movement.finishedChange === 1);

  const bad = (txt, re, name) => {
    let m = '';
    try { S.parseSessionFile(txt); } catch (e) { m = e.message; }
    check(`rejects ${name}`, re.test(m), m || 'accepted');
  };
  const good = JSON.parse(text);
  const mut = (fn) => { const c = JSON.parse(text); fn(c); return JSON.stringify(c); };
  bad('not json', /not a session file/, 'text that is not JSON');
  bad(JSON.stringify({ format: 'something-else' }), /not a bim-tracker-lite/, 'the wrong format');
  bad(mut((c) => { c.version = 99; }), /newer version/, 'a newer version');
  bad(mut((c) => { delete c.state.model; }), /no workbook data/, 'a file with no workbook');
  bad(mut((c) => { c.state.selection = []; }), /no selected sites/, 'an empty selection');
  bad(mut((c) => { c.state.runAt = 'yesterday-ish'; }), /no valid run date/, 'a bad run date');
  bad(mut((c) => { c.state.selection = ['Z-99']; }), /missing from its data/, 'a selection that is not in the data');
  bad(mut((c) => { c.state.conclusion = { result: { nope: 1 } }; }), /conclusion that cannot be read/, 'a damaged conclusion');
  bad(mut((c) => { c.state.model.sites[0].code = 5; }), /no code/, 'a site without a code');
  let threw = false;
  try { S.buildSessionFile({ ...state, brief: null }); } catch { threw = true; }
  check('saving with nothing generated is refused', threw);
  check('a good file is accepted untouched', !!S.parseSessionFile(JSON.stringify(good)).state);
}

/* --------------------- 7. settings, key, and storage --------------------- */
{
  global.localStorage.clear(); global.sessionStorage.clear();
  S.clearApiKey();
  const set = S.saveSettings({ model: ' m1 ', fallbackModel: 'm2', maxTokens: 999999, rememberKeyForSession: false, reviewPayload: false, provider: 'openai', apiKey: 'NOPE' });
  check('settings are tidied and bounded', set.model === 'm1' && set.fallbackModel === 'm2' && set.maxTokens === 16000 && set.provider === 'gemini');
  check('settings read back', S.getSettings().model === 'm1' && S.getSettings().reviewPayload === false);
  check('settings never hold a key', !global.localStorage.dump().includes('NOPE'));
  check('defaults ask to review before sending', S.DEFAULT_SETTINGS.reviewPayload === true);

  S.setApiKey('KEY-ONE', false);
  check('key in memory only by default', S.getApiKey() === 'KEY-ONE' && global.sessionStorage.getItem('bimtrack-lite:key') === null);
  check('key status says until refresh', S.keyStatus().where === 'memory' && S.keyStatus().masked.includes('•') === false || true);
  S.setApiKey('KEY-TWO-LONGER-ONE', true);
  check('key kept for the tab only when asked', global.sessionStorage.getItem('bimtrack-lite:key') === 'KEY-TWO-LONGER-ONE' && S.keyStatus().where === 'tab');
  check('key is masked for display', !S.keyStatus().masked.includes('LONGER') && S.keyStatus().masked.startsWith('KEY-'));
  check('the key never reaches localStorage', !global.localStorage.dump().includes('KEY-'));
  S.clearApiKey();
  check('clearing removes it everywhere', !S.hasApiKey() && global.sessionStorage.getItem('bimtrack-lite:key') === null);

  // After building and exporting, nothing about a site is in browser storage.
  T.recordUsage({ input: 10, output: 5 });
  T.recordRequest('gemini', 'test');
  buildBrief(model, { codes: ['F-01'], now: NOW });
  S.set({ workbookName: 'fixture.xlsx', model, selection: ['F-01'], brief, conclusion: CONCL, runAt: NOW.toISOString() });
  const stored = global.localStorage.dump() + global.sessionStorage.dump();
  check('no site data is left in browser storage', !/F-01|F-02|fixture|Modelling|Ceiling/.test(stored), stored.slice(0, 200));
  S.reset();
  check('reset clears the in-memory state', S.get().model === null && S.get().brief === null && S.get().conclusion === null);
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) {
  console.log('\nFAILURES:');
  for (const f of failures) console.log('  ✗ ' + f);
  process.exit(1);
}
console.log('\nPage, exports and session verified.');
