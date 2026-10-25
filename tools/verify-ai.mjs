/**
 * AI layer verification. No network and no key: fetch is stubbed.
 *
 * Checks the boundary the app depends on:
 *   - the model is given figures for the selected sites only, nothing else;
 *   - its reply is cut back to the selected sites before it reaches the page;
 *   - one request covers every site;
 *   - the retry, rate-limit, bad-key, fallback and cancel rules behave;
 *   - the API key travels in a header, never in the request body.
 *
 * Run: node tools/verify-ai.mjs
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
  return { getItem: (k) => (m.has(k) ? m.get(k) : null), setItem: (k, v) => m.set(k, String(v)), removeItem: (k) => m.delete(k), clear: () => m.clear() };
};
global.localStorage = memStore();
global.sessionStorage = memStore();

const { readWorkbook, parseWorkbook, STATUS } = await import('../js/core/parser.js');
const { buildBrief } = await import('../js/core/brief.js');
const AI = await import('../js/core/ai.js');
const S = await import('../js/core/session.js');
const T = await import('../js/core/tokens.js');

let pass = 0, fail = 0;
const failures = [];
const check = (name, cond, detail = '') => {
  if (cond) pass++;
  else { fail++; failures.push(`${name}${detail ? ` — ${detail}` : ''}`); }
};

const buf = XLSX.write(fixtureWorkbook(XLSX, STATUS), { type: 'array', bookType: 'xlsx' });
const model = parseWorkbook(await readWorkbook({ name: 'f.xlsx', arrayBuffer: async () => buf }));
const NOW = new Date(2026, 1, 16, 9, 30);
const briefOne = buildBrief(model, { codes: ['F-01'], now: NOW });
const briefBoth = buildBrief(model, { codes: ['F-01', 'F-02'], now: NOW });

/* --------------------------- fetch stub --------------------------- */
const okBody = (obj) => ({
  candidates: [{ content: { parts: [{ text: typeof obj === 'string' ? obj : JSON.stringify(obj) }] } }],
  usageMetadata: { promptTokenCount: 100, candidatesTokenCount: 50, totalTokenCount: 150 },
});
const res = (status, body, headers = {}) => ({
  ok: status >= 200 && status < 300, status,
  headers: { get: (k) => headers[String(k).toLowerCase()] ?? null },
  text: async () => (typeof body === 'string' ? body : JSON.stringify(body)),
  json: async () => body,
});
const GOOD = { headline: 'F-01 is behind.', sites: [{ site: 'F-01', verdict: 'Will miss target', reason: 'Pace is 0.61.', action: 'Clear the two stuck tasks.' }], reallocation: null };
let calls = [];
const stub = (handler) => {
  calls = [];
  global.fetch = async (url, opts) => { calls.push({ url, opts, body: opts?.body ? JSON.parse(opts.body) : null }); return handler(calls.length, url, opts); };
};
const SETTINGS = { provider: 'gemini', model: 'model-a', fallbackModel: '', maxTokens: 2048 };
const reset = () => { AI.clearOverloadMemory(); AI._setOverloadWaits([1, 1, 1, 1]); S.setApiKey('TEST-KEY-123', false); T.resetUsage(); T.resetRequests('gemini'); };

/* ---------------------------- 1. the payload ---------------------------- */
{
  const p = AI.buildConclusionPayload(briefOne);
  const s = p.sites[0];
  check('payload names the site', s.site === 'F-01' && p.siteCodes.join() === 'F-01');
  check('payload carries the counts', s.tasks.finished === 4 && s.tasks.wip === 2 && s.tasks.blocked === 1 && s.tasks.waiting === 1);
  check('payload carries the pace and rating', p.overall.pace === 'Poor' && s.rating === 'RED' && s.paceVsPlan === 0.61);
  check('payload carries the stuck tasks with reasons', s.blockedOrWaitingTasks.length === 2 && s.blockedOrWaitingTasks.some((t) => t.reason === 'Ceiling void zones not confirmed'));
  check('payload carries the pending prerequisite', s.prerequisitesPending.length === 1 && s.prerequisitesPending[0].id === 'P02');
  check('payload lists only unfinished groups', s.unfinishedGroups.length === 1 && s.unfinishedGroups[0].group === 'Modelling');
  check('a one-site payload does not contain the other site', !JSON.stringify(p).includes('F-02'));
  const both = AI.buildConclusionPayload(briefBoth);
  check('a two-site payload covers both', both.sites.length === 2 && both.siteCodes.join() === 'F-01,F-02');
  check('a site with no data is passed as such, not invented', both.sites[1].noData === true && /no detail sheet/.test(both.sites[1].reason));
  const prev = AI.previewPayload({ payload: p, settings: SETTINGS });
  check('preview shows destination and model', prev.destination === 'Google Gemini' && prev.model === 'model-a');
  check('preview prompt contains the data', prev.prompt.includes('"F-01"') && prev.prompt.includes('DATA'));
  check('system prompt forbids recalculation', /Do NOT recalculate/.test(prev.system));
  check('system prompt confines the model to the listed sites', /Mention ONLY the sites listed/.test(prev.system));
  check('preview estimates size', prev.estimatedInputTokens > 100 && prev.bytes > 500);
}

/* ------------------------ 2. checking the reply ------------------------ */
{
  const others = ['F-02'];
  const v = AI.validateConclusion(GOOD, ['F-01'], others);
  check('a good reply passes untouched', v.issues.length === 0 && v.sites.length === 1 && v.headline === 'F-01 is behind.');

  const stray = AI.validateConclusion({ ...GOOD, sites: [...GOOD.sites, { site: 'F-02', verdict: 'On track', reason: 'fine', action: null }] }, ['F-01'], others);
  check('an entry for an unselected site is dropped', stray.sites.length === 1 && stray.sites[0].site === 'F-01');
  check('and the drop is reported', stray.issues.some((i) => /F-02/.test(i) && /dropped/.test(i)));

  const leakText = AI.validateConclusion({ ...GOOD, reallocation: 'Move R2 from F-02 to F-01.' }, ['F-01'], others);
  check('advice that names an unselected site is removed', leakText.reallocation === null && leakText.issues.some((i) => /reallocation/.test(i) && /F-02/.test(i)));
  const leakHead = AI.validateConclusion({ ...GOOD, headline: 'Unlike F-02, F-01 is late.' }, ['F-01'], others);
  check('a headline that names an unselected site is removed', leakHead.headline === '');
  const leakReason = AI.validateConclusion({ headline: 'x', sites: [{ site: 'F-01', verdict: 'At risk', reason: 'Worse than F-02.', action: null }] }, ['F-01'], others);
  check('a reason that names an unselected site is blanked', leakReason.sites[0].reason === '');
  const near = AI.validateConclusion({ headline: 'Item F-021 is unrelated.', sites: GOOD.sites }, ['F-01'], others);
  check('a longer code that merely starts with another is not a leak', near.headline === 'Item F-021 is unrelated.');

  const missing = AI.validateConclusion({ headline: 'x', sites: [] }, ['F-01'], others);
  check('a site the model skipped is reported', missing.issues.some((i) => /no conclusion for F-01/.test(i)));

  const odd = AI.validateConclusion({ headline: 'x', sites: [{ site: 'F-01', verdict: 'running late', reason: 'r', action: 'null' }] }, ['F-01'], others);
  check('an unexpected verdict is mapped, and says so', odd.sites[0].verdict === 'Will miss target' && odd.issues.some((i) => /running late/.test(i)));
  check('the text "null" as an action means no action', odd.sites[0].action === null);
  const vague = AI.validateConclusion({ headline: 'x', sites: [{ site: 'F-01', verdict: '???', reason: 'r' }] }, ['F-01'], others);
  check('an unreadable verdict falls back to "Cannot tell"', vague.sites[0].verdict === 'Cannot tell');

  const long = AI.validateConclusion({ headline: 'h'.repeat(900), sites: [{ site: 'F-01', verdict: 'On track', reason: 'r'.repeat(900) }] }, ['F-01'], others);
  check('over-long text is clipped', long.headline.length <= 300 && long.sites[0].reason.length <= 400);
  const dup = AI.validateConclusion({ headline: 'x', sites: [GOOD.sites[0], { ...GOOD.sites[0], reason: 'second' }] }, ['F-01'], others);
  check('a duplicate site entry is ignored', dup.sites.length === 1 && dup.sites[0].reason === 'Pace is 0.61.');
  let threw = false;
  try { AI.validateConclusion('nonsense', ['F-01'], others); } catch { threw = true; }
  check('a reply that is not an object is rejected', threw);
  check('parseModelJson reads fenced JSON', AI.parseModelJson('```json\n{"a":1}\n```').a === 1);
  check('parseModelJson reads JSON wrapped in prose', AI.parseModelJson('Here you go: {"a":2} thanks').a === 2);
}

/* ----------------------------- 3. calling ----------------------------- */
{
  reset();
  stub(() => res(200, okBody(GOOD)));
  const payload = AI.buildConclusionPayload(briefBoth);
  const out = await AI.run({ payload, settings: SETTINGS });
  check('one request covers both sites', calls.length === 1, String(calls.length));
  check('the request goes to the chosen model', calls[0].url.includes('/models/model-a:generateContent'));
  check('the key is sent as a header', calls[0].opts.headers['x-goog-api-key'] === 'TEST-KEY-123');
  check('the key is not in the request body', !JSON.stringify(calls[0].body).includes('TEST-KEY-123'));
  check('the key is not in the URL', !calls[0].url.includes('TEST-KEY-123'));
  check('the body carries both sites', JSON.stringify(calls[0].body).includes('F-02'));
  check('the reply is parsed', out.result.headline === 'F-01 is behind.' && out.model === 'model-a' && out.fellBack === false);
  check('provider-reported tokens are used', out.tokens.input === 100 && out.tokens.output === 50 && out.tokens.estimated === false, JSON.stringify(out.tokens));
  check('usage is recorded', T.getUsage().total === 150 && T.getUsage().calls === 1, JSON.stringify(T.getUsage()));
}
{
  reset();
  S.clearApiKey();
  stub(() => res(200, okBody(GOOD)));
  let msg = '';
  try { await AI.run({ payload: {}, settings: SETTINGS }); } catch (e) { msg = e.message; }
  check('no key: refuses before any request', /No API key/.test(msg) && calls.length === 0);
}
{
  reset();
  stub(() => res(429, { error: { message: 'quota' } }, { 'retry-after': '30' }));
  let e1;
  try { await AI.run({ payload: {}, settings: SETTINGS }); } catch (e) { e1 = e; }
  check('429 fails at once, with no retry', calls.length === 1 && e1?.status === 429, `${calls.length} calls`);
  check('429 says it is your limit and honours Retry-After', /your account's limit/.test(e1?.message) && /30s/.test(e1?.message), e1?.message);
}
{
  reset();
  stub(() => res(401, { error: { message: 'bad key' } }));
  let e1;
  try { await AI.run({ payload: {}, settings: SETTINGS }); } catch (e) { e1 = e; }
  check('401 says the key was rejected, no retry', calls.length === 1 && /rejected the API key/.test(e1?.message));
  stub(() => res(404, { error: { message: 'no model' } }));
  try { await AI.run({ payload: {}, settings: SETTINGS }); } catch (e) { e1 = e; }
  check('404 says the model name is stale', /does not recognise the model name/.test(e1?.message) && e1.status === 404);
}
{
  reset();
  stub(() => res(503, { error: { message: 'overloaded' } }));
  let e1;
  try { await AI.run({ payload: {}, settings: SETTINGS }); } catch (e) { e1 = e; }
  check('overload retries five times in all', calls.length === 5, String(calls.length));
  check('overload error is flagged and explained', e1?.overloaded === true && /not your quota/.test(e1.message));
  check('every retry counts as a request', T.requestWindows('gemini').lastMinute === 5, JSON.stringify(T.requestWindows('gemini')));
}
{
  reset();
  stub((n) => (n < 3 ? res(503, { error: { message: 'busy' } }) : res(200, okBody(GOOD))));
  const statuses = [];
  const out = await AI.run({ payload: {}, settings: SETTINGS, onStatus: (s) => { if (s) statuses.push(s); } });
  check('it recovers when the server comes back', out.result.headline === 'F-01 is behind.' && calls.length === 3);
  check('progress is reported while waiting', statuses.length >= 2 && statuses[0].attempt === 2 && /busy/.test(statuses[0].why), JSON.stringify(statuses[0]));
}
{
  // Fallback model
  reset();
  const withFb = { ...SETTINGS, fallbackModel: 'model-b' };
  stub((n, url) => (url.includes('model-a') ? res(503, { error: { message: 'busy' } }) : res(200, okBody(GOOD))));
  const out = await AI.run({ payload: {}, settings: withFb });
  const aCalls = calls.filter((c) => c.url.includes('model-a')).length;
  check('with a fallback, the main model gets only two tries', aCalls === 2, String(aCalls));
  check('the fallback answers and the page is told', out.model === 'model-b' && out.fellBack === true && out.requestedModel === 'model-a');
  const marks = AI.overloadedModels();
  check('the overloaded model is remembered', marks.length === 1 && marks[0].model === 'model-a');
  // second run skips the primary
  stub(() => res(200, okBody(GOOD)));
  const out2 = await AI.run({ payload: {}, settings: withFb });
  check('a remembered overload skips straight to the fallback', calls.length === 1 && calls[0].url.includes('model-b') && out2.skippedPrimary === true);
}
{
  reset();
  const withFb = { ...SETTINGS, fallbackModel: 'model-b' };
  stub(() => res(503, { error: { message: 'busy' } }));
  let e1;
  try { await AI.run({ payload: {}, settings: withFb }); } catch (e) { e1 = e; }
  check('when both models are down it says both were tried', /Both models are overloaded/.test(e1?.message) && e1.bothModels?.join() === 'model-a,model-b', e1?.message);
}
{
  reset();
  stub((n, url) => (url.includes('model-a') ? res(401, { error: { message: 'bad' } }) : res(200, okBody(GOOD))));
  let e1;
  try { await AI.run({ payload: {}, settings: { ...SETTINGS, fallbackModel: 'model-b' } }); } catch (e) { e1 = e; }
  check('a bad key does not switch models', /rejected the API key/.test(e1?.message) && calls.length === 1);
}
{
  reset();
  stub(() => res(200, okBody(GOOD)));
  const ctrl = new AbortController();
  ctrl.abort();
  let e1;
  try { await AI.run({ payload: {}, settings: SETTINGS, signal: ctrl.signal }); } catch (e) { e1 = e; }
  check('a cancel before sending makes no request', e1 instanceof AI.CancelledError && calls.length === 0);
}
{
  reset();
  stub(() => res(200, { candidates: [{ content: { parts: [{ text: '' }] } }] }));
  let e1;
  try { await AI.run({ payload: {}, settings: SETTINGS }); } catch (e) { e1 = e; }
  check('an empty reply is reported, not treated as success', /empty response/.test(e1?.message));
  stub(() => res(200, okBody('this is not json at all')));
  try { await AI.run({ payload: {}, settings: SETTINGS }); } catch (e) { e1 = e; }
  check('a reply that is not JSON is reported', /valid JSON/.test(e1?.message));
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) {
  console.log('\nFAILURES:');
  for (const f of failures) console.log('  ✗ ' + f);
  process.exit(1);
}
console.log('\nAI layer verified.');
