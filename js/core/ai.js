/**
 * AI layer (Gemini only).
 *
 * The model receives COMPUTED RESULTS, never the spreadsheet. Every count,
 * percentage, forecast and rating is calculated in brief.js / engine.js and
 * passed in as a fixed fact. The model writes the CONCLUSION only: will the
 * sites finish on time, and where is a change of strategy or a move of
 * resources needed.
 *
 * Asking a model to compute a forecast from a status grid produces numbers that
 * look right and are not, and someone would act on them. So the boundary is
 * enforced here (the prompt forbids recalculation) and after the call
 * (validateConclusion drops anything about a site that was not selected).
 *
 * Transport, retry, overload memory and the fallback model are the full app's,
 * unchanged apart from the Gemini-only trim. One request covers all selected
 * sites, so a run costs one request, not one per site.
 */

import { getApiKey } from './session.js';
import {
  recordRequest, recordUsage, readProviderUsage, estimateTokens,
  readRateLimitHeaders,
} from './tokens.js';

export const PROVIDERS = {
  gemini: {
    label: 'Google Gemini',
    keyHint: 'Paste your Google AI Studio key',
    keyUrl: 'https://aistudio.google.com/apikey',
    defaultModel: 'gemini-2.5-flash',
    suggestions: ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.0-flash'],
    note: 'Calls go directly from your browser to Google.',
  },
};

/* ============================ system prompt ============================ */

const SYSTEM = `You are a delivery manager reviewing progress on several construction or BIM-delivery sites. Your reader is a project manager or coordinator: short of time, accountable for the dates, not necessarily a BIM specialist.

Everything you are given was calculated by software from the project's tracker workbook. Treat every number as fixed fact.
- Do NOT recalculate, re-derive or second-guess any figure, date, count or rating.
- Do NOT invent a figure, a date, a person, a task or a cause that is not in the data.
- Mention ONLY the sites listed in the data. Never refer to, compare with or speculate about any other site.
- Resources appear as codes such as R1. Use the codes as given.
- If the data is too thin to judge a site, say "Cannot tell" and say what is missing.
- Plain, short sentences. No jargon without a word of explanation.

Your job is the conclusion only: will each site finish on time, and is a change of strategy or a move of resources needed, and where.`;

/* ============================ the payload ============================ */

const CAP = { groups: 15, stuck: 12, prereq: 10, resources: 8 };

/**
 * The facts the model is given: figures from the brief, trimmed. Free text
 * (task names, blocked reasons) is included; the privacy scan runs before this
 * on the same workbook, and the user reviews this exact text before it is sent.
 */
export function buildConclusionPayload(brief) {
  const sites = brief.sites.map((s) => {
    if (s.noData) return { site: s.code, noData: true, reason: s.reason };
    return {
      site: s.code,
      description: s.description || undefined,
      priority: s.priority || undefined,
      rating: s.overall,
      statusAsAt: s.statusDate,
      week: s.reportingWeek,
      tasks: { ...s.counts },
      percentCompleteByWeight: s.pctByWeight,
      percentTimeElapsed: s.timeElapsed,
      paceVsPlan: s.spiProxy,
      taskGroups: { total: s.groups.total, finished: s.groups.complete },
      unfinishedGroups: s.groups.rows.filter((g) => !g.complete).slice(0, CAP.groups).map((g) => ({
        group: g.name, done: g.done, of: g.live, wip: g.wip, blockedOrWaiting: g.blocked + g.waiting,
      })),
      target: s.schedule.target,
      forecast: s.schedule.forecastDate,
      forecastMethod: s.schedule.forecastMethod,
      daysForecastIsAfterTarget: s.schedule.varianceDays,
      prerequisitesPending: s.prerequisites.pending.slice(0, CAP.prereq).map((p) => ({
        id: p.id, name: p.name, from: p.provider, neededByWeek: p.byWeek, status: p.status,
      })),
      prerequisitesWithNoRecord: s.prerequisites.notTracked.length,
      blockedOrWaitingTasks: s.blocked.slice(0, CAP.stuck).map((t) => ({
        task: t.taskId, name: t.name, status: t.status, weeksStuck: t.weeksStuck,
        owner: t.resource, reason: t.reason, waitingOn: t.waitingOn,
      })),
      resourceLoad: s.resources.rows.slice(0, CAP.resources),
      sinceLastWeek: s.movement?.text || null,
    };
  });
  return {
    asAt: brief.title.reportDate,
    siteCodes: brief.title.sites,
    overall: {
      pace: brief.portfolio.pace.word,
      paceRatio: brief.portfolio.pace.ratio,
      percentCompleteAllSites: brief.portfolio.overallPctByWeight,
    },
    sites,
  };
}

const SCHEMA = `{
  "headline": "one sentence: the overall position and the single most important thing to do",
  "sites": [
    {
      "site": "the site code exactly as given",
      "verdict": "On track | At risk | Will miss target | Cannot tell",
      "reason": "one or two short sentences, grounded in the data",
      "action": "the one action that would help most, or null if none is needed"
    }
  ],
  "reallocation": "if moving resources between the listed sites would help, say which resource codes, from which site to which, and why; otherwise null. Never suggest a move the data does not support."
}`;

export function buildPrompt(payload) {
  return `Write the conclusion for this one-page site progress brief.

Return ONLY a JSON object of exactly this shape, with one entry in "sites" for every site in the data, in the same order:
${SCHEMA}

Keep "reason" and "action" under 35 words each and "headline" under 40 words.

DATA (calculated by software; fixed facts):
${JSON.stringify(payload)}`;
}

export function previewPayload({ payload, settings }) {
  const prompt = buildPrompt(payload);
  return {
    destination: PROVIDERS.gemini.label,
    model: (settings.model || '').trim() || PROVIDERS.gemini.defaultModel,
    system: SYSTEM,
    prompt,
    payload,
    estimatedInputTokens: estimateTokens(SYSTEM) + estimateTokens(prompt),
    bytes: new Blob([prompt]).size,
  };
}

/* ====================== checking what comes back ====================== */

const VERDICTS = ['On track', 'At risk', 'Will miss target', 'Cannot tell'];
const clip = (v, n) => {
  const s = typeof v === 'string' ? v.trim() : '';
  return s.length > n ? `${s.slice(0, n - 1).trimEnd()}…` : s;
};
const normVerdict = (v) => {
  const t = String(v || '').toLowerCase();
  if (/on\s*track|on\s*time|green/.test(t)) return 'On track';
  if (/will\s*miss|late|behind|miss|red/.test(t)) return 'Will miss target';
  if (/at\s*risk|amber|risk|watch/.test(t)) return 'At risk';
  return 'Cannot tell';
};
const hasCode = (text, code) =>
  new RegExp(`(^|[^A-Za-z0-9])${String(code).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![A-Za-z0-9])`, 'i').test(text);

/**
 * The model is not trusted to stay in scope. This keeps only what was asked for:
 * entries for the selected sites, text that names no unselected site, bounded
 * lengths. Anything removed is reported in `issues` so the page can say so.
 *
 * @param {object} result       parsed model JSON
 * @param {string[]} selected   site codes in the brief
 * @param {string[]} others     every other site code in the workbook
 */
export function validateConclusion(result, selected, others = []) {
  const issues = [];
  if (!result || typeof result !== 'object' || Array.isArray(result)) {
    throw new Error('The model\'s reply was not in the expected shape. Try again.');
  }
  const leaks = (text) => others.find((c) => hasCode(text || '', c));
  const safeText = (text, where, max) => {
    const t = clip(text, max);
    const leak = t && leaks(t);
    if (leak) { issues.push(`${where} mentioned a site that is not in this brief (${leak}) and was removed.`); return ''; }
    return t;
  };

  const headline = safeText(result.headline, 'The headline', 300);
  const seen = new Set();
  const sites = [];
  for (const e of Array.isArray(result.sites) ? result.sites : []) {
    const code = String(e?.site ?? '').trim();
    if (!selected.includes(code)) {
      if (code) issues.push(`The model wrote about ${code}, which is not in this brief. That entry was dropped.`);
      continue;
    }
    if (seen.has(code)) continue;
    seen.add(code);
    const rawVerdict = String(e.verdict ?? '').trim();
    const verdict = VERDICTS.includes(rawVerdict) ? rawVerdict : normVerdict(rawVerdict);
    if (verdict !== rawVerdict) issues.push(`${code}: the verdict "${rawVerdict || 'blank'}" was read as "${verdict}".`);
    sites.push({
      site: code, verdict,
      reason: safeText(e.reason, `${code} reason`, 400),
      action: e.action == null || String(e.action).trim().toLowerCase() === 'null' ? null : (safeText(e.action, `${code} action`, 400) || null),
    });
  }
  for (const code of selected) if (!seen.has(code)) issues.push(`The model gave no conclusion for ${code}.`);
  const realloc = result.reallocation == null ? null : (safeText(result.reallocation, 'The reallocation advice', 600) || null);
  return { headline, sites, reallocation: realloc, issues };
}

/* ============================== transport ============================== */
function makeSignal(external, ms) {
  const ctrl = new AbortController();
  // If the caller already cancelled, abort before fetch attaches its listener
  // and fetch never rejects — so bail out explicitly instead.
  if (external?.aborted) { ctrl.abort(); return { signal: ctrl.signal, done: () => {}, preAborted: true }; }
  const timer = setTimeout(() => ctrl.abort(new Error('timeout')), ms);
  if (external) external.addEventListener('abort', () => ctrl.abort(new Error('cancelled')), { once: true });
  return { signal: ctrl.signal, done: () => clearTimeout(timer), preAborted: false };
}

export class CancelledError extends Error {
  constructor() { super('Cancelled.'); this.name = 'CancelledError'; }
}

function retryAfterSeconds(response, bodyText) {
  const h = response?.headers?.get?.('retry-after');
  if (h && Number.isFinite(Number(h))) return Number(h);
  const m = String(bodyText || '').match(/retry in ([\d.]+)\s*s/i);
  return m ? Math.ceil(Number(m[1])) : null;
}

const sleep = (ms, signal) => new Promise((res, rej) => {
  const t = setTimeout(res, ms);
  signal?.addEventListener('abort', () => { clearTimeout(t); rej(new CancelledError()); }, { once: true });
});

/**
 * One request, with retry.
 *
 * WHY 503s HAPPEN: a 503 (or Anthropic's 529) means the provider's servers for
 * that model are overloaded. It is not your quota and not your token count —
 * the request never got far enough to be counted against either. The newest
 * models are the most crowded, and free-tier traffic is the first to be turned
 * away when they are busy.
 *
 * The old version retried three times over about 3.6 seconds, which is far too
 * short: overload clears in tens of seconds, not three. This waits roughly 5s,
 * 12s, 25s, 45s, with random jitter so many clients don't all retry in lockstep,
 * and honours Retry-After when the server sends one. Every retry is a real
 * request against your per-minute limit, which is why this caps at five.
 *
 * 429 is different: it is YOUR rate limit, and retrying immediately just burns
 * more of it. That fails fast with an explanation.
 */
let OVERLOAD_WAITS = [5000, 12000, 25000, 45000];

/** Test hook only: the real waits total over a minute, far too slow for a suite. */
export function _setOverloadWaits(arr) { OVERLOAD_WAITS = arr; }

/** Pull the provider's own error text out of a response body, if it sent one. */
function providerMessage(body) {
  try {
    const j = JSON.parse(body);
    return j?.error?.message || j?.message || '';
  } catch { return String(body || '').slice(0, 200); }
}

async function callWithRetry(provider, doFetch, { signal, onStatus, maxAttempts: cap } = {}) {
  const maxAttempts = Math.min(cap || Infinity, OVERLOAD_WAITS.length + 1);
  let lastStatus = null;
  let lastMessage = '';

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    if (signal?.aborted) throw new CancelledError();
    recordRequest(provider, attempt === 0 ? 'first attempt' : `retry after ${lastStatus || 'error'}`);

    let response;
    try {
      response = await doFetch();
    } catch (e) {
      if (signal?.aborted) throw new CancelledError();
      if (e?.name === 'AbortError') throw new Error('The request timed out. Try selecting fewer sites.');
      lastStatus = 'network error';
      if (attempt === maxAttempts - 1) throw new Error(`Could not reach ${provider}. Check your connection. (${e.message})`);
      await waitWithStatus(OVERLOAD_WAITS[attempt], attempt, maxAttempts, 'connection failed', onStatus, signal);
      continue;
    }

    if (response.ok) { onStatus?.(null); return response; }

    const body = await response.text().catch(() => '');
    if (response.status === 429) {
      const wait = retryAfterSeconds(response, body);
      const e = new Error(
        `Rate limit reached (429) — this is your account's limit, not the server being busy.${wait ? ` The provider asked to wait about ${wait}s.` : ''} ` +
        'Retrying straight away would only burn more of your request quota. Wait a minute or two, then try again.');
      e.status = 429;
      throw e;
    }
    if (response.status === 401 || response.status === 403) {
      throw new Error(`${provider} rejected the API key (${response.status}). Check it is correct, active, and allowed for this model.`);
    }
    if (response.status === 404) {
      const e = new Error(`${provider} does not recognise the model name (404). Model names change — check the current list in Settings.`);
      e.status = 404;
      throw e;
    }
    if ([500, 502, 503, 504, 529].includes(response.status)) {
      lastStatus = response.status;
      lastMessage = providerMessage(body);
      if (attempt === maxAttempts - 1) {
        const waited = Math.round(OVERLOAD_WAITS.slice(0, maxAttempts - 1).reduce((a, b) => a + b, 0) / 1000);
        const e = new Error(
          `${provider} is still overloaded (${response.status}) after ${maxAttempts} attempt${maxAttempts === 1 ? '' : 's'}${waited ? ` over about ${waited}s` : ''}.` +
          `${lastMessage ? ` The provider said: "${lastMessage}"` : ''} ` +
          'This is the provider\'s servers being busy for this model, not your quota or tokens. Try again later, or switch to a more established model in Settings.');
        e.status = response.status;
        e.overloaded = true;
        e.providerMessage = lastMessage;
        throw e;
      }
      const hinted = retryAfterSeconds(response, body);
      const base = hinted ? hinted * 1000 : OVERLOAD_WAITS[attempt];
      await waitWithStatus(base, attempt, maxAttempts, `server busy (${response.status})`, onStatus, signal, lastMessage);
      continue;
    }
    let msg = body.slice(0, 300);
    try { msg = JSON.parse(body)?.error?.message || msg; } catch { /* keep raw */ }
    throw new Error(`${provider} returned ${response.status}: ${msg}`);
  }
  throw new Error('Request failed.');
}

/** Waits with jitter, reporting a live countdown so the user can see it is working. */
async function waitWithStatus(baseMs, attempt, maxAttempts, why, onStatus, signal, message = '') {
  const ms = Math.round(baseMs * (0.8 + Math.random() * 0.4));
  const until = Date.now() + ms;
  while (Date.now() < until) {
    if (signal?.aborted) { onStatus?.(null); throw new CancelledError(); }
    onStatus?.({
      why,
      attempt: attempt + 2,
      maxAttempts,
      secondsLeft: Math.ceil((until - Date.now()) / 1000),
      providerMessage: message,
    });
    await sleep(Math.min(1000, until - Date.now()), signal);
  }
}

async function callGemini({ model, prompt, signal, maxTokens, onStatus, maxAttempts }) {
  const key = getApiKey();
  const { signal: s, done } = makeSignal(signal, 120000);
  try {
    const res = await callWithRetry('gemini', () => fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`,
      {
        method: 'POST', signal: s,
        headers: { 'Content-Type': 'application/json', 'x-goog-api-key': key },
        body: JSON.stringify({
          systemInstruction: { parts: [{ text: SYSTEM }] },
          contents: [{ role: 'user', parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.3, responseMimeType: 'application/json', maxOutputTokens: maxTokens },
        }),
      }), { signal, onStatus, maxAttempts });
    const data = await res.json();
    const text = data?.candidates?.[0]?.content?.parts?.map((p) => p.text).filter(Boolean).join('') || '';
    if (!text && data?.promptFeedback?.blockReason) {
      throw new Error(`Gemini declined the request (${data.promptFeedback.blockReason}).`);
    }
    return { text, usage: data?.usageMetadata || null, headers: null };
  } finally { done(); }
}

/* ============================== parsing ============================== */
/* ============================== parsing ============================== */

export function parseModelJson(text) {
  if (!text) throw new Error('The model returned an empty response. This usually means the reply was cut short — try again, or reduce how much you run at once.');
  let t = String(text).trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
  try { return JSON.parse(t); } catch { /* try harder */ }
  const a = t.indexOf('{'), b = t.lastIndexOf('}');
  if (a >= 0 && b > a) {
    try { return JSON.parse(t.slice(a, b + 1)); } catch { /* fall through */ }
  }
  throw new Error('The model did not return valid JSON. Try again, or switch to a stronger model in Settings.');
}

/* ============================== orchestration ============================== */

/*
 * Overload memory ("circuit breaker").
 *
 * Without this, every site in a batch retried an overloaded model five times
 * before trying the fallback. On a ten-site run that spent about fifty of a
 * 200-a-day free allowance, and a minute and a half per site, on a model that
 * was never going to answer. Once a model has proved overloaded, the rest of
 * the run goes straight to the fallback, and it is only tried again once the
 * cool-off has passed.
 */
const OVERLOAD_COOLOFF_MS = 10 * 60 * 1000;
const overloadedUntil = new Map();   // model -> timestamp

export function isMarkedOverloaded(model, now = Date.now()) {
  const t = overloadedUntil.get(model);
  if (!t) return false;
  if (now >= t) { overloadedUntil.delete(model); return false; }
  return true;
}
export function overloadedModels(now = Date.now()) {
  return [...overloadedUntil.entries()]
    .filter(([m]) => isMarkedOverloaded(m, now))
    .map(([m, t]) => ({ model: m, minutesLeft: Math.ceil((t - now) / 60000) }));
}
function markOverloaded(model) { overloadedUntil.set(model, Date.now() + OVERLOAD_COOLOFF_MS); }

export function clearOverloadMemory(model) {
  if (model) overloadedUntil.delete(model); else overloadedUntil.clear();
}
/** Test hook. */
export function _markOverloaded(model, ms) { overloadedUntil.set(model, Date.now() + ms); }

/**
 * Run one analysis. Returns a report object ready to store.
 */
export async function run({ payload, settings, signal, onStatus }) {
  if (!getApiKey()) throw new Error('No API key set. Add one in Settings — it stays in this browser and is never written to a session file.');
  const provider = PROVIDERS.gemini;
  const model = (settings.model || '').trim() || provider.defaultModel;
  const maxTokens = Number(settings.maxTokens) || 4096;

  const prompt = buildPrompt(payload);
  const started = Date.now();

  let usedModel = model;
  let fellBack = false;
  let skippedPrimary = false;
  let out;
  const fallback = (settings.fallbackModel || '').trim();
  const hasFallback = !!fallback && fallback !== model;

  // A model already seen overloaded gets two tries, not five. Without this a
  // run where BOTH models are down spends about 95 seconds per site proving
  // something it already knew.
  const capFor = (mdl, withAlternative) =>
    (isMarkedOverloaded(mdl) ? 2 : (withAlternative ? 2 : undefined));

  if (hasFallback && isMarkedOverloaded(model) && !isMarkedOverloaded(fallback)) {
    skippedPrimary = true;
    onStatus?.({ why: `${model} is overloaded (seen in the last few minutes) — using ${fallback}`, attempt: 1, maxAttempts: 1, secondsLeft: 0 });
    try {
      out = await callGemini({ model: fallback, prompt, signal, maxTokens, onStatus, maxAttempts: capFor(fallback, false) });
    } catch (e) {
      if (e?.overloaded) markOverloaded(fallback);
      throw e;
    }
    usedModel = fallback;
    fellBack = true;
  } else {
    try {
      out = await callGemini({
        model, prompt, signal, maxTokens, onStatus, maxAttempts: capFor(model, hasFallback),
      });
    } catch (e) {
      if (e?.overloaded) markOverloaded(model);
      // Only an overload is worth a second model. A bad key, a 429 or a 404
      // would fail identically on any model, so those propagate unchanged.
      if (!e?.overloaded || !hasFallback || signal?.aborted) throw e;
      onStatus?.({ why: `${model} overloaded — switching to ${fallback}`, attempt: 1, maxAttempts: 1, secondsLeft: 0, providerMessage: e.providerMessage });
      try {
        out = await callGemini({ model: fallback, prompt, signal, maxTokens, onStatus, maxAttempts: capFor(fallback, false) });
      } catch (e2) {
        if (e2?.overloaded) markOverloaded(fallback);
        if (e2?.overloaded) {
          // Say plainly that BOTH were tried, so the reader does not think the
          // fallback was never reached.
          const both = new Error(
            `Both models are overloaded right now: ${model} and the fallback ${fallback}.`
            + `${e2.providerMessage ? ` The provider said: "${e2.providerMessage}".` : ''}`
            + ' This is the provider\'s capacity for these models, not your key, your quota or your token count.'
            + ' Wait a few minutes, or set a more established model in Settings.');
          both.overloaded = true;
          both.status = e2.status;
          both.bothModels = [model, fallback];
          both.providerMessage = e2.providerMessage;
          throw both;
        }
        throw e2;
      }
      usedModel = fallback;
      fellBack = true;
    }
  }
  // A success clears any stale mark on whichever model answered.
  overloadedUntil.delete(usedModel);
  const { text, usage, headers } = out;
  const result = parseModelJson(text);

  const reported = readProviderUsage(usage);
  const tokens = reported || {
    input: estimateTokens(SYSTEM) + estimateTokens(prompt),
    output: estimateTokens(text),
  };
  recordUsage({ ...tokens, estimated: !reported });

  return {
    id: `conclusion-${Date.now()}`,
    result,
    provider: 'gemini',
    providerLabel: provider.label,
    model: usedModel,
    requestedModel: model,
    fellBack,
    skippedPrimary,
    at: new Date().toISOString(),
    ms: Date.now() - started,
    tokens: { ...tokens, estimated: !reported },
    rateLimit: headers || null,
  };
}

