/**
 * Session, settings and the API key.
 *
 * WHAT PERSISTS, AND WHERE
 *   Project data   Nowhere. It lives in this module's memory only. A page reload
 *                  or closing the tab leaves nothing behind. The only way to keep
 *                  it is "Save session", which writes a file you control.
 *   Settings       localStorage: model names, token limit, review-before-send.
 *                  None of it is project data.
 *   API key        Memory only, unless the user ticks "keep through a refresh",
 *                  which uses sessionStorage (gone when the tab closes). Never in
 *                  a session file, never in an export, never in localStorage.
 *
 * The session file holds only the sites that were selected for the brief, so a
 * workbook of twenty sites does not leave a file about twenty.
 */

import { downloadJSON, pickFile } from './util.js';

export const SESSION_FORMAT = 'bim-tracker-lite.session';
export const SESSION_VERSION = 1;

/** Any key matching these never reaches a portable file. */
const FORBIDDEN = [/api[_-]?key/i, /\bsecret\b/i, /password/i, /bearer/i, /\bcredential/i];

export function stripSecrets(obj) {
  if (obj == null) return obj;
  if (Array.isArray(obj)) return obj.map(stripSecrets);
  if (typeof obj !== 'object') return obj;
  const out = {};
  for (const k of Object.keys(obj)) {
    if (FORBIDDEN.some((re) => re.test(k))) continue;
    out[k] = stripSecrets(obj[k]);
  }
  return out;
}

/* ------------------------------ live state ------------------------------ */

export function blank() {
  return {
    workbookName: null, model: null,
    previousName: null, previousModel: null,
    selection: [], runAt: null,
    brief: null,          // derived; never saved
    conclusion: null,     // { result (validated), meta }
  };
}

let state = blank();
const listeners = new Set();
export const get = () => state;
export const subscribe = (fn) => { listeners.add(fn); return () => listeners.delete(fn); };
export function set(patch) {
  state = { ...state, ...patch };
  for (const fn of listeners) fn(state);
  return state;
}
export function reset() { return set(blank()); }

/* ------------------------------ session file ------------------------------ */

/** A model reduced to the selected sites. Drops the register and other sites. */
export function scopedModel(model, codes) {
  if (!model) return null;
  const keep = new Set(codes);
  return {
    file: model.file, parsedAt: model.parsedAt,
    template: model.template,
    sites: (model.sites || []).filter((s) => keep.has(s.code)),
    problems: [], notes: [], orphanSheets: [],
  };
}

export function buildSessionFile(s = state, now = new Date()) {
  if (!s.model || !s.brief) throw new Error('There is nothing to save yet. Generate a brief first.');
  const codes = s.selection;
  return {
    format: SESSION_FORMAT,
    version: SESSION_VERSION,
    app: 'PM-Site-tracker-Lite',
    savedAt: now.toISOString(),
    note: 'Contains project data for the selected sites. Keep it private. The API key is never included.',
    state: stripSecrets({
      workbookName: s.workbookName,
      model: scopedModel(s.model, codes),
      previousName: s.previousName,
      previousModel: scopedModel(s.previousModel, codes),
      selection: codes,
      runAt: s.runAt,
      conclusion: s.conclusion,
    }),
  };
}

/** Throws a plain-language error if the text is not a usable session. */
export function parseSessionFile(text) {
  let parsed;
  try { parsed = JSON.parse(text); } catch {
    throw new Error('That is not a session file. Session files are the JSON made by "Save session". An export cannot be loaded back.');
  }
  if (!parsed || parsed.format !== SESSION_FORMAT) {
    throw new Error(`That file is not a ${SESSION_FORMAT} session. Check you picked the right one.`);
  }
  if (parsed.version > SESSION_VERSION) {
    throw new Error(`That session was saved by a newer version of the app (v${parsed.version}). Update before opening it.`);
  }
  const st = parsed.state;
  if (!st || typeof st !== 'object') throw new Error('That session file has no state in it.');
  const m = st.model;
  if (!m || !Array.isArray(m.sites) || !m.sites.length || !m.template) {
    throw new Error('That session file has no workbook data in it.');
  }
  if (!m.sites.every((s) => s && typeof s.code === 'string')) throw new Error('That session file is damaged: a site has no code.');
  if (!Array.isArray(st.selection) || !st.selection.length) throw new Error('That session file records no selected sites.');
  if (Number.isNaN(Date.parse(st.runAt))) throw new Error('That session file has no valid run date.');
  const have = new Set(m.sites.map((s) => s.code));
  if (!st.selection.every((c) => have.has(c))) throw new Error('That session file is damaged: a selected site is missing from its data.');
  if (st.conclusion && !(st.conclusion.result && Array.isArray(st.conclusion.result.sites))) {
    throw new Error('That session file has a conclusion that cannot be read.');
  }
  return { state: stripSecrets(st), savedAt: parsed.savedAt };
}

export function saveSessionToFile() {
  const payload = buildSessionFile();
  const stamp = new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-');
  const name = `site-brief-session-${stamp}.session.json`;
  downloadJSON(payload, name);
  return name;
}

export async function loadSessionFromFile() {
  const file = await pickFile('.json,application/json');
  if (!file) return { cancelled: true };
  const { state: st, savedAt } = parseSessionFile(await file.text());
  return { ok: true, filename: file.name, savedAt, state: st };
}

/* -------------------------------- settings -------------------------------- */

const SETTINGS_KEY = 'bimtrack-lite:settings:v1';
export const DEFAULT_SETTINGS = {
  provider: 'gemini',
  model: '',
  fallbackModel: '',
  maxTokens: 4096,
  rememberKeyForSession: false,
  reviewPayload: true,
};

export function getSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (raw) return { ...DEFAULT_SETTINGS, ...JSON.parse(raw), provider: 'gemini' };
  } catch { /* private mode or corrupt: defaults */ }
  return { ...DEFAULT_SETTINGS };
}

export function saveSettings(patch) {
  const next = { ...getSettings(), ...patch, provider: 'gemini' };
  // Only these fields. Nothing else, and never a key.
  const clean = {
    model: String(next.model || '').trim(),
    fallbackModel: String(next.fallbackModel || '').trim(),
    maxTokens: Math.min(16000, Math.max(512, Number(next.maxTokens) || 4096)),
    rememberKeyForSession: !!next.rememberKeyForSession,
    reviewPayload: next.reviewPayload !== false,
  };
  try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(clean)); } catch { /* memory only */ }
  return { ...DEFAULT_SETTINGS, ...clean };
}

/* ------------------------------- API key ------------------------------- */

const KEY_SESSION = 'bimtrack-lite:key';
let memoryKey = '';

export function setApiKey(key, remember) {
  memoryKey = key || '';
  try {
    if (remember && memoryKey) sessionStorage.setItem(KEY_SESSION, memoryKey);
    else sessionStorage.removeItem(KEY_SESSION);
  } catch { /* private mode: memory only */ }
}

export function getApiKey() {
  if (memoryKey) return memoryKey;
  try {
    const k = sessionStorage.getItem(KEY_SESSION);
    if (k) { memoryKey = k; return k; }
  } catch { /* ignore */ }
  return '';
}

export function clearApiKey() {
  memoryKey = '';
  try { sessionStorage.removeItem(KEY_SESSION); } catch { /* ignore */ }
}

export const hasApiKey = () => !!getApiKey();

/** Where the key lives, in plain words, so a blank-looking box is not mistaken for "no key". */
export function keyStatus() {
  const key = getApiKey();
  if (!key) return { set: false, where: 'none', lasts: 'No key saved. Getting a conclusion will ask for one.' };
  let inSession = false;
  try { inSession = !!sessionStorage.getItem(KEY_SESSION); } catch { /* private mode */ }
  return inSession
    ? { set: true, where: 'tab', masked: maskKey(key), lasts: 'Kept until you close this tab. Survives a page refresh.' }
    : { set: true, where: 'memory', masked: maskKey(key), lasts: 'Kept until you refresh or close the page.' };
}

export function maskKey(key) {
  if (!key) return '';
  if (key.length <= 10) return '•'.repeat(key.length);
  return `${key.slice(0, 4)}${'•'.repeat(Math.min(16, key.length - 8))}${key.slice(-4)}`;
}
