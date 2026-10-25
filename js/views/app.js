/**
 * The page. Load a workbook, check what was read, pick sites, get the brief,
 * optionally ask an AI for the conclusion, export.
 *
 * Nothing here calculates a figure: that is brief.js. Nothing here writes
 * project data to disk except the explicit "Save session" and the exports.
 */

import { el, mount, toast, confirmDialog, openModal, clamp } from '../core/util.js';
import { readWorkbook, parseWorkbook } from '../core/parser.js';
import { buildBrief, engineSummary } from '../core/brief.js';
import { renderBriefHTML, BRIEF_CSS } from '../core/briefview.js';
import {
  run, previewPayload, buildConclusionPayload, validateConclusion, PROVIDERS,
  overloadedModels, clearOverloadMemory, CancelledError,
} from '../core/ai.js';
import * as T from '../core/tokens.js';
import * as S from '../core/session.js';
import { exportPDF, exportWord, exportExcel } from '../core/exports.js';

const root = document.getElementById('app');
let preview = null;            // brief over every site, used only for the "what was read" checks
let pendingSel = null;         // Set of ticked site codes before Generate is pressed
let busy = null;               // { ctrl, status } while an AI call is running
let lastError = '';
let orientation = 'portrait';

if (!document.getElementById('brief-css')) {
  const st = document.createElement('style');
  st.id = 'brief-css';
  st.textContent = BRIEF_CSS;
  document.head.appendChild(st);
}

/* ============================== helpers ============================== */

const btn = (label, onclick, cls = '', attrs = {}) =>
  el('button', { class: `btn ${cls}`.trim(), type: 'button', onclick, ...attrs }, [label]);

const problemText = (p) => (typeof p === 'string' ? p : (p?.message || p?.text || JSON.stringify(p)));

async function parseFile(file) {
  const wb = await readWorkbook(file);
  const model = parseWorkbook(wb);
  if (!model.sites?.length) {
    const why = (model.problems || []).map(problemText).slice(0, 3).join(' ');
    throw new Error(`No sites could be read from "${file.name}". ${why || 'Check it is a tracker workbook with a Sites sheet.'}`);
  }
  return model;
}

/* ============================== loading ============================== */

async function loadMain(file) {
  if (!file) return;
  lastError = '';
  try {
    const model = await parseFile(file);
    preview = buildBrief(model, { now: new Date(), workbookName: file.name });
    pendingSel = new Set(preview.sites.filter((s) => !s.noData).map((s) => s.code));
    S.set({ workbookName: file.name, model, selection: [], brief: null, conclusion: null, runAt: null });
  } catch (e) {
    lastError = e.message;
    toast(e.message, 'survey', 7000);
    render();
  }
}

async function loadPrevious(file) {
  if (!file) return;
  try {
    const model = await parseFile(file);
    S.set({ previousName: file.name, previousModel: model });
    toast(`Previous week loaded: ${file.name}. Generate the brief again to use it.`, 'sign');
  } catch (e) { toast(e.message, 'survey', 7000); }
}

function dropZone({ title, hint, accept, onFile, current }) {
  const input = el('input', { type: 'file', accept, 'aria-label': title });
  const zone = el('label', { class: 'drop' }, [
    el('strong', { text: title }),
    el('span', { class: 'small muted', text: hint }),
    current ? el('div', { class: 'file', text: current }) : null,
    input,
  ]);
  input.addEventListener('change', () => { onFile(input.files?.[0]); });
  zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('is-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('is-over'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault(); zone.classList.remove('is-over');
    onFile(e.dataTransfer?.files?.[0]);
  });
  return zone;
}

function loadPanel(st) {
  return el('section', { class: 'panel' }, [
    el('h2', {}, [el('span', { class: 'n', text: '1' }), 'Load your tracker workbook']),
    el('div', { class: 'drops' }, [
      dropZone({
        title: 'Tracker workbook (required)', hint: 'Drop an .xlsx here or click to choose. It is read in this browser and not uploaded.',
        accept: '.xlsx,.xlsm,.xls', onFile: loadMain, current: st.workbookName,
      }),
      el('div', {}, [
        dropZone({
          title: 'Previous week’s workbook (optional)', hint: 'Used only to show what changed since last week.',
          accept: '.xlsx,.xlsm,.xls', onFile: loadPrevious, current: st.previousName,
        }),
        st.previousModel ? el('div', { class: 'row small', style: { marginTop: '6px' } }, [
          el('span', { class: 'muted', text: 'Without it, the previous week inside the main workbook is used.' }),
          btn('Remove', () => S.set({ previousName: null, previousModel: null }), 'btn--ghost', { style: { color: 'var(--ink)', borderColor: '#B5ACE9' } }),
        ]) : el('div', { class: 'small muted', style: { marginTop: '6px' }, text: 'Not loaded: the previous week inside the main workbook is used.' }),
      ]),
    ]),
    lastError ? el('div', { class: 'notice notice--bad', role: 'alert', text: lastError }) : null,
  ]);
}

/* ============================ what was read ============================ */

function readPanel(st) {
  const m = st.model;
  if (!m) return null;
  const byCode = new Map((preview?.sites || []).map((s) => [s.code, s]));
  const withData = [...byCode.values()].filter((s) => !s.noData).length;

  const rows = m.sites.map((s) => {
    const pv = byCode.get(s.code) || {};
    const checks = [];
    if (pv.noData) checks.push(el('span', { class: 'tag tag--warn', text: 'No weekly status' }));
    if (pv.privacy?.flagged) checks.push(el('span', { class: 'tag tag--warn', title: 'Possible personal or place names. Check before generating.', text: 'Possible names' }));
    if (s.matchedBy && s.matchedBy !== 'named in register') checks.push(el('span', { class: 'tag tag--warn', title: 'The detail sheet was matched by a guess. Confirm it is the right sheet.', text: `Sheet matched: ${s.matchedBy}` }));
    if (pv.dataFindings) checks.push(el('span', { class: 'tag', text: `${pv.dataFindings} data check${pv.dataFindings === 1 ? '' : 's'}` }));
    const cb = el('input', { type: 'checkbox', 'aria-label': `Include ${s.code}`, checked: pendingSel?.has(s.code) ? true : null });
    cb.addEventListener('change', () => { cb.checked ? pendingSel.add(s.code) : pendingSel.delete(s.code); });
    return el('tr', {}, [
      el('td', {}, [cb]),
      el('td', {}, [el('strong', { text: s.code })]),
      el('td', { text: s.description || '' }),
      el('td', { text: s.sheet || '—' }),
      el('td', { class: 'num', text: String(s.taskCount ?? '—') }),
      el('td', { class: 'num', text: String(s.weekCount ?? '—') }),
      el('td', {}, checks.length ? checks : [el('span', { class: 'muted', text: 'OK' })]),
    ]);
  });

  const sitePoints = m.sites.flatMap((s) => (s.detail?.problems || []).map((p) => `${s.code}: ${problemText(p)}`));
  const problems = [...(m.problems || []).map(problemText), ...(m.notes || []).map(problemText), ...sitePoints,
    ...((m.orphanSheets || []).map((o) => `Sheet "${o.name || o}" is in the workbook but matches no site, so it was ignored.`))];

  const setAll = (on) => {
    pendingSel = new Set(on ? m.sites.map((s) => s.code) : []);
    render();
  };

  return el('section', { class: 'panel' }, [
    el('h2', {}, [el('span', { class: 'n', text: '2' }), 'Check what was read, then choose sites']),
    el('p', { class: 'small muted', text: `${st.workbookName}: ${m.sites.length} site${m.sites.length === 1 ? '' : 's'} in the register, ${withData} with weekly status.` }),
    preview?.privacy.flagged ? el('div', { class: 'notice notice--warn', role: 'alert' }, [
      el('strong', { text: 'Possible personal or place names found in ' + preview.privacy.sites.join(', ') + '. ' }),
      'The workbook should use codes only. This check is a guess and only warns; look before you generate or send anything to an AI.',
    ]) : null,
    el('div', { style: { overflowX: 'auto' } }, [el('table', { class: 't' }, [
      el('thead', {}, [el('tr', {}, ['', 'Site', 'Description', 'Detail sheet', 'Tasks', 'Weeks', 'Checks'].map((h, i) => el('th', { class: i > 3 && i < 6 ? 'num' : '', text: h })))]),
      el('tbody', {}, rows),
    ])]),
    problems.length ? el('div', { class: 'notice notice--warn' }, [
      el('strong', { text: `${problems.length} thing${problems.length === 1 ? '' : 's'} to look at: ` }),
      el('ul', { style: { margin: '4px 0 0 18px', padding: 0 } }, problems.slice(0, 8).map((t) => el('li', { text: t }))),
      problems.length > 8 ? el('div', { class: 'xs muted', text: `…and ${problems.length - 8} more.` }) : null,
    ]) : null,
    el('div', { class: 'row', style: { marginTop: '10px' } }, [
      btn('Generate brief', generate, 'btn--primary'),
      btn('Select all', () => setAll(true)), btn('Select none', () => setAll(false)),
      el('span', { class: 'small muted', text: 'Sites with no weekly status can be included; they appear as “no data”.' }),
    ]),
  ]);
}

/* ============================== generating ============================== */

function currentModelCodes() { return S.get().model.sites.map((s) => s.code); }

function generate() {
  const st = S.get();
  const codes = currentModelCodes().filter((c) => pendingSel?.has(c));
  if (!codes.length) { toast('Tick at least one site first.', 'survey'); return; }
  const now = new Date();
  const brief = buildBrief(st.model, { codes, previousModel: st.previousModel, now, workbookName: st.workbookName });
  lastError = '';
  S.set({ brief, selection: codes, runAt: now.toISOString(), conclusion: null });
  setTimeout(() => document.getElementById('brief-panel')?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50);
}

/* ============================== AI conclusion ============================== */

function showStatus(s) {
  if (!busy) return;
  busy.status = s;
  const box = document.getElementById('ai-status-text');
  if (!box) return;
  box.textContent = s
    ? `Gemini is busy (${s.why}). Attempt ${s.attempt} of ${s.maxAttempts}, retrying in ${s.secondsLeft}s.${s.providerMessage ? ` It said: “${s.providerMessage}”` : ''}`
    : 'Asking Gemini…';
}

function reviewDialog(prev, cap) {
  return new Promise((resolve) => {
    let decided = false;
    const done = (v) => { if (!decided) { decided = true; resolve(v); } };
    const { body, foot, close } = openModal({ title: 'Review what will be sent', wide: true, onClose: () => done(false) });
    mount(body,
      el('p', { class: 'small', text: `This exact text goes from your browser to ${prev.destination} (${prev.model}). It is the calculated figures for the selected sites, not the workbook.` }),
      el('div', { class: 'row small muted' }, [
        el('span', { text: `About ${T.formatTokens(prev.estimatedInputTokens)} input tokens (estimate)` }),
        el('span', { text: `${(prev.bytes / 1024).toFixed(1)} KB` }),
        el('span', { text: '1 request' }),
      ]),
      cap.level !== 'low' ? el('div', { class: `notice ${cap.level === 'high' ? 'notice--bad' : 'notice--warn'}`, text: cap.message }) : null,
      el('div', { class: 'notice notice--warn', text: T.getLimits('gemini').tier === 'free'
        ? 'On the free tier, Google may use what you send to improve its products, and people may read it. Only send what you are comfortable sharing.'
        : 'Check your provider’s data terms for the plan you are on.' }),
      el('pre', { class: 'payload', tabindex: '0', text: prev.prompt }),
    );
    mount(foot,
      btn('Cancel', () => { done(false); close(); }),
      btn(cap.level === 'high' ? 'Send anyway' : 'Send', () => { done(true); close(); }, 'btn--primary'),
    );
  });
}

async function getConclusion() {
  const st = S.get();
  if (!st.brief || busy) return;
  if (!S.hasApiKey()) {
    toast('Add your Gemini API key first.');
    openSettings(() => getConclusion());
    return;
  }
  const settings = S.getSettings();
  const payload = buildConclusionPayload(st.brief);
  const prev = previewPayload({ payload, settings });
  const cap = T.capacityCheck('gemini', { total: prev.estimatedInputTokens + Math.min(settings.maxTokens, 1500), calls: 1 });

  if (settings.reviewPayload || cap.level === 'high') {
    if (!(await reviewDialog(prev, cap))) return;
  }

  const ctrl = new AbortController();
  busy = { ctrl, status: null };
  lastError = '';
  render();
  try {
    const out = await run({ payload, settings, signal: ctrl.signal, onStatus: showStatus });
    const selected = st.brief.title.sites;
    const others = currentModelCodes().filter((c) => !selected.includes(c));
    const result = validateConclusion(out.result, selected, others);
    S.set({ conclusion: { result, meta: {
      model: out.model, requestedModel: out.requestedModel, fellBack: out.fellBack,
      at: out.at, providerLabel: out.providerLabel, tokens: out.tokens,
    } } });
    if (out.fellBack) toast(`${out.requestedModel} was overloaded, so ${out.model} answered.`);
  } catch (e) {
    if (e instanceof CancelledError || e?.name === 'CancelledError') toast('Cancelled. Nothing was changed.');
    else lastError = e.message;
  } finally {
    busy = null;
    render();
  }
}

/* ============================== settings ============================== */

function openSettings(after) {
  const cfg = S.getSettings();
  const prov = PROVIDERS.gemini;
  const { body, foot, close } = openModal({ title: 'AI settings (Google Gemini)', wide: true });

  const model = el('input', { type: 'text', value: cfg.model, placeholder: prov.defaultModel });
  const fallback = el('input', { type: 'text', value: cfg.fallbackModel, placeholder: `e.g. ${prov.suggestions[2]}` });
  const key = el('input', { type: 'password', autocomplete: 'off', spellcheck: 'false',
    placeholder: S.hasApiKey() ? 'Paste a new key only to replace the saved one' : prov.keyHint });
  const maxTok = el('input', { type: 'number', value: cfg.maxTokens, min: 512, max: 16000 });
  const remember = el('input', { type: 'checkbox', checked: cfg.rememberKeyForSession ? true : null });
  const review = el('input', { type: 'checkbox', checked: cfg.reviewPayload ? true : null });
  const lim = T.getLimits('gemini');
  const tier = el('select', {}, [
    el('option', { value: 'free', selected: lim.tier === 'free' ? true : null, text: 'Free / evaluation' }),
    el('option', { value: 'paid', selected: lim.tier === 'paid' ? true : null, text: 'Paid' }),
  ]);
  const rpm = el('input', { type: 'number', value: lim.rpm ?? '' });
  const rpd = el('input', { type: 'number', value: lim.rpd ?? '' });
  const tpm = el('input', { type: 'number', value: lim.tpm ?? '' });
  const ks = S.keyStatus();

  mount(body,
    el('div', { class: 'notice notice--warn' }, [
      el('strong', { text: 'Before you paste a key. ' }),
      'This app has no server. Your key stays in this browser tab and goes only to Google. It is never written to a session file or an export. Anything with access to this tab can read it, so use a key made for this tool, limited to one model, with a spend cap.',
    ]),
    el('div', { class: 'formgrid', style: { marginTop: '10px' } }, [
      el('div', { class: 'field' }, [el('label', { text: 'Model' }), model,
        el('span', { class: 'hint', text: `Blank uses ${prov.defaultModel}. Model names change; a 404 means the name is out of date.` })]),
      el('div', { class: 'field' }, [el('label', { text: 'Fallback model if overloaded' }), fallback,
        el('span', { class: 'hint', text: 'Used only when the main model still returns 503 after all retries. Pick an established model. Blank means never switch.' })]),
      el('div', { class: 'field span-2' }, [
        el('label', { text: S.hasApiKey() ? 'API key (saved: leave blank to keep it)' : 'API key' }),
        el('div', { class: `notice ${ks.set ? 'notice--ok' : 'notice--warn'}` }, [
          el('strong', { text: ks.set ? `Key saved: ${ks.masked}. ` : 'No key saved. ' }), ks.lasts]),
        key,
        el('span', { class: 'hint' }, ['Get one from ', el('a', { href: prov.keyUrl, target: '_blank', rel: 'noopener noreferrer', text: 'Google AI Studio' }), '.']),
      ]),
      el('div', { class: 'field' }, [el('label', { text: 'Max output tokens' }), maxTok,
        el('span', { class: 'hint', text: 'Lower is cheaper and less likely to be cut short.' })]),
    ]),
    el('label', { class: 'check' }, [remember, el('span', {}, [
      el('strong', { text: 'Keep the key if I refresh the page. ' }),
      'Without this, a refresh clears the key. With it, the key survives refreshes but goes when the tab closes. Leave off on a shared computer.'])]),
    el('label', { class: 'check' }, [review, el('span', {}, [
      el('strong', { text: 'Show me what will be sent before every request. ' }), 'Recommended.'])]),
    el('h3', { text: 'Your account limits', style: { marginTop: '14px' } }),
    el('p', { class: 'small muted', text: lim.caution || 'These drive the capacity warning. They are defaults, not read from your account: edit them to match what Google gives you.' }),
    el('div', { class: 'formgrid' }, [
      el('div', { class: 'field' }, [el('label', { text: 'Tier' }), tier]),
      el('div', { class: 'field' }, [el('label', { text: 'Requests / minute' }), rpm]),
      el('div', { class: 'field' }, [el('label', { text: 'Requests / day' }), rpd]),
      el('div', { class: 'field' }, [el('label', { text: 'Tokens / minute' }), tpm]),
    ]),
  );
  mount(foot,
    S.hasApiKey() ? btn('Clear key', () => { S.clearApiKey(); toast('Key cleared from this browser.'); close(); render(); }, '', { style: { marginRight: 'auto' } }) : null,
    btn('Cancel', close),
    btn('Save', () => {
      S.saveSettings({ model: model.value, fallbackModel: fallback.value, maxTokens: clamp(Number(maxTok.value) || 4096, 512, 16000),
        rememberKeyForSession: remember.checked, reviewPayload: review.checked });
      if (key.value.trim()) S.setApiKey(key.value.trim(), remember.checked);
      else if (S.hasApiKey()) S.setApiKey(S.getApiKey(), remember.checked);
      T.saveLimits('gemini', { tier: tier.value, rpm: rpm.value, rpd: rpd.value, tpm: tpm.value });
      close();
      toast('Settings saved.', 'sign');
      render();
      if (after && S.hasApiKey()) after();
    }, 'btn--primary'),
  );
}

function usageCard() {
  const usage = T.getUsage();
  const win = T.requestWindows('gemini');
  const lim = T.getLimits('gemini');
  const bar = (used, max) => {
    const pc = max ? clamp((used / max) * 100, 0, 100) : 0;
    return el('div', { class: `meter${pc > 85 ? ' hot' : pc > 60 ? ' warm' : ''}` }, [el('div', { style: { width: `${pc}%` } })]);
  };
  const kv = (k, v) => el('div', { class: 'kv' }, [el('span', { text: k }), el('strong', { text: v })]);
  const ks = S.keyStatus();
  return el('details', { class: 'usage' }, [
    el('summary', { text: 'Usage and limits' }),
    el('div', { style: { padding: '8px 0', maxWidth: '420px' } }, [
      kv('Requests this minute', `${win.lastMinute} / ${lim.rpm ?? '—'}`), lim.rpm ? bar(win.lastMinute, lim.rpm) : null,
      kv('Requests today', `${win.lastDay} / ${lim.rpd ?? '—'}`), lim.rpd ? bar(win.lastDay, lim.rpd) : null,
      kv('Tokens used', `${T.formatTokens(usage.total)}${usage.estimated ? ' (estimate)' : ''}`),
      kv('In / out', `${T.formatTokens(usage.input)} / ${T.formatTokens(usage.output)}`),
      kv('Calls made', String(usage.calls)),
      kv('API key', ks.set ? (ks.where === 'tab' ? 'saved, this tab' : 'saved, until refresh') : 'not set'),
      ...overloadedModels().map((o) => el('div', { class: 'small', style: { color: 'var(--warn)' } }, [
        `${o.model} overloaded: skipped for ${o.minutesLeft} more min. `,
        el('a', { href: '#', onclick: (e) => { e.preventDefault(); clearOverloadMemory(o.model); render(); } }, ['Try it again']),
      ])),
      el('div', { class: 'row', style: { marginTop: '6px' } }, [
        btn('Reset tokens', () => { T.resetUsage(); render(); }), btn('Reset requests', () => { T.resetRequests('gemini'); render(); }),
      ]),
    ]),
  ]);
}

/* ============================== the brief ============================== */

function briefPanel(st) {
  if (!st.brief) return null;
  const lines = engineSummary(st.brief);
  const hasKey = S.hasApiKey();
  const page = el('div', { class: 'page', id: 'brief-page' });
  // The renderer escapes every piece of workbook text (tested), so this is safe.
  page.innerHTML = renderBriefHTML(st.brief, st.conclusion, { engineLines: lines, screen: true });

  const orient = el('select', { 'aria-label': 'Paper orientation for PDF' }, [
    el('option', { value: 'portrait', selected: orientation === 'portrait' ? true : null, text: 'Portrait' }),
    el('option', { value: 'landscape', selected: orientation === 'landscape' ? true : null, text: 'Landscape' }),
  ]);
  orient.addEventListener('change', () => { orientation = orient.value; });

  const safe = (fn, label) => () => { try { fn(); } catch (e) { toast(`${label} failed: ${e.message}`, 'survey', 7000); } };
  return el('section', { class: 'panel', id: 'brief-panel' }, [
    el('h2', {}, [el('span', { class: 'n', text: '3' }), 'Brief']),
    el('div', { class: 'toolbar' }, [
      btn(busy ? 'Working…' : (st.conclusion ? 'Redo AI conclusion' : 'Get AI conclusion'), getConclusion, 'btn--primary', { disabled: busy ? true : null }),
      !hasKey ? el('span', { class: 'xs muted', text: 'needs a Gemini key' }) : null,
      el('span', { class: 'sp' }),
      el('span', { class: 'xs muted', text: 'Export' }),
      orient,
      btn('PDF', safe(() => exportPDF(st.brief, st.conclusion, { engineLines: lines, orientation }), 'PDF export')),
      btn('Word', safe(() => exportWord(st.brief, st.conclusion, lines), 'Word export')),
      btn('Excel', safe(() => exportExcel(st.brief, st.conclusion, lines), 'Excel export')),
    ]),
    busy ? el('div', { class: 'notice', role: 'status' }, [
      el('span', { class: 'spin' }), el('span', { id: 'ai-status-text', text: 'Asking Gemini…' }), ' ',
      btn('Cancel', () => busy?.ctrl.abort(), '', { style: { marginLeft: '8px' } }),
    ]) : null,
    lastError ? el('div', { class: 'notice notice--bad', role: 'alert' }, [
      el('strong', { text: 'The AI conclusion could not be made. ' }), lastError,
      el('div', { class: 'small', text: 'Everything else on the page is unaffected, and the calculated summary is shown instead.' }),
    ]) : null,
    page,
    usageCard(),
  ]);
}

/* ============================== session ============================== */

async function saveSession() {
  try {
    const name = S.saveSessionToFile();
    toast(`Saved ${name}. It contains project data for the selected sites. Keep it private.`, 'sign', 7000);
  } catch (e) { toast(e.message, 'survey', 6000); }
}

async function loadSession() {
  const st = S.get();
  if ((st.model || st.brief) && !(await confirmDialog({
    title: 'Replace current work?', message: 'Loading a session replaces what is on the page now. Anything not saved is lost.', confirmLabel: 'Load session',
  }))) return;
  try {
    const out = await S.loadSessionFromFile();
    if (out.cancelled) return;
    const s = out.state;
    const brief = buildBrief(s.model, {
      codes: s.selection, previousModel: s.previousModel || null,
      now: new Date(s.runAt), workbookName: s.workbookName,
    });
    preview = buildBrief(s.model, { now: new Date(s.runAt), workbookName: s.workbookName });
    pendingSel = new Set(s.selection);
    lastError = '';
    S.set({ ...S.blank(), workbookName: s.workbookName, model: s.model, previousName: s.previousName || null,
      previousModel: s.previousModel || null, selection: s.selection, runAt: s.runAt, brief, conclusion: s.conclusion || null });
    toast(`Loaded ${out.filename}.`, 'sign');
  } catch (e) { toast(e.message, 'survey', 8000); }
}

async function startOver() {
  if (!(await confirmDialog({ title: 'Start over?', message: 'This clears the workbook and brief from the page. Anything not saved is lost.', confirmLabel: 'Start over', tone: 'danger' }))) return;
  preview = null; pendingSel = null; lastError = '';
  S.reset();
}

/* ============================== render ============================== */

function render() {
  const st = S.get();
  mount(root,
    el('header', { class: 'top' }, [
      el('h1', {}, ['Site progress brief', el('small', { text: 'Your workbook is read in this browser and not uploaded. Figures are calculated here; only the conclusion is AI-written.' })]),
      btn('AI settings', () => openSettings(), 'btn--ghost'),
      btn('Save session', saveSession, 'btn--ghost', { disabled: st.brief ? null : true }),
      btn('Load session', loadSession, 'btn--ghost'),
      btn('Start over', startOver, 'btn--ghost', { disabled: st.model ? null : true }),
    ]),
    el('main', { class: 'wrap' }, [loadPanel(st), readPanel(st), briefPanel(st)]),
    el('footer', { class: 'foot' }, [
      'Reloading the page or closing the tab clears everything. Use “Save session” to keep your work. Workbooks should use codes only (for example site A-01, person R1).',
    ]),
  );
}

S.subscribe(render);
render();
