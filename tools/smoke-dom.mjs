/**
 * Headless UI smoke test.
 *
 * A passing syntax check says nothing about whether the page renders: the commonest
 * failure is a thrown error inside a render function, giving a blank page with a
 * green build. This boots the real page under jsdom and walks the whole workflow:
 * load, check what was read, pick sites, generate, get the AI conclusion (against a
 * stubbed Gemini), export, save the session, start over, load it back.
 *
 * jsdom has no layout engine, so this proves behaviour, not appearance.
 *
 * Run: node tools/smoke-dom.mjs
 */

import { JSDOM } from 'jsdom';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { fixtureWorkbook, variantWorkbook } from './fixture.mjs';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const XLSX = createRequire(import.meta.url)(path.join(root, 'assets/vendor/xlsx.full.min.js'));

const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8').replace(/<script[\s\S]*?<\/script>/g, '');
const dom = new JSDOM(html, { url: 'http://localhost/', pretendToBeVisual: true });
const { window } = dom;

global.window = window;
global.document = window.document;
for (const k of ['navigator', 'location', 'localStorage', 'sessionStorage']) {
  Object.defineProperty(globalThis, k, { value: window[k], writable: true, configurable: true });
}
global.HTMLElement = window.HTMLElement;
global.Node = window.Node;
global.Blob = window.Blob;
global.URL = window.URL;
global.AbortController = window.AbortController || AbortController;
global.requestAnimationFrame = (fn) => setTimeout(fn, 0);
window.XLSX = XLSX;
window.HTMLElement.prototype.scrollIntoView = function () {};
window.scrollTo = () => {};

// jsdom does not implement <dialog> behaviour. Closing fires its event as a queued task, as browsers do.
Object.defineProperty(window.HTMLElement.prototype, 'showModal', { value() { this.setAttribute('open', ''); }, writable: true, configurable: true });
Object.defineProperty(window.HTMLElement.prototype, 'close', {
  value() { this.removeAttribute('open'); setTimeout(() => this.dispatchEvent(new window.Event('close')), 0); },
  writable: true, configurable: true,
});

// Downloads: capture what would be saved.
const downloads = [];
window.URL.createObjectURL = (b) => { downloads.push({ blob: b }); return `blob:${downloads.length}`; };
window.URL.revokeObjectURL = () => {};
const origAClick = window.HTMLAnchorElement.prototype.click;
window.HTMLAnchorElement.prototype.click = function () {
  const d = downloads[downloads.length - 1];
  if (d && this.hasAttribute('download')) d.name = this.getAttribute('download');
};
const writeFileCalls = [];
XLSX.writeFile = (wb, name) => { writeFileCalls.push({ wb, name }); };

// File picker: deliver whatever the test queued.
let nextPick = null;
window.HTMLInputElement.prototype.click = function () {
  if (this.type === 'file' && nextPick) {
    const f = nextPick; nextPick = null;
    Object.defineProperty(this, 'files', { value: [f], configurable: true });
    setTimeout(() => this.dispatchEvent(new window.Event('change')), 0);
  }
};

const readBlob = (blob) => new Promise((res, rej) => {
  const r = new window.FileReader();
  r.onload = () => res(String(r.result)); r.onerror = rej; r.readAsText(blob);
});

// Stubbed Gemini.
let fetchCalls = [];
let fetchHandler = null;
const okBody = (obj) => ({
  candidates: [{ content: { parts: [{ text: JSON.stringify(obj) }] } }],
  usageMetadata: { promptTokenCount: 800, candidatesTokenCount: 200, totalTokenCount: 1000 },
});
const mkRes = (status, body) => ({
  ok: status >= 200 && status < 300, status, headers: { get: () => null },
  text: async () => JSON.stringify(body), json: async () => body,
});
global.fetch = async (url, opts) => {
  fetchCalls.push({ url, opts, body: opts?.body ? JSON.parse(opts.body) : null });
  return fetchHandler(fetchCalls.length, url, opts);
};
const GOOD = { headline: 'F-01 is behind plan; clear its stuck tasks first.', sites: [{ site: 'F-01', verdict: 'Will miss target', reason: 'Pace is 0.61 of plan.', action: 'Clear the two stuck tasks.' }], reallocation: null };

// An error thrown inside an event handler or a floating promise would otherwise vanish.
const swallowed = [];
process.on('unhandledRejection', (e) => { swallowed.push(e?.stack || String(e)); });
window.addEventListener('error', (ev) => { swallowed.push(ev.error?.stack || ev.message); });

let pass = 0, fail = 0;
const failures = [];
const check = (name, cond, detail = '') => {
  if (cond) pass++;
  else { fail++; failures.push(`${name}${detail ? ` — ${detail}` : ''}`); }
};

const tick = (ms = 15) => new Promise((r) => setTimeout(r, ms));
const waitFor = async (fn, ms = 4000) => {
  const t = Date.now();
  while (Date.now() - t < ms) { const v = fn(); if (v) return v; await tick(10); }
  return null;
};
const $ = (sel, r = document) => r.querySelector(sel);
const btn = (text, r = document) => [...r.querySelectorAll('button')].find((b) => b.textContent.trim() === text);
const click = (el) => {
  if (!el) { check('a button the test needed was on the page', false, `missing at: ${new Error().stack.split('\n')[2].trim()}`); return; }
  el.dispatchEvent(new window.MouseEvent('click', { bubbles: true, cancelable: true }));
};
const setFile = (input, file) => {
  Object.defineProperty(input, 'files', { value: [file], configurable: true });
  input.dispatchEvent(new window.Event('change', { bubbles: true }));
};
const openDialog = () => $('dialog[open]');
const bodyText = () => document.body.textContent.replace(/\s+/g, ' ');

const wbBuf = XLSX.write(fixtureWorkbook(XLSX, (await import('../js/core/parser.js')).STATUS), { type: 'array', bookType: 'xlsx' });
const fakeWb = (name) => ({ name, arrayBuffer: async () => wbBuf });

/* ================================ boot ================================ */
let bootError = null;
const origErr = console.error;
try {
  await import('../js/views/app.js');
} catch (e) { bootError = e; }
check('the page boots without throwing', !bootError, bootError?.stack?.split('\n').slice(0, 3).join(' | '));
if (bootError) { console.log(failures.join('\n')); process.exit(1); }

const S = await import('../js/core/session.js');
const AI = await import('../js/core/ai.js');
AI._setOverloadWaits([1, 1, 1, 1]);

check('header and title render', /Site progress brief/.test(bodyText()) && !!btn('AI settings'));
check('two separate upload zones', document.querySelectorAll('input[type=file]').length === 2);
check('main upload is labelled required, previous optional', /Tracker workbook \(required\)/.test(bodyText()) && /Previous week.s workbook \(optional\)/.test(bodyText()));
check('Save session is off until there is a brief', btn('Save session').disabled === true);
check('Start over is off until a workbook is loaded', btn('Start over').disabled === true);

/* ============================ layout (the upload boxes) ============================ */
// jsdom does not load linked stylesheets, so put the real one in the page and ask for computed styles.
{
  const style = document.createElement('style');
  style.textContent = fs.readFileSync(path.join(root, 'assets/css/lite.css'), 'utf8');
  document.head.appendChild(style);
  const zones = [...document.querySelectorAll('label.drop')];
  check('both upload boxes are block-level, so title, hint and border stay together', zones.length === 2 && zones.every((z) => window.getComputedStyle(z).display === 'block'),
    zones.map((z) => window.getComputedStyle(z).display).join());
  check('the optional box sits in the same grid as the required one', zones.length === 2 && zones[0].closest('.drops') === zones[1].closest('.drops'));
}

/* ================================ load ================================ */
const [mainInput, prevInput] = document.querySelectorAll('input[type=file]');
setFile(mainInput, fakeWb('fixture.xlsx'));
await waitFor(() => /Check what was read/.test(bodyText()));
check('the "what was read" panel appears', /Check what was read, then choose sites/.test(bodyText()));
const rows = [...document.querySelectorAll('table.t tbody tr')];
check('both register sites are listed', rows.length === 2 && /F-01/.test(rows[0].textContent) && /F-02/.test(rows[1].textContent));
check('a site with no weekly status is flagged', /No weekly status/.test(rows[1].textContent));
check('the site with data is ticked, the empty one is not', rows[0].querySelector('input').checked === true && rows[1].querySelector('input').checked === false);
check('workbook name is shown', /fixture\.xlsx/.test(bodyText()));
check('Generate is offered', !!btn('Generate brief'));
check('Start over is now on', btn('Start over').disabled === false);

// A file that is not a tracker must say so, plainly, and leave the page working.
setFile(mainInput, { name: 'nonsense.xlsx', arrayBuffer: async () => new ArrayBuffer(8) });
await tick(200);
check('an unreadable file gives a message, not a blank page', /nonsense\.xlsx|could not|spreadsheet/i.test(bodyText()) && !!btn('AI settings'));
check('and the earlier, good workbook stays on screen after the bad one', document.querySelectorAll('table.t tbody tr').length === 2);
setFile(document.querySelectorAll('input[type=file]')[0], fakeWb('fixture.xlsx'));
await tick(300);   // the table is already on screen, so there is nothing to wait for: let the re-load finish

/* ============================ generate: one site ============================ */
click(btn('Generate brief'));
const page = await waitFor(() => $('#brief-page'));
check('the brief appears', !!page);
const pt = page.textContent.replace(/\s+/g, ' ');
check('title block, overall status and site are there', /Site progress brief/.test(pt) && /Overall: Poor/.test(pt) && /F-01/.test(pt));
check('task groups and counts are there', /Modelling/.test(pt) && /1 of 2/.test(pt));
check('prerequisite and blocked task are there', /P02/.test(pt) && /Ceiling void zones not confirmed/.test(pt));
check('without AI the page shows the calculated summary', /Calculated, no AI/.test(pt));
check('the unselected site is not on the page', !/F-02/.test(pt));
check('the chart is drawn', !!page.querySelector('svg[role=img]'));
const hasHandler = [...page.querySelectorAll('*')].some((e) => [...e.attributes].some((a) => /^on/i.test(a.name)));
check('no element on the page carries an event-handler attribute', !hasHandler);
check('no script element is on the page', !page.querySelector('script'));
check('Save session is now on', btn('Save session').disabled === false);
check('export buttons are offered', !!btn('PDF') && !!btn('Word') && !!btn('Excel'));
check('the AI button asks for a key', /needs a Gemini key/.test(bodyText()));

/* ======================= AI conclusion: key, review, send ======================= */
fetchHandler = () => mkRes(200, okBody(GOOD));
S.saveSettings({ reviewPayload: true });
click(btn('Get AI conclusion'));
const settingsDlg = await waitFor(openDialog);
check('with no key, the settings dialog opens first', !!settingsDlg && /AI settings/.test(settingsDlg.textContent));
check('the dialog warns about the key before asking for it', /Before you paste a key/.test(settingsDlg.textContent));
check('the key box is a password field', !!settingsDlg.querySelector('input[type=password]'));
check('the Gemini key link opens safely', (() => { const a = settingsDlg.querySelector('a[href*=aistudio]'); return a && /noopener/.test(a.rel) && a.target === '_blank'; })());
settingsDlg.querySelector('input[type=password]').value = 'AIza-TEST-SECRET-KEY';
settingsDlg.querySelector('input[type=text]').value = 'model-x';
click(btn('Save', settingsDlg));
const review = await waitFor(() => { const d = openDialog(); return d && /Review what will be sent/.test(d.textContent) ? d : null; });
check('saving the key goes on to the review step', !!review);
check('nothing was sent before the review', fetchCalls.length === 0);
if (!review) {
  console.log('swallowed so far:', swallowed.slice(0, 2).join('\n') || '(none)');
  console.log('dialogs:', [...document.querySelectorAll('dialog')].map((d) => `${d.hasAttribute('open') ? 'OPEN' : 'closed'}: ${d.textContent.slice(0, 80)}`));
  console.log('toasts:', [...document.querySelectorAll('.toast')].map((t) => t.textContent));
  console.log('state has brief:', !!S.get().brief, '| model:', !!S.get().model, '| #brief-page:', !!$('#brief-page'));
  console.log('fetch calls:', fetchCalls.length, '| hasKey:', S.hasApiKey(), '| settings:', JSON.stringify(S.getSettings()));
}
check('the review shows the exact text, for the selected site only', review && /"F-01"/.test(review.querySelector('pre.payload').textContent) && !/F-02/.test(review.querySelector('pre.payload').textContent));
check('the review names the destination and model', !!review && /Google Gemini/.test(review.textContent) && /model-x/.test(review.textContent));
check('the review warns about free-tier data use', !!review && /improve its products/.test(review.textContent));
check('the key is not in the reviewed text', !!review && !review.textContent.includes('AIza-TEST-SECRET-KEY'));
check('the key is held in memory, not in browser storage', S.getApiKey() === 'AIza-TEST-SECRET-KEY' && !window.localStorage.getItem('bimtrack-lite:settings:v1')?.includes('AIza'));

// Cancel at the review: nothing is sent.
click(btn('Cancel', review));
await tick(40);
check('cancelling the review sends nothing', fetchCalls.length === 0);

click(btn('Get AI conclusion'));
const review2 = await waitFor(() => { const d = openDialog(); return d && /Review what will be sent/.test(d.textContent) ? d : null; });
click(btn('Send', review2));
await waitFor(() => /AI-written · /.test($('#brief-page')?.textContent || ''));
const pt2 = $('#brief-page').textContent.replace(/\s+/g, ' ');
if (!/AI-written · /.test(pt2)) console.log('DEBUG notices:', [...document.querySelectorAll('.notice')].map((n) => n.textContent.replace(/\s+/g, ' ').slice(0, 220)));
check('exactly one request was made', fetchCalls.length === 1, String(fetchCalls.length));
check('the key went in a header, not the body or URL', fetchCalls[0].opts.headers['x-goog-api-key'] === 'AIza-TEST-SECRET-KEY' && !JSON.stringify(fetchCalls[0].body).includes('AIza') && !fetchCalls[0].url.includes('AIza'));
check('the request went to the chosen model', fetchCalls[0].url.includes('model-x'));
check('the conclusion is on the page, labelled AI-written with the model', /AI-written · model-x/.test(pt2) && /clear its stuck tasks first/.test(pt2));
check('the verdict and action are shown', /Will miss target/.test(pt2) && /Clear the two stuck tasks/.test(pt2));
check('the calculated figures are still there beside it', /Overall: Poor/.test(pt2) && /Calculated/.test(pt2));
check('the usage meter recorded the call', /Calls made/.test(bodyText()) && AI && true);
check('the button now offers a redo', !!btn('Redo AI conclusion'));

/* ======================== the model strays out of scope ======================== */
fetchCalls = [];
fetchHandler = () => mkRes(200, okBody({
  headline: 'Unlike F-02, F-01 is late.',
  sites: [GOOD.sites[0], { site: 'F-02', verdict: 'On track', reason: 'Fine.', action: null }],
  reallocation: 'Move R2 from F-02 to F-01.',
}));
S.saveSettings({ reviewPayload: false });
click(btn('Redo AI conclusion'));
await waitFor(() => /Checked on receipt/.test($('#brief-page')?.textContent || ''));
const pt3 = $('#brief-page').textContent;
check('with review off, it sends without asking', fetchCalls.length === 1 && !openDialog());
check('a stray site never reaches the page', !/F-02/.test(pt3.replace(/Checked on receipt:[\s\S]*?(?=Written by an AI)/, '')) || /not in this brief/.test(pt3));
check('the page says what was removed', /Checked on receipt/.test(pt3) && /removed|dropped/.test(pt3));
check('the advice naming the other site is gone', !/Move R2 from F-02/.test(pt3));

/* =============================== failure paths =============================== */
fetchCalls = [];
fetchHandler = () => mkRes(503, { error: { message: 'busy' } });
click(btn('Redo AI conclusion'));
await waitFor(() => /could not be made/.test(bodyText()), 6000);
check('overload retried, then failed clearly', fetchCalls.length === 5 && /could not be made/.test(bodyText()) && /overloaded/i.test(bodyText()), `${fetchCalls.length} calls`);
check('the rest of the page is unaffected by the failure', /Overall: Poor/.test($('#brief-page').textContent) && !!btn('PDF'));
fetchCalls = [];
fetchHandler = () => mkRes(401, { error: { message: 'bad key' } });
click(btn('Redo AI conclusion'));
await waitFor(() => /rejected the API key/.test(bodyText()));
check('a rejected key is explained, without retrying', fetchCalls.length === 1 && /rejected the API key/.test(bodyText()));

/* ============================ generate: two sites ============================ */
fetchCalls = [];
fetchHandler = () => mkRes(200, okBody(GOOD));
click(btn('Select all'));
await tick(40);
click(btn('Generate brief'));
await waitFor(() => /Sites tracked \(2\)/.test($('#brief-page')?.textContent || ''));
check('two sites appear, one with no data, in one brief', /Sites tracked \(2\)/.test($('#brief-page').textContent) && /no detail sheet/.test($('#brief-page').textContent));
check('regenerating clears the old conclusion', !/AI-written · /.test($('#brief-page').textContent) && /Calculated, no AI/.test($('#brief-page').textContent));
click(btn('Get AI conclusion'));
await waitFor(() => /AI-written · /.test($('#brief-page')?.textContent || ''));
check('both sites went in ONE request', fetchCalls.length === 1 && fetchCalls[0].body.contents[0].parts[0].text.includes('"F-02"'), String(fetchCalls.length));
check('a site the model skipped is reported', /no conclusion for F-02/.test($('#brief-page').textContent));

// back to one site for the rest
click(btn('Select none'));
await tick(40);
document.querySelectorAll('table.t tbody tr')[0].querySelector('input').checked = true;
document.querySelectorAll('table.t tbody tr')[0].querySelector('input').dispatchEvent(new window.Event('change', { bubbles: true }));
click(btn('Generate brief'));
await waitFor(() => /Sites tracked \(1\)/.test($('#brief-page')?.textContent || ''));
fetchHandler = () => mkRes(200, okBody(GOOD));
click(btn('Get AI conclusion'));
await waitFor(() => /AI-written · /.test($('#brief-page')?.textContent || ''));

/* ================================== exports ================================== */
click(btn('Excel'));
check('Excel export is named by date and holds the sheets', writeFileCalls.length === 1 && /^site-brief-\d{4}-\d{2}-\d{2}\.xlsx$/.test(writeFileCalls[0].name) && writeFileCalls[0].wb.SheetNames.includes('Task groups'), JSON.stringify(writeFileCalls.map((c) => c.name)));
check('Excel export holds the AI conclusion, labelled', JSON.stringify(XLSX.utils.sheet_to_json(writeFileCalls[0].wb.Sheets.Conclusion, { header: 1 })).includes('AI-written (model-x)'));
const before = downloads.length;
click(btn('Word'));
await tick(30);
const wordDl = downloads[downloads.length - 1];
check('Word export downloads an .rtf', downloads.length === before + 1 && /^site-brief-.*\.rtf$/.test(wordDl.name || ''), wordDl?.name);
const rtf = await readBlob(wordDl.blob);
check('the Word file holds the figures and the AI label', rtf.startsWith('{\\rtf1') && rtf.includes('F-01') && rtf.includes('[AI-written (model-x)]'));
click(btn('PDF'));
await tick(30);
const frame = $('iframe');
check('PDF export builds a print document in a hidden frame', !!frame && frame.getAttribute('aria-hidden') === 'true' && frame.srcdoc.startsWith('<!doctype html>'));
check('the print document is self-contained: no scripts', !/<script/i.test(frame.srcdoc));
check('the print document is portrait A4 by default', /@page\{size:A4 portrait/.test(frame.srcdoc));
const orient = $('select[aria-label="Paper orientation for PDF"]');
orient.value = 'landscape'; orient.dispatchEvent(new window.Event('change', { bubbles: true }));
frame.remove();
click(btn('PDF'));
await tick(30);
check('landscape can be chosen for the PDF', /@page\{size:A4 landscape/.test($('iframe').srcdoc));
$('iframe').remove();

/* ================================== session ================================== */
const shown = $('#brief-page').innerHTML;
click(btn('Save session'));
await tick(30);
const sessDl = downloads[downloads.length - 1];
check('a session file is downloaded', /\.session\.json$/.test(sessDl.name || ''), sessDl?.name);
const sessText = await readBlob(sessDl.blob);
const sess = JSON.parse(sessText);
check('the session has the right format', sess.format === 'bim-tracker-lite.session' && sess.state.selection.join() === 'F-01');
check('the session never holds the API key', !sessText.includes('AIza-TEST-SECRET-KEY'));
check('the session holds only the selected site', !sessText.includes('F-02'));
check('the session keeps the AI conclusion', sess.state.conclusion?.result?.headline === GOOD.headline);

click(btn('Start over'));
const conf = await waitFor(() => { const d = openDialog(); return d && /Start over\?/.test(d.textContent) ? d : null; });
check('start over asks first', !!conf);
click(btn('Start over', conf));
await waitFor(() => !$('#brief-page'));
check('starting over clears the brief and the workbook', !$('#brief-page') && !document.querySelector('table.t') && btn('Save session').disabled === true);
check('the key survives start over (it is not project data)', S.hasApiKey());

nextPick = { name: 'saved.session.json', text: async () => sessText };
click(btn('Load session'));
await waitFor(() => $('#brief-page'));
check('loading the session restores the page', !!$('#brief-page'));
check('the restored page is identical to the saved one, AI conclusion included', $('#brief-page').innerHTML === shown, `${$('#brief-page')?.innerHTML.length} vs ${shown.length}`);
check('the restored page shows the site selected and the others absent', /Sites tracked \(1\)/.test($('#brief-page').textContent) && !/F-02/.test($('#brief-page').textContent));

nextPick = { name: 'junk.json', text: async () => '{"hello":1}' };
click(btn('Load session'));
const conf2 = await waitFor(() => { const d = openDialog(); return d && /Replace current work\?/.test(d.textContent) ? d : null; });
click(btn('Load session', conf2));
await tick(120);
check('a file that is not a session is refused, with a reason', /not a bim-tracker-lite\.session/.test(bodyText()));
check('and the current page is untouched', !!$('#brief-page'));

/* ================== problems found in a real workbook ================== */
const variantBuf = XLSX.write(variantWorkbook(XLSX, (await import('../js/core/parser.js')).STATUS), { type: 'array', bookType: 'xlsx' });
click(btn('Start over'));
const conf3 = await waitFor(() => { const d = openDialog(); return d && /Start over\?/.test(d.textContent) ? d : null; });
click(btn('Start over', conf3));
await waitFor(() => !document.querySelector('table.t'));
setFile(document.querySelectorAll('input[type=file]')[0], { name: 'variant.xlsx', arrayBuffer: async () => variantBuf });
await waitFor(() => $('#unregistered-notice'));
const un = $('#unregistered-notice');
check('a tab the register does not list is called out, by name, as NOT read', !!un && /NOT read: F-09/.test(un.textContent), un?.textContent);
check('and the fix is stated in plain words', /add a row for it in the Sites sheet/.test(un?.textContent || ''));
const wrong = $('#wrong-site-notice');
check('a copied tab is called out as possibly the wrong site', !!wrong && /may belong to the wrong site/.test(wrong.textContent) && /F-03: Sheet header code = F-01; register row = F-03/.test(wrong.textContent), wrong?.textContent);
const rowF03 = [...document.querySelectorAll('table.t tbody tr')].find((r) => /F-03/.test(r.textContent));
check('the site row carries the checks', !!rowF03 && /Sheet header disagrees/.test(rowF03.textContent) && /Register date missing/.test(rowF03.textContent) && /Column left out/.test(rowF03.textContent), rowF03?.textContent);
check('the summary counts listed sites without a tab', /3 sites in the register, 2 with weekly status/.test(bodyText()) && /1 listed with no detail tab \(F-02\)/.test(bodyText()), bodyText().match(/variant\.xlsx:[^.]*\.[^.]*\./)?.[0]);
check('the milestone column is explained among the things to look at', /looks like a milestone, not a week/.test(bodyText()));
check('sites with data are ticked, the unlisted tab is not a site', [...document.querySelectorAll('table.t tbody tr')].length === 3 && !/F-09/.test(document.querySelector('table.t').textContent));
click(btn('Generate brief'));
await waitFor(() => $('#brief-page'));
const vp = $('#brief-page');
check('the brief shows the "check before relying" box', !!vp.querySelector('.bf-attn') && /Check before relying on this brief/.test(vp.textContent));
check('the brief names the wrong-site risk on the page', /Sheet header code = F-01; register row = F-03/.test(vp.textContent));
check('pace says it rests on 1 of 2 sites and who was left out', /counting 1 of 2 sites/.test(vp.textContent) && /Not counted in pace: F-03/.test(vp.textContent), vp.textContent.match(/Pace[^·]*/)?.[0]);
check('F-03 reads its real last week, not the milestone column', /W05/.test(vp.textContent) && !/Submission · /.test(vp.textContent));
check('a tab with no register row does not appear in the brief', !/F-09/.test(vp.textContent));

/* ======================= nothing about a site is persisted ======================= */
const stored = JSON.stringify({ ...window.localStorage }) + JSON.stringify({ ...window.sessionStorage });
check('no site data is in browser storage', !/F-01|F-02|fixture|Modelling|Ceiling|Walls/.test(stored), stored.slice(0, 160));
check('the API key is not in browser storage (not asked to keep it)', !stored.includes('AIza'));
check('settings are stored, and are not project data', /model-x/.test(stored));

check('no error was thrown and swallowed along the way', swallowed.length === 0, swallowed[0]?.split('\n').slice(0, 3).join(' | '));
console.log(`\n${pass} passed, ${fail} failed`);
if (fail) {
  console.log('\nFAILURES:');
  for (const f of failures) console.log('  ✗ ' + f);
  process.exit(1);
}
console.log('\nUI smoke test passed.');
process.exit(0);
