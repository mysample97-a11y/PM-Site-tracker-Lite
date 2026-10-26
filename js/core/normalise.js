/**
 * Lite-only tidy-ups applied to the parser's output.
 *
 * The parser (shared with the full tracker, and deliberately left unchanged)
 * treats ANY column with a date in its header as a week. A workbook that puts a
 * milestone column such as "Submission 01-Jan" in the weekly grid therefore gets
 * a fake final "week". Every task reads "Not started" there, the site's latest
 * status becomes that column, and the site shows 0% complete.
 *
 * Doing the fix here, not in parser.js, keeps the shared code identical. The
 * cost is that Lite and the full app can differ on a workbook like that, so the
 * brief always says when a column was left out.
 */

// Built from a string so the quote marks inside it stay plain text.
const ORPHAN_SHEET = new RegExp('^Sheet "(.+)" holds (\\d+) tasks? but no register row points at it');
const problemText = (p) => (typeof p === 'string' ? p : (p?.message || ''));

/** A column whose LABEL names a milestone or plan, not a reporting week. */
const MILESTONE_LABEL = /\b(submission|submit|submitted|target|deadline|due|milestone|planned?)\b/i;

/**
 * Removes milestone columns from every site's weekly grid, in place.
 * Records what was removed on the site (detail.ignoredColumns) and in model.notes.
 * Safe to run twice.
 */
export function dropMilestoneColumns(model) {
  const notes = [];
  for (const site of model.sites || []) {
    const d = site.detail;
    if (!d?.weeks?.length) continue;
    const keep = d.weeks.map((w) => !MILESTONE_LABEL.test(String(w.label || '')));
    if (keep.every(Boolean)) continue;

    const pick = (arr) => (Array.isArray(arr) ? arr.filter((_, i) => keep[i]) : arr);
    const removed = d.weeks.filter((_, i) => !keep[i]).map((w) => ({ label: w.label, end: w.end }));
    d.ignoredColumns = [...(d.ignoredColumns || []), ...removed];
    d.weeks = d.weeks.filter((_, i) => keep[i]).map((w, i) => ({ ...w, number: i + 1 }));
    for (const t of d.tasks || []) {
      t.weekly = pick(t.weekly);
      if (t.rawWeekly) t.rawWeekly = pick(t.rawWeekly);
    }
    for (const c of d.categories || []) c.weekly = pick(c.weekly);
    site.weekCount = d.weeks.length;
    for (const r of removed) {
      notes.push(`${site.code}: the column "${r.label}"${r.end ? ` (${r.end})` : ''} looks like a milestone, not a week, so it was left out of the weekly status.`);
    }
  }
  model.notes = [...(model.notes || []), ...notes];
  return model;
}

/**
 * Sheets that hold tasks but that no row in the Sites register points at. The
 * parser reports these as text; this reads them back out so the page can show
 * them as a clear list instead of a line in a long warnings list.
 * @returns {{name: string, tasks: number}[]}
 */
export function unregisteredSheets(model) {
  const out = [];
  for (const p of model.problems || []) {
    const m = ORPHAN_SHEET.exec(problemText(p));
    if (m) out.push({ name: m[1], tasks: Number(m[2]) });
  }
  return out;
}

/** The parser problems that are NOT about unregistered sheets (those get their own notice). */
export function otherProblems(model) {
  return (model.problems || []).filter((p) => !ORPHAN_SHEET.test(problemText(p)));
}
