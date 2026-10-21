/* Shared hand-checked fixture (site F-01 at W06). Used by verify-core and verify-brief. */
export const W = (n) => new Date(2026, 0, 4 + 7 * n);   // week-ending Sundays


export function fixtureWorkbook(XLSX, STATUS) {
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['Site code', 'Description (generic)', 'Wave / batch', 'Start date', 'Target submission',
     'Status', 'Coordinator (R code)', 'Detail sheet'],
    ['F-01', 'Type A', 'Wave 1', new Date(2026, 0, 5), new Date(2026, 2, 1), 'In progress', 'R1', 'F-01'],
    ['F-02', 'Type A', 'Wave 1', new Date(2026, 0, 5), new Date(2026, 3, 1), 'Planned', 'R1', ''],
  ]), 'Sites');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['Order', 'ID', 'Category', 'Standard task / scope item', 'Definition of done (category)',
     'Typical duration (days)', 'Applies to', 'Default prerequisites (PT IDs)'],
    [10, 'C010', 'Initial setup', '', 'Project set up', 3, 'All', 'PT01'],
    ['', 'T010-01', 'Initial setup', 'Create folders', '', '', 'All', ''],
    [20, 'C020', 'Modelling', '', 'Modelled to LOD', 10, 'All', ''],
    ['', 'T020-01', 'Modelling', 'Foundations', '', '', 'Struct', ''],
    [],
    ['Prereq ID', 'Prerequisite', 'Why it is needed', 'Usually provided by', 'Needed before (category ID)'],
    ['PT01', 'Template file', 'Sets standards', 'BIM lead', 'C010'],
  ]), 'Task Template');

  const head = ['Task ID', 'Category', 'Discipline', 'Task / component', 'Weight', 'Type',
    'Date added', 'Target week'];
  const tail = ['Completed - week', 'Completed - date', 'Responsible (R code)',
    'Prerequisite (P IDs)', 'Depends on (Task IDs)'];
  const wLabels = ['W01', 'W02', 'W03', 'W04', 'W05', 'W06'];

  const F = STATUS.FINISHED, P = STATUS.WIP, B = STATUS.BLOCKED,
        Wt = STATUS.WAITING, N = STATUS.NOT_STARTED, NA = STATUS.NA;

  const rows = [
    ['C010', 'Initial setup', 'All', '[Category] Initial setup', '', '', '', '',
      F, F, F, F, F, F, '', '', '', '', ''],
    ['T010-01', 'Initial setup', 'All', 'Create folders', 1, 'Regular', '', 1,
      F, F, F, F, F, F, 1, new Date(2026, 0, 11), 'R1', 'P01', ''],
    ['T010-02', 'Initial setup', 'All', 'Apply template', 2, 'Regular', '', 2,
      N, F, F, F, F, F, 2, new Date(2026, 0, 18), 'R1', 'P01', 'T010-01'],
    ['C020', 'Modelling', 'All', '[Category] Modelling', '', '', '', '',
      N, N, P, P, P, P, '', '', '', '', ''],
    ['T020-01', 'Modelling', 'Struct', 'Foundations', 1, 'Regular', '', 3,
      N, N, F, F, F, F, 3, new Date(2026, 0, 25), 'R2', '', 'T010-02'],
    ['T020-02', 'Modelling', 'Struct', 'Columns', 1, 'Regular', '', 4,
      N, N, N, N, F, F, 5, new Date(2026, 1, 8), 'R2', '', 'T020-01'],
    ['T020-03', 'Modelling', 'Struct', 'Beams', 1, 'Regular', '', 5,
      N, N, N, P, P, P, '', '', 'R2', '', 'T020-02'],
    ['T020-04', 'Modelling', 'Arch', 'Walls', 1, 'Regular', '', 5,
      N, N, N, N, P, P, '', '', 'R3', '', ''],
    ['T020-05', 'Modelling', 'MEP', 'Ductwork', 1, 'Regular', '', 6,
      N, N, N, B, B, B, '', '', 'R4', 'P02', ''],
    ['T020-06', 'Modelling', 'MEP', 'Pipework', 1, 'Regular', '', 6,
      N, N, N, N, Wt, Wt, '', '', 'R4', 'P02', ''],
    ['T020-07', 'Modelling', 'Arch', 'Ceilings', 1, 'Regular', '', 7,
      N, N, N, N, N, N, '', '', 'R3', '', ''],
    ['T020-08', 'Modelling', 'Arch', 'Roof', 1, 'Regular', '', 7,
      NA, NA, NA, NA, NA, NA, '', '', '', '', ''],
    ['T020-09', 'Modelling', 'Struct', 'Extra bracing', 1, 'Additional', new Date(2026, 1, 2), 8,
      '', '', '', '', N, N, '', '', 'R2', '', ''],
  ];

  const aoa = [
    ['SITE F-01'],
    ['Site code', 'Start date', 'Target submission'],
    ['F-01', new Date(2026, 0, 5), new Date(2026, 2, 1)],
    [],
    ['TABLE 1 - TASK LIST AND WEEKLY STATUS'],
    ['', '', '', '', '', '', '', '', ...wLabels],
    [...head, W(1), W(2), W(3), W(4), W(5), W(6), ...tail],
    ...rows,
    [],
    ['TABLE 2 - PREREQUISITES FOR THIS SITE'],
    ['Prereq ID', 'Prerequisite', 'From template (PT ID)', 'Needed for (Task IDs)',
     'Provided by (R code / party)', 'Required by week', 'Status', 'Date received', 'Notes'],
    ['P01', 'Template file', 'PT01', 'T010-01,T010-02', 'R1', 1, 'Received', new Date(2026, 0, 6), ''],
    ['P02', 'Ceiling void zones', 'PT06', 'T020-05,T020-06', 'R3', 3, 'Outstanding', '', ''],
    [],
    ['TABLE 3 - WAITING ON / BLOCKED LOG'],
    ['Log ID', 'Week (W no)', 'Task ID', 'Status raised', 'Reason',
     'Waiting on (R code / party)', 'Raised date', 'Expected clear date', 'Cleared date', 'Notes'],
    ['L01', 5, 'T020-06', 'Waiting on', 'Ceiling void zones not confirmed', 'R3',
     new Date(2026, 1, 2), new Date(2026, 1, 9), '', ''],
  ];

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'F-01');
  return wb;
}

