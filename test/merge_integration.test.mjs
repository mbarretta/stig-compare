// Characterization test: runs the real merge (DISA baseline XLSX × Vulcan CSV
// export) through merge_logic.js and pins the verified outcome — 17 conflicts,
// 97 concurred rows, 47 umbrella auto-keeps, 14 new rows. If the merge logic
// changes these numbers, the change must be intentional.
//
// The data files are local-only (gitignored); the test skips when they are
// absent so CI/fresh clones still pass.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import ExcelJS from 'exceljs';
import Papa from 'papaparse';
import { runMerge, worksheetToRows, buildKey } from '../merge_logic.js';

const here = dirname(fileURLToPath(import.meta.url));
const BASELINE_XLSX = join(here, 'U_GPOS_SRG_STIG_ChainguardOS_BS_20260514.xlsx');
const VULCAN_CSV = join(here, 'Chainguard-OS-CGOS-01-6.csv');
const haveData = existsSync(BASELINE_XLSX) && existsSync(VULCAN_CSV);

test('real-data merge reproduces the verified baseline outcome', { skip: !haveData && 'local STIG data files not present' }, async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(BASELINE_XLSX);
  const ws = wb.getWorksheet('DataTable') || wb.worksheets[0];
  const xlsxRows = worksheetToRows(ws);

  const csvRows = Papa.parse(readFileSync(VULCAN_CSV, 'utf8'), {
    header: true, skipEmptyLines: false,
  }).data;

  assert.equal(xlsxRows.length, 248);
  assert.equal(csvRows.length, 253);

  const res = runMerge(csvRows, xlsxRows);

  assert.deepEqual(res.metadata, {
    totalConflicts: 17,
    totalComments: 17,
    concurredCount: 97,
    autoResolvedCount: 202,
    newRowsCount: 14,
    unmatchedCount: 5,
    ambiguousCount: 4,
    oneToOneCount: 188,
    subMatchCount: 51,
    umbrellaAutoCount: 47,
    umbrellaConflictCount: 0,
    concurredConflictCount: 8,
  });

  // Latest government iteration with data is the 3rd.
  assert.equal(res.iteration.iteration, 3);
  assert.equal(res.iteration.vendorResponseColumn, '3rd Vendor Response');

  // No CSV row is lost: every row is either paired to a baseline row or new.
  const pairedCsv = new Set(Object.values(res.pairings));
  assert.equal(pairedCsv.size + res.newCsvIdx.length, csvRows.length);
  assert.ok(res.newCsvIdx.every((ci) => !pairedCsv.has(ci)));

  // Every pairing joins rows that share the SRGID+CCI key.
  for (const [xi, ci] of Object.entries(res.pairings)) {
    assert.equal(buildKey(xlsxRows[xi]), buildKey(csvRows[ci]));
  }

  // Conflicts on concurred rows are confined to the audit-critical columns.
  for (const c of res.conflicts.filter((c) => c.concurred)) {
    assert.ok(['Check', 'Fix', 'Status'].includes(c.column));
  }
});
