// Inspect the exported XLSX to assert concurrence gating held on export.
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(join(here, '..', '.playwright-mcp', 'working-merged.xlsx'));
const ws = wb.getWorksheet('DataTable') || wb.worksheets[0];

const headers = {};
ws.getRow(1).eachCell((c, n) => { headers[c.value] = n; });
const get = (r, name) => ws.getRow(r).getCell(headers[name]).value;
const fill = (r, name) => ws.getRow(r).getCell(headers[name]).fill?.fgColor?.argb;

const checks = [];
const assert = (label, cond) => checks.push(`${cond ? 'PASS' : 'FAIL'}  ${label}`);

// Row 2 = V-1 (concurred) — must be untouched
assert('V-1 Requirement untouched (req one)', get(2, 'Requirement') === 'req one');
assert('V-1 Check untouched (check one)', get(2, 'Check') === 'check one');
assert('V-1 gov fill still green', fill(2, '1st Government Comments') === 'FFC6EFCE');
assert('V-1 vendor still Concur', get(2, '1st Vendor Response') === 'Concur');

// Open rows with no decision → XLSX kept
assert('V-2 Check kept (check two)', get(3, 'Check') === 'check two');
assert('V-4 Requirement kept (req four)', get(5, 'Requirement') === 'req four');

// V-6 (row 6): Status was "Not Yet Determined" -> whole row taken from CSV
assert('V-6 Status taken from CSV (Open)', get(6, 'Status') === 'Open');
assert('V-6 Check taken from CSV', get(6, 'Check') === 'check six CHANGED');
assert('V-6 Requirement taken from CSV', get(6, 'Requirement') === 'req six CHANGED');

// V-7 (row 7): only OS-phrasing diff in Requirement -> Chainguard OS wins
assert('V-7 Requirement -> Chainguard OS', get(7, 'Requirement') === 'The Chainguard OS must lock the session.');

// V-8 (row 8): OS-phrasing diff in Check; XLSX "Chainguard OS" wins over CSV generic
assert('V-8 Check keeps Chainguard OS', get(8, 'Check') === 'Inspect the Chainguard OS audit config.');

// New row V-5 appended
let v5 = null;
ws.eachRow((row, n) => { if (row.getCell(headers['STIGID']).value === 'V-5') v5 = n; });
assert('V-5 new row appended', v5 !== null);

console.log(checks.join('\n'));
console.log(checks.every((c) => c.startsWith('PASS')) ? '\nALL PASS' : '\nSOME FAILED');
