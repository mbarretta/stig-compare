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

// New row V-5 appended
let v5 = null;
ws.eachRow((row, n) => { if (row.getCell(headers['STIGID']).value === 'V-5') v5 = n; });
assert('V-5 new row appended', v5 !== null);

console.log(checks.join('\n'));
console.log(checks.every((c) => c.startsWith('PASS')) ? '\nALL PASS' : '\nSOME FAILED');
