// Generates synthetic CSV + XLSX fixtures to exercise concurrence gating.
import ExcelJS from 'exceljs';
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));

const HEADERS = [
  'SRGID', 'CCI', 'STIGID', 'Requirement', 'Check', 'Fix', 'VulDiscussion',
  '1st Government Comments', '1st Vendor Response',
];

// xlsx rows — the source of truth
const xlsx = [
  // A: concurred (gov Concur + vendor Concur) -> SKIP entirely
  { SRGID: 'SRG-1', CCI: 'CCI-1', STIGID: 'V-1', Requirement: 'req one', Check: 'check one', Fix: 'fix one', VulDiscussion: 'disc one',
    '1st Government Comments': 'Concur with status', '1st Vendor Response': 'Concur' },
  // B: open + gov asks for a change -> conflict + comment
  { SRGID: 'SRG-2', CCI: 'CCI-2', STIGID: 'V-2', Requirement: 'req two', Check: 'check two', Fix: 'fix two', VulDiscussion: 'disc two',
    '1st Government Comments': 'Please tighten the check procedure.', '1st Vendor Response': '' },
  // C: open, no comment -> conflict only
  { SRGID: 'SRG-3', CCI: 'CCI-3', STIGID: 'V-3', Requirement: 'req three', Check: 'check three', Fix: 'fix three', VulDiscussion: 'disc three',
    '1st Government Comments': '', '1st Vendor Response': '' },
  // D: gov Concur but vendor blank -> NOT settled, still open -> conflict
  { SRGID: 'SRG-4', CCI: 'CCI-4', STIGID: 'V-4', Requirement: 'req four', Check: 'check four', Fix: 'fix four', VulDiscussion: 'disc four',
    '1st Government Comments': 'Concur', '1st Vendor Response': '' },
];

// csv rows — fresh from Vulcan, each differs from xlsx in some field
const csv = [
  { SRGID: 'SRG-1', CCI: 'CCI-1', STIGID: 'V-1', Requirement: 'req one CHANGED', Check: 'check one CHANGED', Fix: 'fix one', VulDiscussion: 'disc one STALE' },
  { SRGID: 'SRG-2', CCI: 'CCI-2', STIGID: 'V-2', Requirement: 'req two', Check: 'check two CHANGED', Fix: 'fix two', VulDiscussion: 'disc two' },
  { SRGID: 'SRG-3', CCI: 'CCI-3', STIGID: 'V-3', Requirement: 'req three', Check: 'check three CHANGED', Fix: 'fix three', VulDiscussion: 'disc three' },
  { SRGID: 'SRG-4', CCI: 'CCI-4', STIGID: 'V-4', Requirement: 'req four CHANGED', Check: 'check four', Fix: 'fix four', VulDiscussion: 'disc four' },
  // E: brand-new vuln not in xlsx -> appended as new row
  { SRGID: 'SRG-5', CCI: 'CCI-5', STIGID: 'V-5', Requirement: 'req five', Check: 'check five', Fix: 'fix five', VulDiscussion: 'disc five' },
];

// write CSV
const csvText = [HEADERS.join(',')].concat(
  csv.map((r) => HEADERS.map((h) => JSON.stringify(r[h] ?? '')).join(','))
).join('\n');
writeFileSync(join(here, 'vulcan.csv'), csvText);

// write XLSX with green/yellow fills to mimic the real sheet
const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('DataTable');
ws.addRow(HEADERS);
const GREEN = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
const YELLOW = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
xlsx.forEach((r, i) => {
  const row = ws.addRow(HEADERS.map((h) => r[h] ?? ''));
  const govCell = row.getCell(HEADERS.indexOf('1st Government Comments') + 1);
  const venCell = row.getCell(HEADERS.indexOf('1st Vendor Response') + 1);
  const gov = r['1st Government Comments'];
  const ven = r['1st Vendor Response'];
  if (/concur/i.test(gov)) govCell.fill = GREEN; else if (gov) govCell.fill = YELLOW;
  if (/concur/i.test(ven)) venCell.fill = GREEN;
});
await wb.xlsx.writeFile(join(here, 'working.xlsx'));

console.log('Wrote fixtures/vulcan.csv and fixtures/working.xlsx');
console.log('Expected: A skipped (concurred), B conflict+comment, C conflict, D conflict (vendor blank), E new row.');
