// Unit tests for the pure merge logic (merge_logic.js), promoted from the
// ad-hoc fixtures/*.mjs verification harnesses. Run with `npm test`.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  clean, norm, buildKey, extractSignature, jaccard,
  canonicalizeOs, resolveOsPhrasing, isMeaningfulComment, ordinal, diffWords,
  findLatestGovIteration, govConcurs, vendorConcurs, collectIterationColumns,
  latestNonEmpty, rowConcurred, parseSatisfiedBy, setsEqual, setDelta,
  cellText, runMerge, SOURCE,
} from '../merge_logic.js';

// ── string/cell helpers ──────────────────────────────────────────────────────

test('clean strips blanks, nan/none placeholders, and whitespace', () => {
  assert.equal(clean(null), '');
  assert.equal(clean(undefined), '');
  assert.equal(clean('  '), '');
  assert.equal(clean('NaN'), '');
  assert.equal(clean('None'), '');
  assert.equal(clean('  keep me  '), 'keep me');
  assert.equal(clean(42), '42');
});

test('norm lowercases and collapses whitespace', () => {
  assert.equal(norm('  The   Operating\n System '), 'the operating system');
});

test('buildKey joins SRGID and CCI', () => {
  assert.equal(buildKey({ SRGID: ' SRG-1 ', CCI: 'CCI-1' }), 'SRG-1|||CCI-1');
  assert.equal(buildKey({}), '|||');
});

test('cellText coerces every ExcelJS value shape to a string', () => {
  assert.equal(cellText(null), '');
  assert.equal(cellText('x'), 'x');
  assert.equal(cellText(42), '42');
  assert.equal(cellText({ richText: [{ text: 'a' }, { text: 'b' }] }), 'ab');
  assert.equal(cellText({ hyperlink: 'http://x', text: 'label' }), 'label');
  assert.equal(cellText({ formula: 'A1', result: 5 }), '5');
  assert.equal(cellText({ error: '#N/A' }), '');
  assert.equal(cellText(new Date('2026-01-02T00:00:00Z')), '2026-01-02T00:00:00.000Z');
});

test('ordinal', () => {
  assert.deepEqual([1, 2, 3, 4, 11].map(ordinal), ['1st', '2nd', '3rd', '4th', '11th']);
});

test('diffWords partitions both strings into eq/del/add runs', () => {
  const a = 'the quick brown fox';
  const b = 'the slow brown fox jumps';
  const ops = diffWords(a, b);
  const joined = (types) => ops.filter((o) => types.includes(o.type)).map((o) => o.text).join('');
  assert.equal(ops.filter((o) => o.type === 'eq' || o.type === 'del').map((o) => o.text).join(''), a);
  assert.equal(ops.filter((o) => o.type === 'eq' || o.type === 'add').map((o) => o.text).join(''), b);
  assert.ok(joined(['del']).includes('quick'));
  assert.ok(joined(['add']).includes('slow'));
  assert.deepEqual(diffWords('same', 'same'), [{ type: 'eq', text: 'same' }]);
});

// ── signature matching ───────────────────────────────────────────────────────

test('extractSignature collects paths, quoted strings, and technical tokens', () => {
  const sig = extractSignature({
    Check: 'Run chmod 0644 on /etc/audit/rules.d and verify "audit=1" via systemctl',
  });
  assert.ok(sig.has('/etc/audit/rules.d'));
  assert.ok(sig.has('audit=1'));
  assert.ok(sig.has('chmod'));
  assert.ok(sig.has('systemctl'));
  assert.ok(sig.has('0644'));
});

test('jaccard similarity', () => {
  assert.equal(jaccard(new Set(), new Set()), 0);
  assert.equal(jaccard(new Set(['a']), new Set(['a'])), 1);
  assert.equal(jaccard(new Set(['a', 'b']), new Set(['b', 'c'])), 1 / 3);
});

// ── OS-phrasing auto-resolve ─────────────────────────────────────────────────

test('resolveOsPhrasing auto-picks the Chainguard OS variant when that is the only diff', () => {
  const x = 'The operating system must lock the session.';
  const c = 'The Chainguard OS must lock the session.';
  assert.deepEqual(resolveOsPhrasing(x, c).chosen, c);
  assert.equal(resolveOsPhrasing(x, c).source, SOURCE.CSV);
  assert.equal(resolveOsPhrasing(c, x).source, SOURCE.XLSX);
  // any other difference disqualifies the auto-merge
  assert.equal(resolveOsPhrasing(x, 'The Chainguard OS must lock the screen.'), null);
});

test('canonicalizeOs removes both phrasings so equal remainders compare equal', () => {
  assert.equal(
    canonicalizeOs('The operating system locks.'),
    canonicalizeOs('The Chainguard OS locks.')
  );
});

// ── concurrence gate ─────────────────────────────────────────────────────────

test('govConcurs requires a leading "concur" and excludes non-concur', () => {
  assert.ok(govConcurs('Concur with status'));
  assert.ok(govConcurs('concur.'));
  assert.ok(!govConcurs('Non-concur'));
  assert.ok(!govConcurs('non concur'));
  assert.ok(!govConcurs('We concur')); // must start with it
  assert.ok(!govConcurs(''));
});

test('vendorConcurs matches "concur" anywhere but excludes non-concur', () => {
  assert.ok(vendorConcurs('We concur with the finding'));
  assert.ok(!vendorConcurs('We non-concur'));
  assert.ok(!vendorConcurs(''));
});

test('collectIterationColumns parses and sorts numbered gov/vendor columns', () => {
  const cols = collectIterationColumns([
    '2nd Government Comments', 'STIGID', '1st Government Comments',
    '1st Vendor Response', '3rd Vendor Response',
  ]);
  assert.deepEqual(cols.gov.map((c) => c.n), [1, 2]);
  assert.deepEqual(cols.vendor.map((c) => c.n), [1, 3]);
});

test('rowConcurred uses the LATEST non-empty cell per track', () => {
  const cols = collectIterationColumns([
    '1st Government Comments', '2nd Government Comments',
    '1st Vendor Response', '2nd Vendor Response',
  ]);
  // concurred at iteration 1, reopened at iteration 2 → not settled
  assert.ok(!rowConcurred({
    '1st Government Comments': 'Concur', '2nd Government Comments': 'Non-concur, see notes',
  }, cols));
  // gov track silent, vendor concurs → settled
  assert.ok(rowConcurred({ '2nd Vendor Response': 'We concur.' }, cols));
  assert.equal(latestNonEmpty({ '1st Government Comments': 'a', '2nd Government Comments': 'b' }, cols.gov), 'b');
});

test('findLatestGovIteration picks the highest iteration with data', () => {
  const rows = [
    { '1st Government Comments': 'x', '2nd Government Comments': '', '2nd Vendor Response': '', __wsRow: 2 },
    { '1st Government Comments': '', '2nd Government Comments': 'y', '2nd Vendor Response': '', __wsRow: 3 },
  ];
  const iter = findLatestGovIteration(rows);
  assert.equal(iter.iteration, 2);
  assert.equal(iter.label, '2nd');
  assert.equal(iter.govCommentColumn, '2nd Government Comments');
  assert.equal(iter.vendorResponseColumn, '2nd Vendor Response');
  assert.equal(findLatestGovIteration([{ STIGID: 'V-1' }]), null);
});

test('isMeaningfulComment filters bare concurrences', () => {
  assert.ok(!isMeaningfulComment('Concur'));
  assert.ok(!isMeaningfulComment('Concur with status.'));
  assert.ok(isMeaningfulComment('Concur, but please clarify the check text.'));
  assert.ok(isMeaningfulComment('Please tighten the check procedure.'));
});

// ── umbrella (Satisfied By) parsing ──────────────────────────────────────────

test('parseSatisfiedBy extracts a normalized CGOS id set', () => {
  const set = parseSatisfiedBy({
    'Vendor Comments': 'Satisfied By: CGOS-000123, cgos-000456;\nCGOS-000789.',
  });
  assert.deepEqual([...set].sort(), ['CGOS-000123', 'CGOS-000456', 'CGOS-000789']);
  assert.equal(parseSatisfiedBy({ 'Vendor Comments': 'Just a note' }), null);
  assert.equal(parseSatisfiedBy({ 'Vendor Comments': 'Satisfied By: nothing valid' }), null);
  assert.equal(parseSatisfiedBy({}), null);
});

test('setsEqual and setDelta', () => {
  const a = new Set(['A', 'B']);
  assert.ok(setsEqual(a, new Set(['B', 'A'])));
  assert.ok(!setsEqual(a, new Set(['A'])));
  assert.ok(!setsEqual(null, a));
  assert.deepEqual(setDelta(new Set(['A', 'B']), new Set(['B', 'C'])), { added: ['C'], removed: ['A'] });
});

// ── runMerge scenarios ───────────────────────────────────────────────────────
// Rows mirror the synthetic fixture set in fixtures/make_fixtures.mjs.

const xrow = (wsRow, props) => ({ __wsRow: wsRow, ...props });

test('runMerge: 1:1 matching, conflicts, auto-resolves, comments, new rows', () => {
  const xlsxRows = [
    // A: concurred → skipped, but the changed Check (audit-critical) surfaces
    xrow(2, { SRGID: 'SRG-1', CCI: 'CCI-1', STIGID: 'V-1', Requirement: 'req one', Check: 'check one',
      '1st Government Comments': 'Concur with status', '1st Vendor Response': 'Concur' }),
    // B: open, gov comment + Check conflict
    xrow(3, { SRGID: 'SRG-2', CCI: 'CCI-2', STIGID: 'V-2', Requirement: 'req two', Check: 'check two',
      '1st Government Comments': 'Please tighten the check procedure.', '1st Vendor Response': '' }),
    // F: Status placeholder → whole row taken from CSV
    xrow(4, { SRGID: 'SRG-6', CCI: 'CCI-6', STIGID: 'V-6', Status: 'Not Yet Determined',
      Requirement: 'req six', Check: 'check six' }),
    // G: OS-phrasing-only diff on Requirement → auto, CSV wins
    xrow(5, { SRGID: 'SRG-7', CCI: 'CCI-7', STIGID: 'V-7',
      Requirement: 'The operating system must lock the session.', Check: 'check seven' }),
    // H: OS-phrasing-only diff on Check → auto, XLSX (Chainguard OS side) wins
    xrow(6, { SRGID: 'SRG-8', CCI: 'CCI-8', STIGID: 'V-8', Requirement: 'req eight',
      Check: 'Inspect the Chainguard OS audit config.' }),
  ];
  const csvRows = [
    { SRGID: 'SRG-1', CCI: 'CCI-1', STIGID: 'V-1', Requirement: 'req one CHANGED', Check: 'check one CHANGED' },
    { SRGID: 'SRG-2', CCI: 'CCI-2', STIGID: 'V-2', Requirement: 'req two', Check: 'check two CHANGED' },
    { SRGID: 'SRG-6', CCI: 'CCI-6', STIGID: 'V-6', Status: 'Open', Requirement: 'req six CHANGED', Check: 'check six CHANGED' },
    { SRGID: 'SRG-7', CCI: 'CCI-7', STIGID: 'V-7', Requirement: 'The Chainguard OS must lock the session.', Check: 'check seven' },
    { SRGID: 'SRG-8', CCI: 'CCI-8', STIGID: 'V-8', Requirement: 'req eight', Check: 'Inspect the operating system audit config.' },
    // E: brand-new control → new row
    { SRGID: 'SRG-5', CCI: 'CCI-5', STIGID: 'V-5', Requirement: 'req five', Check: 'check five' },
  ];

  const res = runMerge(csvRows, xlsxRows);

  assert.equal(res.metadata.oneToOneCount, 5);
  assert.equal(res.metadata.newRowsCount, 1);
  assert.equal(csvRows[res.newCsvIdx[0]].STIGID, 'V-5');
  assert.equal(res.metadata.concurredCount, 1);

  // A: Requirement change suppressed (settled row), Check surfaced as concurred
  const aConflicts = res.conflicts.filter((c) => c.stigid === 'V-1');
  assert.deepEqual(aConflicts.map((c) => c.column), ['Check']);
  assert.ok(aConflicts[0].concurred);

  // B: ordinary conflict + a meaningful gov comment
  const bConflicts = res.conflicts.filter((c) => c.stigid === 'V-2');
  assert.deepEqual(bConflicts.map((c) => c.column), ['Check']);
  assert.ok(!bConflicts[0].concurred);
  assert.deepEqual(res.comments.map((c) => c.srgid), ['SRG-2']);
  assert.equal(res.comments[0].iterationLabel, '1st');

  // F: every differing column auto-taken from CSV
  const fAuto = res.autoResolved.filter((a) => a.stigid === 'V-6');
  assert.ok(fAuto.length >= 2);
  assert.ok(fAuto.every((a) => a.chosenSource === SOURCE.CSV && /Not Yet Determined/.test(a.reason)));
  assert.equal(fAuto.find((a) => a.column === 'Status').resolvedValue, 'Open');

  // G: OS phrasing → CSV's "Chainguard OS" wording chosen
  const gAuto = res.autoResolved.find((a) => a.stigid === 'V-7');
  assert.equal(gAuto.chosenSource, SOURCE.CSV);
  assert.equal(gAuto.resolvedValue, 'The Chainguard OS must lock the session.');

  // H: OS phrasing → XLSX's "Chainguard OS" wording kept
  const hAuto = res.autoResolved.find((a) => a.stigid === 'V-8');
  assert.equal(hAuto.chosenSource, SOURCE.XLSX);
  assert.equal(hAuto.resolvedValue, 'Inspect the Chainguard OS audit config.');

  // no conflicts beyond A and B
  assert.equal(res.conflicts.length, 2);
});

test('runMerge: umbrella controls diff the Satisfied-By set, not the text', () => {
  const xlsxRows = [
    // set unchanged → keep XLSX text, auto-resolved
    xrow(2, { SRGID: 'SRG-1', CCI: 'CCI-1', STIGID: 'V-1', Check: 'rendering of CGOS-000001',
      'Vendor Comments': 'Satisfied By: CGOS-000001, CGOS-000002' }),
    // set changed → real conflict with the delta attached
    xrow(3, { SRGID: 'SRG-2', CCI: 'CCI-2', STIGID: 'V-2', Check: 'rendering of CGOS-000003',
      'Vendor Comments': 'Satisfied By: CGOS-000003, CGOS-000004' }),
  ];
  const csvRows = [
    { SRGID: 'SRG-1', CCI: 'CCI-1', STIGID: 'V-1', Check: 'rendering of CGOS-000002',
      'Vendor Comments': 'Satisfied By: CGOS-000002, CGOS-000001' },
    { SRGID: 'SRG-2', CCI: 'CCI-2', STIGID: 'V-2', Check: 'rendering of CGOS-000005',
      'Vendor Comments': 'Satisfied By: CGOS-000003, CGOS-000005' },
  ];

  const res = runMerge(csvRows, xlsxRows);

  assert.equal(res.metadata.umbrellaAutoCount, 1);
  const auto = res.autoResolved.find((a) => a.stigid === 'V-1');
  assert.equal(auto.chosenSource, SOURCE.XLSX);
  assert.equal(auto.resolvedValue, 'rendering of CGOS-000001');

  assert.equal(res.metadata.umbrellaConflictCount, 1);
  const conflict = res.conflicts.find((c) => c.stigid === 'V-2');
  assert.deepEqual(conflict.umbrella.added, ['CGOS-000005']);
  assert.deepEqual(conflict.umbrella.removed, ['CGOS-000004']);
});

test('runMerge: STIGID-first pairing inside a shared SRGID+CCI group', () => {
  // Two controls under the same key, listed in a different order on each side.
  const xlsxRows = [
    xrow(2, { SRGID: 'SRG-1', CCI: 'CCI-1', STIGID: 'V-1a', Check: 'alpha check' }),
    xrow(3, { SRGID: 'SRG-1', CCI: 'CCI-1', STIGID: 'V-1b', Check: 'beta check' }),
  ];
  const csvRows = [
    { SRGID: 'SRG-1', CCI: 'CCI-1', STIGID: 'V-1b', Check: 'beta check CHANGED' },
    { SRGID: 'SRG-1', CCI: 'CCI-1', STIGID: 'V-1a', Check: 'alpha check CHANGED' },
  ];

  const res = runMerge(csvRows, xlsxRows);

  assert.deepEqual(res.pairings, { 0: 1, 1: 0 }); // paired by STIGID, not by order
  assert.equal(res.matchMethod[0], 'STIGID within SRGID+CCI');
  assert.equal(res.metadata.newRowsCount, 0); // nothing misread as new
  assert.equal(res.metadata.subMatchCount, 2);
});

test('runMerge: Jaccard sub-match pairs rows by content signature; blank signatures are ambiguous', () => {
  const xlsxRows = [
    xrow(2, { SRGID: 'SRG-1', CCI: 'CCI-1', Check: 'Verify permissions with chmod 0644 on /etc/audit/auditd.conf' }),
    xrow(3, { SRGID: 'SRG-1', CCI: 'CCI-1', Check: 'Verify ownership with chown on /etc/ssh/sshd_config' }),
    xrow(4, { SRGID: 'SRG-1', CCI: 'CCI-1', Check: 'no technical tokens here at all' }), // empty signature
  ];
  const csvRows = [
    { SRGID: 'SRG-1', CCI: 'CCI-1', Check: 'Verify ownership with chown on /etc/ssh/sshd_config UPDATED' },
    { SRGID: 'SRG-1', CCI: 'CCI-1', Check: 'Verify permissions with chmod 0644 on /etc/audit/auditd.conf UPDATED' },
  ];

  const res = runMerge(csvRows, xlsxRows);

  assert.deepEqual(res.pairings, { 0: 1, 1: 0 });
  assert.match(res.matchMethod[0], /sub-match \(Jaccard/);
  assert.deepEqual(res.ambiguousXlsxIdx, [2]);
  assert.deepEqual(res.unmatchedXlsxIdx, []); // ambiguous rows are excluded from plain-unmatched
});
