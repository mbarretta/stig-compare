// Pure merge logic shared by the app (merge_review.jsx) and the test suite
// (test/*.test.mjs). No React, no DOM, no imports — safe to load in plain Node.

export const SHARED_COLS = [
  'IA Control', 'CCI', 'SRGID', 'STIGID', 'SRG Requirement', 'Requirement',
  'SRG VulDiscussion', 'VulDiscussion', 'Status', 'SRG Check', 'Check',
  'SRG Fix', 'Fix', 'Severity', 'Mitigation', 'Artifact Description',
  'Status Justification',
];
export const CSV_ONLY_COLS = ['Vendor Comments', 'InSpec Control Body'];

export const SIDE = { XLSX: 'xlsx', CSV: 'csv' };
export const SOURCE = { XLSX: 'XLSX', CSV: 'CSV' };

export const cellKey = (xi, col) => `${xi}|${col}`;
export const conflictId = (xi, col) => `r${xi}_${col.replace(/\s+/g, '_')}`;

export function groupByKey(rows) {
  const m = new Map();
  rows.forEach((r, i) => {
    const k = buildKey(r);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(i);
  });
  return m;
}

export function bucketByXlsxRow(byRow, row, seed, defaults) {
  if (!byRow.has(row)) {
    byRow.set(row, {
      xlsxRow: row,
      stigid: seed?.stigid || '',
      srgid: seed?.srgid || '',
      cci: seed?.cci || '',
      ...defaults,
    });
  }
  return byRow.get(row);
}

export function clean(s) {
  if (s === null || s === undefined) return '';
  const str = String(s).trim();
  const lower = str.toLowerCase();
  if (str === '' || lower === 'nan' || lower === 'none') return '';
  return str;
}

export function norm(s) {
  return clean(s).toLowerCase().replace(/\s+/g, ' ');
}

export function buildKey(row) {
  return clean(row['SRGID']) + '|||' + clean(row['CCI']);
}

export function extractSignature(row) {
  const text = ['Requirement', 'Check', 'Fix', 'VulDiscussion']
    .map((f) => clean(row[f] || ''))
    .join(' ');
  const sig = new Set();
  // file paths
  for (const m of text.matchAll(/\/[a-zA-Z0-9_\-./]+(?:\/|\b)/g)) {
    let v = m[0].replace(/[/.]+$/, '');
    if (v.length > 3) sig.add(v.toLowerCase());
  }
  // quoted strings
  for (const m of text.matchAll(/"([^"]{2,40})"/g)) {
    sig.add(m[1].toLowerCase());
  }
  // technical tokens
  for (const m of text.matchAll(
    /\b(?:syscall|chmod|chown|chgrp|chage|umask|mode\s+\d+|0[0-7]{3,4}|systemctl|systemd-[a-z]+)\b/gi
  )) {
    sig.add(m[0].toLowerCase());
  }
  return sig;
}

export function jaccard(a, b) {
  if (a.size === 0 && b.size === 0) return 0;
  let inter = 0;
  for (const x of a) if (b.has(x)) inter++;
  const union = a.size + b.size - inter;
  return union > 0 ? inter / union : 0;
}

export function isChainguardPreferred(val) {
  return /\bchainguard\s*os\b/i.test(val);
}
export function isGenericPhrasing(val) {
  return /\boperating\s+system\b/i.test(val) && !isChainguardPreferred(val);
}

// If the ONLY difference between two cells is "operating system" ⟷ "Chainguard
// OS" phrasing, return the "Chainguard OS" variant to auto-merge; else null.
export function canonicalizeOs(s) {
  return norm(s).replace(/chainguard\s*os/g, '').replace(/operating\s+system/g, '');
}
export function resolveOsPhrasing(xv, cv) {
  if (canonicalizeOs(xv) !== canonicalizeOs(cv)) return null; // other differences too
  const reason = 'Only difference is "operating system" → "Chainguard OS"';
  if (isChainguardPreferred(xv)) return { chosen: xv, source: SOURCE.XLSX, reason };
  if (isChainguardPreferred(cv)) return { chosen: cv, source: SOURCE.CSV, reason };
  return null;
}

export function isMeaningfulComment(text) {
  const t = clean(text);
  if (!t) return false;
  if (/^concur\b[\s.,;:]*(with\s+status[\s.,;:]*)?$/i.test(t)) return false;
  return true;
}

export function ordinal(n) {
  if (n === 1) return '1st';
  if (n === 2) return '2nd';
  if (n === 3) return '3rd';
  return n + 'th';
}

// Word-level LCS diff
export function diffWords(a, b) {
  if (a === b) return [{ type: 'eq', text: a }];
  const tokenize = (s) => (s || '').split(/(\s+)/).filter((t) => t !== '');
  const aT = tokenize(a);
  const bT = tokenize(b);
  const n = aT.length;
  const m = bT.length;
  if (n === 0) return bT.length ? [{ type: 'add', text: b }] : [];
  if (m === 0) return aT.length ? [{ type: 'del', text: a }] : [];

  const dp = [];
  for (let i = 0; i <= n; i++) dp.push(new Array(m + 1).fill(0));
  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      if (aT[i - 1] === bT[j - 1]) dp[i][j] = dp[i - 1][j - 1] + 1;
      else dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
    }
  }
  const ops = [];
  let i = n, j = m;
  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && aT[i - 1] === bT[j - 1]) {
      ops.unshift({ type: 'eq', text: aT[i - 1] });
      i--; j--;
    } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
      ops.unshift({ type: 'add', text: bT[j - 1] });
      j--;
    } else {
      ops.unshift({ type: 'del', text: aT[i - 1] });
      i--;
    }
  }
  const out = [];
  for (const op of ops) {
    const last = out[out.length - 1];
    if (last && last.type === op.type) last.text += op.text;
    else out.push({ ...op });
  }
  return out;
}

export function findLatestGovIteration(xlsxRows) {
  if (xlsxRows.length === 0) return null;
  const headers = Object.keys(xlsxRows[0]);
  const govCols = [];
  for (const h of headers) {
    const m = h.match(/^(\d+)\w*\s+Government\s+Comments?$/i);
    if (m) govCols.push({ n: parseInt(m[1], 10), name: h });
  }
  govCols.sort((a, b) => a.n - b.n);

  let latest = null;
  for (const { n, name } of govCols) {
    const hasData = xlsxRows.some((r) => clean(r[name]));
    if (hasData) latest = { n, name };
  }
  if (!latest) return null;

  const venName = headers.find((h) =>
    new RegExp('^' + latest.n + '\\w*\\s+Vendor\\s+Response$', 'i').test(h)
  );
  return {
    iteration: latest.n,
    label: ordinal(latest.n),
    govCommentColumn: latest.name,
    vendorResponseColumn: venName || '',
  };
}

// Concurrence: a row is settled when the LATEST non-empty Government Comments
// cell starts with "concur", OR the LATEST non-empty Vendor Response cell
// contains "concur". Text-only (no fill color needed); "non-concur" excluded.
// Settled rows are skipped entirely — Vulcan's data for them is stale.
export function govConcurs(text) {
  const t = clean(text);
  return /^concur/i.test(t) && !/^non[\s-]*concur/i.test(t);
}
export function vendorConcurs(text) {
  const t = clean(text);
  return /concur/i.test(t) && !/non[\s-]*concur/i.test(t);
}

export function collectIterationColumns(headers) {
  const gov = [];
  const vendor = [];
  for (const h of headers) {
    let m = h.match(/^(\d+)\w*\s+Government\s+Comments?$/i);
    if (m) { gov.push({ n: parseInt(m[1], 10), name: h }); continue; }
    m = h.match(/^(\d+)\w*\s+Vendor\s+Response$/i);
    if (m) { vendor.push({ n: parseInt(m[1], 10), name: h }); }
  }
  gov.sort((a, b) => a.n - b.n);
  vendor.sort((a, b) => a.n - b.n);
  return { gov, vendor };
}

// cols sorted ascending by iteration — the last non-empty cell wins.
export function latestNonEmpty(row, cols) {
  let val = '';
  for (const c of cols) { const t = clean(row[c.name]); if (t) val = t; }
  return val;
}

export function rowConcurred(row, cols) {
  return govConcurs(latestNonEmpty(row, cols.gov)) ||
         vendorConcurs(latestNonEmpty(row, cols.vendor));
}

// ── Umbrella ("Satisfied By") awareness ─────────────────────────────────────
// Many SRG controls are umbrella rows: their Vendor Comments cell carries a
// "Satisfied By: CGOS-…, CGOS-…" list naming the concrete base rules that
// satisfy them. For those rows the Check/Fix cell is NOT authored content — it
// is a *rendering* of one of the satisfying base rules (the first one listed),
// so it changes between exports purely because that list gets reordered, even
// when the control's actual requirements are identical. We therefore compare
// the Satisfied-By SET (order-independent) instead of the Check/Fix text:
//   · set unchanged → keep the baseline XLSX text, no conflict (avoids churn);
//   · set changed   → real requirement change, surfaced as a conflict with the
//                     added/removed base rules shown.
// Leaf controls (no Satisfied-By list) keep the normal text comparison.
export const UMBRELLA_DERIVED_COLS = new Set(['Check', 'Fix']);

// A "concurred" row is otherwise settled and skipped, but the Vulcan export can
// still carry a live, unresolved divergence in the audit-critical columns (a row
// can read "Concur…" while the export's Check/Fix/Status disagree with the
// baseline). For concurred rows we therefore re-examine ONLY these columns and
// surface any post-auto-resolve divergence as a conflict tagged "previously
// concurred"; every other column on a settled row stays untouched.
export const CONCURRED_REVIEW_COLS = new Set(['Check', 'Fix', 'Status']);

export function parseSatisfiedBy(row) {
  const t = clean(row['Vendor Comments']);
  if (!t) return null;
  const m = t.match(/Satisfied\s+By\s*:\s*([\s\S]+)/i);
  if (!m) return null;
  const ids = m[1]
    .split(/[,\n;]+/)
    .map((x) => x.trim().replace(/\.+$/, '').toUpperCase())
    .filter((x) => /^CGOS-/.test(x));
  return ids.length ? new Set(ids) : null;
}

export function setsEqual(a, b) {
  if (!a || !b || a.size !== b.size) return false;
  for (const x of a) if (!b.has(x)) return false;
  return true;
}

export function setDelta(xSet, cSet) {
  const added = [...cSet].filter((x) => !xSet.has(x)).sort();   // in CSV, not XLSX
  const removed = [...xSet].filter((x) => !cSet.has(x)).sort(); // in XLSX, not CSV
  return { added, removed };
}

export function runMerge(csvRows, xlsxRows) {
  const log = [];
  log.push(`CSV: ${csvRows.length} rows · XLSX: ${xlsxRows.length} rows`);

  const csvByKey = groupByKey(csvRows);
  const xlsxByKey = groupByKey(xlsxRows);
  log.push(`Unique SRGID+CCI keys: CSV=${csvByKey.size}, XLSX=${xlsxByKey.size}, shared=${[...csvByKey.keys()].filter((k) => xlsxByKey.has(k)).length}`);

  const xlsxToCsv = new Map();
  const matchMethod = new Map();
  const ambiguousXlsx = new Set();
  const usedCsv = new Set();

  let countOneOne = 0;
  let countSubMatch = 0;

  const allKeys = new Set([...csvByKey.keys(), ...xlsxByKey.keys()]);
  for (const key of allKeys) {
    const csvIdx = csvByKey.get(key) || [];
    const xlsxIdx = xlsxByKey.get(key) || [];
    if (csvIdx.length === 1 && xlsxIdx.length === 1) {
      xlsxToCsv.set(xlsxIdx[0], csvIdx[0]);
      matchMethod.set(xlsxIdx[0], '1:1 SRGID+CCI');
      usedCsv.add(csvIdx[0]);
      countOneOne++;
    } else if (csvIdx.length >= 1 && xlsxIdx.length >= 1) {
      const usedX = new Set();
      const usedC = new Set();

      // Pass 1 — pair rows that share a STIGID within this SRGID+CCI group.
      // Same STIGID under the same key is the same control, so this prevents a
      // CSV row from being treated as "new" (and inserted as a duplicate) when an
      // identically-numbered baseline row already exists. STIGID is only trusted
      // as a positive signal here; renumbered rows still fall through to Jaccard.
      const csvByStig = new Map();
      for (const ci of csvIdx) {
        const sid = clean(csvRows[ci]['STIGID']);
        if (sid && !csvByStig.has(sid)) csvByStig.set(sid, ci);
      }
      for (const xi of xlsxIdx) {
        const sid = clean(xlsxRows[xi]['STIGID']);
        if (!sid) continue;
        const ci = csvByStig.get(sid);
        if (ci === undefined || usedC.has(ci)) continue;
        xlsxToCsv.set(xi, ci);
        matchMethod.set(xi, 'STIGID within SRGID+CCI');
        usedX.add(xi);
        usedC.add(ci);
        usedCsv.add(ci);
        countSubMatch++;
      }

      // Pass 2 — Jaccard signature match for whatever is still unpaired.
      const csvSigs = new Map(csvIdx.filter((i) => !usedC.has(i)).map((i) => [i, extractSignature(csvRows[i])]));
      const xlsxSigs = new Map(xlsxIdx.filter((i) => !usedX.has(i)).map((i) => [i, extractSignature(xlsxRows[i])]));
      const candidates = [];
      for (const [xi, xs] of xlsxSigs) {
        if (xs.size === 0) continue;
        for (const [ci, cs] of csvSigs) {
          if (cs.size === 0) continue;
          const score = jaccard(xs, cs);
          if (score >= 0.3) candidates.push({ score, xi, ci });
        }
      }
      candidates.sort((a, b) => b.score - a.score);
      for (const { score, xi, ci } of candidates) {
        if (usedX.has(xi) || usedC.has(ci)) continue;
        xlsxToCsv.set(xi, ci);
        matchMethod.set(xi, `sub-match (Jaccard ${score.toFixed(2)})`);
        usedX.add(xi);
        usedC.add(ci);
        usedCsv.add(ci);
        countSubMatch++;
      }
      for (const xi of xlsxIdx) {
        if (xlsxToCsv.has(xi)) continue;
        const xs = xlsxSigs.get(xi);
        if (xs && xs.size === 0) ambiguousXlsx.add(xi);
      }
    }
  }
  log.push(`Matched: ${xlsxToCsv.size} (${countOneOne} one-to-one, ${countSubMatch} sub-match)`);

  const newCsvIdx = [];
  for (let i = 0; i < csvRows.length; i++) if (!usedCsv.has(i)) newCsvIdx.push(i);
  const unmatchedXlsxIdxAll = [];
  for (let i = 0; i < xlsxRows.length; i++) if (!xlsxToCsv.has(i)) unmatchedXlsxIdxAll.push(i);
  log.push(`New CSV rows: ${newCsvIdx.length} · Unmatched XLSX: ${unmatchedXlsxIdxAll.length} (${ambiguousXlsx.size} ambiguous)`);

  // Concurrence gate: settled rows (gov + vendor both concurred) are skipped.
  const iterCols = collectIterationColumns(xlsxRows.length ? Object.keys(xlsxRows[0]) : []);
  const concurredXlsxIdx = new Set();
  xlsxRows.forEach((r, i) => { if (rowConcurred(r, iterCols)) concurredXlsxIdx.add(i); });
  let concurredMatched = 0;
  for (const xi of xlsxToCsv.keys()) if (concurredXlsxIdx.has(xi)) concurredMatched++;
  log.push(`Concurred rows skipped: ${concurredMatched} of ${xlsxToCsv.size} matched (${concurredXlsxIdx.size} concurred total)`);

  const conflicts = [];
  const autoResolved = [];
  const resolvedCells = new Map();
  const conflictCellSet = new Set();

  const pushAuto = (xi, ci, col, chosen, source, reason, xv, cv) => {
    resolvedCells.set(cellKey(xi, col), chosen);
    autoResolved.push({
      xlsxRow: xlsxRows[xi].__wsRow,
      stigid: clean(csvRows[ci]['STIGID']),
      srgid: clean(xlsxRows[xi]['SRGID']),
      cci: clean(xlsxRows[xi]['CCI']),
      column: col,
      chosenSource: source,
      reason,
      resolvedValue: chosen,
      csvValue: cv,
      xlsxValue: xv,
    });
  };

  for (const [xi, ci] of xlsxToCsv) {
    const isConcurred = concurredXlsxIdx.has(xi);
    const xr = xlsxRows[xi];
    const cr = csvRows[ci];

    // Row-level auto-merge: XLSX Status was a placeholder ("Not Yet Determined")
    // and Vulcan now has a real determination — take the whole row from CSV.
    // Never bulk-replace a settled (concurred) row this way.
    const wholeRowToCsv =
      !isConcurred &&
      norm(xr['Status']) === 'not yet determined' &&
      clean(cr['Status']) !== '' &&
      norm(cr['Status']) !== 'not yet determined';

    // Umbrella detection: both sides name a "Satisfied By" base-rule set.
    const xSat = parseSatisfiedBy(xr);
    const cSat = parseSatisfiedBy(cr);
    const isUmbrella = !!(xSat && cSat);
    const satUnchanged = isUmbrella && setsEqual(xSat, cSat);

    for (const col of SHARED_COLS) {
      // Settled rows: only the audit-critical columns are re-examined.
      if (isConcurred && !CONCURRED_REVIEW_COLS.has(col)) continue;
      const cv = clean(cr[col]);
      const xv = clean(xr[col]);
      if (cv === '' || xv === '') continue;
      if (norm(cv) === norm(xv)) continue;

      // Row-level: Status placeholder resolved — take CSV for the whole row.
      if (wholeRowToCsv) {
        pushAuto(xi, ci, col, cv, SOURCE.CSV, 'Status was "Not Yet Determined" — whole row taken from CSV', xv, cv);
        continue;
      }

      // Umbrella control: Check/Fix is a representative of the Satisfied-By set,
      // not authored text — diff the SET, not the words.
      if (UMBRELLA_DERIVED_COLS.has(col) && isUmbrella) {
        if (satUnchanged) {
          // Same requirements; text differs only by which base rule is shown.
          // Keep the approved baseline (XLSX) verbatim — no churn, no conflict.
          pushAuto(
            xi, ci, col, xv, SOURCE.XLSX,
            `Umbrella control — "Satisfied By" set unchanged (${xSat.size} base rules); ${col} differs only by which base rule is surfaced`,
            xv, cv
          );
          continue;
        }
        // Set changed → a genuine requirement change. Flag for review with the
        // actual base-rule delta attached.
        const delta = setDelta(xSat, cSat);
        conflictCellSet.add(cellKey(xi, col));
        conflicts.push({
          id: conflictId(xi, col),
          xlsxRow: xlsxRows[xi].__wsRow,
          xlsxIndex: xi,
          csvIndex: ci,
          stigid: clean(cr['STIGID']),
          srgid: clean(xr['SRGID']),
          cci: clean(xr['CCI']),
          matchMethod: matchMethod.get(xi),
          column: col,
          csvValue: cv,
          xlsxValue: xv,
          concurred: isConcurred,
          umbrella: { added: delta.added, removed: delta.removed, xCount: xSat.size, cCount: cSat.size },
        });
        continue;
      }

      // Cell-level: only difference is "operating system" → "Chainguard OS".
      const os = resolveOsPhrasing(xv, cv);
      if (os) {
        pushAuto(xi, ci, col, os.chosen, os.source, os.reason, xv, cv);
        continue;
      }

      // Auto-resolve: Requirement column "Chainguard OS" preference (broader)
      if (col === 'Requirement') {
        const xC = isChainguardPreferred(xv);
        const cC = isChainguardPreferred(cv);
        const xG = isGenericPhrasing(xv);
        const cG = isGenericPhrasing(cv);
        let chosen = null;
        let source = null;
        let reason = null;
        if (xC && cG) { chosen = xv; source = SOURCE.XLSX; reason = "XLSX has 'Chainguard OS', CSV has 'operating system'"; }
        else if (cC && xG) { chosen = cv; source = SOURCE.CSV; reason = "CSV has 'Chainguard OS', XLSX has 'operating system'"; }
        else if (xC && !cC) { chosen = xv; source = SOURCE.XLSX; reason = "XLSX has 'Chainguard OS', CSV does not"; }
        else if (cC && !xC) { chosen = cv; source = SOURCE.CSV; reason = "CSV has 'Chainguard OS', XLSX does not"; }
        if (chosen !== null) {
          pushAuto(xi, ci, col, chosen, source, reason, xv, cv);
          continue;
        }
      }

      conflictCellSet.add(cellKey(xi, col));
      conflicts.push({
        id: conflictId(xi, col),
        xlsxRow: xlsxRows[xi].__wsRow,
        xlsxIndex: xi,
        csvIndex: ci,
        stigid: clean(cr['STIGID']),
        srgid: clean(xr['SRGID']),
        cci: clean(xr['CCI']),
        matchMethod: matchMethod.get(xi),
        column: col,
        csvValue: cv,
        xlsxValue: xv,
        concurred: isConcurred,
      });
    }
  }
  const umbrellaAutoCount = autoResolved.filter((a) => /Umbrella control/.test(a.reason)).length;
  const umbrellaConflictCount = conflicts.filter((c) => c.umbrella).length;
  const concurredConflictCount = conflicts.filter((c) => c.concurred).length;
  log.push(`Conflicts: ${conflicts.length} need review · Auto-resolved: ${autoResolved.length}`);
  log.push(`Umbrella (Satisfied-By) controls: ${umbrellaAutoCount} Check/Fix kept (set unchanged), ${umbrellaConflictCount} flagged (set changed)`);
  log.push(`Concurred rows: ${concurredMatched} settled · ${concurredConflictCount} Check/Fix/Status conflicts surfaced (export still differed)`);

  const iter = findLatestGovIteration(xlsxRows);
  const comments = [];
  if (iter) {
    const csvStigLookup = new Map();
    for (const r of csvRows) {
      const k = buildKey(r);
      if (!csvStigLookup.has(k)) csvStigLookup.set(k, clean(r['STIGID']));
    }
    xlsxRows.forEach((r, i) => {
      if (concurredXlsxIdx.has(i)) return;
      const gc = clean(r[iter.govCommentColumn]);
      if (!isMeaningfulComment(gc)) return;
      comments.push({
        id: `comment_r${r.__wsRow}`,
        xlsxRow: r.__wsRow,
        xlsxIndex: i,
        stigid: csvStigLookup.get(buildKey(r)) || '',
        srgid: clean(r['SRGID']),
        cci: clean(r['CCI']),
        iteration: iter.iteration,
        iterationLabel: iter.label,
        govComment: gc,
        existingVendorResponse: clean(r[iter.vendorResponseColumn] || ''),
      });
    });
    log.push(`Government comments needing response: ${comments.length} (iteration ${iter.label})`);
  } else {
    log.push('No government comments found');
  }

  return {
    conflicts,
    autoResolved,
    comments,
    newCsvIdx,
    unmatchedXlsxIdx: unmatchedXlsxIdxAll.filter((i) => !ambiguousXlsx.has(i)),
    ambiguousXlsxIdx: [...ambiguousXlsx],
    concurredXlsxIdx: [...concurredXlsxIdx],
    xlsxRowNums: xlsxRows.map((r) => r.__wsRow),
    pairings: Object.fromEntries(xlsxToCsv),
    matchMethod: Object.fromEntries(matchMethod),
    resolvedCellsMap: Object.fromEntries(resolvedCells),
    conflictCellSet,
    iteration: iter,
    log,
    metadata: {
      totalConflicts: conflicts.length,
      totalComments: comments.length,
      concurredCount: concurredMatched,
      autoResolvedCount: autoResolved.length,
      newRowsCount: newCsvIdx.length,
      unmatchedCount: unmatchedXlsxIdxAll.length - ambiguousXlsx.size,
      ambiguousCount: ambiguousXlsx.size,
      oneToOneCount: countOneOne,
      subMatchCount: countSubMatch,
      umbrellaAutoCount,
      umbrellaConflictCount,
      concurredConflictCount,
    },
  };
}

// ExcelJS cell.value can be string | number | Date | { richText: [...] } |
// { hyperlink, text } | { formula, result } | { error } — coerce to string.
export function cellText(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map(cellText).join('');
  if (typeof value === 'object') {
    if (value.richText) return value.richText.map((t) => t.text).join('');
    if (value.hyperlink) return value.text || '';
    if (value.formula) return value.result != null ? cellText(value.result) : '';
    if (value.error) return '';
  }
  return String(value);
}

export function readHeaders(ws) {
  const byName = new Map();
  const byCol = {};
  ws.getRow(1).eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const name = cellText(cell.value).trim();
    if (name) {
      byName.set(name, colNumber);
      byCol[colNumber] = name;
    }
  });
  return { byName, byCol };
}

// Reads worksheet data rows into plain objects keyed by header name — the same
// shape handleXlsxFile produces in the app, reused by the tests.
export function worksheetToRows(ws) {
  const headers = readHeaders(ws).byCol;
  const rows = [];
  const lastCol = Math.max(...Object.keys(headers).map(Number));
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const obj = {};
    let nonEmpty = false;
    for (let c = 1; c <= lastCol; c++) {
      const h = headers[c];
      if (!h) continue;
      const v = cellText(row.getCell(c).value);
      obj[h] = v;
      if (v !== '') nonEmpty = true;
    }
    // Track the true worksheet row: the baseline has ~796 blank rows between
    // the original table and a prior bottom-append, so compact index !== row.
    obj.__wsRow = r;
    if (nonEmpty) rows.push(obj);
  }
  return rows;
}
