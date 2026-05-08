import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import ExcelJS from 'exceljs/dist/exceljs.min.js';
import Papa from 'papaparse';
import {
  Download, Upload, List, RotateCcw, FileText, AlertCircle, AlertTriangle,
  Keyboard, ArrowLeft, ArrowRight, MessageSquare, Save, FileSpreadsheet,
  FileCode, Play, Loader, X, Check,
} from 'lucide-react';

const SERIF = 'ui-serif, Georgia, "Times New Roman", serif';
const MONO = 'ui-monospace, "SF Mono", Menlo, Consolas, monospace';
const COLORS = {
  bg: '#f4ede0',
  paper: '#fbf7ed',
  ink: '#1a1a1a',
  inkSoft: '#5a5247',
  inkFaint: '#8a8275',
  rule: '#d8cdb6',
  ruleSoft: '#ebe2cf',
  xlsx: '#1f4e3d',
  xlsxSoft: '#d8e6dd',
  csv: '#7a3a1f',
  csvSoft: '#efddc8',
  delBg: '#f4c5c5',
  delFg: '#6b1c1c',
  addBg: '#c8e6c9',
  addFg: '#1f4f23',
  warn: '#b8742d',
  warnBg: '#fae0c5',
};

const SHARED_COLS = [
  'IA Control', 'CCI', 'SRGID', 'STIGID', 'SRG Requirement', 'Requirement',
  'SRG VulDiscussion', 'VulDiscussion', 'Status', 'SRG Check', 'Check',
  'SRG Fix', 'Fix', 'Severity', 'Mitigation', 'Artifact Description',
  'Status Justification',
];
const CSV_ONLY_COLS = ['Vendor Comments', 'InSpec Control Body'];

const SIDE = { XLSX: 'xlsx', CSV: 'csv' };
const SOURCE = { XLSX: 'XLSX', CSV: 'CSV' };

const PILL_BASE = {
  padding: '1px 7px',
  borderRadius: 2,
  fontSize: 10,
  letterSpacing: '0.06em',
  textTransform: 'uppercase',
  fontWeight: 600,
  justifySelf: 'start',
};

const DEC_PILL = {
  [SIDE.XLSX]: { background: COLORS.xlsxSoft, color: COLORS.xlsx, label: 'XLSX kept' },
  [SIDE.CSV]:  { background: COLORS.csvSoft, color: COLORS.csv,  label: 'CSV used' },
};

const SOURCE_PALETTE = {
  [SOURCE.XLSX]: { fg: COLORS.xlsx, bg: COLORS.xlsxSoft, label: 'XLSX kept' },
  [SOURCE.CSV]:  { fg: COLORS.csv,  bg: COLORS.csvSoft,  label: 'CSV used' },
};

const STATUS_PILL = {
  done:    { background: COLORS.addBg, color: COLORS.addFg },
  partial: { background: COLORS.csvSoft, color: COLORS.csv },
  open:    { background: COLORS.ruleSoft, color: COLORS.inkFaint },
};

function SummaryLine({ column, pill, value, valueNode, valueTitle }) {
  return (
    <div style={{ display: 'grid', gridTemplateColumns: '160px 70px 1fr', gap: 12, padding: '4px 0', alignItems: 'baseline' }}>
      <span style={{ color: COLORS.inkSoft }}>{column}</span>
      <span style={{ ...PILL_BASE, background: pill.background, color: pill.color }}>{pill.label}</span>
      <span style={{ color: COLORS.ink, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={valueTitle}>
        {valueNode || value}
      </span>
    </div>
  );
}

const cellKey = (xi, col) => `${xi}|${col}`;
const conflictId = (xi, col) => `r${xi + 2}_${col.replace(/\s+/g, '_')}`;
const xlsxIndexToRow = (xi) => xi + 2;

function groupByKey(rows) {
  const m = new Map();
  rows.forEach((r, i) => {
    const k = buildKey(r);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(i);
  });
  return m;
}

function bucketByXlsxRow(byRow, row, seed, defaults) {
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

function clean(s) {
  if (s === null || s === undefined) return '';
  const str = String(s).trim();
  const lower = str.toLowerCase();
  if (str === '' || lower === 'nan' || lower === 'none') return '';
  return str;
}

function norm(s) {
  return clean(s).toLowerCase().replace(/\s+/g, ' ');
}

function buildKey(row) {
  return clean(row['SRGID']) + '|||' + clean(row['CCI']);
}

function extractSignature(row) {
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

function jaccard(a, b) {
  if (a.size === 0 && b.size === 0) return 0;
  let inter = 0;
  for (const x of a) if (b.has(x)) inter++;
  const union = a.size + b.size - inter;
  return union > 0 ? inter / union : 0;
}

function isChainguardPreferred(val) {
  return /\bchainguard\s*os\b/i.test(val);
}
function isGenericPhrasing(val) {
  return /\boperating\s+system\b/i.test(val) && !isChainguardPreferred(val);
}

function isMeaningfulComment(text) {
  const t = clean(text);
  if (!t) return false;
  if (/^concur\b[\s.,;:]*(with\s+status[\s.,;:]*)?$/i.test(t)) return false;
  return true;
}

function ordinal(n) {
  if (n === 1) return '1st';
  if (n === 2) return '2nd';
  if (n === 3) return '3rd';
  return n + 'th';
}

// Word-level LCS diff
function diffWords(a, b) {
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

function findLatestGovIteration(xlsxRows) {
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

function runMerge(csvRows, xlsxRows) {
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
      const csvSigs = new Map(csvIdx.map((i) => [i, extractSignature(csvRows[i])]));
      const xlsxSigs = new Map(xlsxIdx.map((i) => [i, extractSignature(xlsxRows[i])]));
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
      const usedX = new Set();
      const usedC = new Set();
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
        if (xlsxSigs.get(xi).size === 0) ambiguousXlsx.add(xi);
      }
    }
  }
  log.push(`Matched: ${xlsxToCsv.size} (${countOneOne} one-to-one, ${countSubMatch} sub-match)`);

  const newCsvIdx = [];
  for (let i = 0; i < csvRows.length; i++) if (!usedCsv.has(i)) newCsvIdx.push(i);
  const unmatchedXlsxIdxAll = [];
  for (let i = 0; i < xlsxRows.length; i++) if (!xlsxToCsv.has(i)) unmatchedXlsxIdxAll.push(i);
  log.push(`New CSV rows: ${newCsvIdx.length} · Unmatched XLSX: ${unmatchedXlsxIdxAll.length} (${ambiguousXlsx.size} ambiguous)`);

  const conflicts = [];
  const autoResolved = [];
  const resolvedCells = new Map();
  const conflictCellSet = new Set();

  for (const [xi, ci] of xlsxToCsv) {
    const xr = xlsxRows[xi];
    const cr = csvRows[ci];
    for (const col of SHARED_COLS) {
      const cv = clean(cr[col]);
      const xv = clean(xr[col]);
      if (cv === '' || xv === '') continue;
      if (norm(cv) === norm(xv)) continue;

      // Auto-resolve: Requirement column "Chainguard OS" preference
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
          resolvedCells.set(cellKey(xi, col), chosen);
          autoResolved.push({
            xlsxRow: xlsxIndexToRow(xi),
            stigid: clean(cr['STIGID']),
            srgid: clean(xr['SRGID']),
            cci: clean(xr['CCI']),
            column: col,
            chosenSource: source,
            reason,
            resolvedValue: chosen,
            csvValue: cv,
            xlsxValue: xv,
          });
          continue;
        }
      }

      conflictCellSet.add(cellKey(xi, col));
      conflicts.push({
        id: conflictId(xi, col),
        xlsxRow: xlsxIndexToRow(xi),
        xlsxIndex: xi,
        csvIndex: ci,
        stigid: clean(cr['STIGID']),
        srgid: clean(xr['SRGID']),
        cci: clean(xr['CCI']),
        matchMethod: matchMethod.get(xi),
        column: col,
        csvValue: cv,
        xlsxValue: xv,
      });
    }
  }
  log.push(`Conflicts: ${conflicts.length} need review · Auto-resolved: ${autoResolved.length}`);

  const iter = findLatestGovIteration(xlsxRows);
  const comments = [];
  if (iter) {
    const csvStigLookup = new Map();
    for (const r of csvRows) {
      const k = buildKey(r);
      if (!csvStigLookup.has(k)) csvStigLookup.set(k, clean(r['STIGID']));
    }
    xlsxRows.forEach((r, i) => {
      const gc = clean(r[iter.govCommentColumn]);
      if (!isMeaningfulComment(gc)) return;
      comments.push({
        id: `comment_r${xlsxIndexToRow(i)}`,
        xlsxRow: xlsxIndexToRow(i),
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
    pairings: Object.fromEntries(xlsxToCsv),
    matchMethod: Object.fromEntries(matchMethod),
    resolvedCellsMap: Object.fromEntries(resolvedCells),
    conflictCellSet,
    iteration: iter,
    log,
    metadata: {
      totalConflicts: conflicts.length,
      totalComments: comments.length,
      autoResolvedCount: autoResolved.length,
      newRowsCount: newCsvIdx.length,
      unmatchedCount: unmatchedXlsxIdxAll.length - ambiguousXlsx.size,
      ambiguousCount: ambiguousXlsx.size,
      oneToOneCount: countOneOne,
      subMatchCount: countSubMatch,
    },
  };
}

// ExcelJS cell.value can be string | number | Date | { richText: [...] } |
// { hyperlink, text } | { formula, result } | { error } — coerce to string.
function cellText(value) {
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

function readHeaders(ws) {
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

async function applyAndExport({ workbook, csvRows, mergeResult, decisions, responses, originalFileName }) {
  const ws = workbook.getWorksheet('DataTable') || workbook.worksheets[0];
  const colMap = readHeaders(ws).byName;

  let nextCol = ws.columnCount + 1;
  for (const colName of CSV_ONLY_COLS) {
    if (!colMap.has(colName)) {
      colMap.set(colName, nextCol);
      ws.getRow(1).getCell(nextCol).value = colName;
      ws.getColumn(nextCol).width = 40;
      nextCol++;
    }
  }

  for (const [xiStr, ci] of Object.entries(mergeResult.pairings)) {
    const xi = parseInt(xiStr, 10);
    const cr = csvRows[ci];
    const row = ws.getRow(xlsxIndexToRow(xi));

    if (colMap.has('STIGID')) {
      row.getCell(colMap.get('STIGID')).value = clean(cr['STIGID']);
    }

    for (const col of SHARED_COLS) {
      if (col === 'STIGID') continue;
      if (!colMap.has(col)) continue;
      const cIdx = colMap.get(col);
      const csvVal = clean(cr[col]);
      const cell = row.getCell(cIdx);
      const existingVal = clean(cellText(cell.value));

      const key = cellKey(xi, col);

      if (mergeResult.resolvedCellsMap[key] !== undefined) {
        cell.value = mergeResult.resolvedCellsMap[key];
        continue;
      }

      if (mergeResult.conflictCellSet.has(key)) {
        if (decisions[conflictId(xi, col)] === SIDE.CSV) {
          cell.value = csvVal;
        }
        continue;
      }

      if (existingVal === '' && csvVal !== '') {
        cell.value = csvVal;
      }
    }

    for (const col of CSV_ONLY_COLS) {
      if (!colMap.has(col)) continue;
      const csvVal = clean(cr[col]);
      if (csvVal) row.getCell(colMap.get(col)).value = csvVal;
    }
  }

  if (mergeResult.iteration && mergeResult.iteration.vendorResponseColumn) {
    const venCol = colMap.get(mergeResult.iteration.vendorResponseColumn);
    if (venCol !== undefined) {
      for (const cmt of mergeResult.comments) {
        const resp = (responses[cmt.id] || '').trim();
        if (resp) {
          ws.getRow(xlsxIndexToRow(cmt.xlsxIndex)).getCell(venCol).value = resp;
        }
      }
    }
  }

  let newRowR = ws.rowCount + 1;
  for (const ci of mergeResult.newCsvIdx) {
    const cr = csvRows[ci];
    const row = ws.getRow(newRowR);
    for (const [colName, cIdx] of colMap) {
      const val = clean(cr[colName]);
      if (val) row.getCell(cIdx).value = val;
    }
    newRowR++;
  }

  const buf = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = (originalFileName || 'merged').replace(/\.xlsx$/i, '') + '_merged.xlsx';
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 200);
}

export default function App() {
  const [stage, setStage] = useState('import');
  const [view, setView] = useState('review');

  const [csvFile, setCsvFile] = useState(null);
  const [xlsxFile, setXlsxFile] = useState(null);
  const [csvRows, setCsvRows] = useState(null);
  const [xlsxRows, setXlsxRows] = useState(null);
  const [workbook, setWorkbook] = useState(null);
  const [mergeResult, setMergeResult] = useState(null);
  const [importError, setImportError] = useState(null);
  const [mergeError, setMergeError] = useState(null);

  // Review state
  const [decisions, setDecisions] = useState({});
  const [responses, setResponses] = useState({});
  const [currentIndex, setCurrentIndex] = useState(0);
  const [storageReady, setStorageReady] = useState(false);
  const [exportNotice, setExportNotice] = useState(null);

  const csvInputRef = useRef(null);
  const xlsxInputRef = useRef(null);

  const conflicts = mergeResult?.conflicts || [];
  const comments = mergeResult?.comments || [];

  const reviewItems = useMemo(() => {
    if (!mergeResult) return [];
    const byRow = new Map();
    for (const c of mergeResult.conflicts) {
      const it = bucketByXlsxRow(byRow, c.xlsxRow, c, {
        key: `row_${c.xlsxRow}`,
        matchMethod: c.matchMethod || null,
        conflicts: [],
        comment: null,
      });
      it.conflicts.push(c);
      if (!it.matchMethod && c.matchMethod) it.matchMethod = c.matchMethod;
    }
    for (const cm of mergeResult.comments) {
      const it = bucketByXlsxRow(byRow, cm.xlsxRow, cm, {
        key: `row_${cm.xlsxRow}`,
        matchMethod: null,
        conflicts: [],
        comment: null,
      });
      it.comment = cm;
    }
    return [...byRow.values()].sort((a, b) => a.xlsxRow - b.xlsxRow);
  }, [mergeResult]);

  const datasetKey = useMemo(() => {
    if (!csvFile || !xlsxFile) return null;
    return `merge:${csvFile.name}:${csvFile.size}|${xlsxFile.name}:${xlsxFile.size}`;
  }, [csvFile, xlsxFile]);

  // ── persistence
  useEffect(() => {
    if (!datasetKey || stage !== 'review') {
      setStorageReady(true);
      return;
    }
    let cancelled = false;
    (async () => {
      try {
        const r = await window.storage.get(datasetKey);
        if (!cancelled && r && r.value) {
          const d = JSON.parse(r.value);
          setDecisions(d.decisions || {});
          setResponses(d.responses || {});
          setCurrentIndex(Math.min(d.currentIndex || 0, Math.max(reviewItems.length - 1, 0)));
        }
      } catch (e) { /* fresh */ }
      if (!cancelled) setStorageReady(true);
    })();
    return () => { cancelled = true; };
  }, [datasetKey, stage, reviewItems.length]);

  useEffect(() => {
    if (!storageReady || !datasetKey) return;
    const t = setTimeout(() => {
      try {
        window.storage.set(datasetKey, JSON.stringify({
          decisions, responses, currentIndex,
        }), false).catch(() => {});
      } catch (e) {/*ignore*/}
    }, 200);
    return () => clearTimeout(t);
  }, [decisions, responses, currentIndex, datasetKey, storageReady]);

  // ── Import handlers
  const handleCsvFile = (file) => {
    if (!file) return;
    setImportError(null);
    setCsvFile(file);
    Papa.parse(file, {
      header: true,
      skipEmptyLines: false,
      complete: (results) => {
        if (results.errors && results.errors.length > 0) {
          // Only warn for serious errors
          const fatal = results.errors.find((e) => e.type !== 'FieldMismatch');
          if (fatal) {
            setImportError(`CSV parse error: ${fatal.message}`);
            setCsvFile(null);
            return;
          }
        }
        setCsvRows(results.data);
      },
      error: (err) => {
        setImportError(`CSV read error: ${err.message}`);
        setCsvFile(null);
      },
    });
  };

  const handleXlsxFile = async (file) => {
    if (!file) return;
    setImportError(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const ws = wb.getWorksheet('DataTable') || wb.worksheets[0];

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
        if (nonEmpty) rows.push(obj);
      }

      setXlsxFile(file);
      setWorkbook(wb);
      setXlsxRows(rows);
    } catch (e) {
      console.error('XLSX read error', e);
      setImportError(`XLSX read error: ${e.message}${e.stack ? ' — ' + e.stack.split('\n')[1] : ''}`);
      setXlsxFile(null);
    }
  };

  const runTheMerge = () => {
    setMergeError(null);
    setStage('merging');
    setTimeout(() => {
      try {
        const result = runMerge(csvRows, xlsxRows);
        setMergeResult(result);
        setDecisions({});
        setResponses({});
        setCurrentIndex(0);
        setStage('review');
        setView('review');
      } catch (e) {
        setMergeError(`Merge failed: ${e.message}`);
        setStage('import');
      }
    }, 50); // allow UI to repaint
  };

  const startOver = () => {
    setStage('import');
    setCsvFile(null);
    setXlsxFile(null);
    setCsvRows(null);
    setXlsxRows(null);
    setWorkbook(null);
    setMergeResult(null);
    setDecisions({});
    setResponses({});
    setCurrentIndex(0);
    setView('review');
  };

  // ── Review actions (unified — operate on reviewItems)
  const currentItem = reviewItems[currentIndex] || null;
  const rowConflicts = currentItem?.conflicts || [];
  const currentRowComment = currentItem?.comment || null;
  const rowDecidedCount = rowConflicts.filter((c) => decisions[c.id]).length;

  const decidedCount = Object.keys(decisions).length;
  const xlsxKept = Object.values(decisions).filter((v) => v === SIDE.XLSX).length;
  const csvChosen = Object.values(decisions).filter((v) => v === SIDE.CSV).length;
  const remaining = conflicts.length - decidedCount;

  const respondedCount = Object.values(responses).filter((v) => (v || '').trim()).length;
  const commentRemaining = comments.length - respondedCount;

  const advance = useCallback(
    () => setCurrentIndex((i) => Math.min(i + 1, Math.max(reviewItems.length - 1, 0))),
    [reviewItems.length]
  );
  const goBack = useCallback(() => setCurrentIndex((i) => Math.max(i - 1, 0)), []);

  const choose = useCallback((id, side) => {
    setDecisions((d) => ({ ...d, [id]: side }));
    // Auto-advance: find the next undecided conflict on this row, scroll to it.
    // If none remain, jump to the next row.
    const remaining = rowConflicts.filter((c) => c.id !== id && !decisions[c.id]);
    if (remaining.length > 0) {
      setTimeout(() => {
        const el = document.getElementById(`conflict-${remaining[0].id}`);
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }, 60);
    } else {
      advance();
    }
  }, [decisions, rowConflicts, advance]);

  const clearDecision = useCallback((id) => {
    setDecisions((d) => { const n = { ...d }; delete n[id]; return n; });
  }, []);

  const clearRowDecisions = useCallback(() => {
    if (!rowConflicts.length) return;
    setDecisions((d) => {
      const n = { ...d };
      for (const c of rowConflicts) delete n[c.id];
      return n;
    });
  }, [rowConflicts]);

  const setResponseFor = useCallback((commentId, text) => {
    setResponses((r) => ({ ...r, [commentId]: text }));
  }, []);

  const clearResponse = useCallback(() => {
    if (!currentRowComment) return;
    setResponses((r) => { const n = { ...r }; delete n[currentRowComment.id]; return n; });
  }, [currentRowComment]);

  // Jump to next row that still has an undecided conflict OR an unanswered comment
  const jumpToNextOpen = useCallback(() => {
    const idx = reviewItems.findIndex((it) => {
      if (it.conflicts.some((c) => !decisions[c.id])) return true;
      if (it.comment && !(responses[it.comment.id] || '').trim()) return true;
      return false;
    });
    if (idx >= 0) { setCurrentIndex(idx); setView('review'); }
  }, [reviewItems, decisions, responses]);

  // ── Scroll to top when the active row changes (so advancing rows starts you at the top)
  useEffect(() => {
    if (stage === 'review' && view === 'review') {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  }, [currentIndex, stage, view]);

  // ── Keyboard
  useEffect(() => {
    if (stage !== 'review') return;
    if (view !== 'review') return;
    const onKey = (e) => {
      const tag = (e.target.tagName || '').toUpperCase();
      if (tag === 'TEXTAREA' || tag === 'INPUT') {
        // Allow ⌘/Ctrl+Enter to advance from inside the response textarea
        if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') {
          e.preventDefault();
          advance();
        }
        return;
      }
      if (e.key === 'ArrowDown' || e.key === 's') { e.preventDefault(); advance(); }
      else if (e.key === 'ArrowUp' || e.key === 'w') { e.preventDefault(); goBack(); }
      else if (e.key === '?') { e.preventDefault(); setView('help'); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [stage, view, advance, goBack]);

  const handleExport = async () => {
    try {
      await applyAndExport({
        workbook,
        csvRows,
        mergeResult,
        decisions,
        responses,
        originalFileName: xlsxFile.name,
      });
      setExportNotice('✓ Exported with cell formatting preserved.');
      setView('summary');
    } catch (e) {
      setExportNotice(`× Export failed: ${e.message}`);
    }
  };

  const summary = useMemo(() => {
    if (!mergeResult) return null;
    const rowMap = new Map();
    const seedDefaults = { decided: [], skipped: [], autoResolved: [], response: null };
    for (const c of mergeResult.conflicts) {
      const dec = decisions[c.id];
      const e = bucketByXlsxRow(rowMap, c.xlsxRow, c, seedDefaults);
      if (dec === SIDE.CSV) e.decided.push({ column: c.column, side: SOURCE.CSV, value: c.csvValue, replaced: c.xlsxValue });
      else if (dec === SIDE.XLSX) e.decided.push({ column: c.column, side: SOURCE.XLSX, value: c.xlsxValue, replaced: c.csvValue });
      else e.skipped.push({ column: c.column, value: c.xlsxValue });
    }
    for (const ar of mergeResult.autoResolved || []) {
      const e = bucketByXlsxRow(rowMap, ar.xlsxRow, ar, seedDefaults);
      e.autoResolved.push({ column: ar.column, source: ar.chosenSource, value: ar.resolvedValue, reason: ar.reason });
    }
    for (const cm of mergeResult.comments) {
      const resp = (responses[cm.id] || '').trim();
      if (!resp) continue;
      const e = bucketByXlsxRow(rowMap, cm.xlsxRow, cm, seedDefaults);
      e.response = { iterationLabel: cm.iterationLabel, govComment: cm.govComment, text: resp };
    }
    const rows = [...rowMap.values()].sort((a, b) => a.xlsxRow - b.xlsxRow);

    const newRows = (mergeResult.newCsvIdx || []).map((ci) => {
      const r = csvRows[ci];
      return { stigid: clean(r['STIGID']), srgid: clean(r['SRGID']), cci: clean(r['CCI']) };
    });

    let totalCsvChosen = 0;
    let totalXlsxChosen = 0;
    for (const r of rows) {
      for (const d of r.decided) {
        if (d.side === SOURCE.CSV) totalCsvChosen++;
        else if (d.side === SOURCE.XLSX) totalXlsxChosen++;
      }
    }
    const totalSkipped = rows.reduce((n, r) => n + r.skipped.length, 0);
    const totalResponses = rows.filter((r) => r.response).length;
    const totalAutoResolved = mergeResult.autoResolved?.length || 0;

    return {
      rows,
      newRows,
      stats: { totalCsvChosen, totalXlsxChosen, totalSkipped, totalResponses, totalAutoResolved, totalNewRows: newRows.length },
    };
  }, [mergeResult, decisions, responses, csvRows]);

  // ── Style helpers
  const iconBtn = (active = false) => ({
    background: active ? COLORS.ink : 'transparent',
    border: '1px solid ' + (active ? COLORS.ink : COLORS.rule),
    color: active ? COLORS.paper : COLORS.inkSoft,
    padding: '6px 10px',
    cursor: 'pointer',
    fontFamily: MONO,
    fontSize: 11,
    letterSpacing: '0.05em',
    textTransform: 'uppercase',
    display: 'inline-flex',
    gap: 6,
    alignItems: 'center',
    borderRadius: 0,
    transition: 'all 0.15s',
  });

  const renderDiffSide = (op, side) => {
    if (op.type === 'eq') return <span style={{ color: COLORS.ink }}>{op.text}</span>;
    if (side === SIDE.XLSX && op.type === 'del')
      return <span style={{ background: COLORS.delBg, color: COLORS.delFg, padding: '1px 2px', borderRadius: 2, fontWeight: 500 }}>{op.text}</span>;
    if (side === SIDE.CSV && op.type === 'add')
      return <span style={{ background: COLORS.addBg, color: COLORS.addFg, padding: '1px 2px', borderRadius: 2, fontWeight: 500 }}>{op.text}</span>;
    return null;
  };

  return (
    <div style={{ background: COLORS.bg, minHeight: '100vh', fontFamily: SERIF, color: COLORS.ink, lineHeight: 1.5 }}>
      <div style={{ maxWidth: 1280, margin: '0 auto', padding: '32px 28px 80px' }}>

        {/* MASTHEAD */}
        <div style={{ borderTop: '1px solid ' + COLORS.ink, borderBottom: '1px solid ' + COLORS.rule, padding: '14px 0', display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: 24, flexWrap: 'wrap' }}>
          <div style={{ fontWeight: 600, fontSize: 22, letterSpacing: '-0.01em' }}>
            The <em style={{ fontWeight: 400, fontStyle: 'italic', color: COLORS.inkSoft }}>Merge</em> Review
          </div>
          <div style={{ fontFamily: MONO, fontSize: 11, letterSpacing: '0.05em', textTransform: 'uppercase', color: COLORS.inkFaint, display: 'flex', gap: 14, alignItems: 'center', flexWrap: 'wrap' }}>
            {stage === 'import' && <span>step 1 · import files</span>}
            {stage === 'merging' && <span>merging…</span>}
            {stage === 'review' && (
              <>
                <span>{xlsxFile?.name}</span>
                <span style={{ opacity: 0.4 }}>·</span>
                <span>{conflicts.length} conflicts · {comments.length} comments · {reviewItems.length} rows</span>
                <span style={{ opacity: 0.4 }}>·</span>
                <div style={{ display: 'flex', gap: 4 }}>
                  <button style={iconBtn(view === 'review')} onClick={() => setView('review')}>
                    <FileText size={12} /> Review
                  </button>
                  <button style={iconBtn(view === 'list')} onClick={() => setView('list')}>
                    <List size={12} /> List
                  </button>
                  <button style={iconBtn(view === 'summary')} onClick={() => setView('summary')} title="View change summary">
                    <FileText size={12} /> Summary
                  </button>
                  <button
                    style={iconBtn(false)}
                    onClick={handleExport}
                    title="Export merged XLSX (formatting preserved)"
                  >
                    <Download size={12} />
                    Export XLSX
                  </button>
                  <button style={iconBtn(view === 'help')} onClick={() => setView('help')}>
                    <Keyboard size={12} /> ?
                  </button>
                  <button style={iconBtn(false)} onClick={startOver} title="Start over with new files">
                    <RotateCcw size={12} /> New
                  </button>
                </div>
              </>
            )}
          </div>
        </div>

        {/* PROGRESS — row-based: conflicts + comments */}
        {stage === 'review' && view === 'review' && reviewItems.length > 0 && (
          <div style={{ marginTop: 10, display: 'flex', gap: 16, alignItems: 'center', flexWrap: 'wrap', fontFamily: MONO, fontSize: 11, letterSpacing: '0.04em' }}>
            <span style={{ color: COLORS.inkSoft }}>
              row <strong style={{ color: COLORS.ink }}>{currentIndex + 1}</strong>/{reviewItems.length}
            </span>
            <div style={{ flex: '1 1 200px', minWidth: 100, height: 4, background: COLORS.ruleSoft, position: 'relative', overflow: 'hidden' }}>
              <div style={{ position: 'absolute', top: 0, bottom: 0, width: 2, background: COLORS.csv, left: ((currentIndex / Math.max(reviewItems.length, 1)) * 100) + '%' }} />
            </div>
            {rowConflicts.length > 0 && (
              <span style={{ color: COLORS.inkSoft }}>
                this row <strong style={{ color: COLORS.ink }}>{rowDecidedCount}</strong>/{rowConflicts.length}
              </span>
            )}
            {conflicts.length > 0 && (
              <span style={{ color: COLORS.inkSoft }}>
                total conflicts <strong style={{ color: COLORS.ink }}>{decidedCount}</strong>/{conflicts.length}
                {' '}<span style={{ opacity: 0.6 }}>({xlsxKept}× XLSX, {csvChosen}× CSV)</span>
              </span>
            )}
            {comments.length > 0 && (
              <span style={{ color: COLORS.inkSoft }}>
                responses <strong style={{ color: COLORS.ink }}>{respondedCount}</strong>/{comments.length}
              </span>
            )}
            {(remaining > 0 || commentRemaining > 0) && (
              <span onClick={jumpToNextOpen} style={{ cursor: 'pointer', textDecoration: 'underline', color: COLORS.csv }}>jump to next open →</span>
            )}
          </div>
        )}

        {exportNotice && (
          <div style={{ marginTop: 12, padding: '10px 14px', background: COLORS.warnBg, borderLeft: '3px solid ' + COLORS.warn, fontFamily: MONO, fontSize: 12, color: '#5e3a16', display: 'flex', justifyContent: 'space-between', gap: 12 }}>
            <span>{exportNotice}</span>
            <span onClick={() => setExportNotice(null)} style={{ cursor: 'pointer', opacity: 0.7 }}><X size={12} /></span>
          </div>
        )}

        {stage === 'import' && (
          <div style={{ marginTop: 36 }}>
            <div style={{ marginBottom: 28 }}>
              <div style={{ fontFamily: MONO, fontSize: 11, letterSpacing: '0.12em', textTransform: 'uppercase', color: COLORS.inkFaint }}>Step 1 of 3</div>
              <h1 style={{ fontSize: 36, fontWeight: 600, letterSpacing: '-0.015em', margin: '6px 0 4px', lineHeight: 1.05 }}>
                Bring in <em style={{ fontStyle: 'italic', fontWeight: 400, color: COLORS.inkSoft }}>both</em> files
              </h1>
              <p style={{ fontSize: 15, color: COLORS.inkSoft, maxWidth: 640 }}>
                The CSV is your fresh export from MITRE Vulcan. The XLSX is your working spreadsheet —
                its cell formatting (color coding, fonts, borders) will be preserved through the merge.
              </p>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
              {/* CSV upload */}
              <div
                onClick={() => csvInputRef.current?.click()}
                onDragOver={(e) => { e.preventDefault(); }}
                onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) handleCsvFile(f); }}
                style={{
                  border: '2px dashed ' + (csvFile ? COLORS.xlsx : COLORS.rule),
                  background: csvFile ? COLORS.xlsxSoft : COLORS.paper,
                  padding: '32px 24px',
                  cursor: 'pointer',
                  transition: 'all 0.15s',
                  minHeight: 200,
                  display: 'flex',
                  flexDirection: 'column',
                  justifyContent: 'center',
                  alignItems: 'center',
                  textAlign: 'center',
                  gap: 12,
                }}
              >
                <FileCode size={32} color={csvFile ? COLORS.xlsx : COLORS.inkFaint} />
                <div style={{ fontFamily: MONO, fontSize: 11, letterSpacing: '0.1em', textTransform: 'uppercase', color: COLORS.inkFaint }}>CSV from Vulcan</div>
                {csvFile ? (
                  <>
                    <div style={{ fontSize: 17, fontWeight: 500 }}>{csvFile.name}</div>
                    <div style={{ fontFamily: MONO, fontSize: 11, color: COLORS.inkSoft }}>
                      {csvRows ? `${csvRows.length} rows parsed` : 'parsing…'}
                    </div>
                  </>
                ) : (
                  <>
                    <div style={{ fontSize: 17, fontStyle: 'italic', color: COLORS.inkSoft }}>Drop or click to choose</div>
                    <div style={{ fontFamily: MONO, fontSize: 10, color: COLORS.inkFaint, letterSpacing: '0.06em' }}>.csv</div>
                  </>
                )}
                <input ref={csvInputRef} type="file" accept=".csv" style={{ display: 'none' }} onChange={(e) => handleCsvFile(e.target.files?.[0])} />
              </div>

              {/* XLSX upload */}
              <div
                onClick={() => xlsxInputRef.current?.click()}
                onDragOver={(e) => { e.preventDefault(); }}
                onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) handleXlsxFile(f); }}
                style={{
                  border: '2px dashed ' + (xlsxFile ? COLORS.csv : COLORS.rule),
                  background: xlsxFile ? COLORS.csvSoft : COLORS.paper,
                  padding: '32px 24px',
                  cursor: 'pointer',
                  transition: 'all 0.15s',
                  minHeight: 200,
                  display: 'flex',
                  flexDirection: 'column',
                  justifyContent: 'center',
                  alignItems: 'center',
                  textAlign: 'center',
                  gap: 12,
                }}
              >
                <FileSpreadsheet size={32} color={xlsxFile ? COLORS.csv : COLORS.inkFaint} />
                <div style={{ fontFamily: MONO, fontSize: 11, letterSpacing: '0.1em', textTransform: 'uppercase', color: COLORS.inkFaint }}>Working XLSX</div>
                {xlsxFile ? (
                  <>
                    <div style={{ fontSize: 17, fontWeight: 500 }}>{xlsxFile.name}</div>
                    <div style={{ fontFamily: MONO, fontSize: 11, color: COLORS.inkSoft }}>
                      {xlsxRows ? `${xlsxRows.length} rows · sheet "${workbook?.worksheets?.[0]?.name || ''}"` : 'reading…'}
                    </div>
                  </>
                ) : (
                  <>
                    <div style={{ fontSize: 17, fontStyle: 'italic', color: COLORS.inkSoft }}>Drop or click to choose</div>
                    <div style={{ fontFamily: MONO, fontSize: 10, color: COLORS.inkFaint, letterSpacing: '0.06em' }}>.xlsx</div>
                  </>
                )}
                <input ref={xlsxInputRef} type="file" accept=".xlsx,.xlsm,.xlsb" style={{ display: 'none' }} onChange={(e) => handleXlsxFile(e.target.files?.[0])} />
              </div>
            </div>

            {importError && (
              <div style={{ marginTop: 16, padding: '12px 16px', background: COLORS.warnBg, borderLeft: '3px solid ' + COLORS.warn, fontFamily: MONO, fontSize: 13, color: '#5e3a16' }}>
                <AlertCircle size={14} style={{ display: 'inline', marginRight: 6, verticalAlign: '-2px' }} />
                {importError}
              </div>
            )}

            <div style={{ marginTop: 28, display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 16, flexWrap: 'wrap' }}>
              <div style={{ fontFamily: MONO, fontSize: 12, color: COLORS.inkFaint }}>
                {csvRows && xlsxRows ? <>Both files loaded · CSV {csvRows.length} rows · XLSX {xlsxRows.length} rows</> : <>Load both files to continue</>}
              </div>
              <button
                onClick={runTheMerge}
                disabled={!(csvRows && xlsxRows)}
                style={{
                  background: csvRows && xlsxRows ? COLORS.ink : COLORS.ruleSoft,
                  color: csvRows && xlsxRows ? COLORS.paper : COLORS.inkFaint,
                  border: 0,
                  padding: '14px 28px',
                  fontFamily: SERIF,
                  fontSize: 17,
                  cursor: csvRows && xlsxRows ? 'pointer' : 'not-allowed',
                  display: 'flex',
                  alignItems: 'center',
                  gap: 12,
                }}
              >
                <Play size={18} />
                Run merge
              </button>
            </div>

            {/* Format preservation note */}
            <div style={{ marginTop: 36, padding: '16px 20px', background: COLORS.paper, border: '1px solid ' + COLORS.rule, fontSize: 13, color: COLORS.inkSoft, lineHeight: 1.6 }}>
              <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10 }}>
                <Check size={16} color={COLORS.xlsx} style={{ marginTop: 3, flexShrink: 0 }} />
                <div>
                  <strong style={{ color: COLORS.ink }}>Cell formatting will be preserved.</strong>{' '}
                  The exported XLSX keeps your original green / yellow review markers, fonts, borders,
                  and column widths.
                </div>
              </div>
            </div>
          </div>
        )}

        {stage === 'merging' && (
          <div style={{ marginTop: 80, textAlign: 'center' }}>
            <Loader size={32} color={COLORS.inkSoft} style={{ animation: 'spin 1s linear infinite' }} />
            <h2 style={{ fontSize: 28, fontWeight: 500, fontStyle: 'italic', margin: '16px 0 4px', color: COLORS.inkSoft }}>Merging…</h2>
            <p style={{ color: COLORS.inkFaint, fontSize: 14 }}>Categorizing groups, computing similarity scores, detecting conflicts.</p>
            <style>{`@keyframes spin { from {transform:rotate(0)} to {transform:rotate(360deg)} }`}</style>
          </div>
        )}

        {mergeError && (
          <div style={{ marginTop: 12, padding: '10px 14px', background: COLORS.warnBg, borderLeft: '3px solid ' + COLORS.warn, fontFamily: MONO, fontSize: 12, color: '#5e3a16' }}>
            <AlertCircle size={14} style={{ display: 'inline', marginRight: 6, verticalAlign: '-2px' }} />
            {mergeError}
          </div>
        )}

        {stage === 'review' && view === 'help' && (
          <div style={{ marginTop: 36, padding: '28px 32px', background: COLORS.paper, border: '1px solid ' + COLORS.rule, maxWidth: 760 }}>
            <h2 style={{ fontSize: 24, fontWeight: 600, margin: '0 0 16px', letterSpacing: '-0.01em' }}>How this works</h2>
            <p style={{ color: COLORS.inkSoft, fontSize: 15, marginTop: 0 }}>
              Differences between XLSX and CSV are highlighted word-by-word. Pick a side or type a response;
              the app saves your decisions, and on Export it writes a merged XLSX with your choices applied.
            </p>

            <h3 style={{ fontSize: 13, fontWeight: 600, margin: '20px 0 10px', fontFamily: MONO, textTransform: 'uppercase', letterSpacing: '0.06em', color: COLORS.inkSoft }}>Merge summary</h3>
            <div style={{ fontFamily: MONO, fontSize: 12, color: COLORS.inkSoft, lineHeight: 1.8 }}>
              {mergeResult?.log.map((l, i) => <div key={i}>· {l}</div>)}
            </div>

            <h3 style={{ fontSize: 13, fontWeight: 600, margin: '20px 0 10px', fontFamily: MONO, textTransform: 'uppercase', letterSpacing: '0.06em', color: COLORS.inkSoft }}>Auto-resolved</h3>
            <p style={{ fontSize: 14, color: COLORS.inkSoft }}>
              {mergeResult?.autoResolved.length || 0} conflicts on the Requirement column auto-resolved by the
              "prefer 'Chainguard OS' over 'operating system'" rule. They're applied directly without review.
            </p>

            <h3 style={{ fontSize: 13, fontWeight: 600, margin: '20px 0 10px', fontFamily: MONO, textTransform: 'uppercase', letterSpacing: '0.06em', color: COLORS.inkSoft }}>Keyboard — Review</h3>
            <p style={{ fontSize: 13, color: COLORS.inkSoft, margin: '0 0 10px' }}>
              Pick <strong>Keep XLSX</strong> or <strong>Use CSV</strong> per conflict by clicking — every conflict on a row is on the same screen, so each one needs its own click.
            </p>
            {[['↓ or S','Next row'],['↑ or W','Previous row'],['⌘/Ctrl + Enter','Advance from inside the response box'],['?','This help']].map(([k,v]) => (
              <div key={k} style={{ display: 'grid', gridTemplateColumns: '180px 1fr', gap: 16, padding: '4px 0', fontSize: 14, alignItems: 'baseline' }}>
                <kbd style={{ fontFamily: MONO, background: COLORS.bg, border: '1px solid ' + COLORS.rule, padding: '1px 8px', borderRadius: 2, fontSize: 12, width: 'fit-content' }}>{k}</kbd>
                <span>{v}</span>
              </div>
            ))}

            <button style={{ ...iconBtn(false), marginTop: 24 }} onClick={() => setView('review')}>← Back</button>
          </div>
        )}

        {stage === 'review' && view === 'summary' && summary && (
          <div style={{ marginTop: 36 }}>
            <div style={{ borderBottom: '1px solid ' + COLORS.rule, paddingBottom: 18 }}>
              <div style={{ fontFamily: MONO, fontSize: 11, letterSpacing: '0.12em', textTransform: 'uppercase', color: COLORS.inkFaint }}>Export Summary</div>
              <h1 style={{ fontSize: 36, fontWeight: 600, letterSpacing: '-0.015em', margin: '6px 0 4px', lineHeight: 1.05 }}>
                {(xlsxFile?.name || 'merged').replace(/\.xlsx$/i, '')}<em style={{ fontStyle: 'italic', fontWeight: 400, color: COLORS.inkSoft }}>_merged.xlsx</em>
              </h1>
              <div style={{ fontFamily: MONO, fontSize: 12, color: COLORS.inkSoft, display: 'flex', flexWrap: 'wrap', gap: 14, marginTop: 6 }}>
                <span><strong style={{ color: COLORS.csv }}>{summary.stats.totalCsvChosen}</strong> CSV chosen</span>
                <span><strong style={{ color: COLORS.xlsx }}>{summary.stats.totalXlsxChosen}</strong> XLSX kept</span>
                {summary.stats.totalSkipped > 0 && <span><strong style={{ color: COLORS.warn }}>{summary.stats.totalSkipped}</strong> skipped (XLSX kept by default)</span>}
                <span><strong>{summary.stats.totalAutoResolved}</strong> auto-resolved</span>
                <span><strong>{summary.stats.totalResponses}</strong> responses entered</span>
                <span><strong>{summary.stats.totalNewRows}</strong> new rows added</span>
              </div>
              <div style={{ marginTop: 12, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <button style={iconBtn(false)} onClick={handleExport}><Download size={12} /> Re-export</button>
                <button style={iconBtn(false)} onClick={() => setView('review')}><ArrowLeft size={12} /> Back to review</button>
              </div>
            </div>

            {/* Per-row breakdown */}
            <h2 style={{ fontSize: 13, fontWeight: 600, margin: '28px 0 10px', fontFamily: MONO, textTransform: 'uppercase', letterSpacing: '0.06em', color: COLORS.inkSoft }}>
              Changes by row ({summary.rows.length})
            </h2>
            {summary.rows.length === 0 && (
              <div style={{ padding: '16px 18px', background: COLORS.paper, border: '1px solid ' + COLORS.rule, fontStyle: 'italic', color: COLORS.inkSoft }}>
                No per-row changes recorded — only auto-resolved cells and/or new rows below.
              </div>
            )}
            {summary.rows.map((r, idx) => {
              const rowReviewIdx = reviewItems.findIndex((it) => it.xlsxRow === r.xlsxRow);
              return (
                <div key={r.xlsxRow} style={{ marginBottom: 12, border: '1px solid ' + COLORS.rule, background: COLORS.paper }}>
                  <div style={{ padding: '10px 14px', borderBottom: '1px solid ' + COLORS.rule, background: COLORS.bg, display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: 12, flexWrap: 'wrap' }}>
                    <div>
                      <span style={{ fontFamily: SERIF, fontSize: 17, fontWeight: 600 }}>{r.stigid || '—'}</span>
                      <span style={{ fontFamily: MONO, fontSize: 11, color: COLORS.inkSoft, marginLeft: 12 }}>row {r.xlsxRow} · {r.srgid} · {r.cci}</span>
                    </div>
                    {rowReviewIdx >= 0 && (
                      <span onClick={() => { setCurrentIndex(rowReviewIdx); setView('review'); }}
                            style={{ fontFamily: MONO, fontSize: 11, color: COLORS.csv, cursor: 'pointer', textDecoration: 'underline' }}>open in review →</span>
                    )}
                  </div>
                  <div style={{ padding: '10px 14px', fontFamily: MONO, fontSize: 12, lineHeight: 1.6 }}>
                    {r.decided.map((d, i) => {
                      const pal = SOURCE_PALETTE[d.side];
                      return (
                        <SummaryLine
                          key={'d' + i}
                          column={d.column}
                          pill={{ background: pal.bg, color: pal.fg, label: pal.label }}
                          value={(d.value || '').slice(0, 200).replace(/\s+/g, ' ')}
                          valueTitle={d.value}
                        />
                      );
                    })}
                    {r.skipped.map((s, i) => (
                      <SummaryLine
                        key={'s' + i}
                        column={s.column}
                        pill={{ background: COLORS.warnBg, color: COLORS.warn, label: 'skipped' }}
                        valueNode={<span style={{ color: COLORS.inkFaint, fontStyle: 'italic' }}>no decision · XLSX kept by default</span>}
                      />
                    ))}
                    {r.autoResolved.map((a, i) => (
                      <SummaryLine
                        key={'a' + i}
                        column={a.column}
                        pill={{ background: COLORS.addBg, color: COLORS.addFg, label: `auto · ${a.source}` }}
                        valueTitle={a.value}
                        valueNode={<span style={{ color: COLORS.inkSoft, fontStyle: 'italic' }}>{a.reason}</span>}
                      />
                    ))}
                    {r.response && (
                      <div style={{ marginTop: 8, paddingTop: 8, borderTop: '1px dashed ' + COLORS.ruleSoft }}>
                        <div style={{ display: 'flex', alignItems: 'baseline', gap: 8, marginBottom: 4 }}>
                          <MessageSquare size={12} color={COLORS.xlsx} />
                          <span style={{ color: COLORS.xlsx, fontWeight: 600 }}>{r.response.iterationLabel} Government Comment</span>
                        </div>
                        <div style={{ color: COLORS.inkSoft, paddingLeft: 20, marginBottom: 6, whiteSpace: 'pre-wrap' }}>{r.response.govComment}</div>
                        <div style={{ display: 'flex', alignItems: 'baseline', gap: 8, marginBottom: 4 }}>
                          <Save size={12} color={COLORS.csv} />
                          <span style={{ color: COLORS.csv, fontWeight: 600 }}>Vendor Response</span>
                        </div>
                        <div style={{ color: COLORS.ink, paddingLeft: 20, whiteSpace: 'pre-wrap' }}>{r.response.text}</div>
                      </div>
                    )}
                  </div>
                </div>
              );
            })}

            {/* New rows from CSV */}
            {summary.newRows.length > 0 && (
              <>
                <h2 style={{ fontSize: 13, fontWeight: 600, margin: '28px 0 10px', fontFamily: MONO, textTransform: 'uppercase', letterSpacing: '0.06em', color: COLORS.inkSoft }}>
                  New rows added from CSV ({summary.newRows.length})
                </h2>
                <div style={{ border: '1px solid ' + COLORS.rule, background: COLORS.paper, fontFamily: MONO, fontSize: 12 }}>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 12, padding: '10px 14px', background: COLORS.bg, fontSize: 10, letterSpacing: '0.05em', textTransform: 'uppercase', color: COLORS.inkFaint, fontWeight: 600, borderBottom: '1px solid ' + COLORS.rule }}>
                    <span>STIGID</span><span>SRGID</span><span>CCI</span>
                  </div>
                  {summary.newRows.map((nr, i) => (
                    <div key={i} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 12, padding: '7px 14px', borderBottom: i === summary.newRows.length - 1 ? 0 : '1px solid ' + COLORS.ruleSoft }}>
                      <span>{nr.stigid || '—'}</span>
                      <span style={{ color: COLORS.inkSoft }}>{nr.srgid || '—'}</span>
                      <span style={{ color: COLORS.inkSoft }}>{nr.cci || '—'}</span>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        )}

        {stage === 'review' && view === 'list' && (
          <div style={{ marginTop: 36, border: '1px solid ' + COLORS.rule, background: COLORS.paper, maxHeight: '70vh', overflowY: 'auto' }}>
            <div style={{ display: 'grid', gridTemplateColumns: '50px 70px 130px 1fr 70px 80px', gap: 12, padding: '10px 16px', background: COLORS.bg, fontFamily: MONO, fontSize: 10, letterSpacing: '0.05em', textTransform: 'uppercase', color: COLORS.inkFaint, fontWeight: 600, borderBottom: '1px solid ' + COLORS.rule, position: 'sticky', top: 0 }}>
              <span>#</span><span>Row</span><span>STIGID</span><span>Conflicts</span><span>Comment</span><span>Status</span>
            </div>
            {reviewItems.map((item, idx) => {
              const total = item.conflicts.length;
              const decided = item.conflicts.filter((c) => decisions[c.id]).length;
              const hasComment = !!item.comment;
              const hasResp = hasComment && (responses[item.comment.id] || '').trim();
              const conflictsDone = total > 0 && decided === total;
              const conflictsPartial = decided > 0 && decided < total;
              const allDone = (total === 0 || conflictsDone) && (!hasComment || hasResp);
              const noneDone = decided === 0 && !hasResp;
              const status = allDone ? 'done' : noneDone ? 'open' : 'partial';
              const statusPill = STATUS_PILL[status];
              const isCurrent = idx === currentIndex;
              const cols = item.conflicts.map((c) => c.column).join(', ');
              const conflictText = total === 0 ? '—' : `${decided}/${total} · ${cols}`;
              return (
                <div key={item.key} onClick={() => { setCurrentIndex(idx); setView('review'); }}
                     style={{ display: 'grid', gridTemplateColumns: '50px 70px 130px 1fr 70px 80px', gap: 12, padding: '10px 16px', borderBottom: '1px solid ' + COLORS.ruleSoft, fontFamily: MONO, fontSize: 12, cursor: 'pointer', alignItems: 'center', background: isCurrent ? COLORS.ruleSoft : 'transparent' }}>
                  <span style={{ color: COLORS.inkFaint }}>{idx + 1}</span>
                  <span>{item.xlsxRow}</span>
                  <span>{item.stigid || '—'}</span>
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', color: COLORS.inkSoft, fontSize: 11 }}>
                    {conflictsPartial && <span style={{ color: COLORS.csv, marginRight: 6 }}>●</span>}
                    {conflictText}
                  </span>
                  <span style={{ color: hasComment ? (hasResp ? COLORS.xlsx : COLORS.warn) : COLORS.inkFaint, fontSize: 13 }}>
                    {hasComment ? (hasResp ? '✓' : '…') : ''}
                  </span>
                  <span>
                    <span style={{ ...statusPill, padding: '2px 7px', borderRadius: 2, fontSize: 10, letterSpacing: '0.06em', textTransform: 'uppercase', fontWeight: 500 }}>
                      {status}
                    </span>
                  </span>
                </div>
              );
            })}
          </div>
        )}

        {stage === 'review' && view === 'review' && currentItem && (
          <>
            {/* Row header */}
            <div style={{ marginTop: 36, paddingBottom: 18, borderBottom: '1px solid ' + COLORS.rule }}>
              <div style={{ fontFamily: MONO, fontSize: 11, letterSpacing: '0.12em', textTransform: 'uppercase', color: COLORS.inkFaint }}>
                Row {currentIndex + 1} of {reviewItems.length}
                {rowConflicts.length > 0 && (
                  <span style={{ marginLeft: 12 }}>
                    · {rowConflicts.length} conflict{rowConflicts.length === 1 ? '' : 's'}
                    {rowDecidedCount > 0 && <span style={{ color: COLORS.csv }}> ({rowDecidedCount} decided)</span>}
                  </span>
                )}
                {currentRowComment && (
                  <span style={{ marginLeft: 12 }}>
                    · iteration {currentRowComment.iterationLabel} comment
                  </span>
                )}
              </div>
              <h1 style={{ fontSize: 36, fontWeight: 600, letterSpacing: '-0.015em', margin: '6px 0 4px', lineHeight: 1.05 }}>
                {currentItem.stigid || '—'}
              </h1>
              <div style={{ fontFamily: MONO, fontSize: 12, color: COLORS.inkSoft, letterSpacing: '0.02em' }}>
                Row {currentItem.xlsxRow}
                <span style={{ display: 'inline-block', margin: '0 8px', opacity: 0.4 }}>·</span>
                {currentItem.srgid}
                <span style={{ display: 'inline-block', margin: '0 8px', opacity: 0.4 }}>·</span>
                {currentItem.cci}
                {currentItem.matchMethod && (
                  <>
                    <span style={{ display: 'inline-block', margin: '0 8px', opacity: 0.4 }}>·</span>
                    {currentItem.matchMethod}
                  </>
                )}
              </div>
            </div>

            {/* Conflicts — stacked, label-on-left layout */}
            {rowConflicts.length > 0 && (
              <div style={{ marginTop: 24, border: '1px solid ' + COLORS.rule, background: COLORS.paper }}>
                {rowConflicts.map((c, ci) => {
                  const dec = decisions[c.id];
                  const diff = diffWords(c.xlsxValue || '', c.csvValue || '');
                  const decPill = DEC_PILL[dec] || null;
                  return (
                    <div key={c.id}
                         id={`conflict-${c.id}`}
                         style={{ display: 'grid', gridTemplateColumns: '180px 1fr 1fr', borderTop: ci === 0 ? 'none' : '1px solid ' + COLORS.rule, scrollMarginTop: 80 }}>
                      {/* Gutter: column name + status pill, spans both content rows */}
                      <div style={{ gridRow: 'span 3', padding: '14px 16px', background: COLORS.bg, borderRight: '1px solid ' + COLORS.rule, display: 'flex', flexDirection: 'column', gap: 8 }}>
                        <div style={{ fontFamily: SERIF, fontStyle: 'italic', fontSize: 18, fontWeight: 500, color: COLORS.ink, lineHeight: 1.2 }}>
                          {c.column}
                        </div>
                        <div style={{ fontFamily: MONO, fontSize: 10, letterSpacing: '0.06em', textTransform: 'uppercase', color: COLORS.inkFaint }}>
                          conflict {ci + 1}/{rowConflicts.length}
                        </div>
                        {decPill && (
                          <span style={{ background: decPill.background, color: decPill.color, padding: '3px 8px', borderRadius: 2, fontSize: 10, letterSpacing: '0.06em', textTransform: 'uppercase', fontWeight: 600, alignSelf: 'flex-start' }}>
                            {decPill.label}
                          </span>
                        )}
                        {dec && (
                          <span onClick={() => clearDecision(c.id)} style={{ fontFamily: MONO, fontSize: 11, color: COLORS.inkSoft, cursor: 'pointer', textDecoration: 'underline' }}>clear</span>
                        )}
                      </div>

                      {/* XLSX header */}
                      <div style={{ padding: '8px 14px', borderBottom: '1px solid ' + COLORS.rule, borderRight: '1px solid ' + COLORS.rule, background: COLORS.xlsxSoft, display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: 12 }}>
                        <span style={{ fontStyle: 'italic', fontSize: 15, fontWeight: 500, color: COLORS.xlsx }}>XLSX</span>
                        <span style={{ fontFamily: MONO, fontSize: 9, letterSpacing: '0.1em', textTransform: 'uppercase', color: COLORS.inkFaint }}>current</span>
                      </div>
                      {/* CSV header */}
                      <div style={{ padding: '8px 14px', borderBottom: '1px solid ' + COLORS.rule, background: COLORS.csvSoft, display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: 12 }}>
                        <span style={{ fontStyle: 'italic', fontSize: 15, fontWeight: 500, color: COLORS.csv }}>CSV</span>
                        <span style={{ fontFamily: MONO, fontSize: 9, letterSpacing: '0.1em', textTransform: 'uppercase', color: COLORS.inkFaint }}>proposed</span>
                      </div>

                      {/* XLSX diff */}
                      <div style={{ padding: '14px 14px', borderRight: '1px solid ' + COLORS.rule, fontFamily: MONO, fontSize: 13, lineHeight: 1.6, whiteSpace: 'pre-wrap', wordBreak: 'break-word', overflowY: 'auto', maxHeight: 360 }}>
                        {diff.map((op, i) => <React.Fragment key={i}>{renderDiffSide(op, SIDE.XLSX)}</React.Fragment>)}
                      </div>
                      {/* CSV diff */}
                      <div style={{ padding: '14px 14px', fontFamily: MONO, fontSize: 13, lineHeight: 1.6, whiteSpace: 'pre-wrap', wordBreak: 'break-word', overflowY: 'auto', maxHeight: 360 }}>
                        {diff.map((op, i) => <React.Fragment key={i}>{renderDiffSide(op, SIDE.CSV)}</React.Fragment>)}
                      </div>

                      <button onClick={() => choose(c.id, SIDE.XLSX)}
                              style={{ padding: '14px 16px', border: 0, borderTop: '1px solid ' + COLORS.rule, borderRight: '1px solid ' + COLORS.rule, background: dec === SIDE.XLSX ? COLORS.ink : COLORS.paper, color: dec === SIDE.XLSX ? COLORS.paper : COLORS.ink, cursor: 'pointer', fontFamily: SERIF, fontSize: 15, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10 }}>
                        <ArrowLeft size={16} /><span>Keep XLSX</span>
                      </button>
                      <button onClick={() => choose(c.id, SIDE.CSV)}
                              style={{ padding: '14px 16px', border: 0, borderTop: '1px solid ' + COLORS.rule, background: dec === SIDE.CSV ? COLORS.ink : COLORS.paper, color: dec === SIDE.CSV ? COLORS.paper : COLORS.ink, cursor: 'pointer', fontFamily: SERIF, fontSize: 15, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10 }}>
                        <span>Use CSV</span><ArrowRight size={16} />
                      </button>
                    </div>
                  );
                })}
              </div>
            )}

            {/* Comment-only placeholder when no conflicts */}
            {rowConflicts.length === 0 && currentRowComment && (
              <div style={{ marginTop: 24, padding: '24px 20px', background: COLORS.paper, border: '1px solid ' + COLORS.rule, fontStyle: 'italic', color: COLORS.inkSoft, fontSize: 15 }}>
                No merge conflicts on this row — only a government comment to address.
              </div>
            )}

            {/* Government comment + response (once per row, if any) */}
            {currentRowComment && (
              <div style={{ marginTop: rowConflicts.length > 0 ? 28 : 16 }}>
                <div style={{ border: '1px solid ' + COLORS.rule, background: COLORS.paper }}>
                  <div style={{ padding: '10px 16px', borderBottom: '1px solid ' + COLORS.rule, background: COLORS.xlsxSoft, display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: 12 }}>
                    <span style={{ fontStyle: 'italic', fontSize: 17, fontWeight: 500, color: COLORS.xlsx }}>
                      <MessageSquare size={14} style={{ display: 'inline', verticalAlign: '-2px', marginRight: 6 }} />
                      {currentRowComment.iterationLabel} Government Comment
                    </span>
                    <span style={{ fontFamily: MONO, fontSize: 10, letterSpacing: '0.1em', textTransform: 'uppercase', color: COLORS.inkFaint }}>from reviewer</span>
                  </div>
                  <div style={{ padding: '16px', fontFamily: MONO, fontSize: 13, lineHeight: 1.65, whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>
                    {currentRowComment.govComment}
                  </div>
                </div>
                <div style={{ marginTop: 8, border: '1px solid ' + COLORS.rule, borderTop: 0, background: COLORS.paper }}>
                  <div style={{ padding: '10px 16px', borderBottom: '1px solid ' + COLORS.rule, background: COLORS.csvSoft, display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: 12 }}>
                    <span style={{ fontStyle: 'italic', fontSize: 17, fontWeight: 500, color: COLORS.csv }}>
                      {currentRowComment.iterationLabel} Vendor Response
                    </span>
                    <span style={{ fontFamily: MONO, fontSize: 10, letterSpacing: '0.1em', textTransform: 'uppercase', color: COLORS.inkFaint }}>
                      {(responses[currentRowComment.id] || '').trim() ? 'auto-saved ✓' : 'auto-saves as you type'}
                    </span>
                  </div>
                  <textarea
                    value={responses[currentRowComment.id] !== undefined ? responses[currentRowComment.id] : (currentRowComment.existingVendorResponse || '')}
                    onChange={(e) => setResponseFor(currentRowComment.id, e.target.value)}
                    placeholder="Type your response…  (⌘/Ctrl+Enter to advance)"
                    style={{ width: '100%', minHeight: 120, padding: '14px 16px', border: 0, outline: 'none', background: 'transparent', resize: 'vertical', fontFamily: MONO, fontSize: 13, lineHeight: 1.65, color: COLORS.ink, display: 'block', boxSizing: 'border-box' }}
                  />
                </div>
              </div>
            )}

            {/* Sub-actions */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 14, fontFamily: MONO, fontSize: 11, color: COLORS.inkFaint, flexWrap: 'wrap', gap: 12 }}>
              <div style={{ display: 'flex', gap: 18, flexWrap: 'wrap' }}>
                <span onClick={goBack} style={{ color: COLORS.inkSoft, cursor: 'pointer' }}>← previous row</span>
                <span onClick={advance} style={{ color: COLORS.inkSoft, cursor: 'pointer' }}>next row →</span>
                {rowDecidedCount > 0 && <span onClick={clearRowDecisions} style={{ color: COLORS.inkSoft, cursor: 'pointer' }}>clear all decisions on this row</span>}
                {currentRowComment && responses[currentRowComment.id] !== undefined && (
                  <span onClick={clearResponse} style={{ color: COLORS.inkSoft, cursor: 'pointer' }}>clear response</span>
                )}
              </div>
            </div>
          </>
        )}

        {stage === 'review' && view === 'review' && !currentItem && (
          <div style={{ marginTop: 60, textAlign: 'center' }}>
            <h2 style={{ fontSize: 28, fontWeight: 500, fontStyle: 'italic', margin: '0 0 10px', color: COLORS.inkSoft }}>Nothing to review.</h2>
            <p style={{ color: COLORS.inkFaint, maxWidth: 480, margin: '8px auto 24px' }}>
              No conflicts and no government comments. Click <strong>Export XLSX</strong> to download the merged file.
            </p>
          </div>
        )}

      </div>
    </div>
  );
}
