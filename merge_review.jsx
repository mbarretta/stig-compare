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

import {
  SHARED_COLS, CSV_ONLY_COLS, SIDE, SOURCE, CONCURRED_REVIEW_COLS,
  cellKey, conflictId, bucketByXlsxRow, clean, diffWords, runMerge,
  cellText, readHeaders, worksheetToRows,
} from './merge_logic.js';

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

  const concurredSet = new Set(mergeResult.concurredXlsxIdx || []);
  for (const [xiStr, ci] of Object.entries(mergeResult.pairings)) {
    const xi = parseInt(xiStr, 10);
    const isConcurred = concurredSet.has(xi);
    const cr = csvRows[ci];
    const row = ws.getRow(mergeResult.xlsxRowNums[xi]);

    // Settled rows stay untouched except for the Check/Fix/Status cells that
    // auto-resolved or that the reviewer explicitly took from CSV.
    if (!isConcurred && colMap.has('STIGID')) {
      row.getCell(colMap.get('STIGID')).value = clean(cr['STIGID']);
    }

    for (const col of SHARED_COLS) {
      if (col === 'STIGID') continue;
      if (!colMap.has(col)) continue;
      if (isConcurred && !CONCURRED_REVIEW_COLS.has(col)) continue;
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

      if (!isConcurred && existingVal === '' && csvVal !== '') {
        cell.value = csvVal;
      }
    }

    if (!isConcurred) {
      for (const col of CSV_ONLY_COLS) {
        if (!colMap.has(col)) continue;
        const csvVal = clean(cr[col]);
        if (csvVal) row.getCell(colMap.get(col)).value = csvVal;
      }
    }
  }

  if (mergeResult.iteration && mergeResult.iteration.vendorResponseColumn) {
    const venCol = colMap.get(mergeResult.iteration.vendorResponseColumn);
    if (venCol !== undefined) {
      for (const cmt of mergeResult.comments) {
        const resp = (responses[cmt.id] || '').trim();
        if (resp) {
          ws.getRow(mergeResult.xlsxRowNums[cmt.xlsxIndex]).getCell(venCol).value = resp;
        }
      }
    }
  }

  // Full cleanup. The baseline can carry a large block of blank rows followed by
  // a strand of rows appended below them by a prior merge. We rebuild a single
  // contiguous, SRGID/CCI/STIGID-ordered sheet: keep the canonical top block in
  // place, lift every stranded row and every genuinely-new CSV row into its
  // correct sorted position, and drop all blank rows. The sort key
  // (primary SRGID, primary CCI, STIGID suffix) reproduces the approved DISA top
  // block exactly, so existing rows are never reordered.
  const srgCol = colMap.get('SRGID');
  const cciCol = colMap.get('CCI');
  const stigCol = colMap.get('STIGID');
  const primary = (v) => clean(v).split(/[,\n;]+/)[0].trim(); // first token of a multi-value cell
  const suffixOf = (v) => {
    const m = clean(v).match(/(\d{3,})\s*$/);
    return m ? parseInt(m[1], 10) : -1; // umbrella rows (blank STIGID) sort first within their group
  };
  const rowText = (r, col) => (col !== undefined ? cellText(ws.getRow(r).getCell(col).value) : '');
  const tupleOfRow = (r) => [primary(rowText(r, srgCol)), primary(rowText(r, cciCol)), suffixOf(rowText(r, stigCol))];
  const tupleOfVals = (v) => [primary(v['SRGID']), primary(v['CCI']), suffixOf(v['STIGID'])];
  const cmpTuple = (a, b) =>
    a[0] !== b[0] ? (a[0] < b[0] ? -1 : 1) : a[1] !== b[1] ? (a[1] < b[1] ? -1 : 1) : a[2] - b[2];

  // Locate the first blank gap: rows before it are the canonical block; the
  // compact rows at or after that index are stranded.
  const rowNums = mergeResult.xlsxRowNums;
  let firstGap = rowNums.length;
  for (let i = 1; i < rowNums.length; i++) {
    if (rowNums[i] - rowNums[i - 1] > 1) { firstGap = i; break; }
  }
  const lastTopRow = firstGap > 0 ? rowNums[firstGap - 1] : 1;

  // Capture final (post-edit) values for stranded rows, then the new CSV rows.
  const toPlace = []; // { values: {col: val}, tuple }
  for (let i = firstGap; i < rowNums.length; i++) {
    const r = rowNums[i];
    const values = {};
    for (const [colName, cIdx] of colMap) values[colName] = clean(cellText(ws.getRow(r).getCell(cIdx).value));
    toPlace.push({ values, tuple: tupleOfVals(values) });
  }
  for (const ci of mergeResult.newCsvIdx) {
    const cr = csvRows[ci];
    const values = {};
    for (const [colName] of colMap) values[colName] = clean(cr[colName]);
    toPlace.push({ values, tuple: tupleOfVals(values) });
  }

  // Cut everything below the canonical block (blank gap + strand). A single
  // large-count splice is a no-op here because the blank rows are unmaterialized;
  // splicing one row at a time reliably collapses both the gap and the strand.
  let guard = ws.rowCount + 1;
  while (ws.rowCount > lastTopRow && guard-- > 0) ws.spliceRows(lastTopRow + 1, 1);

  // Inserted rows inherit fill via 'i+' (from the row above) or from the row
  // below in the header-adjacent case — either of which can carry a stale
  // changed-cell highlight (e.g. green FF92D050). Normalize every inserted row to
  // the baseline data-row convention: the five SRG-prefix reference columns keep
  // the gray fill, every other cell is plain white. Only the fill is overridden,
  // so inherited font/alignment/wrap/border (Arial 11, top, wrap) are preserved.
  const SRG_GRAY_COLS = new Set(['SRGID', 'SRG Requirement', 'SRG VulDiscussion', 'SRG Check', 'SRG Fix']);
  const GRAY_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC0C0C0' }, bgColor: { argb: 'FFC0C0C0' } };
  const WHITE_FILL = { type: 'pattern', pattern: 'none' };

  // Insert each captured row at its sorted position. Inserting in key order means
  // same-group rows land in suffix order next to one another.
  toPlace.sort((a, b) => cmpTuple(a.tuple, b.tuple));
  for (const item of toPlace) {
    let target = ws.rowCount + 1;
    for (let r = 2; r <= ws.rowCount; r++) {
      if (cmpTuple(tupleOfRow(r), item.tuple) > 0) { target = r; break; }
    }
    // 'i+' inherits styling from the row above; for the header-adjacent case copy
    // style from the row below instead so we don't inherit the header's style.
    const row = ws.insertRow(target, [], target <= 2 ? 'n' : 'i+');
    if (target <= 2 && ws.rowCount > target) {
      const below = ws.getRow(target + 1);
      row.eachCell({ includeEmpty: true }, (cell, c) => { cell.style = { ...below.getCell(c).style }; });
    }
    for (const [colName, cIdx] of colMap) {
      const cell = row.getCell(cIdx);
      const val = item.values[colName];
      if (val) cell.value = val;
      cell.fill = SRG_GRAY_COLS.has(colName) ? { ...GRAY_FILL } : { ...WHITE_FILL };
    }
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
      const rows = worksheetToRows(ws);

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

  // Bulk: apply one side to every conflict on the current row.
  const chooseAllInRow = useCallback((side) => {
    if (!rowConflicts.length) return;
    setDecisions((d) => {
      const n = { ...d };
      for (const c of rowConflicts) n[c.id] = side;
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
    // Fresh arrays per call — a shared defaults object would alias decided/
    // skipped/autoResolved across every row.
    const mkSeed = () => ({ decided: [], skipped: [], autoResolved: [], response: null });
    for (const c of mergeResult.conflicts) {
      const dec = decisions[c.id];
      const e = bucketByXlsxRow(rowMap, c.xlsxRow, c, mkSeed());
      if (dec === SIDE.CSV) e.decided.push({ column: c.column, side: SOURCE.CSV, value: c.csvValue, replaced: c.xlsxValue });
      else if (dec === SIDE.XLSX) e.decided.push({ column: c.column, side: SOURCE.XLSX, value: c.xlsxValue, replaced: c.csvValue });
      else e.skipped.push({ column: c.column, value: c.xlsxValue });
    }
    for (const ar of mergeResult.autoResolved || []) {
      const e = bucketByXlsxRow(rowMap, ar.xlsxRow, ar, mkSeed());
      e.autoResolved.push({ column: ar.column, source: ar.chosenSource, value: ar.resolvedValue, reason: ar.reason });
    }
    for (const cm of mergeResult.comments) {
      const resp = (responses[cm.id] || '').trim();
      if (!resp) continue;
      const e = bucketByXlsxRow(rowMap, cm.xlsxRow, cm, mkSeed());
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
                <span>{conflicts.length} conflicts · {comments.length} comments · {reviewItems.length} rows{mergeResult?.metadata?.concurredCount ? ` · ${mergeResult.metadata.concurredCount} concurred` : ''}</span>
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

            <h3 style={{ fontSize: 13, fontWeight: 600, margin: '20px 0 10px', fontFamily: MONO, textTransform: 'uppercase', letterSpacing: '0.06em', color: COLORS.inkSoft }}>Concurred rows</h3>
            <p style={{ fontSize: 14, color: COLORS.inkSoft }}>
              {mergeResult?.metadata?.concurredCount || 0} matched rows are already settled — the latest Government
              Comments cell starts with "Concur" <em>or</em> the latest Vendor Response contains "Concur". These are
              left untouched in the export <em>except</em> for the audit-critical columns (<strong>Check</strong>,
              {' '}<strong>Fix</strong>, <strong>Status</strong>): if Vulcan's export still disagrees with the baseline
              there, the cell is surfaced as a conflict tagged "previously concurred" so a stray live change can't be
              silently dropped. {mergeResult?.metadata?.concurredConflictCount || 0} such conflicts were surfaced.
            </p>

            <h3 style={{ fontSize: 13, fontWeight: 600, margin: '20px 0 10px', fontFamily: MONO, textTransform: 'uppercase', letterSpacing: '0.06em', color: COLORS.inkSoft }}>Auto-resolved</h3>
            <p style={{ fontSize: 14, color: COLORS.inkSoft }}>
              {mergeResult?.autoResolved.length || 0} cells auto-resolved without review, by four rules:
              (1) when XLSX <em>Status</em> was "Not Yet Determined", the whole row is taken from CSV;
              (2) when a cell's only difference is "operating system" → "Chainguard OS", the Chainguard OS
              wording wins; (3) on the Requirement column, the side that says "Chainguard OS" is preferred;
              (4) on umbrella controls (those with a "Satisfied By" list), <em>Check</em> and <em>Fix</em> are
              compared by their base-rule SET rather than their text — if the set is unchanged the baseline
              XLSX text is kept (the differing text is only a different representative of the same rules).
            </p>

            <h3 style={{ fontSize: 13, fontWeight: 600, margin: '20px 0 10px', fontFamily: MONO, textTransform: 'uppercase', letterSpacing: '0.06em', color: COLORS.inkSoft }}>Umbrella ("Satisfied By") controls</h3>
            <p style={{ fontSize: 14, color: COLORS.inkSoft }}>
              {mergeResult?.metadata?.umbrellaAutoCount || 0} Check/Fix cells were kept because the control's
              "Satisfied By" base-rule set was unchanged (the text only differs by which base rule is surfaced).
              {' '}{mergeResult?.metadata?.umbrellaConflictCount || 0} umbrella controls were flagged as real
              conflicts because their base-rule set actually changed — those show the added/removed rules.
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
                {mergeResult?.metadata?.concurredCount > 0 && (
                  <span><strong style={{ color: COLORS.xlsx }}>{mergeResult.metadata.concurredCount}</strong> concurred (Check/Fix/Status re-checked){mergeResult?.metadata?.concurredConflictCount ? ` · ${mergeResult.metadata.concurredConflictCount} surfaced` : ''}</span>
                )}
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
                {rowConflicts.length > 1 && (
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12, flexWrap: 'wrap', padding: '10px 14px', borderBottom: '1px solid ' + COLORS.rule, background: COLORS.bg }}>
                    <span style={{ fontFamily: MONO, fontSize: 10, letterSpacing: '0.08em', textTransform: 'uppercase', color: COLORS.inkFaint }}>
                      {rowConflicts.length} conflicts · apply one side to all
                    </span>
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button onClick={() => chooseAllInRow(SIDE.XLSX)}
                              title="Keep the XLSX value for every conflict on this row"
                              style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '6px 12px', cursor: 'pointer', fontFamily: MONO, fontSize: 11, letterSpacing: '0.04em', textTransform: 'uppercase', border: '1px solid ' + COLORS.xlsx, background: COLORS.xlsxSoft, color: COLORS.xlsx }}>
                        <ArrowLeft size={13} /> Keep all XLSX
                      </button>
                      <button onClick={() => chooseAllInRow(SIDE.CSV)}
                              title="Use the CSV value for every conflict on this row"
                              style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '6px 12px', cursor: 'pointer', fontFamily: MONO, fontSize: 11, letterSpacing: '0.04em', textTransform: 'uppercase', border: '1px solid ' + COLORS.csv, background: COLORS.csvSoft, color: COLORS.csv }}>
                        Use all CSV <ArrowRight size={13} />
                      </button>
                    </div>
                  </div>
                )}
                {rowConflicts.map((c, ci) => {
                  const dec = decisions[c.id];
                  const diff = diffWords(c.xlsxValue || '', c.csvValue || '');
                  const decPill = DEC_PILL[dec] || null;
                  return (
                    <div key={c.id}
                         id={`conflict-${c.id}`}
                         style={{ display: 'grid', gridTemplateColumns: '180px 1fr 1fr', borderTop: ci === 0 ? 'none' : '1px solid ' + COLORS.rule, scrollMarginTop: 80 }}>
                      {/* Gutter: column name + status pill. Spans every content row —
                          headers + diffs + buttons (3) plus one per optional banner —
                          so the buttons stay in the XLSX/CSV columns, not under the gutter. */}
                      <div style={{ gridRow: `span ${3 + (c.umbrella ? 1 : 0) + (c.concurred ? 1 : 0)}`, padding: '14px 16px', background: COLORS.bg, borderRight: '1px solid ' + COLORS.rule, display: 'flex', flexDirection: 'column', gap: 8 }}>
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

                      {/* Umbrella set-change banner — the real signal is the
                          base-rule delta, not the representative text diff. */}
                      {c.umbrella && (
                        <div style={{ gridColumn: '2 / 4', padding: '10px 14px', borderBottom: '1px solid ' + COLORS.rule, background: COLORS.warnBg, fontFamily: MONO, fontSize: 11, lineHeight: 1.6, color: '#5e3a16' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontWeight: 700, letterSpacing: '0.04em', textTransform: 'uppercase', marginBottom: 4 }}>
                            <AlertTriangle size={12} /> "Satisfied By" set changed · {c.umbrella.xCount} → {c.umbrella.cCount} base rules
                          </div>
                          <div>This is an umbrella control — its {c.column} text is just a representative of its base rules. The substantive change is which base rules satisfy it:</div>
                          {c.umbrella.added.length > 0 && (
                            <div style={{ marginTop: 3 }}><span style={{ color: COLORS.addFg, fontWeight: 600 }}>+ added (CSV):</span> {c.umbrella.added.join(', ')}</div>
                          )}
                          {c.umbrella.removed.length > 0 && (
                            <div style={{ marginTop: 3 }}><span style={{ color: COLORS.delFg, fontWeight: 600 }}>− removed (XLSX):</span> {c.umbrella.removed.join(', ')}</div>
                          )}
                          <div style={{ marginTop: 4, fontStyle: 'italic', opacity: 0.85 }}>The text panels below differ only by representative selection — decide based on the set change above.</div>
                        </div>
                      )}

                      {/* Previously-concurred banner — this row was marked settled
                          ("Concur…") yet the export still diverges here. */}
                      {c.concurred && (
                        <div style={{ gridColumn: '2 / 4', padding: '10px 14px', borderBottom: '1px solid ' + COLORS.rule, background: COLORS.warnBg, fontFamily: MONO, fontSize: 11, lineHeight: 1.6, color: '#5e3a16' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontWeight: 700, letterSpacing: '0.04em', textTransform: 'uppercase' }}>
                            <AlertTriangle size={12} /> Previously concurred — export still differs
                          </div>
                          <div style={{ marginTop: 3 }}>This row was marked settled ("Concur"), so it was otherwise skipped — but Vulcan's {c.column} still disagrees with the baseline. Decide whether this is a live change to merge.</div>
                        </div>
                      )}

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
