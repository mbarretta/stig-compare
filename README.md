# STIG Merge Review

A browser-based tool for merging a fresh CSV export from MITRE Vulcan into a working STIG XLSX spreadsheet. Detects conflicts cell-by-cell, shows word-level diffs, lets you pick a side for each one, handles government comment responses, and exports a merged XLSX.

## Requirements

- Node.js 18+
- npm 9+

## Install

```bash
npm install
```

## Run locally

```bash
npm run dev
```

Opens at `http://localhost:5173` by default.

## Build for deployment

```bash
npm run build
```

Output goes to `dist/`. Serve that directory from any static host.

To preview the production build locally before deploying:

```bash
npm run preview
```

## Tests

```bash
npm test
```

The merge logic (matching, concurrence gate, umbrella handling, auto-resolve rules) lives in `merge_logic.js` — a plain-JS module with no React or DOM dependencies, imported by both the app (`merge_review.jsx`) and the test suite:

- `test/merge_logic.test.mjs` — unit tests for the helpers plus `runMerge` scenarios on synthetic rows (concurrence gating, umbrella Satisfied-By sets, STIGID-first pairing, Jaccard sub-matching, OS-phrasing auto-resolve, Not-Yet-Determined row replacement).
- `test/merge_integration.test.mjs` — characterization test that runs the real DISA baseline × Vulcan export merge and pins the verified outcome (17 conflicts, 97 concurred, 47 umbrella auto-keeps, 14 new rows). The real data files under `test/` are local-only (gitignored); the test skips when they're absent.

`fixtures/make_fixtures.mjs` generates a small synthetic CSV/XLSX pair for exercising the app end-to-end in the browser, and `fixtures/check_export.mjs` asserts on the exported result.

## Usage

1. Open the app in a browser.
2. Drop in your **CSV** (Vulcan export) and **XLSX** (working spreadsheet).
3. Click **Run merge**.
4. Step through conflicts — use arrow keys or the buttons to keep XLSX or use CSV.
5. If the XLSX has government comments, switch to the **Comments** tab and type vendor responses.
6. Click **Export XLSX** to download the merged file.

Progress is saved automatically to `localStorage` keyed by filename + size, so you can close the tab and resume.

## How matching and merging work

The XLSX is treated as the **trusted, DISA-approved baseline**; the CSV is a fresh Vulcan export whose changes are merged *into* it. The whole pipeline lives in `merge_review.jsx` (`runMerge` for analysis, `applyAndExport` for writing).

### 1. Join key

Rows are joined on **`SRGID + CCI`** — not STIGID. STIGID renumbers between exports (a large fraction of rules get a new 6-digit suffix each round), so it is unreliable as a global identifier. `SRGID + CCI` is stable, but it is **not unique**: several rules can share one key (e.g. the audit-syscall family under `SRG-OS-000037-GPOS-00015 / CCI-000130` holds many rows), so a single key can map a *group* of CSV rows to a *group* of XLSX rows.

### 2. Matching within a key group

For each `SRGID + CCI` key:

- **1 × 1** — exactly one row on each side → matched directly.
- **N × M** — multiple rows on either side are sub-matched in two passes:
  1. **STIGID pairing.** Rows that share the same STIGID under the key are the same control, so they are paired first. This stops a CSV row from being treated as "new" (and inserted as a duplicate) when an identically-numbered baseline row already exists. STIGID is used only as a *positive* signal — renumbered rows simply fall through to the next pass.
  2. **Jaccard signature.** Remaining rows are paired by content similarity. A signature is extracted from `Requirement / Check / Fix / VulDiscussion` — file paths, quoted strings, and technical tokens (`syscall`, `chmod`, octal modes like `0644`, `systemctl`, …) — and rows are matched greedily, highest score first, at a similarity threshold of **0.30**.

CSV rows that never match become **new rows** (inserted in place — see §5). Unmatched XLSX rows whose signature is empty are flagged **ambiguous** rather than silently dropped.

### 3. Concurrence gate

A row is considered **settled** when its review history shows concurrence: the latest non-empty *Government Comments* cell starts with "concur" (but not "non-concur"), **or** the latest *Vendor Response* contains "concur" (not "non-concur"). Iteration columns are `Nth Government Comments` / `Nth Vendor Response`; the highest-numbered non-empty cell wins.

Settled rows are **not** skipped wholesale — the Vulcan export can still carry a live divergence on a row that already reads "Concur." Instead, only the audit-critical columns **Check / Fix / Status** are re-examined on settled rows; every other column is left as the approved baseline.

### 4. Per-column resolution

For every matched pair, each shared column is compared (blank-vs-blank and whitespace-only differences are ignored). Differences are auto-resolved in this order, and whatever is left becomes a **conflict** for manual review:

1. **Status placeholder → whole row from CSV.** If the XLSX `Status` is `Not Yet Determined` and the CSV has a real determination, the entire row is taken from the CSV. (Never applied to a settled row.)
2. **Umbrella controls.** Many SRG rows carry a `Satisfied By: CGOS-…, CGOS-…` list in *Vendor Comments*; for these, `Check` / `Fix` is a *rendering* of the first-listed base rule, not authored text, so it changes purely from list reordering. These columns are compared by the **set** of base rules, not the text:
   - set unchanged → keep the baseline XLSX verbatim (no churn, no conflict);
   - set changed → real requirement change, surfaced as a conflict with the added/removed base rules shown.
3. **OS phrasing.** If the only difference is `operating system` → `Chainguard OS`, the Chainguard wording wins automatically.
4. **Requirement preference.** In the `Requirement` column, the side that says `Chainguard OS` is preferred over generic `operating system` phrasing.

Undecided conflicts default to keeping the baseline (XLSX) on export until you pick a side in the UI.

### 5. Writing the merged XLSX

Export writes decisions back into the **original workbook** with ExcelJS, so cell styling (font, alignment, borders, fills) is preserved.

- **True-row mapping.** The baseline can contain a large block of blank rows and a strand of rows appended by a prior merge, so the compact parse index is not the worksheet row number. Each parsed row carries its real worksheet row, and every read/write/insert uses it.
- **In-place insertion.** New rows are inserted into their `SRGID + CCI` group in STIGID-suffix order — sorted by `(primary SRGID token, primary CCI token, STIGID 6-digit suffix)` — rather than appended at the bottom. This sort reproduces the approved DISA ordering exactly, so existing rows are never reordered.
- **Full cleanup.** Blank-row gaps are collapsed and any stranded rows from a previous merge are lifted back into their correct sorted position, producing one contiguous, ordered sheet.
- **Fill normalization.** Inserted rows are reset to the baseline convention: the five SRG-prefix reference columns (`SRGID`, `SRG Requirement`, `SRG VulDiscussion`, `SRG Check`, `SRG Fix`) keep the gray fill, every other cell is white — so an inserted row never inherits a stale changed-cell highlight from its neighbor.

## Notes

- Reading and writing both use [ExcelJS](https://github.com/exceljs/exceljs) in the browser, which **preserves** the original workbook's cell formatting (fonts, alignment, borders, fills) in the exported file.
- The app is entirely client-side — uploaded files never leave the browser. CSV parsing uses PapaParse; XLSX read/write uses ExcelJS.
