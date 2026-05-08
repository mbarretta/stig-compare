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

## Deploy

### Netlify

1. Push the repo to GitHub/GitLab.
2. In Netlify: **Add new site → Import an existing project**.
3. Set build command: `npm run build`
4. Set publish directory: `dist`
5. Deploy.

Or via CLI:

```bash
npm install -g netlify-cli
netlify deploy --prod --dir dist
```

### Vercel

```bash
npm install -g vercel
vercel --prod
```

Vercel auto-detects Vite. No configuration needed.

### GitHub Pages

1. Install the deploy plugin:

```bash
npm install -D gh-pages
```

2. Add to `package.json` scripts:

```json
"deploy": "npm run build && gh-pages -d dist"
```

3. Run:

```bash
npm run deploy
```

### Any static host (S3, nginx, Caddy, etc.)

Run `npm run build`, then serve the contents of `dist/` as static files. The app is entirely client-side — no server required.

If you're serving from a sub-path (e.g. `https://example.com/stig-compare/`), set the base in `vite.config.js`:

```js
export default defineConfig({
  base: '/stig-compare/',
  plugins: [react()],
});
```

## Usage

1. Open the app in a browser.
2. Drop in your **CSV** (Vulcan export) and **XLSX** (working spreadsheet).
3. Click **Run merge**.
4. Step through conflicts — use arrow keys or the buttons to keep XLSX or use CSV.
5. If the XLSX has government comments, switch to the **Comments** tab and type vendor responses.
6. Click **Export XLSX** to download the merged file.

Progress is saved automatically to `localStorage` keyed by filename + size, so you can close the tab and resume.

## Notes

- Cell formatting (colors, styles) in the original XLSX is not preserved in the exported file. This is a limitation of the SheetJS community edition used in the browser. If preserving review color-coding matters, run the merge server-side with openpyxl.
- The `xlsx` package has known audit warnings in its community edition. The app uses it only to read and write user-supplied local files, so there is no server-side exposure.
