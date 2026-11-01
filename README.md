# XLSM Cloud Studio

> Modern, real-time spreadsheet analytics dashboard and VBA macro engineering studio. Ingests legacy `.xlsm`, `.xlsx`, `.xls`, and `.csv` files, automatically stripping workbook/sheet password protection locks, compiling reactive charts and dynamic pivot matrices, simulating legacy VBA procedures, and exporting clean RFC-4180 CSV and Excel files.

**Developed by**: Ammar Thaqif ([ammarthaqif.ar@gmail.com](mailto:ammarthaqif.ar@gmail.com))

---

## 🚀 Instant Deployment

### Option 1: Deploy to Vercel (Recommended — Zero Config)

This repository is pre-configured for Vercel with `vercel.json` and a serverless `/api/health` handler.

1. Push this repository to your GitHub account:
   ```bash
   git init
   git add .
   git commit -m "feat: initial commit of XLSM Cloud Studio"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git push -u origin main
   ```
2. Go to [vercel.com/new](https://vercel.com/new).
3. Import your GitHub repository.
4. Framework Preset will auto-detect as **Vite**.
5. Click **Deploy**. Vercel will build and launch the application globally in seconds!

---

### Option 2: Automated CI with GitHub Actions

The repository includes a ready-to-use GitHub Actions workflow at [`.github/workflows/ci.yml`](.github/workflows/ci.yml).
- Automatically triggers on every push and pull request to `main` and `master`.
- Validates dependencies with `npm ci`.
- Runs TypeScript type checking (`npm run lint`).
- Builds the production bundle (`npm run build`) and uploads production build artifacts.

---

## 🛠️ Local Development & Scripts

### Prerequisites
- Node.js 18+ (Node 20+ recommended)
- npm 9+

### Installation
```bash
npm install
```

### Development Server
```bash
npm run dev
```
Starts the local development server on [http://localhost:3000](http://localhost:3000).

### Type Checking & Linting
```bash
npm run lint
```
Runs `tsc --noEmit` across all TypeScript modules.

### Production Build
```bash
npm run build
```
Generates production assets in the `dist/` directory.

### Preview Production Build
```bash
npm run preview
```

---

## ✨ Features

- **Macro & VBA Project Extraction**: Parses `vbaProject.bin` from `.xlsm` archives, displaying modules, procedures, triggers, line counts, and modernized TypeScript equivalents.
- **Automated Sheet & Workbook Unprotection**: Automatically identifies and strips legacy XOR and SHA-512 sheet password locks and workbook protection tags.
- **VBA Simulation & Runtime Mock**: Simulates VBA macro execution in the browser with variable tracking, syntax validation, and line-by-line runtime diagnostics.
- **Customizable Dashboard**: Drag-and-drop widget layout with direct reordering controls, adjustable card widths (`1/3`, `1/2`, `2/3`, `Full`), and persistent localStorage state across sessions.
- **RFC-4180 CSV Export**: One-click download of processed spreadsheet data with UTF-8 BOM encoding for seamless Microsoft Excel and Google Sheets compatibility.
- **Dynamic Pivot Matrix**: Multi-dimensional pivot table with custom aggregations (Sum, Avg, Min, Max, Count).

---

## 📄 License

Apache-2.0
