# VAT Forensic Workbench

Deployable Vite + React + TypeScript app for workbook-driven VAT review, routing, reconciliation, and release gating.

## Features

- primary and optional secondary workbook upload
- sheet/header/column auto-detection with manual override
- workbook-drift handling with graceful fallbacks
- routing buckets for report / defer / exclude / manual review
- summary bucket detection from shifting control blocks
- detail-to-summary reconciliation
- revenue vs purchase code parity checks
- audit workbook export and release workbook export
- browser local storage for last-used mappings and settings

## Local run

```bash
npm install
npm run dev
```

## Build

```bash
npm install
npm run build
npm run start
```

## Railway

- push to GitHub
- create a Railway project from the repo
- Railway will use `railway.toml`
- build command: `npm run build`
- start command: `npm run start`

No API is required for the current version. Files are processed in the browser. For shared approvals, saved audit history, or centralized config, add an API later.
