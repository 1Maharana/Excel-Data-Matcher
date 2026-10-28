# Excel Data Matcher

A standalone Next.js application for comparing Excel/CSV files, filtering records, finding duplicate candidates, reviewing differences, and exporting results.

## Requirements

- Node.js 20.9 or newer (recommended)
- npm
- Windows + VS Code works fine

## Install and run

Open a terminal in this project folder:

```powershell
npm install
npm run dev
```

Open http://localhost:3000.

To check a production build:

```powershell
npm run build
npm start
```

## Features

- Import multiple `.xlsx`, `.xls`, and `.csv` files (first worksheet for Excel files)
- Consolidated record table with source file and original row number
- Search and filter by source and missing fields
- Common header normalization and editable column mapping
- Cross-file duplicate candidate matching using indexed identifiers and fuzzy text similarity
- Background Web Worker for matching, keeping the interface responsive during large comparisons
- Paginated record tables (100 rows per page) and capped on-screen match cards for large result sets
- Adjustable matching threshold
- Side-by-side comparison with field differences highlighted
- Manual match/not-a-match review decisions during the current session
- Separate views for missing information and unmatched records
- Export filtered records to Excel/CSV and all match candidates plus review decisions to Excel
- No database, account, server-side upload, or external matching API

## Privacy and limitations

File contents are processed in the browser and are not sent to a matching service by this app. Do not deploy publicly with sensitive data without reviewing your hosting, access, and privacy requirements.

- Matching uses indexed candidate blocking rather than comparing every possible pair. To prevent common names/tokens from causing quadratic memory use, very large index buckets are skipped and candidate generation is capped at 100,000 pairs. This improves stability but can miss matches in extremely common buckets or datasets that exceed the cap; validate results before making identity decisions.
- The app is designed to handle larger datasets more safely, but 100,000-row performance depends on the number of files, columns, browser memory, and candidate density. Test with representative files before relying on it for production decisions.
- Match review displays at most 200 cards at a time for responsiveness; use search/filter or Export matches to inspect the full report.
- Only the first worksheet in each Excel file is imported.
- Matching currently compares records across different source files, not duplicates within the same file.
- Review decisions are held in React memory and reset when the page is refreshed. Export match review before closing.
- Scores are heuristic similarity scores, not probabilities or proof of identity.
- Excel files with unusual formatting, merged header rows, or multiple header lines may need preprocessing.

## Main files

- `src/app/page.tsx` — UI, imports, filters, review, exports
- `src/app/globals.css` — global styles
- `src/lib/matcher.ts` — column normalization and match scoring

This project is independent from any other application and contains no database configuration.


## Union Master export

Use **Export Union Master** to create one consolidated workbook from all imported records. Only pairs explicitly marked **Confirmed same person** in Match Review are merged; pending or uncertain candidates remain separate to avoid losing distinct alumni. The export records source files and row numbers, fills blank fields from confirmed duplicates, and adds conflict details when confirmed records disagree on a field.

## Duplicate detection improvements

- Candidate generation checks records both **across files** and **within the same file**, so repeated entries inside one workbook can appear in Match Review.
- Comparison-only normalization handles common degree abbreviations such as `M.Sc. (Engineering)` and `M.Sc. (Engg)` without rewriting the source values.
- Candidate generation uses roll number, other identifiers, normalized full names, name prefixes, and shared name tokens.
- Match Review labels candidates as cross-file or within-file and supports manual confirmation before Union Master merges them.
- Candidate generation is bounded (20,000 pairs per index bucket and 250,000 candidate pairs total) to protect browser memory. Very large common-name buckets or datasets exceeding these limits can still omit candidates; always validate on a known sample.

## Compare using the most complete submission per person

The import panel includes an optional toggle: **Compare using the most complete submission per person, per file**.

When enabled, records within the same source file are grouped when their normalized name, graduation year, and department match. The row with the most populated fields is used as the representative for matching against other files, and blank fields in that representative are filled from the other submissions. Original imported rows remain available in the All Records view. The Union Master export consolidates those same-file groups and records source rows/field conflicts.

This is a heuristic, not proof of identity. Two different alumni can share a name, graduation year, and department, so review groups and conflicting identifiers before relying on a merged export. Turn the toggle off to compare every original row independently.

## Most-complete-per-file toggle behavior

When **Compare using the most complete submission per person, per file** is enabled, the app now uses the per-file representative dataset throughout the workspace: All Records, Missing Information, Unmatched, matching candidates, filtered Excel/CSV exports, unmatched exports, and Union Master. Likely repeated submissions are represented by the row with the most populated fields, with blank fields filled from other rows in that same group. Filtered exports include the representative row plus submission count and source-row provenance. Union Master chooses the most complete row as the primary version and preserves conflicting non-empty values in conflict columns.

Turn the toggle off to inspect and export original rows individually. Grouping is heuristic and depends on name, graduation year, and department; review likely duplicates and conflicts before relying on a consolidated export.
