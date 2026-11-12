# PMRS -> Excel v1.0

Chrome / Edge (Manifest V3) extension that exports the PMRS table you are
looking at to a CSV file that opens directly in Excel.

Works only on `https://pmrs.digitalsp.gov.kh/*`.

## Folder structure

```
extension/
├── manifest.json      MV3 manifest (name, permissions, content script, worker)
├── popup.html         popup markup (header, info box, checkboxes, buttons)
├── popup.css          dark-green theme
├── popup.js           popup UI + start/stop + progress rendering
├── background.js      service worker: auto-sync pagination loop (survives reloads)
├── content.js         table parser, pagination helpers, sync bridge, XLSX builder
├── README.md
└── icons/
    ├── icon16.png
    ├── icon32.png
    ├── icon48.png
    └── icon128.png
```

## Install (unpacked)

1. Open `edge://extensions` (Edge) or `chrome://extensions` (Chrome).
2. Turn on **Developer mode**.
3. Click **Load unpacked** and select this `extension` folder.
4. Pin **PMRS -> Excel v1.0** to the toolbar.
5. Open the PMRS rejected/approved list, then click the extension icon.
   (If the popup says "Reload", refresh the PMRS tab once so the content
   script is injected.)

> A previously generated `extension.crx` in the project root is an old build.
> Re-pack it or keep using the unpacked folder above.

## Use

- **Info box** shows the detected table size, e.g. `10 rows x 18 columns`.
- **Include header row** – adds the site's header row as the first CSV row.
- **Go through all pages** – when checked, the primary button walks every page.
- **Download all pages** – collects page 1, clicks the next-page arrow
  (`>>`, `›`, `Next`, or PMRS's `aria-label="Next page"` button), waits for the
  table to change, repeats until the last page, then downloads one CSV.
  If the checkbox is off it simply exports the visible page.
- **Save Rejections to Logs** (`btn-sync-reject`) – reads the rows on the
  current page, extracts the health-card number (`បណ្ណសុខភាព`), submission
  status (`ស្ថានភាពបញ្ជូន`) and reject reason (`ប្រភេទកំហុស`), then POSTs
  `{"sync_type": "rejection", "records": [...]}` to
  `http://localhost:8000/api/save_pmrs_records`. The backend UPSERTs them into
  `npca_validation_logs`. Khmer statuses are converted to `REJECTED`,
  `APPROVED`, `SUBMITTED`, `SAVED` or `PENDING`.
- **Save Approved (Gold) to DB** (`btn-sync-approve`) – scrapes the current
  page and uses the approved list's `លេខបណ្ណ` (card number) as `pt_code`,
  plus patient name, gender and age. The approved list has no health-card
  column, so its identifier is used directly. It POSTs
  `{"sync_type": "approved", "records": [...]}` and the backend UPSERTs them
  into `pmrs_approved_gold` by `pt_code`.
- **Deep scrape detail modals** (`deepScrape`) – when ticked, the content
  script clicks the blue eye button on each row, waits for the detail modal
  using `MutationObserver` + polling, extracts diagnoses, vitals, HPI and the
  labelled modal fields, closes the modal, then moves to the next row. The
  result is sent as `detailed_clinical_payload` and stored as JSON text. The
  modal open/close and view-button selectors live in `DEEP_SCRAPE` at the top
  of `content.js`; add exact inner-element selectors under
  `DEEP_SCRAPE.selectors` to override the generic label/value parser.
  Deep scraping is slow (one modal per row), so use `rows per page = 10` for
  testing and keep the popup open for this-page-only runs.
  When **Go through all pages** is ticked, `background.js` drives either sync:
  scrape page → POST → click Next → wait for the reload → repeat. The content
  script only handles one page at a time, so a full portal reload can no longer
  break the popup's message port. While it runs, the active button becomes
  **Stop sync**; progress ("Saving page 2…", "Finished syncing 5 pages…")
  appears below the buttons.
- **This page only** – exports exactly what is on screen.
- **Diagnose page** – prints a JSON report (detected tables, selected table,
  headers, pagination candidates and their scores) into the popup and the
  DevTools console. Use it if pagination is not found on a new screen.
- **Minimize** (dash icon) collapses the popup body; click again to restore.

Files are named like `pmrs-rejected-list-20261025-1534.csv`.

## How it works

- `content.js` scores every visible `<table>` by column count (weighted),
  body rows, and presence of `<thead>`, then picks the winner. PMRS renders
  multiple `<thead>` rows, so the header row with the most text is used.
- Cells keep embedded line breaks (e.g. `19/09/1954\n72 ឆ្នាំ`); the CSV
  writer quotes them so Excel shows two lines in one cell.
- CSV is RFC 4180 escaped and prefixed with a UTF-8 BOM (`\uFEFF`) so Khmer
  text opens correctly in Excel. Line endings are CRLF.
- Pagination detection tries `aria-label="Next page"`, `[data-slot=next]`,
  DataTables `.paginate_button.next`, `.pagination .next`, plus the literal
  glyphs `>>`, `»`, `›`, `>` and the word `Next`. Disabled state is read from
  the `disabled` attribute, `aria-disabled`, and exact `disabled` class
  tokens — Tailwind's `disabled:` utility classes are ignored.
- After clicking Next it polls the table until the row signature changes and
  stays stable, then collects. Duplicate page signatures are skipped, and a
  300-page safety limit prevents infinite loops.
- All exports happen inside the page, so closing the popup mid-run does not
  cancel the download.

## Tested

Verified against the live PMRS rejected-claims list (18 columns):

- popup info: `10 rows x 18 columns`, `Page 1 · more pages available`
- "This page only": 10 data rows + header, Khmer headers and values intact
- "Download all pages": page 1 + page 2 = 20 rows, `pages: 2`, progress
  messages emitted per page, single CSV downloaded
- diagnose: finds `Next page` button (score 160, enabled) and the 18 headers
