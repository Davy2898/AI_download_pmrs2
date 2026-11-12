/* ────────────────────────────────────────────────────────────────────────────
   PMRS -> Excel v1.0 — content script
   Runs only on https://pmrs.digitalsp.gov.kh/*

   Exports the table currently on screen to an Excel .xlsx file, or walks the
   pagination (next-page arrow ">>" / Next / ›) and downloads every page as one
   .xlsx file (saved by popup.js into Downloads/Report_download).

   Messages
     PMRS_GET_INFO     -> { ok, rows, columns, page, pageCount, hasNext }
     PMRS_EXPORT_PAGE  -> { ok, rows, file, ... }        (current page only)
     PMRS_EXPORT_ALL   -> { ok, rows, pages, file, ... } (walk all pages)
     PMRS_STOP_DOWNLOAD -> { ok: true }  (flag the running loop to stop)
     PMRS_COLLECT_RECORDS -> { ok, records: [{pt_code, status}], rows, page }
     PMRS_DIAGNOSE     -> { ok, selected, next, candidates, ... }

   One-page auto-sync bridge used by background.js
     PMRS_PING         -> { ok, build }
     PMRS_SYNC_SCRAPE  -> { ok, records, page, hasNext, signature }
     PMRS_SYNC_STATE   -> { ok, page, hasNext, signature }  (polling)
     PMRS_SYNC_NEXT    -> { ok, clicked }  (reply first, then click Next)
     PMRS_SYNC_STOP    -> { ok: true }
   ──────────────────────────────────────────────────────────────────────────── */
(() => {
  "use strict";

  if (window.__PMRS_EXCEL_V1__) return;
  window.__PMRS_EXCEL_V1__ = true;

  const BUILD = "PMRS -> Excel v1.0";
  const MAX_PAGES = 300;
  const NEXT_TIMEOUT_MS = 25000;
  const POLL_MS = 300;

  /* Deep-scrape configuration. The `selectors` block is intentionally empty:
     when the exact modal elements are known, add CSS selectors here and they
     take priority over the generic label/value fallback. */
  const DEEP_SCRAPE = {
    modalSelectors: [
      '[role="dialog"][aria-modal="true"]',
      '[aria-modal="true"]',
      '[role="dialog"]',
      ".modal.show",
      ".modal.in",
    ],
    viewButtonSelectors: [
      'button:has(.i-solar\\:eye-outline)',
      'button:has([class*="eye"])',
      'button[aria-label*="view" i]',
      'button[title*="view" i]',
      'a[aria-label*="view" i]',
    ],
    closeButtonSelectors: [
      '[role="dialog"] button[aria-label*="close" i]',
      '[role="dialog"] button[title*="close" i]',
      '.modal button[aria-label*="close" i]',
      '.modal button[title*="close" i]',
    ],
    selectors: {
      primaryDiagnosis: [],
      differentialDiagnosis: [],
      chiefComplaint: [],
      temperature: [],
      weight: [],
      bloodPressure: [],
      historyOfPresentIllness: [],
    },
    openTimeoutMs: 15000,
    closeTimeoutMs: 6000,
    pollMs: 120,
    settleMs: 200,
    rowDelayMs: 100,
    maxRows: 200,
    maxRawText: 120000,
  };

  /* Column names used by the PMRS claims table (Khmer + English fallbacks). */
  const PT_CODE_HEADERS = [
    "បណ្ណសុខភាព",
    "លេខបណ្ណសុខភាព",
    "លេខសម្គាល់បណ្ណសុខភាព",
    /* Approved list: the only identifier column is "លេខបណ្ណ" (card number). */
    "លេខបណ្ណ",
    "health card",
    "health card id",
    "health card no",
    "health card number",
    "health_card_id",
    "pt_code",
    "pt code",
    "pid",
    "patient code",
    "patient id",
    "card number",
    "card no",
  ];
  const PMRS_STATUS_HEADERS = [
    "ស្ថានភាពបញ្ជូន",
    "ស្ថានភាព",
    "submission status",
    "claim status",
    "pmrs_status",
    "status",
  ];
  /* Reject reason column, shown next to the status in the rejected list.
     The PMRS 2.0 header is "ប្រភេទកំហុស" (error type). */
  const PMRS_REJECT_REASON_HEADERS = [
    "ប្រភេទកំហុស",
    "មូលហេតុបដិសេធ",
    "ហេតុផលបដិសេធ",
    "មូលហេតុ",
    "reject reason",
    "rejection reason",
    "rejected reason",
    "error type",
    "error_type",
    "rejection_reason",
    "reject_reason",
    "pmrs_reject_reason",
    "reason",
  ];

  /* Approved (gold-standard) columns. The PMRS approved list exposes patient
     and visit details next to the health-card number; missing columns are
     simply sent as empty strings and the backend keeps the old value. */
  const PATIENT_NAME_HEADERS = [
    "ឈ្មោះអ្នកជំងឺ",
    "ឈ្មោះ",
    "patient name",
    "patient_name",
    "full name",
    "fullname",
  ];
  const GENDER_HEADERS = ["ភេទ", "gender", "sex"];
  const AGE_HEADERS = ["អាយុ", "age", "age_years"];
  const DOB_HEADERS = [
    "ថ្ងៃខែឆ្នាំកំណើត",
    "date of birth",
    "dob",
    "birth date",
  ];
  const VISIT_TYPE_HEADERS = [
    "ប្រភេទនៃការចូល",
    "ប្រភេទការចូល",
    "ប្រភេទនៃការមកពិនិត្យ",
    "visit type",
    "visit_type",
    "type of visit",
    "admission type",
  ];
  const ICD10_DIAGNOSIS_HEADERS = [
    "icd10_diagnosis",
    "icd10",
    "icd-10",
    "icd_10",
    "icd 10",
    "រោគវិនិច្ឆ័យចេញ",
    "រោគវិនិច្ឆ័យចូល",
    "រោគវិនិច្ឆ័យ",
    "diagnosis",
  ];
  const APPROVAL_DATE_HEADERS = [
    "approval_date",
    "approval date",
    "approved date",
    "approved at",
    "approved on",
    "approval_dt",
    "ថ្ងៃអនុម័ត",
    "កាលបរិច្ឆេទអនុម័ត",
  ];
  const PMRS_STATUS_TOKENS = [
    ["បដិសេធ", "REJECTED"],
    ["reject", "REJECTED"],
    ["អនុម័ត", "APPROVED"],
    ["approv", "APPROVED"],
    ["បញ្ជូន", "SUBMITTED"],
    ["submit", "SUBMITTED"],
    ["រក្សាទុក", "SAVED"],
    ["saved", "SAVED"],
    ["រង់ចាំ", "PENDING"],
    ["pending", "PENDING"],
  ];

  /* Set true by PMRS_STOP_DOWNLOAD; the pagination loop breaks on it. */
  let stopRequested = false;

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  /* ══════════════════════════════════════════════════════════════════════
     DOM / TABLE HELPERS
     ══════════════════════════════════════════════════════════════════════ */
  function clean(value) {
    return String(value == null ? "" : value)
      .replace(/\u200b/g, "")
      .replace(/[ \t\u00a0]+/g, " ")
      .trim();
  }

  function isVisible(node) {
    if (!node || !node.isConnected) return false;
    const style = getComputedStyle(node);
    if (style.display === "none" || style.visibility === "hidden" || style.opacity === "0") {
      return false;
    }
    const rect = node.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  /** Cell text with line breaks preserved (Excel handles them inside quotes). */
  function cellText(cell) {
    let raw = cell.innerText;
    if (!raw) raw = cell.textContent || "";
    return String(raw)
      .replace(/\u200b/g, "")
      .replace(/\r\n?/g, "\n")
      .split("\n")
      .map((line) => line.replace(/[ \t\u00a0]+/g, " ").trim())
      .filter(Boolean)
      .join("\n");
  }

  function singleLine(value) {
    return clean(String(value || "").replace(/\s*\n+\s*/g, " "));
  }

  /** Read a <tr> including colspan padding so columns stay aligned. */
  function rowCells(row) {
    const out = [];
    for (const cell of Array.from(row.cells)) {
      out.push(cellText(cell));
      const span = Math.min(Math.max(parseInt(cell.getAttribute("colspan") || "1", 10), 1), 50);
      for (let i = 1; i < span; i += 1) out.push("");
    }
    return out;
  }

  /**
   * Pick the main data table: visible, at least 2 columns, most columns wins
   * (PMRS has 18), with a small bonus for tables that have a real <thead>.
   */
  function pickTable() {
    let best = null;
    let bestScore = -1;

    for (const table of Array.from(document.querySelectorAll("table"))) {
      if (!isVisible(table)) continue;

      const rows = Array.from(table.rows);
      if (rows.length < 1) continue;

      let columns = 0;
      for (const row of rows.slice(0, 5)) columns = Math.max(columns, row.cells.length);
      if (columns < 2) continue;

      const bodyRows = rows.filter((row) => !row.closest("thead")).length;
      if (bodyRows < 1) continue;

      const headerBonus = table.tHead && table.tHead.rows.length ? 40 : 0;
      const score = columns * 100 + bodyRows * 2 + headerBonus;
      if (score > bestScore) {
        best = table;
        bestScore = score;
      }
    }
    return best;
  }

  function readTable(table) {
    const rows = Array.from(table.rows);
    let headerRow = null;
    let headers = [];

    if (table.tHead && table.tHead.rows.length) {
      // PMRS renders more than one <thead> row; pick the one with the most
      // non-empty text instead of blindly using the last one.
      let bestRow = null;
      let bestCount = 0;
      for (const candidate of Array.from(table.tHead.rows)) {
        const cells = rowCells(candidate).map(singleLine);
        const filled = cells.filter(Boolean).length;
        if (filled > bestCount) {
          bestCount = filled;
          bestRow = candidate;
        }
      }
      if (bestRow && bestCount > 0) {
        headerRow = bestRow;
        headers = rowCells(bestRow).map(singleLine);
      }
    }

    if (!headers.some(Boolean) && rows.length) {
      const firstRow = rows[0];
      if (Array.from(firstRow.cells).every((c) => c.tagName === "TH")) {
        headerRow = firstRow;
        headers = rowCells(firstRow).map(singleLine);
      }
    }

    const bodyRows = [];
    for (const row of rows) {
      if (row === headerRow || row.closest("thead")) continue;
      if (!isVisible(row)) continue;
      const cells = rowCells(row);
      if (!cells.some((value) => value)) continue;
      bodyRows.push(cells);
    }

    const width = Math.max(headers.length, ...bodyRows.map((r) => r.length), 0);
    const pad = (source) => {
      const out = source.slice(0, width);
      while (out.length < width) out.push("");
      return out;
    };

    return {
      headers: pad(headers),
      rows: bodyRows.map(pad),
      columns: width,
    };
  }

  /* ── rows for the NPCA log (pt_code + PMRS status) ───────────────────── */
  function normaliseHeader(value) {
    return singleLine(value)
      .toLowerCase()
      .replace(/[_-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function findHeaderIndex(headers, aliases) {
    const keys = headers.map(normaliseHeader);

    for (const alias of aliases) {
      const needle = normaliseHeader(alias);
      if (!needle) continue;
      const exact = keys.indexOf(needle);
      if (exact !== -1) return exact;
    }

    for (const alias of aliases) {
      const needle = normaliseHeader(alias);
      if (needle.length < 4) continue;
      // A short header such as "ID" or "No" must not match an alias such
      // as "health card id"; that can silently select the row-number column.
      const partial = keys.findIndex((key) => key.length >= 4 && key.includes(needle));
      if (partial !== -1) return partial;
    }
    return -1;
  }

  function isSequentialRowNumberColumn(values) {
    const numbers = values.map((value) => singleLine(value)).filter(Boolean);
    if (numbers.length < 2 || numbers.some((value) => !/^\d+$/.test(value))) return false;

    const parsed = numbers.map(Number);
    const startsAtRowNumber = parsed[0] === 0 || parsed[0] === 1;
    return (
      startsAtRowNumber &&
      parsed.every((value, index) => index === 0 || value === parsed[index - 1] + 1)
    );
  }

  /** "បានបដិសេធ" -> REJECTED, "approved" -> APPROVED, unknown labels kept. */
  function normalisePmrsStatus(value) {
    const text = singleLine(value);
    if (!text) return "";

    const lowered = text.toLowerCase();
    for (const [token, code] of PMRS_STATUS_TOKENS) {
      if (lowered.includes(token)) return code;
    }

    if (/^[\x20-\x7e]+$/.test(text)) {
      return text.replace(/\s+/g, "_").toUpperCase().slice(0, 64);
    }
    return text.slice(0, 64);
  }

  /** Reject reason keeps its Khmer text; only whitespace is normalised. */
  function normalisePmrsRejectReason(value) {
    return singleLine(value).slice(0, 512);
  }

  function valueAt(row, index) {
    return index >= 0 ? singleLine(row[index]) : "";
  }

  /** PMRS often shows "1954-09-19 (72 ឆ្នាំ 1 ខែ)" in the DOB column. */
  function ageFromDob(value) {
    const text = singleLine(value);
    const khmer = text.match(/(\d{1,3})\s*ឆ្នាំ/);
    if (khmer) return khmer[1];
    const english = text.match(/(\d{1,3})\s*(?:years?|yrs?|y\/o)\b/i);
    return english ? english[1] : "";
  }

  /** Approved list: age cell is "13/04/1961\n65 ឆ្នាំ"; keep "65".
      Rejected list: age cell is "1 ឆ្នាំ 6 ខែ"; keep it unchanged. */
  function patientAge(value, dobValue) {
    const text = singleLine(value);
    const startsWithDate = /^\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}\b/.test(text);
    if (text && !startsWithDate) return text;
    return ageFromDob(text) || text || ageFromDob(dobValue);
  }

  function recordsFromTable(info) {
    const ptColumn = findHeaderIndex(info.headers, PT_CODE_HEADERS);
    if (ptColumn < 0) {
      throw new Error(
        "Could not find the patient-card column (បណ្ណសុខភាព / លេខបណ្ណ / Health card)."
      );
    }

    const statusColumn = findHeaderIndex(info.headers, PMRS_STATUS_HEADERS);
    const reasonColumn = findHeaderIndex(info.headers, PMRS_REJECT_REASON_HEADERS);

    /* Approved-list columns. Every record carries them, so the popup can send
       the same scrape to either table depending on the button pressed. */
    const nameColumn = findHeaderIndex(info.headers, PATIENT_NAME_HEADERS);
    const genderColumn = findHeaderIndex(info.headers, GENDER_HEADERS);
    const ageColumn = findHeaderIndex(info.headers, AGE_HEADERS);
    const dobColumn = findHeaderIndex(info.headers, DOB_HEADERS);
    const visitTypeColumn = findHeaderIndex(info.headers, VISIT_TYPE_HEADERS);
    const icd10Column = findHeaderIndex(info.headers, ICD10_DIAGNOSIS_HEADERS);
    const approvalDateColumn = findHeaderIndex(info.headers, APPROVAL_DATE_HEADERS);

    const fallbackStatus = /\/rejected\//i.test(location.pathname)
      ? "REJECTED"
      : "UNKNOWN";

    const cardValues = info.rows.map((row) => singleLine(row[ptColumn]).replace(/\s+/g, ""));
    if (isSequentialRowNumberColumn(cardValues)) {
      throw new Error(
        `The detected patient-card column ("${info.headers[ptColumn] || ptColumn + 1}") ` +
          "contains sequential row numbers. No records were sent; check the PMRS table headers and reload the page."
      );
    }

    const byCode = new Map();
    for (const row of info.rows) {
      const ptCode = singleLine(row[ptColumn])
        .replace(/\s+/g, "")
        .toUpperCase()
        .slice(0, 64);
      if (!ptCode) continue;

      const status =
        normalisePmrsStatus(statusColumn >= 0 ? row[statusColumn] : "") ||
        fallbackStatus;
      const rawReason = reasonColumn >= 0 ? row[reasonColumn] : "";
      const rejectReason =
        status === "REJECTED" ? normalisePmrsRejectReason(rawReason) : "";

      const rawAge = valueAt(row, ageColumn);
      byCode.set(ptCode, {
        pt_code: ptCode,
        status,
        pmrs_reject_reason: rejectReason,
        patient_name: valueAt(row, nameColumn),
        gender: valueAt(row, genderColumn),
        age: patientAge(rawAge, valueAt(row, dobColumn)),
        visit_type: valueAt(row, visitTypeColumn),
        icd10_diagnosis: valueAt(row, icd10Column),
        approval_date: valueAt(row, approvalDateColumn),
      });
    }
    return Array.from(byCode.values());
  }

  function collectCurrentPageRecords() {
    const table = pickTable();
    if (!table) throw new Error("No table with data was found on this page.");

    const info = readTable(table);
    if (!info.rows.length) throw new Error("The table on this page has no rows.");

    const records = recordsFromTable(info);
    if (!records.length) {
      throw new Error("No health-card numbers were found in the table.");
    }

    const meta = readPageNumber();
    return {
      ok: true,
      records,
      rows: records.length,
      page: meta.page,
      columns: info.columns,
    };
  }

  function readPageNumber() {
    let page = NaN;
    const current = document.querySelector('[aria-current="page"], [data-selected="true"]');
    if (current) page = parseInt(clean(current.innerText).replace(/\D+/g, ""), 10);

    if (!Number.isFinite(page)) {
      const fromUrl = new URLSearchParams(location.search).get("page");
      page = parseInt(fromUrl || "", 10);
    }
    if (!Number.isFinite(page) || page < 1) page = 1;

    const numbers = Array.from(
      document.querySelectorAll('[data-type="page"], [aria-label^="Page "]')
    )
      .map((node) => {
        const aria = node.getAttribute("aria-label") || "";
        const match = aria.match(/Page\s+(\d+)/i);
        if (match) return Number(match[1]);
        const digits = clean(node.innerText).replace(/\D+/g, "");
        return digits ? Number(digits) : NaN;
      })
      .filter((n) => Number.isFinite(n) && n > 0);

    return { page, pageCount: numbers.length ? Math.max(...numbers) : 0 };
  }

  function rowsSignature(info) {
    const rows = info.rows;
    const first = rows.length ? rows[0].join("|") : "";
    const last = rows.length ? rows[rows.length - 1].join("|") : "";
    return `${rows.length}|${first.slice(0, 200)}|${last.slice(0, 200)}`;
  }

  function fullSignature(info, meta) {
    return `${meta.page}|${rowsSignature(info)}`;
  }

  /* ══════════════════════════════════════════════════════════════════════
     PAGINATION
     ══════════════════════════════════════════════════════════════════════ */
  const NEXT_SELECTORS = [
    'button[aria-label="Next page"]',
    'button[aria-label*="next" i]',
    'a[aria-label*="next" i]',
    '[data-slot="next"]',
    'button[data-slot="next"]',
    'li.paginate_button.next > a',
    'a.paginate_button.next',
    '.pagination li.next a',
    '.pagination a.next',
    '[class*="pagination" i] [class*="next" i]',
  ];

  const NEXT_TEXTS = new Set([">>", "»", "›", ">", "next", "next page", "បន្ទាប់"]);

  const DISABLED_TOKEN = /(^|\s)(is-)?disabled($|\s)/i;

  function isDisabled(node) {
    if (!node) return true;
    const target = node.matches("li") ? node.querySelector("a, button, [role=button]") || node : node;

    if (target.disabled) return true;
    if ((target.getAttribute("aria-disabled") || "").toLowerCase() === "true") return true;
    if (DISABLED_TOKEN.test(target.className || "")) return true;
    if (target.parentElement && DISABLED_TOKEN.test(target.parentElement.className || "")) {
      return true;
    }

    const style = getComputedStyle(target);
    if (style.pointerEvents === "none" || style.display === "none" || style.visibility === "hidden") {
      return true;
    }
    return false;
  }

  function nextCandidates() {
    const seen = new Set();
    const out = [];

    const add = (node, baseScore, why) => {
      if (!node || seen.has(node)) return;
      seen.add(node);
      if (!isVisible(node)) return;

      const blob = [
        node.getAttribute("aria-label"),
        node.getAttribute("title"),
        node.getAttribute("data-slot"),
        node.className,
        node.id,
      ]
        .join(" ")
        .toString();

      let score = baseScore;
      if (/next/i.test(blob)) score += 60;
      if (node.closest('[class*="pagination" i], [class*="pager" i]')) score += 30;
      if (node.closest('[class*="breadcrumb" i], [aria-label*="breadcrumb" i]')) score -= 120;
      if (isDisabled(node)) score -= 200;

      out.push({
        el: node,
        score,
        why,
        label: clean(node.getAttribute("aria-label") || node.innerText || node.title || why),
        disabled: isDisabled(node),
      });
    };

    NEXT_SELECTORS.forEach((selector) => {
      document.querySelectorAll(selector).forEach((node) => add(node, 100, selector));
    });

    document.querySelectorAll('button, a, [role="button"], li').forEach((node) => {
      const text = clean(node.innerText).toLowerCase();
      if (!text || text.length > 24) return;
      if (NEXT_TEXTS.has(text)) add(node, 70, `text:${text}`);
    });

    out.sort((a, b) => b.score - a.score);
    return out;
  }

  function findNext() {
    const candidates = nextCandidates();
    const usable = candidates.find((candidate) => candidate.score > 0 && !candidate.disabled);
    return usable || null;
  }

  async function clickNext(candidate) {
    const node = candidate.el;
    const target = node.matches("li")
      ? node.querySelector("a, button, [role=button]") || node
      : node;

    try {
      target.scrollIntoView({ block: "center", inline: "nearest" });
    } catch (_) {
      /* ignore */
    }
    try {
      target.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
    } catch (_) {
      /* ignore */
    }

    target.click();
    await sleep(200);
  }

  async function waitForRowsChange(previousSignature, timeoutMs = NEXT_TIMEOUT_MS) {
    const start = Date.now();
    let changedSignature = null;
    let stablePolls = 0;

    while (Date.now() - start < timeoutMs) {
      if (stopRequested) return false;
      await sleep(POLL_MS);

      const table = pickTable();
      if (!table) continue;

      const info = readTable(table);
      if (!info.rows.length) continue;

      const signature = rowsSignature(info);
      if (signature === previousSignature) continue;

      if (signature !== changedSignature) {
        changedSignature = signature;
        stablePolls = 1;
        continue;
      }

      stablePolls += 1;
      if (stablePolls >= 2) {
        await sleep(250);
        return true;
      }
    }
    return false;
  }

  /* ══════════════════════════════════════════════════════════════════════
     XLSX EXPORT
     An .xlsx file is a ZIP of XML parts. Chrome's Excel accepts an
     uncompressed ("stored") zip, so we build one from scratch — no library.
     The bytes are returned to popup.js as base64; popup.js saves the file
     into the Report_download folder via chrome.downloads.download().
     ══════════════════════════════════════════════════════════════════════ */
  function pad2(number) {
    return String(number).padStart(2, "0");
  }

  /** Current date as YYYY-MM-DD, e.g. 2026-10-25. */
  function reportDate() {
    const now = new Date();
    return `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())}`;
  }

  function reportFilename() {
    // Dynamic filename with the current date, e.g. PMRS_Report_2026-10-25.xlsx
    return `PMRS_Report_${reportDate()}.xlsx`;
  }

  function xmlEscape(value) {
    return String(value == null ? "" : value)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "") // invalid in XML
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  const PLAIN_NUMBER = /^[+-]?\d+(\.\d{1,10})?$/;

  /** Numeric cells as numbers; everything else as inline strings. */
  function sheetCell(value) {
    const text = String(value == null ? "" : value);
    if (text && text.length <= 20 && PLAIN_NUMBER.test(text)) {
      return `<c t="n"><v>${text}</v></c>`;
    }
    return `<c t="inlineStr"><is><t xml:space="preserve">${xmlEscape(text)}</t></is></c>`;
  }

  function buildSheetXml(headers, rows) {
    const parts = [];
    parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    parts.push(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        "<sheetData>"
    );

    // Columns in Excel go A, B, ... AA ... — the sheet uses A1 refs.
    const colLetters = (n) => {
      let out = "";
      let x = n;
      while (x >= 0) {
        out = String.fromCharCode(65 + (x % 26)) + out;
        x = Math.floor(x / 26) - 1;
      }
      return out;
    };

    let rowNumber = 0;
    const emitRow = (cells) => {
      parts.push(`<row r="${rowNumber + 1}">`);
      cells.forEach((value, col) => {
        if (value === "" || value == null) return;
        parts.push(sheetCell(value).replace("<c ", `<c r="${colLetters(col)}${rowNumber + 1}" `));
      });
      parts.push("</row>");
      rowNumber += 1;
    };

    if (headers && headers.some(Boolean)) emitRow(headers);
    rows.forEach((row) => emitRow(row));

    parts.push("</sheetData></worksheet>");
    return parts.join("");
  }

  /* ── minimal STORED-ENTRY zip writer ─────────────────────────────── */
  const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n += 1) {
      let c = n;
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i += 1) {
      crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function zipParts(chunk, push) {
    push(Uint8Array.of(chunk & 0xff, (chunk >>> 8) & 0xff));
  }
  function zipDword(chunk, push) {
    push(
      Uint8Array.of(chunk & 0xff, (chunk >>> 8) & 0xff, (chunk >>> 16) & 0xff, (chunk >>> 24) & 0xff)
    );
  }

  /** entries: Array<{ name: string, data: Uint8Array }> → zip as Blob */
  function makeZipBlob(entries) {
    const chunks = [];
    let offset = 0;
    const central = [];

    const pushByteChunk = (bytes) => {
      chunks.push(bytes);
      offset += bytes.length;
    };

    const encoder = new TextEncoder();
    const fileTime = Uint8Array.of(0x00, 0x00); // 00:00
    const fileDate = Uint8Array.of(0x21, 0x51); // dos date stub (not important)

    for (const entry of entries) {
      const crc = crc32(entry.data);
      const nameBytes = encoder.encode(entry.name);
      const localOffset = offset;

      // local file header
      pushByteChunk(Uint8Array.of(0x50, 0x4b, 0x03, 0x04));
      pushByteChunk(Uint8Array.of(20, 0)); // version needed
      pushByteChunk(Uint8Array.of(0, 0)); // flags
      pushByteChunk(Uint8Array.of(0, 0)); // method: stored
      pushByteChunk(fileTime);
      pushByteChunk(fileDate);
      zipDword(crc, pushByteChunk);
      zipDword(entry.data.length, pushByteChunk); // compressed size
      zipDword(entry.data.length, pushByteChunk); // uncompressed size
      zipParts(nameBytes.length, pushByteChunk);
      zipParts(0, pushByteChunk); // extra length
      pushByteChunk(nameBytes);
      pushByteChunk(entry.data);

      central.push({ crc, nameBytes, localOffset, size: entry.data.length });
    }

    const centralOffset = offset;
    for (const item of central) {
      pushByteChunk(Uint8Array.of(0x50, 0x4b, 0x01, 0x02));
      pushByteChunk(Uint8Array.of(20, 0)); // version made by
      pushByteChunk(Uint8Array.of(20, 0)); // version needed
      pushByteChunk(Uint8Array.of(0, 0)); // flags
      pushByteChunk(Uint8Array.of(0, 0)); // method stored
      pushByteChunk(fileTime);
      pushByteChunk(fileDate);
      zipDword(item.crc, pushByteChunk);
      zipDword(item.size, pushByteChunk);
      zipDword(item.size, pushByteChunk);
      zipParts(item.nameBytes.length, pushByteChunk);
      zipParts(0, pushByteChunk); // extra
      zipParts(0, pushByteChunk); // comment
      zipParts(0, pushByteChunk); // disk start
      zipParts(0, pushByteChunk); // internal attrs
      zipDword(0, pushByteChunk); // external attrs
      zipDword(item.localOffset, pushByteChunk);
      pushByteChunk(item.nameBytes);
    }

    const centralSize = offset - centralOffset;

    // end of central directory
    pushByteChunk(Uint8Array.of(0x50, 0x4b, 0x05, 0x06));
    zipParts(0, pushByteChunk);
    zipParts(0, pushByteChunk);
    zipParts(central.length, pushByteChunk);
    zipParts(central.length, pushByteChunk);
    zipDword(centralSize, pushByteChunk);
    zipDword(centralOffset, pushByteChunk);
    zipParts(0, pushByteChunk); // comment length

    return new Blob(chunks, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  /** Build the xlsx file for one sheet, returning base64 (for the popup). */
  function buildXlsxBase64(sheetXml) {
    const encoder = new TextEncoder();
    const entries = [
      {
        name: "[Content_Types].xml",
        data: encoder.encode(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
            "</Types>"
        ),
      },
      {
        name: "_rels/.rels",
        data: encoder.encode(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            "</Relationships>"
        ),
      },
      {
        name: "xl/workbook.xml",
        data: encoder.encode(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<sheets><sheet name="PMRS" sheetId="1" r:id="rId1"/></sheets>' +
            "</workbook>"
        ),
      },
      {
        name: "xl/_rels/workbook.xml.rels",
        data: encoder.encode(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
            "</Relationships>"
        ),
      },
      { name: "xl/worksheets/sheet1.xml", data: encoder.encode(sheetXml) },
    ];

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const dataUrl = String(reader.result || "");
        const comma = dataUrl.indexOf(",");
        resolve(comma >= 0 ? dataUrl.slice(comma + 1) : "");
      };
      reader.onerror = () => reject(new Error("Could not build the xlsx file."));
      reader.readAsDataURL(makeZipBlob(entries));
    });
  }

  function notifyProgress(payload) {
    try {
      chrome.runtime.sendMessage(payload, () => {
        void chrome.runtime.lastError;
      });
    } catch (_) {
      /* popup closed — export continues anyway */
    }
  }

  /* ══════════════════════════════════════════════════════════════════════
     EXPORTS
     ══════════════════════════════════════════════════════════════════════ */

  async function exportPage(message) {
    const table = pickTable();
    if (!table) throw new Error("No table with data was found on this page.");

    const info = readTable(table);
    if (!info.rows.length) throw new Error("The table on this page has no rows.");

    const sheetXml = buildSheetXml(message.includeHeader !== false ? info.headers : [], info.rows);
    const base64 = await buildXlsxBase64(sheetXml);

    return {
      ok: true,
      rows: info.rows.length,
      columns: info.columns,
      pages: 1,
      includeHeader: message.includeHeader !== false,
      filename: reportFilename(),
      base64,
    };
  }

  async function exportAll(message) {
    const firstTable = pickTable();
    if (!firstTable) throw new Error("No table with data was found on this page.");

    const includeHeader = message.includeHeader !== false;
    const limit = Math.min(Math.max(parseInt(message.maxPages || MAX_PAGES, 10), 1), MAX_PAGES);

    // Fresh run — clear any leftover stop flag from a previous download.
    stopRequested = false;

    const seen = new Set();
    const collected = [];
    let headers = [];
    let columns = 0;
    let pages = 0;
    let warning = "";

    for (let guard = 0; guard < limit; guard += 1) {
      if (stopRequested) {
        warning = `Stopped by user after ${pages} page(s).`;
        break;
      }

      const table = pickTable();
      if (!table) break;

      const info = readTable(table);
      if (!info.rows.length) break;

      const meta = readPageNumber();
      const signature = fullSignature(info, meta);

      if (!seen.has(signature)) {
        seen.add(signature);
        pages += 1;
        if (!headers.some(Boolean)) headers = info.headers;
        columns = Math.max(columns, info.columns);
        collected.push(...info.rows);

        notifyProgress({
          type: "PMRS_PROGRESS",
          page: pages,
          rows: info.rows.length,
          totalRows: collected.length,
        });
      } else if (pages === 0) {
        break;
      }

      if (guard === limit - 1) {
        warning = `Stopped at the ${limit}-page safety limit.`;
        break;
      }

      const next = findNext();
      if (!next) break;

      const before = rowsSignature(info);
      await clickNext(next);

      if (stopRequested) {
        warning = `Stopped by user after ${pages} page(s).`;
        break;
      }

      if (await waitForRowsChange(before)) continue;

      // One retry — Vue re-renders the pager too, so re-query it.
      const retry = findNext();
      if (!retry) {
        warning = "Stopped: no further Next button was found.";
        break;
      }
      await clickNext(retry);
      if (!(await waitForRowsChange(before))) {
        warning = "Stopped: the table did not change after clicking Next (last page?).";
        break;
      }
    }

    if (!collected.length) throw new Error("No rows were collected.");

    const filename = reportFilename(); // PMRS_Report_YYYY-MM-DD.xlsx
    const base64 = await buildXlsxBase64(
      buildSheetXml(includeHeader ? headers : [], collected)
    );

    const stoppedNote = stopRequested ? " (stopped early by user)" : "";
    notifyProgress({
      type: "PMRS_PROGRESS",
      phase: "done",
      page: pages,
      rows: collected.length,
      totalRows: collected.length,
      warning,
      message: `${filename} — ${collected.length} rows from ${pages} page(s)${stoppedNote}`,
    });

    return {
      ok: true,
      rows: collected.length,
      columns,
      pages,
      includeHeader,
      warning,
      filename,
      base64,
    };
  }

  /* ══════════════════════════════════════════════════════════════════════
     DIAGNOSE
     ══════════════════════════════════════════════════════════════════════ */
  function diagnose() {
    const tables = Array.from(document.querySelectorAll("table")).map((table, index) => {
      const info = readTable(table);
      return {
        index,
        visible: isVisible(table),
        hasThead: Boolean(table.tHead),
        rows: info.rows.length,
        columns: info.columns,
        firstCell: info.rows.length ? info.rows[0][0] : "",
      };
    });

    const table = pickTable();
    const info = table ? readTable(table) : null;
    const meta = readPageNumber();
    const candidates = nextCandidates();

    return {
      ok: true,
      build: BUILD,
      url: location.href,
      title: document.title,
      tableCount: tables.length,
      tables,
      selected: info
        ? {
            rows: info.rows.length,
            columns: info.columns,
            headers: info.headers,
            firstRow: info.rows[0] || [],
            lastRow: info.rows[info.rows.length - 1] || [],
          }
        : null,
      page: meta,
      next: candidates.length
        ? {
            label: candidates[0].label,
            why: candidates[0].why,
            score: candidates[0].score,
            disabled: candidates[0].disabled,
          }
        : null,
      candidates: candidates.slice(0, 8).map((candidate) => ({
        label: candidate.label,
        why: candidate.why,
        score: candidate.score,
        disabled: candidate.disabled,
        ariaLabel: candidate.el.getAttribute("aria-label"),
        dataSlot: candidate.el.getAttribute("data-slot"),
        text: clean(candidate.el.innerText).slice(0, 60),
        html: candidate.el.outerHTML.slice(0, 240),
      })),
    };
  }

  /* ══════════════════════════════════════════════════════════════════════
     DEEP ASYNC SCRAPE
     For each visible data row: click the blue eye / View button, wait for the
     detail modal to settle, read the labelled clinical fields, close it, then
     move to the next row. The first output is a rich JSON object per patient;
     the caller forwards it as `detailed_clinical_payload`.
     ══════════════════════════════════════════════════════════════════════ */

  function uniqueBlocks(values) {
    const seen = new Set();
    const out = [];
    for (const value of values) {
      const text = cleanBlock(value);
      if (text && !seen.has(text)) {
        seen.add(text);
        out.push(text);
      }
    }
    return out;
  }

  function cleanBlock(value) {
    return String(value == null ? "" : value)
      .replace(/\u200b/g, "")
      .replace(/\r\n?/g, "\n")
      .split("\n")
      .map((line) => line.replace(/[ \t\u00a0]+/g, " ").trim())
      .filter(Boolean)
      .join("\n")
      .trim();
  }

  /** MutationObserver + interval fallback; resolves with the first truthy value. */
  function waitUntil(predicate, options = {}) {
    const timeout = Number.isFinite(options.timeout)
      ? options.timeout
      : DEEP_SCRAPE.openTimeoutMs;
    const interval = Number.isFinite(options.interval)
      ? options.interval
      : DEEP_SCRAPE.pollMs;
    const label = options.label || "condition";

    return new Promise((resolve, reject) => {
      const started = Date.now();
      let settled = false;
      let timer = null;
      let observer = null;

      const cleanup = () => {
        if (timer) clearInterval(timer);
        if (observer) observer.disconnect();
        timer = null;
        observer = null;
      };
      const finish = (value, error) => {
        if (settled) return;
        settled = true;
        cleanup();
        if (error) reject(error);
        else resolve(value);
      };
      const check = () => {
        if (settled) return;
        try {
          const value = predicate();
          if (value) {
            finish(value);
            return;
          }
        } catch (_) {
          /* keep polling while Vue replaces nodes */
        }
        if (Date.now() - started >= timeout) {
          finish(null, new Error(`Timed out waiting for ${label}.`));
        }
      };

      check();
      if (settled) return;
      timer = setInterval(check, interval);
      if (document.body && typeof MutationObserver === "function") {
        observer = new MutationObserver(check);
        observer.observe(document.body, {
          childList: true,
          subtree: true,
          characterData: true,
        });
      }
    });
  }

  function visibleModals() {
    const nodes = new Set();
    for (const selector of DEEP_SCRAPE.modalSelectors) {
      try {
        document.querySelectorAll(selector).forEach((node) => {
          if (isVisible(node)) nodes.add(node);
        });
      } catch (_) {
        /* selector may be unsupported on an old browser */
      }
    }
    return Array.from(nodes);
  }

  function findVisibleModal() {
    const modals = visibleModals();
    return modals.length ? modals[modals.length - 1] : null;
  }

  async function waitForModalStable(modal) {
    let lastText = "";
    let stable = 0;

    await waitUntil(
      () => {
        if (!isVisible(modal)) return false;
        const text = String(modal.innerText || "");
        if (text.trim().length < 120) return false;

        // Do not accept the empty shell/header before the detail API answers.
        const looksLoaded =
          text.trim().length >= 300 ||
          /រោគវិនិច្ឆ័យ|ឈ្មោះអ្នកជំងឺ|លេខបណ្ណ|History of Present Illness|diagnosis|សីតុណ្ហភាព/i.test(
            text
          );
        if (!looksLoaded) return false;

        const busy = Array.from(
          modal.querySelectorAll(
            '[aria-busy="true"], [class*="loading" i], [class*="spinner" i]'
          )
        ).some(isVisible);
        if (busy) {
          lastText = "";
          stable = 0;
          return false;
        }

        if (text === lastText) stable += 1;
        else {
          lastText = text;
          stable = 1;
        }
        return stable >= 2;
      },
      { timeout: DEEP_SCRAPE.openTimeoutMs, label: "detail modal content" }
    );
    await sleep(DEEP_SCRAPE.settleMs);
  }

  async function waitForModalOpen(previousModal) {
    const modal = await waitUntil(
      () => {
        const current = findVisibleModal();
        if (!current) return null;
        if (previousModal && current === previousModal) return null;
        return current;
      },
      { timeout: DEEP_SCRAPE.openTimeoutMs, label: "detail modal" }
    );
    await waitForModalStable(modal);
    return modal;
  }

  function findModalCloseButton(modal) {
    for (const selector of DEEP_SCRAPE.closeButtonSelectors) {
      try {
        const node = modal.querySelector(selector);
        if (node && isVisible(node) && !node.disabled) return node;
      } catch (_) {
        /* ignore */
      }
    }

    const buttons = Array.from(modal.querySelectorAll("button"));
    return (
      buttons.find((button) => {
        if (!isVisible(button) || button.disabled) return false;
        const text = clean(button.innerText);
        const attrs = [
          button.getAttribute("aria-label") || "",
          button.getAttribute("title") || "",
        ].join(" ");
        return /^[✕×x]$/i.test(text) || /close|dismiss/i.test(attrs);
      }) ||
      buttons.find((button) => !button.disabled && isVisible(button)) ||
      null
    );
  }

  async function closeModal(modal) {
    if (!modal || !modal.isConnected) return;

    const close = findModalCloseButton(modal);
    if (close) close.click();
    else {
      document.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "Escape",
          code: "Escape",
          keyCode: 27,
          which: 27,
          bubbles: true,
        })
      );
    }

    try {
      await waitUntil(() => !modal.isConnected || !isVisible(modal), {
        timeout: DEEP_SCRAPE.closeTimeoutMs,
        label: "detail modal close",
      });
    } catch (_) {
      document.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "Escape",
          code: "Escape",
          keyCode: 27,
          which: 27,
          bubbles: true,
        })
      );
      await waitUntil(() => !modal.isConnected || !isVisible(modal), {
        timeout: 2000,
        label: "detail modal close",
      }).catch(() => {});
    }
    await sleep(DEEP_SCRAPE.settleMs);
  }

  function tableDataRows(table) {
    return Array.from(table.rows).filter((row) => {
      if (row.closest("thead")) return false;
      if (!isVisible(row)) return false;
      return rowCells(row).some(Boolean);
    });
  }

  function findRowViewButton(row) {
    if (!row) return null;

    const cells = Array.from(row.cells || []);
    const scope = cells.length ? cells[cells.length - 1] : row;

    for (const selector of DEEP_SCRAPE.viewButtonSelectors) {
      try {
        const node = scope.querySelector(selector) || row.querySelector(selector);
        if (node && isVisible(node) && !node.disabled) return node;
      } catch (_) {
        /* ignore selector support */
      }
    }

    const buttons = Array.from(
      scope.querySelectorAll('button, a, [role="button"]')
    );
    if (!buttons.length) return null;

    const scored = buttons
      .map((button) => {
        const blob = [
          button.className,
          button.getAttribute("aria-label"),
          button.getAttribute("title"),
          button.innerText,
        ].join(" ");
        const html = button.innerHTML || "";
        let score = 0;

        if (button.tagName === "BUTTON") score += 20;
        if (/eye/i.test(blob) || /eye/i.test(html)) score += 100;
        if (/view|detail|eye/i.test(blob)) score += 40;
        if (/bg-blue|text-white/.test(String(button.className || ""))) score += 20;
        if (
          button.disabled ||
          (button.getAttribute("aria-disabled") || "").toLowerCase() === "true"
        ) {
          score -= 200;
        }
        return { button, score };
      })
      .filter((item) => item.score > 0);

    scored.sort((a, b) => b.score - a.score);
    return scored.length ? scored[0].button : null;
  }

  function selectorTexts(root, selectors) {
    const out = [];
    for (const selector of selectors || []) {
      try {
        root.querySelectorAll(selector).forEach((node) => {
          const text = cleanBlock(node.innerText || node.textContent || "");
          if (text) out.push(text);
        });
      } catch (_) {
        /* invalid selector until the caller configures it */
      }
    }
    return uniqueBlocks(out);
  }

  function stripLabel(text, label) {
    const cleaned = cleanBlock(text);
    const needle = clean(label);
    if (!needle) return cleaned.replace(/^[៖:\s]+/, "").trim();
    const index = cleaned.indexOf(needle);
    if (index < 0) return cleaned.replace(/^[៖:\s]+/, "").trim();
    return cleaned
      .slice(index + needle.length)
      .replace(/^[៖:\s]+/, "")
      .trim();
  }

  function collectLabelValuePairs(root) {
    const pairs = [];
    const seen = new Set();
    const nodes = root.querySelectorAll("h3, label, strong, dt");

    for (const node of nodes) {
      if (!isVisible(node)) continue;

      let label = clean(node.innerText || node.textContent || "")
        .replace(/[៖:]\s*$/, "")
        .trim();
      if (!label || label.length > 120) continue;

      let value = "";
      if (node.matches("h3") || node.matches("dt")) {
        const next = node.nextElementSibling;
        if (next) value = cleanBlock(next.innerText || next.textContent || "");
      } else if (node.matches("label")) {
        const nested = node.querySelector(":scope > span.text-sm");
        if (nested) {
          const nestedText = clean(nested.innerText || nested.textContent || "");
          label = nestedText.replace(/[៖:]\s*$/, "").trim();
          value = stripLabel(node.innerText || node.textContent || "", nestedText);
        } else {
          const next = node.nextElementSibling;
          if (next) value = cleanBlock(next.innerText || next.textContent || "");
          if (!value && node.parentElement) {
            value = stripLabel(
              node.parentElement.innerText || node.parentElement.textContent || "",
              label
            );
          }
        }
      } else if (node.matches("strong")) {
        if (node.parentElement) {
          value = stripLabel(
            node.parentElement.innerText || node.parentElement.textContent || "",
            label
          );
        }
        if (!value) {
          const next = node.nextElementSibling;
          if (next) value = cleanBlock(next.innerText || next.textContent || "");
        }
      }

      value = String(value || "")
        .replace(/^[៖:\s]+/, "")
        .trim();
      if (!value || value === label) continue;
      if (value.length > 8000) value = value.slice(0, 8000);

      const key = `${label}\u0000${value}`;
      if (seen.has(key)) continue;
      seen.add(key);
      pairs.push({ label, value });
    }
    return pairs;
  }

  function pickField(pairs, aliases) {
    const needles = aliases.map(normaliseHeader).filter(Boolean);
    let partial = "";

    for (const pair of pairs) {
      const label = normaliseHeader(pair.label);
      if (needles.some((needle) => label === needle)) return pair.value;
      if (!partial && needles.some((needle) => label.includes(needle))) {
        partial = pair.value;
      }
    }
    return partial;
  }

  function pickFields(pairs, aliases) {
    const needles = aliases.map(normaliseHeader).filter(Boolean);
    return pairs
      .filter((pair) => {
        const label = normaliseHeader(pair.label);
        return needles.some((needle) => label === needle || label.includes(needle));
      })
      .map((pair) => pair.value)
      .filter(Boolean);
  }

  function splitDiagnosisLines(value) {
    const lines = [];
    for (const rawLine of cleanBlock(value).split(/\n+/)) {
      const line = singleLine(rawLine).replace(/^[-•*]\s*/, "");
      if (!line) continue;
      const pieces = line.split(
        /(?=\b\d{1,2}[.)]\s*[A-TV-Z]\d{2}(?:\.\d{1,4})?\b)/i
      );
      for (const piece of pieces) {
        const cleaned = singleLine(piece).replace(/^[-•*]\s*/, "");
        if (cleaned) lines.push(cleaned);
      }
    }
    return lines;
  }

  function parseDiagnosisLine(line) {
    const text = singleLine(line);
    const match = text.match(/\b([A-TV-Z]\d{2}(?:\.\d{1,4})?)\b/i);
    let code = "";
    let description = text;

    if (match) {
      code = match[1].toUpperCase();
      description = text
        .slice((match.index || 0) + match[0].length)
        .replace(/^[\s.:\-–—]+/, "")
        .trim();
    }

    const markerMatch =
      description.match(/\b(primary|differential)\b/i) ||
      text.match(/\b(primary|differential)\b/i);
    const marker = markerMatch ? markerMatch[1].toLowerCase() : "";
    description = description
      .replace(/\s*\((?:primary|differential)\)\s*/gi, " ")
      .replace(/\s+/g, " ")
      .trim();

    return { code, description, text, marker };
  }

  function classifyDiagnosisItems(items) {
    const unique = [];
    const seen = new Set();
    for (const item of items) {
      const key = `${item.code}\u0000${item.text}`;
      if (seen.has(key)) continue;
      seen.add(key);
      unique.push(item);
    }

    let primarySeen = unique.some((item) => item.marker === "primary");
    for (const item of unique) {
      if (item.marker === "differential") item.classification = "differential";
      else if (item.marker === "primary") {
        item.classification = "primary";
        primarySeen = true;
      } else if (!primarySeen) {
        item.classification = "primary";
        primarySeen = true;
      } else {
        item.classification = "differential";
      }
      delete item.marker;
    }
    return unique;
  }

  function buildDiagnoses(pairs, configuredPrimary, configuredDifferential) {
    const admissionRaw = pickField(pairs, [
      "រោគវិនិច្ឆ័យចូល",
      "diagnosis in",
      "admission diagnosis",
    ]);
    const dischargeValues = pickFields(pairs, [
      "រោគវិនិច្ឆ័យចេញ",
      "diagnosis out",
      "discharge diagnosis",
      "final diagnosis",
    ]);
    const sourceValues = dischargeValues.length
      ? dischargeValues
      : admissionRaw
        ? [admissionRaw]
        : [];
    const source = sourceValues.sort((a, b) => b.length - a.length)[0] || "";

    const items = splitDiagnosisLines(source).map(parseDiagnosisLine);
    for (const value of configuredPrimary || []) {
      const item = parseDiagnosisLine(value);
      item.marker = "primary";
      items.push(item);
    }
    for (const value of configuredDifferential || []) {
      const item = parseDiagnosisLine(value);
      item.marker = "differential";
      items.push(item);
    }

    const all = classifyDiagnosisItems(items);
    return {
      primary: all.filter((item) => item.classification === "primary"),
      differential: all.filter(
        (item) => item.classification === "differential"
      ),
      all,
      admission_raw: admissionRaw,
      discharge_raw: source,
    };
  }

  function firstNumber(value) {
    const text = singleLine(value);
    const match = text.match(/-?\d+(?:\.\d+)?/);
    return match ? match[0] : text;
  }

  function parseBloodPressure(value) {
    const text = singleLine(value);
    if (!text) return { systolic: "", diastolic: "", text: "" };

    let systolic = "";
    let diastolic = "";
    const labeledSystolic = text.match(/systolic[^0-9]*(\d{2,3})/i);
    const labeledDiastolic = text.match(/diastolic[^0-9]*(\d{2,3})/i);
    if (labeledSystolic) systolic = labeledSystolic[1];
    if (labeledDiastolic) diastolic = labeledDiastolic[1];

    if (!systolic || !diastolic) {
      const pair = text.match(
        /(\d{2,3})\s*\/\s*(?:diastolic[^0-9]*)?(\d{2,3})/i
      );
      if (pair) {
        systolic = systolic || pair[1];
        diastolic = diastolic || pair[2];
      }
    }

    return {
      systolic,
      diastolic,
      text: systolic && diastolic ? `${systolic}/${diastolic}` : text,
    };
  }

  function buildClinicalDetail(modal) {
    const pairs = collectLabelValuePairs(modal);
    const configured = DEEP_SCRAPE.selectors;
    const configuredText = (key) => selectorTexts(modal, configured[key] || []);

    const diagnoses = buildDiagnoses(
      pairs,
      configuredText("primaryDiagnosis"),
      configuredText("differentialDiagnosis")
    );

    const chiefComplaint =
      configuredText("chiefComplaint")[0] ||
      pickField(pairs, [
        "មូលហេតុចូល",
        "chief complaint",
        "presenting complaint",
        "reason for admission",
      ]);
    const hpi =
      configuredText("historyOfPresentIllness")[0] ||
      pickField(pairs, [
        "History of Present Illness",
        "HPI",
        "ប្រវត្តិជំងឺបច្ចុប្បន្ន",
        "ប្រវត្តិជំងឺ",
      ]);
    const temperature =
      configuredText("temperature")[0] ||
      pickField(pairs, ["សីតុណ្ហភាព", "temperature", "temp"]);
    const weight =
      configuredText("weight")[0] ||
      pickField(pairs, ["ទម្ងន់ (kg)", "ទម្ងន់(kg)", "weight (kg)", "weight"]);
    const bloodPressureRaw =
      configuredText("bloodPressure")[0] ||
      pickField(pairs, ["សម្ពាធឈាម", "blood pressure", "bp"]);
    const bloodPressure = parseBloodPressure(bloodPressureRaw);
    const height = pickField(pairs, ["កម្ពស់ (cm)", "height (cm)", "height"]);
    const pulse = pickField(pairs, ["ចង្វាក់បេះដូង", "pulse", "heart rate"]);
    const respiratoryRate = pickField(pairs, [
      "ចង្វាក់ដង្ហើម",
      "respiratory rate",
      "resp rate",
      "rr",
    ]);
    const spo2 = pickField(pairs, [
      "អុកស៊ីហ្សែន",
      "spo2",
      "oxygen saturation",
      "o2 sat",
    ]);
    const symptoms = pickField(pairs, ["រោគសញ្ញា", "symptoms", "symptom"]);

    const primaryCodes = diagnoses.primary
      .map((item) => item.code)
      .filter(Boolean);
    const differentialCodes = diagnoses.differential
      .map((item) => item.code)
      .filter(Boolean);

    const fields = pairs.map((pair) => ({
      label: pair.label,
      value: pair.value,
    }));
    const fieldMap = {};
    for (const pair of pairs) {
      if (!fieldMap[pair.label]) fieldMap[pair.label] = [];
      fieldMap[pair.label].push(pair.value);
    }

    return {
      schema_version: 1,
      extracted_at: new Date().toISOString(),
      claim: {
        claim_no: pickField(pairs, ["លេខករណី", "claim no", "claim number"]),
        request_date: pickField(pairs, ["ថ្ងៃស្នើសុំ", "request date"]),
        district: pickField(pairs, ["ស្រុកប្រតិបត្តិ", "district"]),
        facility: pickField(pairs, ["មូលដ្ឋានសុខាភិបាល", "facility"]),
        facility_level: pickField(pairs, [
          "កម្រិតមូលដ្ឋានសុខាភិបាល",
          "facility level",
        ]),
      },
      patient: {
        health_card_id: pickField(pairs, [
          "លេខបណ្ណសុខភាព",
          "health card id",
          "health card",
        ]),
        equity_card_id: pickField(pairs, [
          "លេខបណ្ណសមធម៌",
          "equity card id",
          "equity card",
        ]),
        patient_name: pickField(pairs, [
          "ឈ្មោះអ្នកជំងឺ",
          "patient name",
          "patient",
        ]),
        gender: pickField(pairs, ["ភេទ", "gender", "sex"]),
        dob: pickField(pairs, [
          "ថ្ងៃខែឆ្នាំកំណើត",
          "date of birth",
          "dob",
        ]),
        phone: pickField(pairs, ["លេខទូរសព្ទ", "phone"]),
        nationality: pickField(pairs, ["សញ្ជាតិ", "nationality"]),
        address: pickField(pairs, ["អាសយដ្ឋាន", "address"]),
      },
      diagnoses,
      primary_icd10: primaryCodes,
      differential_icd10: differentialCodes,
      chief_complaint: cleanBlock(chiefComplaint),
      chief_complaints: uniqueBlocks(
        cleanBlock(chiefComplaint).split(/\n+/)
      ),
      symptoms: cleanBlock(symptoms),
      history_of_present_illness: cleanBlock(hpi).slice(0, 8000),
      vitals: {
        temperature_c: firstNumber(temperature),
        weight_kg: firstNumber(weight),
        height_cm: firstNumber(height),
        blood_pressure: bloodPressure.text,
        systolic: bloodPressure.systolic,
        diastolic: bloodPressure.diastolic,
        pulse: firstNumber(pulse),
        respiratory_rate: firstNumber(respiratoryRate),
        spo2: firstNumber(spo2),
      },
      clinical: {
        department_in: pickField(pairs, ["ផ្នែកចូល", "department in"]),
        department_out: pickField(pairs, ["ផ្នែកចេញ", "department out"]),
        admission_date: pickField(pairs, [
          "កាលបរិច្ឆេទចូល",
          "admission date",
        ]),
        discharge_date: pickField(pairs, [
          "កាលបរិច្ឆេទចេញ",
          "discharge date",
        ]),
        visit_type: pickField(pairs, [
          "ប្រភេទនៃការចូល",
          "visit type",
          "admission type",
        ]),
        discharge_type: pickField(pairs, [
          "ប្រភេទនៃការចេញ",
          "discharge type",
        ]),
        disease_status: pickField(pairs, [
          "ស្ថានភាពជំងឺ",
          "disease status",
        ]),
        service_unit: pickField(pairs, [
          "ផ្នែកទទួលសេវា",
          "service unit",
        ]),
      },
      fields,
      field_map: fieldMap,
      raw_text: String(modal.innerText || "")
        .replace(/\u200b/g, "")
        .trim()
        .slice(0, DEEP_SCRAPE.maxRawText),
    };
  }

  function notifyDeepProgress(payload) {
    try {
      chrome.runtime.sendMessage(
        { type: "PMRS_AUTO_SYNC_PROGRESS", phase: "deep", ...payload },
        () => {
          void chrome.runtime.lastError;
        }
      );
    } catch (_) {
      /* popup closed */
    }
  }

  async function deepScrapeCurrentPage(message = {}) {
    const table = pickTable();
    if (!table) throw new Error("No table with data was found on this page.");

    const info = readTable(table);
    if (!info.rows.length) throw new Error("The table on this page has no rows.");

    const ptColumn = findHeaderIndex(info.headers, PT_CODE_HEADERS);
    if (ptColumn < 0) {
      throw new Error(
        "Could not find the patient-card column (បណ្ណសុខភាព / លេខបណ្ណ / Health card)."
      );
    }

    const surface = recordsFromTable(info);
    const byCode = new Map(surface.map((record) => [record.pt_code, record]));
    const parsedLimit = parseInt(message.maxRows || DEEP_SCRAPE.maxRows, 10);
    const limit = Math.min(
      Math.max(Number.isFinite(parsedLimit) ? parsedLimit : DEEP_SCRAPE.maxRows, 1),
      DEEP_SCRAPE.maxRows
    );

    stopRequested = false;
    const meta = readPageNumber();
    const rows = tableDataRows(table);
    const total = Math.min(rows.length, limit);
    const records = [];
    const errors = [];
    let deepScraped = 0;
    let stopped = false;

    for (let index = 0; index < total; index += 1) {
      if (stopRequested) {
        stopped = true;
        break;
      }

      const currentTable = pickTable();
      if (!currentTable) break;
      const currentRows = tableDataRows(currentTable);
      const row = currentRows[index];
      if (!row) continue;

      const cells = rowCells(row);
      const code = singleLine(cells[ptColumn] || "")
        .replace(/\s+/g, "")
        .toUpperCase()
        .slice(0, 64);
      if (!code) continue;

      const fallbackStatus = /\/rejected\//i.test(location.pathname)
        ? "REJECTED"
        : /\/approved\//i.test(location.pathname)
          ? "APPROVED"
          : "UNKNOWN";
      const base = byCode.get(code) || {
        pt_code: code,
        status: fallbackStatus,
        pmrs_reject_reason: "",
        patient_name: "",
        gender: "",
        age: "",
        visit_type: "",
        icd10_diagnosis: "",
        approval_date: "",
      };

      const button = findRowViewButton(row);
      let detail = null;
      let detailError = "";

      if (!button) {
        detailError = "View button not found";
      } else {
        const previousModal = findVisibleModal();
        try {
          try {
            row.scrollIntoView({ block: "center", inline: "nearest" });
          } catch (_) {
            /* ignore */
          }
          button.click();
          const modal = await waitForModalOpen(previousModal);
          detail = buildClinicalDetail(modal);
          deepScraped += 1;
          await closeModal(modal);
        } catch (error) {
          detailError = String((error && error.message) || error);
          const open = findVisibleModal();
          if (open) {
            try {
              await closeModal(open);
            } catch (_) {
              /* leave the next iteration to retry */
            }
          }
        }
      }

      if (detailError) errors.push({ index, pt_code: code, error: detailError });
      const record = { ...base, detailed_clinical_payload: detail };
      if (detailError) record.detail_error = detailError;
      records.push(record);

      notifyDeepProgress({
        page: meta.page,
        processed: records.length,
        total,
        pt_code: code,
        error: detailError,
      });
      await sleep(DEEP_SCRAPE.rowDelayMs);
    }

    if (!records.length) {
      throw new Error("No patient rows were available for deep scraping.");
    }
    if (deepScraped === 0) {
      throw new Error(
        "Deep scrape could not open any detail modal. Check the View button selector."
      );
    }

    const snapshot = pageSnapshot();
    return {
      ok: true,
      records,
      rows: records.length,
      deepScraped,
      errors,
      stopped,
      page: meta.page,
      pageCount: snapshot.meta.pageCount,
      hasNext: Boolean(findNext()),
      signature: snapshot.signature,
    };
  }

  /* ══════════════════════════════════════════════════════════════════════
     AUTO-SYNC BRIDGE (background.js owns the pagination loop)
     The service worker requests ONE page at a time. "Next" is clicked only
     after the reply has been sent, so a full page reload can never cut the
     message port off mid-response.
     ══════════════════════════════════════════════════════════════════════ */

  function pageSnapshot() {
    const table = pickTable();
    const info = table ? readTable(table) : null;
    const meta = readPageNumber();
    const hasRows = Boolean(info && info.rows.length);
    return {
      info,
      meta,
      signature: hasRows ? fullSignature(info, meta) : "",
    };
  }

  function scrapeForSync() {
    const result = collectCurrentPageRecords();
    const snapshot = pageSnapshot();
    return {
      ok: true,
      records: result.records,
      rows: result.rows,
      page: result.page,
      pageCount: snapshot.meta.pageCount,
      hasNext: Boolean(findNext()),
      signature: snapshot.signature,
    };
  }

  function syncPageState() {
    const snapshot = pageSnapshot();
    return {
      ok: Boolean(snapshot.info && snapshot.info.rows.length),
      page: snapshot.meta.page,
      pageCount: snapshot.meta.pageCount,
      hasNext: Boolean(findNext()),
      signature: snapshot.signature,
    };
  }

  /* ══════════════════════════════════════════════════════════════════════
     MESSAGE ROUTER
     ══════════════════════════════════════════════════════════════════════ */

  function getInfo() {
    const table = pickTable();
    if (!table) return { ok: false, error: "No table found on this page." };

    const info = readTable(table);
    if (!info.rows.length) return { ok: false, error: "The table has no rows." };

    const meta = readPageNumber();
    return {
      ok: true,
      rows: info.rows.length,
      columns: info.columns,
      page: meta.page,
      pageCount: meta.pageCount,
      hasNext: Boolean(findNext()),
    };
  }

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    const type = message && message.type;

    try {
      if (type === "PMRS_GET_INFO") {
        sendResponse(getInfo());
        return;
      }

      if (type === "PMRS_EXPORT_PAGE") {
        exportPage(message)
          .then((result) => sendResponse(result))
          .catch((error) =>
            sendResponse({ ok: false, error: String((error && error.message) || error) })
          );
        return true; // async response
      }

      if (type === "PMRS_EXPORT_ALL") {
        exportAll(message)
          .then((result) => sendResponse(result))
          .catch((error) =>
            sendResponse({ ok: false, error: String((error && error.message) || error) })
          );
        return true; // async response
      }

      if (type === "PMRS_STOP_DOWNLOAD") {
        stopRequested = true;
        sendResponse({ ok: true });
        return;
      }

      if (type === "PMRS_COLLECT_RECORDS") {
        if (message.deepScrape) {
          deepScrapeCurrentPage(message)
            .then((result) => sendResponse(result))
            .catch((error) =>
              sendResponse({ ok: false, error: String((error && error.message) || error) })
            );
          return true; // async response
        }
        sendResponse(collectCurrentPageRecords());
        return;
      }

      if (type === "PMRS_SYNC_SCRAPE_DEEP") {
        deepScrapeCurrentPage(message)
          .then((result) => sendResponse(result))
          .catch((error) =>
            sendResponse({ ok: false, error: String((error && error.message) || error) })
          );
        return true; // async response
      }

      if (type === "PMRS_AUTO_SYNC") {
        sendResponse({
          ok: false,
          error:
            "The pagination loop now lives in background.js. Use PMRS_SYNC_SCRAPE / PMRS_SYNC_NEXT.",
        });
        return;
      }

      if (type === "PMRS_PING") {
        sendResponse({ ok: true, build: BUILD });
        return;
      }

      if (type === "PMRS_SYNC_SCRAPE") {
        try {
          sendResponse(scrapeForSync());
        } catch (error) {
          sendResponse({ ok: false, error: String((error && error.message) || error) });
        }
        return;
      }

      if (type === "PMRS_SYNC_STATE") {
        try {
          sendResponse(syncPageState());
        } catch (error) {
          sendResponse({ ok: false, error: String((error && error.message) || error) });
        }
        return;
      }

      if (type === "PMRS_SYNC_NEXT") {
        const next = findNext();
        if (!next) {
          sendResponse({ ok: true, clicked: false });
          return;
        }

        // Reply BEFORE clicking. A full page reload destroys this content
        // script, and any reply sent afterwards would lose the message port.
        sendResponse({ ok: true, clicked: true });
        setTimeout(() => {
          // Re-query: Vue may have replaced the pager nodes in the meantime.
          const fresh = findNext();
          if (fresh) clickNext(fresh).catch(() => {});
        }, 80);
        return;
      }

      if (type === "PMRS_SYNC_STOP") {
        stopRequested = true;
        sendResponse({ ok: true });
        return;
      }

      if (type === "PMRS_DIAGNOSE") {
        sendResponse(diagnose());
        return;
      }
    } catch (error) {
      sendResponse({ ok: false, error: String((error && error.message) || error) });
    }
  });
})();
