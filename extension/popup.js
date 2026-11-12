/* ─────────────────────────────────────────────────────────────
   PMRS -> Excel v1.0 — popup logic
   Talks to content.js on the active PMRS tab.
   ───────────────────────────────────────────────────────────── */
"use strict";

const PMRS_ORIGIN = "https://pmrs.digitalsp.gov.kh/";
const PMRS_LOGS_URL = "http://localhost:8000/api/save_pmrs_records";

const el = (id) => document.getElementById(id);
const dimsEl = el("dims");
const pageInfoEl = el("pageInfo");
const statusEl = el("status");
const diagEl = el("diag");
const includeHeaderEl = el("includeHeader");
const allPagesEl = el("allPages");
const deepScrapeEl = el("deepScrape");
const downloadAllBtn = el("downloadAll");
const downloadPageBtn = el("downloadPage");
const diagnoseBtn = el("diagnose");
const syncRejectBtn = el("btn-sync-reject");
const syncApproveBtn = el("btn-sync-approve");
const syncStatusEl = el("syncStatus");
const minimizeBtn = el("minimize");

const SYNC_BUTTONS = [
  { type: "rejection", el: syncRejectBtn, label: "Save Rejections to Logs" },
  { type: "approved", el: syncApproveBtn, label: "Save Approved (Gold) to DB" },
];

let activeTab = null;
let busy = false;

/* Start/stop toggle state for 'Download all pages'. */
const ALL_LABEL = "Download all pages";
const STOP_LABEL = "Stop Downloading";
let isDownloading = false;

/* Auto-sync state (owned by background.js, mirrored here for the UI). */
const STOP_SYNC_LABEL = "Stop sync";
let isSyncing = false;
let syncingType = null; // "rejection" | "approved"

/* ── tiny UI helpers ─────────────────────────────────────── */
function setInfo(main, sub, bad = false) {
  dimsEl.textContent = main;
  dimsEl.classList.toggle("is-bad", Boolean(bad));
  if (typeof sub === "string") pageInfoEl.textContent = sub;
}

function setStatus(message, kind = "") {
  statusEl.textContent = message;
  statusEl.classList.toggle("is-error", kind === "error");
  statusEl.classList.toggle("is-ok", kind === "ok");
}

/* Separate line for the database-sync result so export messages are untouched. */
function setSyncStatus(message, kind = "") {
  syncStatusEl.textContent = message || "";
  syncStatusEl.hidden = !message;
  syncStatusEl.classList.toggle("is-error", kind === "error");
  syncStatusEl.classList.toggle("is-ok", kind === "ok");
}

function syncButtonFor(type) {
  const entry = SYNC_BUTTONS.find((item) => item.type === type);
  return entry ? entry.el : null;
}

function setBusy(value, label) {
  busy = value;
  deepScrapeEl.disabled = value;
  [downloadAllBtn, downloadPageBtn, diagnoseBtn, ...SYNC_BUTTONS.map((item) => item.el)].forEach(
    (btn) => {
      btn.disabled = value || !activeTab;
    }
  );
  // While downloading, downloadAllBtn stays clickable so it can act as "Stop".
  if (value && isDownloading) downloadAllBtn.disabled = false;
  // While auto-syncing, the active sync button stays clickable as "Stop sync".
  if (value && isSyncing) {
    const active = syncButtonFor(syncingType);
    if (active) active.disabled = false;
  }
  if (value && label) setStatus(label);
}

function setSyncButtonState(syncing, syncType = null) {
  isSyncing = syncing;
  syncingType = syncing ? syncType : null;

  SYNC_BUTTONS.forEach(({ type, el: button, label }) => {
    const active = syncing && type === syncType;
    button.textContent = active ? STOP_SYNC_LABEL : label;
    button.classList.toggle("btn-stop", active);
    button.classList.toggle("btn-outline", !active && type === "rejection");
    button.classList.toggle("btn-gold", !active && type === "approved");
    // Keep only the running action clickable; the other one is frozen.
    button.disabled = syncing ? !active : !activeTab;
  });
}

function setDownloadingState(downloading) {
  isDownloading = downloading;
  downloadAllBtn.textContent = downloading ? STOP_LABEL : ALL_LABEL;
  downloadAllBtn.classList.toggle("btn-stop", downloading);
  downloadAllBtn.classList.toggle("btn-primary", !downloading);
  // While downloading: this button stays enabled (as Stop), others are frozen.
  downloadPageBtn.disabled = downloading || !activeTab;
  diagnoseBtn.disabled = downloading || !activeTab;
  SYNC_BUTTONS.forEach(({ el: button }) => {
    button.disabled = downloading || !activeTab;
  });
  downloadAllBtn.disabled = false;
}

function plural(n, word) {
  return `${n} ${word}${n === 1 ? "" : "s"}`;
}

function friendlyError(err) {
  const text = String((err && err.message) || err || "Unknown error");
  if (/Receiving end does not exist|Could not establish connection/i.test(text)) {
    return "Content script not found. Reload the PMRS page, then reopen this popup.";
  }
  if (/Failed to fetch|NetworkError|Load failed|fetch failed/i.test(text)) {
    return "Could not reach the backend. Start the API on http://localhost:8000 and try again.";
  }
  return text;
}

/* ── save the generated .xlsx into Downloads/Report_download ────────────── */
const DOWNLOAD_FOLDER = "Report_download";
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function base64ToBlob(base64, mime) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes], { type: mime });
}

function saveXlsxFile(base64, filename) {
  return new Promise((resolve, reject) => {
    let url;
    try {
      url = URL.createObjectURL(base64ToBlob(base64, XLSX_MIME));
    } catch (err) {
      reject(err);
      return;
    }

    chrome.downloads.download(
      {
        url,
        // Relative path under the browser's Downloads folder:
        filename: `${DOWNLOAD_FOLDER}/${filename}`,
        saveAs: false,
        conflictAction: "uniquify",
      },
      (downloadId) => {
        const err = chrome.runtime.lastError;
        setTimeout(() => URL.revokeObjectURL(url), 60000);
        if (err) reject(new Error(err.message));
        else if (typeof downloadId !== "number") reject(new Error("Download was cancelled."));
        else resolve(downloadId);
      }
    );
  });
}

/* ── chrome helpers ──────────────────────────────────────── */
function getActiveTab() {
  return new Promise((resolve) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      void chrome.runtime.lastError;
      resolve(tabs && tabs[0] ? tabs[0] : null);
    });
  });
}

function sendToTab(tabId, message) {
  return new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabId, message, (response) => {
      const err = chrome.runtime.lastError;
      if (err) {
        reject(new Error(err.message));
        return;
      }
      resolve(response);
    });
  });
}

function isPmrs(url) {
  return typeof url === "string" && url.startsWith(PMRS_ORIGIN);
}

/* ── preferences ─────────────────────────────────────────── */
function loadPrefs() {
  return new Promise((resolve) => {
    chrome.storage.local.get(
      { includeHeader: true, allPages: true, deepScrape: true },
      (prefs) => {
        void chrome.runtime.lastError;
        resolve(prefs);
      }
    );
  });
}

function savePrefs() {
  chrome.storage.local.set({
    includeHeader: includeHeaderEl.checked,
    allPages: allPagesEl.checked,
    deepScrape: deepScrapeEl.checked,
  });
}

/* ── page info ───────────────────────────────────────────── */
async function refreshInfo(silent = false) {
  if (!activeTab) return;
  try {
    const info = await sendToTab(activeTab.id, { type: "PMRS_GET_INFO" });
    if (!info || !info.ok) throw new Error((info && info.error) || "No table found.");

    setInfo(
      `${info.rows} rows x ${info.columns} columns`,
      `Page ${info.page || 1}${info.hasNext ? " · more pages available" : " · last page"}`
    );
    if (!silent) {
      setStatus(
        info.hasNext
          ? "Ready. Tip: set 'rows per page' to 100 on the site first for fewer requests."
          : "Ready. This looks like the last page.",
        "ok"
      );
    }
  } catch (err) {
    setInfo(friendlyError(err), "Open a page that contains the PMRS table.", true);
    if (!silent) setStatus("Could not read the table on this tab.", "error");
  }
}

/* ── actions ─────────────────────────────────────────────── */
async function runExport(allPages) {
  if (!activeTab || busy) return;
  setBusy(true, allPages ? "Starting — reading page 1…" : "Exporting current page…");

  // 'Download all pages' becomes a Start/Stop toggle while the loop runs.
  if (allPages) setDownloadingState(true);

  try {
    const message = {
      type: allPages ? "PMRS_EXPORT_ALL" : "PMRS_EXPORT_PAGE",
      includeHeader: includeHeaderEl.checked,
    };
    const result = await sendToTab(activeTab.id, message);
    if (!result || !result.ok) throw new Error((result && result.error) || "Export failed.");
    if (!result.base64) throw new Error("The content script returned no file data.");

    // Save the generated .xlsx into Downloads/Report_download/
    await saveXlsxFile(result.base64, result.filename);

    const parts = [`Saved ${result.filename} (${plural(result.rows, "row")})`];
    if (result.pages > 1) parts.push(`${result.pages} pages`);
    if (result.warning) parts.push(result.warning);
    parts.push(`in Downloads/${DOWNLOAD_FOLDER}`);
    setStatus(parts.join(" · "), "ok");
    await refreshInfo();
  } catch (err) {
    setStatus(friendlyError(err), "error");
  } finally {
    // Restore the button label and clear the downloading state either way.
    setBusy(false);
    setDownloadingState(false);
  }
}

/* Read the rows on the current PMRS page and send them to the matching table:

     sync_type = "rejection" -> npca_validation_logs (status + reject reason)
     sync_type = "approved"  -> pmrs_approved_gold  (patient + visit details)

   The Excel export path above is not touched. */
function primaryDiagnosisText(detail) {
  const primary = detail && detail.diagnoses && detail.diagnoses.primary;
  if (!Array.isArray(primary)) return "";
  return primary
    .map((item) => {
      if (!item) return "";
      if (item.code && item.description) {
        return `${item.code}. ${item.description}`;
      }
      return String(item.text || item.description || item.code || "").trim();
    })
    .filter(Boolean)
    .join(" | ");
}

function rejectionRecord(record) {
  return {
    pt_code: String(record.pt_code || "").trim(),
    status: String(record.status || "UNKNOWN").trim(),
    pmrs_reject_reason: String(record.pmrs_reject_reason || "").trim(),
    detailed_clinical_payload: record.detailed_clinical_payload || null,
  };
}

function approvedRecord(record) {
  return {
    pt_code: String(record.pt_code || "").trim(),
    patient_name: String(record.patient_name || "").trim(),
    gender: String(record.gender || "").trim(),
    age: String(record.age || "").trim(),
    visit_type: String(record.visit_type || "").trim(),
    icd10_diagnosis:
      String(record.icd10_diagnosis || "").trim() ||
      primaryDiagnosisText(record.detailed_clinical_payload),
    approval_date: String(record.approval_date || "").trim(),
    detailed_clinical_payload: record.detailed_clinical_payload || null,
  };
}

function syncLabel(syncType) {
  return syncType === "approved" ? "Approved" : "Rejection";
}

async function runSinglePageSync(syncType = "rejection") {
  if (!activeTab || busy) return;
  const label = syncLabel(syncType);
  const deep = deepScrapeEl.checked;
  setBusy(true);
  setSyncStatus(
    `Reading ${label.toLowerCase()} rows from PMRS` +
      (deep ? " · opening each detail modal…" : "…")
  );

  try {
    const result = await sendToTab(activeTab.id, {
      type: "PMRS_COLLECT_RECORDS",
      syncType,
      deepScrape: deepScrapeEl.checked,
    });
    if (!result || !result.ok) {
      throw new Error((result && result.error) || "Could not read the PMRS table.");
    }

    const records = Array.isArray(result.records) ? result.records : [];
    if (!records.length) {
      throw new Error("No rows with a patient-card number were found on this page.");
    }

    const payload =
      syncType === "approved"
        ? records.map(approvedRecord)
        : records.map(rejectionRecord);
    const withReason = payload.filter((record) => record.pmrs_reject_reason).length;

    setSyncStatus(
      `Saving ${plural(payload.length, "record")}` +
        (syncType === "rejection" && withReason
          ? ` · ${plural(withReason, "reject reason")}`
          : "") +
        ` to ${syncType === "approved" ? "pmrs_approved_gold" : "npca_validation_logs"}…`
    );
    const response = await fetch(PMRS_LOGS_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sync_type: syncType, records: payload }),
    });
    const data = await response.json().catch(() => ({}));

    if (!response.ok) {
      if (response.status === 404) {
        throw new Error(
          "Backend endpoint not found. Restart the API so it loads /api/save_pmrs_records."
        );
      }
      const detail = data && data.detail;
      const message = Array.isArray(detail)
        ? detail
            .map((item) => item && item.msg)
            .filter(Boolean)
            .join("; ")
        : detail;
      throw new Error(message || `Backend returned HTTP ${response.status}.`);
    }

    const saved = Number(data && data.saved);
    const count = Number.isFinite(saved) && saved > 0 ? saved : payload.length;
    setSyncStatus(
      `${label} records saved · ${plural(count, "record")}` +
        (data && data.table ? ` → ${data.table}` : ""),
      "ok"
    );
  } catch (err) {
    setSyncStatus(friendlyError(err), "error");
  } finally {
    setBusy(false);
  }
}

/* ── auto-sync: save current page + all subsequent pages ───────────────────────
   background.js owns the pagination loop so a page reload cannot break the
   message port. The popup only starts/stops the run and renders progress. */
async function runAutoSync(syncType = "rejection") {
  if (!activeTab || busy) return;
  const label = syncLabel(syncType);

  setSyncButtonState(true, syncType);
  setBusy(true, `Starting ${label.toLowerCase()} auto-sync — reading page 1…`);
  setSyncStatus(`Starting ${label.toLowerCase()} auto-sync…`);

  try {
    const response = await chrome.runtime.sendMessage({
      type: "PMRS_START_AUTO_SYNC",
      tabId: activeTab.id,
      syncType,
      deepScrape: deepScrapeEl.checked,
    });

    if (!response || !response.ok) {
      throw new Error((response && response.error) || "Could not start auto-sync.");
    }
    // The background answers immediately; progress and completion arrive as
    // PMRS_AUTO_SYNC_PROGRESS messages handled below.
  } catch (err) {
    setSyncButtonState(false);
    setBusy(false);
    setSyncStatus(friendlyError(err), "error");
  }
}

async function stopAutoSync() {
  if (!isSyncing) return;
  setSyncStatus("Stopping auto-sync…");

  try {
    await chrome.runtime.sendMessage({ type: "PMRS_STOP_AUTO_SYNC" });
  } catch (err) {
    setSyncStatus(friendlyError(err), "error");
  }

  // Safety net: if the worker never reports back (e.g. it was recycled),
  // release the UI instead of leaving the button stuck on "Stop sync".
  setTimeout(async () => {
    if (!isSyncing) return;
    try {
      const status = await chrome.runtime.sendMessage({
        type: "PMRS_GET_AUTO_SYNC_STATUS",
      });
      if (status && status.active) return; // still winding down
    } catch (_) {
      /* treat an unreachable worker as stopped */
    }
    setSyncButtonState(false);
    setBusy(false);
    setSyncStatus("Stopped.", "");
    setStatus("Stopped.", "");
  }, 8000);
}

function handleAutoSyncProgress(message) {
  const phase = message.phase;
  const type = message.syncType || syncingType || "rejection";
  const label = syncLabel(type);

  if (phase === "done" || phase === "stopped" || phase === "error") {
    setSyncButtonState(false);
    setBusy(false);

    const kind = phase === "error" ? "error" : phase === "done" ? "ok" : "";
    setSyncStatus(message.message || "Auto-sync finished.", kind);
    setStatus(message.message || "Auto-sync finished.", kind);

    // The portal is now parked on the last synced page — refresh the header
    // without overwriting the "Finished syncing…" status line.
    if (phase !== "error") refreshInfo(true);
    return;
  }

  const totalRows = Number(message.totalRows) || 0;
  const suffix = totalRows ? ` · ${plural(totalRows, "record")} ${label.toLowerCase()} synced` : "";

  if (phase === "deep") {
    const total = Number(message.total) || 0;
    const processed = Number(message.processed) || 0;
    const ptCode = message.pt_code ? ` · ${message.pt_code}` : "";
    const text =
      `Deep scrape — page ${message.page || 1} · row ${processed}` +
      (total ? `/${total}` : "") +
      ptCode +
      (message.error ? ` · skipped: ${message.error}` : "");
    setSyncStatus(text);
    setStatus(text);
    return;
  }

  if (phase === "starting") {
    setSyncStatus(message.message || `Starting ${label.toLowerCase()} auto-sync…`);
    setStatus(message.message || `Starting ${label.toLowerCase()} auto-sync…`);
    return;
  }

  if (phase === "saving") {
    const text = message.message || `Saving page ${message.page}…`;
    setSyncStatus(text);
    setStatus(text);
    return;
  }

  if (phase === "synced") {
    const text = `${label} sync — page ${message.page}…${suffix}`;
    setSyncStatus(text, "ok");
    setStatus(text);
    return;
  }

  if (phase === "loading") {
    const text = message.message || `Waiting for page ${message.page}…`;
    setSyncStatus(text);
    setStatus(text);
  }
}

/* Called when 'Download all pages' is clicked:
   - if idle → start the export (the button label flips inside runExport);
   - if downloading → tell content.js to break its loop; partial data is
     still saved to the file before the loop exits. */
function onDownloadAllClick() {
  if (isDownloading) {
    if (!activeTab) return;
    sendToTab(activeTab.id, { type: "PMRS_STOP_DOWNLOAD" }).catch(() => {
      /* export keeps running if the tab is gone — nothing else to do */
    });
    setStatus("Stop requested — saving collected data…");
    return;
  }
  runExport(allPagesEl.checked);
}

async function runDiagnose() {
  if (!activeTab || busy) return;
  setBusy(true, "Diagnosing the page…");

  try {
    const report = await sendToTab(activeTab.id, { type: "PMRS_DIAGNOSE" });
    if (!report || !report.ok) throw new Error((report && report.error) || "Diagnosis failed.");

    diagEl.hidden = false;
    diagEl.textContent = JSON.stringify(report, null, 2);
    console.log("[PMRS -> Excel] diagnosis", report);

    const selected = report.selected;
    setInfo(
      selected
        ? `${selected.rows} rows x ${selected.columns} columns`
        : "No table selected",
      report.next
        ? `Next button: ${report.next.label}`
        : "No Next button detected"
    );
    setStatus(
      report.next
        ? `Diagnosis ready — Next: ${report.next.label}${report.next.disabled ? " (disabled)" : ""}. Details below + console.`
        : "Diagnosis ready — no Next pagination found. Details below + console.",
      "ok"
    );
  } catch (err) {
    setStatus(friendlyError(err), "error");
  } finally {
    setBusy(false);
  }
}

/* ── minimize ────────────────────────────────────────────── */
function toggleMinimize() {
  const minimized = document.body.classList.toggle("is-minimized");
  minimizeBtn.setAttribute("aria-pressed", String(minimized));
  minimizeBtn.title = minimized ? "Restore" : "Minimize";
  minimizeBtn.setAttribute("aria-label", minimized ? "Restore" : "Minimize");
  minimizeBtn.innerHTML = minimized
    ? '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>'
    : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M5 12h14"/></svg>';
}

/* ── live progress ─────────────────────────────────────────
   PMRS_PROGRESS comes from content.js (Excel export loop).
   PMRS_AUTO_SYNC_PROGRESS comes from background.js (save-to-logs loop). */
chrome.runtime.onMessage.addListener((message) => {
  if (!message) return;

  if (message.type === "PMRS_AUTO_SYNC_PROGRESS") {
    handleAutoSyncProgress(message);
    return;
  }

  if (message.type !== "PMRS_PROGRESS") return;
  if (message.phase === "done") {
    setStatus(message.message || "Done.", message.warning ? "error" : "ok");
    return;
  }
  setStatus(
    `Page ${message.page} — ${plural(message.rows, "row")} · ${plural(message.totalRows, "row")} total…`
  );
});

/* If the popup is reopened while background.js is still syncing, restore the
   in-progress UI instead of showing an idle "Ready" state. */
async function refreshAutoSyncState() {
  try {
    const status = await chrome.runtime.sendMessage({
      type: "PMRS_GET_AUTO_SYNC_STATUS",
    });

    if (!status || !status.active) return false;

    const type = status.syncType === "approved" ? "approved" : "rejection";
    const label = syncLabel(type);
    const deep = status.deepScrape ? "deep " : "";
    setSyncButtonState(true, type);
    setBusy(true);
    setSyncStatus(
      `${label} ${deep}auto-sync in progress — ${plural(status.pagesSynced || 0, "page")} synced…`
    );
    setStatus(
      `${label} ${deep}auto-sync in progress — ${plural(status.pagesSynced || 0, "page")} synced…`
    );
    return true;
  } catch (_) {
    /* background not reachable — no sync running */
    return false;
  }
}

/* ── wiring ──────────────────────────────────────────────── */
document.addEventListener("DOMContentLoaded", async () => {
  const prefs = await loadPrefs();
  includeHeaderEl.checked = Boolean(prefs.includeHeader);
  allPagesEl.checked = Boolean(prefs.allPages);
  deepScrapeEl.checked = Boolean(prefs.deepScrape);

  includeHeaderEl.addEventListener("change", savePrefs);
  allPagesEl.addEventListener("change", savePrefs);
  deepScrapeEl.addEventListener("change", savePrefs);

  downloadAllBtn.addEventListener("click", onDownloadAllClick);
  downloadPageBtn.addEventListener("click", () => runExport(false));
  diagnoseBtn.addEventListener("click", runDiagnose);

  SYNC_BUTTONS.forEach(({ type, el: button }) => {
    button.addEventListener("click", () => {
      if (isSyncing) {
        // Only the running action is enabled, so this stops that run.
        stopAutoSync();
        return;
      }
      if (allPagesEl.checked) {
        runAutoSync(type);
      } else {
        runSinglePageSync(type);
      }
    });
  });

  minimizeBtn.addEventListener("click", toggleMinimize);

  activeTab = await getActiveTab();
  if (!activeTab) {
    setInfo("No active tab", "Open PMRS and try again.", true);
    setStatus("No active tab found.", "error");
    setBusy(false);
    return;
  }

  if (!isPmrs(activeTab.url)) {
    setInfo("Not on PMRS", "Open https://pmrs.digitalsp.gov.kh and reopen this popup.", true);
    setStatus("This extension only runs on pmrs.digitalsp.gov.kh.", "error");
    [downloadAllBtn, downloadPageBtn, diagnoseBtn, ...SYNC_BUTTONS.map((item) => item.el)].forEach(
      (btn) => (btn.disabled = true)
    );
    return;
  }

  const syncInProgress = await refreshAutoSyncState();
  if (!syncInProgress) await refreshInfo();
});
