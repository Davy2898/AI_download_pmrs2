/* ────────────────────────────────────────────────────────────────────────────
   PMRS -> Excel v1.0 — background service worker

   Owns the auto-sync ("Save to logs" + "Go through all pages") state machine
   so it survives the PMRS portal reloading the page after "Next" is clicked.

     popup      ──PMRS_START_AUTO_SYNC──────▶ background
     background ──PMRS_SYNC_SCRAPE─────────▶ content  (records + hasNext + sig)
     background ──POST /api/save_pmrs_records──▶ backend
     background ──PMRS_SYNC_NEXT───────────▶ content  (click the pager)
     chrome.tabs.onUpdated + signature poll  (wait for the next page to load)
     background ──PMRS_AUTO_SYNC_PROGRESS───▶ popup

   The content script never replies after navigation. It sends the records and
   clicks "Next" in separate messages, so a reload can never close a port that
   still owes the background a response.

   Progress is mirrored into chrome.storage.session (Chrome 102+) so a popup
   reopened mid-run can restore the "sync in progress" UI.
   ──────────────────────────────────────────────────────────────────────────── */
"use strict";

const PMRS_LOGS_URL = "http://localhost:8000/api/save_pmrs_records";
const STATE_KEY = "pmrsAutoSyncState";

const MAX_PAGES = 300;
const PAGE_TIMEOUT_MS = 45000;
const POLL_MS = 600;
const DEEP_SCRAPE_PAGE_TIMEOUT_MS = 10 * 60 * 1000;

/* In-memory state of the single running sync (one at a time). */
let runningTabId = null;
let runningSyncType = "rejection"; // "rejection" | "approved"
let runningDeepScrape = false;
let stopRequested = false;

function normaliseSyncType(value) {
  return value === "approved" ? "approved" : "rejection";
}

/* tabId -> async function that checks whether the table has changed. */
const pageWaiters = new Map();

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/* ══════════════════════════════════════════════════════════════════════════
   STATE (survives service-worker restarts within this browser session)
   ══════════════════════════════════════════════════════════════════════════ */
async function loadState() {
  try {
    const data = await chrome.storage.session.get(STATE_KEY);
    return data[STATE_KEY] || null;
  } catch (_) {
    return null;
  }
}

async function saveState(state) {
  try {
    await chrome.storage.session.set({ [STATE_KEY]: state });
  } catch (_) {
    /* storage is a convenience only — the running loop owns the truth */
  }
}

async function clearState() {
  try {
    await chrome.storage.session.remove(STATE_KEY);
  } catch (_) {
    /* ignore */
  }
}

/* Fire-and-forget progress message to the popup. If the popup is closed the
   callback swallows the "Receiving end does not exist" error. */
function broadcast(payload) {
  try {
    chrome.runtime.sendMessage(
      { type: "PMRS_AUTO_SYNC_PROGRESS", ...payload },
      () => {
        void chrome.runtime.lastError;
      }
    );
  } catch (_) {
    /* popup closed */
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   TAB MESSAGING
   ══════════════════════════════════════════════════════════════════════════ */
function sendToTab(tabId, message, timeoutMs = 5000) {
  return new Promise((resolve, reject) => {
    let settled = false;

    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      reject(new Error("Timed out waiting for the content script."));
    }, timeoutMs);

    chrome.tabs.sendMessage(tabId, message, (response) => {
      // Always read lastError, even after a timeout, so Chrome does not log
      // "Unchecked runtime.lastError" to the service worker console.
      const err = chrome.runtime.lastError;

      if (settled) return;
      settled = true;
      clearTimeout(timer);

      if (err) {
        reject(new Error(err.message));
        return;
      }
      resolve(response);
    });
  });
}

/* The manifest injects content.js on every PMRS load, but a tab that was open
   before the extension was installed has no listener. Inject on demand. */
async function ensureContentScript(tabId) {
  try {
    const pong = await sendToTab(tabId, { type: "PMRS_PING" });
    if (pong && pong.ok) return true;
  } catch (_) {
    /* not injected yet — fall through */
  }

  try {
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["content.js"],
    });
    await sleep(300);
    const pong = await sendToTab(tabId, { type: "PMRS_PING" });
    return Boolean(pong && pong.ok);
  } catch (_) {
    return false;
  }
}

async function readPageState(tabId) {
  try {
    const state = await sendToTab(tabId, { type: "PMRS_SYNC_STATE" });
    return state && state.ok ? state : null;
  } catch (_) {
    /* content script is mid-navigation */
    return null;
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   WAIT FOR THE NEXT PAGE
   Two signals race:
     • chrome.tabs.onUpdated  → full reload finished
     • signature polling      → Vue re-rendered the table without a reload
   Both funnel into check(); the signature must be different from the previous
   page and stable for two polls before the loop continues.
   ══════════════════════════════════════════════════════════════════════════ */
function waitForPageChange(tabId, previousSignature, timeoutMs) {
  return new Promise((resolve) => {
    let settled = false;
    let pollTimer = null;
    let timeoutTimer = null;
    let candidateSignature = null;
    let stablePolls = 0;

    const finish = (changed) => {
      if (settled) return;
      settled = true;
      if (pollTimer) clearInterval(pollTimer);
      if (timeoutTimer) clearTimeout(timeoutTimer);
      pageWaiters.delete(tabId);
      resolve(changed);
    };

    const check = async () => {
      if (settled) return;

      if (stopRequested) {
        finish(false);
        return;
      }

      const state = await readPageState(tabId);
      if (!state || !state.signature || state.signature === previousSignature) {
        candidateSignature = null;
        stablePolls = 0;
        return;
      }

      if (state.signature !== candidateSignature) {
        candidateSignature = state.signature;
        stablePolls = 1;
        return;
      }

      stablePolls += 1;
      if (stablePolls >= 2) {
        await sleep(250); // let the table finish painting
        finish(true);
      }
    };

    pageWaiters.set(tabId, check);
    pollTimer = setInterval(() => {
      check().catch(() => {});
    }, POLL_MS);
    timeoutTimer = setTimeout(() => finish(false), timeoutMs);

    // If the DOM already changed, don't wait for the first interval.
    setTimeout(() => {
      check().catch(() => {});
    }, 50);
  });
}

/* after a full reload, let the freshly injected content script wake up */
chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (changeInfo.status !== "complete") return;
  const check = pageWaiters.get(tabId);
  if (!check) return;
  setTimeout(() => {
    check().catch(() => {});
  }, 400);
});

/* ══════════════════════════════════════════════════════════════════════════
   BACKEND
   The fetch runs here, not in the content script: extension-origin requests
   use host_permissions and are not subject to the PMRS page's CORS policy.
   ══════════════════════════════════════════════════════════════════════════ */
async function postRecords(records, syncType = "rejection") {
  const type = normaliseSyncType(syncType);
  const payload = records
    .map((record) =>
      type === "approved"
        ? {
            pt_code: String(record.pt_code || "").trim(),
            patient_name: String(record.patient_name || "").trim(),
            gender: String(record.gender || "").trim(),
            age: String(record.age || "").trim(),
            visit_type: String(record.visit_type || "").trim(),
            icd10_diagnosis: String(record.icd10_diagnosis || "").trim(),
            approval_date: String(record.approval_date || "").trim(),
            detailed_clinical_payload:
              record.detailed_clinical_payload || null,
          }
        : {
            pt_code: String(record.pt_code || "").trim(),
            status: String(record.status || "UNKNOWN").trim(),
            pmrs_reject_reason: String(record.pmrs_reject_reason || "").trim(),
            detailed_clinical_payload:
              record.detailed_clinical_payload || null,
          }
    )
    .filter((record) => record.pt_code);

  if (!payload.length) return { ok: true, saved: 0, syncType: type };

  try {
    const response = await fetch(PMRS_LOGS_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sync_type: type, records: payload }),
    });
    const data = await response.json().catch(() => ({}));

    if (!response.ok) {
      if (response.status === 404) {
        return {
          ok: false,
          error:
            "Backend endpoint not found. Restart the API so it loads /api/save_pmrs_records.",
        };
      }
      const detail = data && data.detail;
      const message = Array.isArray(detail)
        ? detail
            .map((item) => item && item.msg)
            .filter(Boolean)
            .join("; ")
        : detail;
      return {
        ok: false,
        error: message || `Backend returned HTTP ${response.status}.`,
      };
    }

    const saved = Number(data && data.saved);
    return {
      ok: true,
      syncType: type,
      table: (data && data.table) || "",
      saved: Number.isFinite(saved) && saved > 0 ? saved : payload.length,
    };
  } catch (err) {
    return {
      ok: false,
      error: `Could not reach the backend (${(err && err.message) || err}).`,
    };
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   THE PAGINATION LOOP
   One iteration = scrape → POST → click Next → wait for the new table.
   ══════════════════════════════════════════════════════════════════════════ */
async function finishSync(phase, message) {
  const state = await loadState();
  const pages = (state && state.pagesSynced) || 0;
  const totalRows = (state && state.totalRecords) || 0;
  const syncType = normaliseSyncType((state && state.syncType) || runningSyncType);
  const table = syncType === "approved" ? "pmrs_approved_gold" : "npca_validation_logs";

  await clearState();

  broadcast({
    phase,
    pages,
    totalRows,
    syncType,
    table,
    error: phase === "error",
    message: `${message} · destination: ${table}`,
  });
}

async function runLoop(tabId, syncType, deepScrape = false) {
  let previousSignature = null;
  let pages = 0;
  let totalRows = 0;
  let samePageRetries = 0;

  const ready = await ensureContentScript(tabId);
  if (!ready) {
    await finishSync(
      "error",
      "Content script is not available. Reload the PMRS page and try again."
    );
    return;
  }

  for (let guard = 0; guard < MAX_PAGES; guard += 1) {
    if (stopRequested) {
      await finishSync("stopped", `Stopped after syncing ${pages} page(s).`);
      return;
    }

    /* 1 ── scrape the current page ------------------------------------- */
    const scrapeMessage = deepScrape
      ? { type: "PMRS_SYNC_SCRAPE_DEEP", maxRows: 200 }
      : { type: "PMRS_SYNC_SCRAPE" };
    const scrapeTimeout = deepScrape
      ? DEEP_SCRAPE_PAGE_TIMEOUT_MS
      : 5000;

    let scraped = null;
    try {
      scraped = await sendToTab(tabId, scrapeMessage, scrapeTimeout);
    } catch (_) {
      // A reload may still be settling; retry once.
      await sleep(900);
      try {
        scraped = await sendToTab(tabId, scrapeMessage, scrapeTimeout);
      } catch (err) {
        await finishSync(
          "error",
          `Could not read the PMRS page: ${err.message}`
        );
        return;
      }
    }

    if (!scraped || !scraped.ok) {
      await finishSync(
        "error",
        (scraped && scraped.error) || "Could not read the PMRS table."
      );
      return;
    }

    const records = Array.isArray(scraped.records) ? scraped.records : [];
    if (!records.length) {
      await finishSync(
        "done",
        `Finished syncing ${pages} page(s) · ${totalRows} record(s).`
      );
      return;
    }

    /* 2 ── new page? ---------------------------------------------------- */
    if (scraped.signature && scraped.signature === previousSignature) {
      // The click did not move us anywhere. Try once more, then give up.
      samePageRetries += 1;
      if (samePageRetries >= 2) {
        await finishSync(
          "done",
          `Finished syncing ${pages} page(s) · ${totalRows} record(s).`
        );
        return;
      }
    } else {
      samePageRetries = 0;
      previousSignature = scraped.signature;
      pages += 1;
      totalRows += records.length;

      broadcast({
        phase: "saving",
        syncType,
        page: pages,
        rows: records.length,
        totalRows,
        message: `Saving page ${pages} (${records.length} records)…`,
      });

      /* 3 ── POST this page to the backend ----------------------------- */
      const result = await postRecords(records, syncType);
      if (!result.ok) {
        await finishSync("error", `Backend error: ${result.error}`);
        return;
      }

      await saveState({
        active: true,
        tabId,
        syncType,
        deepScrape,
        pagesSynced: pages,
        totalRecords: totalRows,
        phase: "syncing",
        message: `Synced page ${pages}`,
        updatedAt: Date.now(),
      });

      broadcast({
        phase: "synced",
        syncType,
        page: pages,
        rows: records.length,
        totalRows,
        saved: result.saved,
        table: result.table,
        message: `Syncing page ${pages}… · ${totalRows} record(s) synced`,
      });
    }

    /* 4 ── last page? --------------------------------------------------- */
    if (!scraped.hasNext) {
      await finishSync(
        "done",
        `Finished syncing ${pages} page(s) · ${totalRows} record(s).`
      );
      return;
    }

    /* 5 ── click Next (the reply arrives before the click) -------------- */
    let clickResult = null;
    try {
      clickResult = await sendToTab(tabId, { type: "PMRS_SYNC_NEXT" });
    } catch (_) {
      // Expected when the click triggers a full navigation.
    }

    if (clickResult && clickResult.clicked === false) {
      await finishSync(
        "done",
        `Finished syncing ${pages} page(s) · ${totalRows} record(s).`
      );
      return;
    }

    broadcast({
      phase: "loading",
      page: pages + 1,
      totalRows,
      message: `Waiting for page ${pages + 1} to load…`,
    });

    /* 6 ── wait for the table to be replaced ---------------------------- */
    const changed = await waitForPageChange(
      tabId,
      scraped.signature,
      PAGE_TIMEOUT_MS
    );

    if (stopRequested) {
      await finishSync("stopped", `Stopped after syncing ${pages} page(s).`);
      return;
    }

    if (!changed) {
      await finishSync(
        "done",
        `Finished syncing ${pages} page(s) · ${totalRows} record(s).`
      );
      return;
    }
  }

  await finishSync(
    "done",
    `Stopped at the ${MAX_PAGES}-page safety limit.`
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   START / STOP
   ══════════════════════════════════════════════════════════════════════════ */
async function startAutoSync(tabId, syncType = "rejection", deepScrape = false) {
  if (runningTabId !== null) {
    broadcast({
      phase: "error",
      error: true,
      syncType: normaliseSyncType(syncType),
      message: "Auto-sync is already running.",
    });
    return;
  }

  runningTabId = tabId;
  runningSyncType = normaliseSyncType(syncType);
  runningDeepScrape = Boolean(deepScrape);
  stopRequested = false;

  await saveState({
    active: true,
    tabId,
    syncType: runningSyncType,
    deepScrape: runningDeepScrape,
    pagesSynced: 0,
    totalRecords: 0,
    phase: "starting",
    message: "Starting auto-sync…",
    updatedAt: Date.now(),
  });

  broadcast({
    phase: "starting",
    syncType: runningSyncType,
    deepScrape: runningDeepScrape,
    pages: 0,
    totalRows: 0,
    message: `Starting ${runningSyncType === "approved" ? "approved" : "rejection"} ${
      runningDeepScrape ? "deep " : ""
    }auto-sync…`,
  });

  try {
    await runLoop(tabId, runningSyncType, runningDeepScrape);
  } catch (err) {
    await finishSync(
      "error",
      `Auto-sync failed: ${(err && err.message) || err}`
    );
  } finally {
    runningTabId = null;
    runningDeepScrape = false;
    stopRequested = false;
  }
}

async function stopAutoSync() {
  stopRequested = true;

  const state = await loadState();
  const tabId = (state && state.tabId) ?? runningTabId;

  if (typeof tabId === "number") {
    try {
      await sendToTab(tabId, { type: "PMRS_SYNC_STOP" });
    } catch (_) {
      /* the tab may be navigating; the loop's stop flag is enough */
    }
  }

  return { ok: true };
}

/* ══════════════════════════════════════════════════════════════════════════
   MESSAGE ROUTER
   ══════════════════════════════════════════════════════════════════════════ */
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  const type = message && message.type;

  if (type === "PMRS_START_AUTO_SYNC") {
    const tabId = Number(message.tabId);
    const syncType = normaliseSyncType(message.syncType);
    const deepScrape = Boolean(message.deepScrape);

    if (!Number.isFinite(tabId) || tabId < 0) {
      sendResponse({ ok: false, error: "No valid PMRS tab supplied." });
      return;
    }

    if (runningTabId !== null) {
      sendResponse({ ok: false, error: "Auto-sync is already running." });
      return;
    }

    // Answer immediately. The popup must not keep its port open while the
    // service worker runs the long pagination loop — that is exactly what
    // caused "The message port closed before a response was received."
    sendResponse({ ok: true, started: true, syncType, deepScrape });
    startAutoSync(tabId, syncType, deepScrape).catch(() => {});
    return;
  }

  if (type === "PMRS_STOP_AUTO_SYNC") {
    stopAutoSync()
      .then((result) => sendResponse(result))
      .catch((error) =>
        sendResponse({ ok: false, error: String((error && error.message) || error) })
      );
    return true;
  }

  if (type === "PMRS_GET_AUTO_SYNC_STATUS") {
    if (runningTabId === null) {
      // No loop in this worker instance → any stored "active" flag is stale.
      loadState()
        .then(async (state) => {
          if (state && state.active) await clearState();
          sendResponse({
            ok: true,
            active: false,
            syncType: normaliseSyncType(state && state.syncType),
            deepScrape: Boolean(state && state.deepScrape),
            pagesSynced: 0,
            totalRecords: 0,
            phase: "idle",
            message: "",
          });
        })
        .catch(() => sendResponse({ ok: true, active: false }));
      return true;
    }

    loadState()
      .then((state) =>
        sendResponse({
          ok: true,
          active: true,
          syncType: normaliseSyncType((state && state.syncType) || runningSyncType),
          deepScrape: Boolean((state && state.deepScrape) ?? runningDeepScrape),
          pagesSynced: (state && state.pagesSynced) || 0,
          totalRecords: (state && state.totalRecords) || 0,
          phase: (state && state.phase) || "syncing",
          message: (state && state.message) || "",
        })
      )
      .catch(() =>
        sendResponse({
          ok: true,
          active: true,
          syncType: normaliseSyncType(runningSyncType),
          deepScrape: runningDeepScrape,
          pagesSynced: 0,
          totalRecords: 0,
          phase: "syncing",
          message: "",
        })
      );
    return true;
  }
});
