#!/usr/bin/env python3
"""
══════════════════════════════════════════════════════════════════════════
PMRS REJECTED CLAIMS — DEEP DETAIL SCRAPER  →  SQLite
══════════════════════════════════════════════════════════════════════════

For every row of the rejected-claims table this script:

  1. clicks the row's "View details" button (blue eye icon),
  2. waits for the full-screen detail dialog to finish rendering,
  3. extracts every card of that dialog (claim, patient, approval history,
     clinical visit, vitals, history, daily visit + examination, SOAP,
     invoice, equity package),
  4. upserts it into ``patients_data.db`` (parameterized SQL, one
     transaction per patient),
  5. closes the dialog with the ✕ button, waits for it to disappear and
     moves on to the next row.

The PMRS detail dialog is a card layout built with Nuxt UI + Tailwind:

  * ``<strong>label</strong> ៖ value``           (claim / patient cards)
  * ``<h3>label៖</h3><p|div>value</p|div>``      (clinical, vitals, exams)
  * ``<label>label៖</label><article>value``      (SOAP note)
  * ``<div>label៖ <span>value</span></div>``     (invoice summary rows)
  * ``<h3>item</h3> … x1 … 40,000 ៛``            (invoice line items)
  * ``<ol><li>date / status / reason / actor``   (approval history)

Extraction runs page-side in one ``page.evaluate`` call, which is fast and
keeps every value readable even when it is nested several divs deep.
Everything is stored twice: mapped into typed columns *and* verbatim into
``modal_fields`` / ``modal_tables`` / ``modal_cards`` so nothing is lost.

Usage
-----
    python scrape_rejected_claims.py                    # normal run
    python scrape_rejected_claims.py --max-rows 5       # test on 5 rows
    python scrape_rejected_claims.py --max-pages 1      # test on 1 page
    python scrape_rejected_claims.py --start-page 151   # resume at page 151
    python scrape_rejected_claims.py --reset            # wipe DB first
    python scrape_rejected_claims.py --dump-html 3      # save 3 dialogs
    python scrape_rejected_claims.py --no-excel         # DB only
    python scrape_rejected_claims.py --headless         # no window

Rows are upserted by claim number, so re-running (with or without
``--start-page``) never duplicates data.

The browser profile (``.pmrs_profile``) is persistent, so the PMRS login
survives between runs. On startup the script waits (no terminal input needed)
until the claims table is visible, so you can log in in the browser window;
scraping then starts automatically. Use ``--login-timeout`` to change the
wait and ``--no-pause`` to skip it.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sqlite3
import sys
import time
from pathlib import Path
from typing import Any, Iterable, Sequence

from playwright.sync_api import (
    Locator,
    Page,
    sync_playwright,
)


def _configure_console() -> None:
    """Make emoji / Khmer text printable on Windows consoles and when piped."""
    for stream in (sys.stdout, sys.stderr):
        try:
            encoding = (getattr(stream, "encoding", "") or "").lower()
            if encoding and encoding.replace("-", "") not in ("utf8", "utf8mb4"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


_configure_console()

# ══════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════
BASE_URL = "https://pmrs.digitalsp.gov.kh/claims/rejected/list"

PROJECT_DIR = Path(__file__).resolve().parent
DB_PATH = PROJECT_DIR / "patients_data.db"
PROFILE_DIR = PROJECT_DIR / ".pmrs_profile"
DEBUG_DIR = PROJECT_DIR / "debug"

ROW_SELECTOR = "table tbody tr"

# Modal containers, most specific first. PMRS uses role="dialog".
MODAL_SELECTORS = (
    "[role='dialog']",
    ".modal.show",
    ".modal.in",
    ".modal[style*='display: block']",
    ".offcanvas.show",
    ".drawer.show",
    ".modal-dialog",
)

CLOSE_SELECTORS = (
    "button.btn-close",
    ".modal-header .btn-close",
    ".modal-header button.close",
    ".modal-header .close",
    "button[aria-label='Close']",
    "button[aria-label='close']",
    "[data-bs-dismiss='modal']",
    "[data-dismiss='modal']",
    ".offcanvas-header .btn-close",
    "button:has-text('✕')",
    "button:has-text('×')",
    "button:has-text('X')",
)

# ── section hints (card titles) ───────────────────────────────────────────
CLAIM_SECTIONS = ("ករណីបានស្នើសុំ", "claim")
PATIENT_SECTIONS = ("ព័ត៌មានអ្នកជំងឺ", "patient")
CLINICAL_SECTIONS = ("ចូលមកពិនិត្យ", "ព័ត៌មានការពិនិត្យព្យាបាល")
VITAL_SECTIONS = ("វាយតម្លៃជំងឺដំបូង", "vital")
SOAP_SECTIONS = ("SOAP", "soap")
HISTORY_SECTIONS = ("សវតារោគ", "history")
EXAM_SECTIONS = ("Department (Ward)", "department")
WARD_SECTIONS = ("Department (Ward)",)
INVOICE_SECTIONS = ("Invoice", "វិក្កយបត្រ", "វិក័យប័ត្រ", "billing")

# ── label aliases (Khmer + English) ───────────────────────────────────────
CLAIM_NO = ("លេខករណី", "claim number", "claim no", "claim id")
REQUEST_DATE = ("ថ្ងៃស្នើសុំ", "កាលបរិច្ឆេទស្នើសុំ", "request date", "submission date")
DISTRICT = ("ស្រុកប្រតិបត្តិ", "district")
FACILITY = ("មូលដ្ឋានសុខាភិបាល", "facility", "hospital", "health facility")
FACILITY_LEVEL = ("កម្រិតមូលដ្ឋានសុខាភិបាល", "facility level")
STATUS = ("ស្ថានភាព", "status", "claim status")
REASON = ("មូលហេតុ", "ហេតុផល", "reason", "rejection reason")

HEALTH_CARD_ID = ("លេខបណ្ណសុខភាព", "health card id", "health card", "card id")
EQUITY_CARD_ID = ("លេខបណ្ណសមធម៌", "equity card", "equity card id")
PHONE = (
    "លេខទូរសព្ទ",
    "លេខទូរស័ព្ទ",
    "ទូរសព្ទ",
    "ទូរស័ព្ទ",
    "phone",
    "phone number",
    "mobile",
)
PATIENT_NAME = ("ឈ្មោះអ្នកជំងឺ", "patient name", "full name", "name")
GENDER = ("ភេទ", "gender", "sex")
NATIONALITY = ("សញ្ជាតិ", "nationality")
DOB = ("ថ្ងៃខែឆ្នាំកំណើត", "date of birth", "dob", "birth date")
AGE = ("អាយុ", "age")
ADDRESS = ("អាសយដ្ឋាន", "address")

DEPARTMENT_IN = ("ផ្នែកចូល", "department", "department in")
DEPARTMENT_OUT = ("ផ្នែកចេញ", "department out")
ADMISSION_DATE = ("កាលបរិច្ឆេទចូល", "admission date", "admit date")
DISCHARGE_DATE = ("កាលបរិច្ឆេទចេញ", "discharge date")
DIAGNOSIS_IN = ("រោគវិនិច្ឆ័យចូល", "diagnosis in")
DIAGNOSIS_OUT = ("រោគវិនិច្ឆ័យចេញ", "diagnosis out")
VISIT_TYPE = ("ប្រភេទនៃការចូល", "visit type", "admission type")
DISCHARGE_TYPE = ("ប្រភេទនៃការចេញ", "discharge type")
DISEASE_STATUS = ("ស្ថានភាពជំងឺ", "disease status", "condition")
SERVICE_UNIT = ("ផ្នែកទទួលសេវា", "service unit")
APPOINTMENT_DATE = ("កាលបរិច្ឆេទណាត់ជួប", "appointment date")
ADMISSION_REASON = ("មូលហេតុចូល", "admission reason", "chief complaint")

BP = ("សម្ពាធឈាម", "bp", "blood pressure")
SYSTOLIC = ("systolic",)
DIASTOLIC = ("diastolic",)
PULSE = ("ចង្វាក់បេះដូង", "pulse", "heart rate")
TEMP = ("សីតុណ្ហភាព", "temp", "temperature")
SPO2 = ("អុកស៊ីហ្សែនក្នុងឈាម", "អុកស៊ីសែន", "spo2", "oxygen")
WEIGHT = ("ទម្ងន់", "weight")
HEIGHT = ("កម្ពស់", "height")
RESP_RATE = ("ចង្វាក់ដង្ហើម", "respiratory rate", "resp rate")
GLUCOSE = ("គ្លុយកូស", "glucose", "blood sugar")

SUBJECTIVE = ("subjective",)
OBJECTIVE = ("objective",)
ASSESSMENT = ("assessment",)
EVALUATION = ("evaluation",)
PLAN = ("plan",)
HPI = ("history of present illness", "present illness", "hpi")
CURRENT_MEDICATION = ("current medication",)
PAST_MEDICAL_HISTORY = ("past medical history",)
PAST_SURGICAL_HISTORY = ("past surgical history",)
ALLERGIES = ("allergies", "allergy")
IMMUNIZATIONS = ("immunizations", "immunization")
FAMILY_HISTORY = ("family history",)
SOCIAL_HISTORY = ("social history",)
SYMPTOMS = ("រោគសញ្ញា", "symptoms")

INVOICE_NO = ("លេខវិក្កយបត្រ", "លេខវិក័យប័ត្រ", "invoice number", "invoice no", "invoice #")
INVOICE_DATE = ("កាលបរិច្ឆេទ", "invoice date", "billing date")
INVOICE_AMOUNT = (
    "សរុបទាំងអស់",
    "តម្លៃសរុប ថ្លៃសេវាកម្ម",
    "តម្លៃសរុប",
    "amount",
    "total",
    "grand total",
)
INVOICE_DESCRIPTION = ("ថ្លៃសេវាកម្ម", "description", "note")

ICD_RE = re.compile(r"\b([A-Z][0-9]{2}(?:\.[0-9A-Za-z]+)?)\b")
AGE_RE = re.compile(r"\(([^)]*ឆ្នាំ[^)]*)\)")


# ══════════════════════════════════════════════════════════════════════════
# SQLITE SCHEMA
# ══════════════════════════════════════════════════════════════════════════
SCHEMA = """
CREATE TABLE IF NOT EXISTS patients (
    patient_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_key      TEXT    NOT NULL UNIQUE,
    nssf_id          TEXT,
    patient_name     TEXT    NOT NULL DEFAULT '',
    gender           TEXT    NOT NULL DEFAULT '',
    dob              TEXT    NOT NULL DEFAULT '',
    age              TEXT    NOT NULL DEFAULT '',
    phone            TEXT    NOT NULL DEFAULT '',
    health_card_id   TEXT    NOT NULL DEFAULT '',
    card_type        TEXT    NOT NULL DEFAULT '',
    created_at       TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at       TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS claims (
    claim_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    record_key       TEXT    NOT NULL UNIQUE,
    claim_no         TEXT,
    request_date     TEXT    NOT NULL DEFAULT '',
    facility         TEXT    NOT NULL DEFAULT '',
    status           TEXT    NOT NULL DEFAULT '',
    reason           TEXT    NOT NULL DEFAULT '',
    patient_id       INTEGER,
    source_page      INTEGER NOT NULL DEFAULT 0,
    source_row       INTEGER NOT NULL DEFAULT 0,
    list_row_json    TEXT    NOT NULL DEFAULT '{}',
    raw_json         TEXT    NOT NULL DEFAULT '{}',
    created_at       TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at       TEXT    NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (patient_id) REFERENCES patients (patient_id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS clinical_records (
    clinical_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    claim_id         INTEGER NOT NULL UNIQUE,
    department       TEXT    NOT NULL DEFAULT '',
    admission_date   TEXT    NOT NULL DEFAULT '',
    discharge_date   TEXT    NOT NULL DEFAULT '',
    diagnosis        TEXT    NOT NULL DEFAULT '',
    icd10            TEXT    NOT NULL DEFAULT '',
    treatment_type   TEXT    NOT NULL DEFAULT '',
    raw_json         TEXT    NOT NULL DEFAULT '{}',
    updated_at       TEXT    NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (claim_id) REFERENCES claims (claim_id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS vitals (
    vital_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    claim_id         INTEGER NOT NULL UNIQUE,
    bp               TEXT    NOT NULL DEFAULT '',
    systolic         TEXT    NOT NULL DEFAULT '',
    diastolic        TEXT    NOT NULL DEFAULT '',
    pulse            TEXT    NOT NULL DEFAULT '',
    temp             TEXT    NOT NULL DEFAULT '',
    spo2             TEXT    NOT NULL DEFAULT '',
    weight           TEXT    NOT NULL DEFAULT '',
    height           TEXT    NOT NULL DEFAULT '',
    resp_rate        TEXT    NOT NULL DEFAULT '',
    bmi              TEXT    NOT NULL DEFAULT '',
    glucose          TEXT    NOT NULL DEFAULT '',
    measured_at      TEXT    NOT NULL DEFAULT '',
    raw_json         TEXT    NOT NULL DEFAULT '{}',
    updated_at       TEXT    NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (claim_id) REFERENCES claims (claim_id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS soap_notes (
    soap_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    claim_id         INTEGER NOT NULL UNIQUE,
    subjective       TEXT    NOT NULL DEFAULT '',
    objective        TEXT    NOT NULL DEFAULT '',
    assessment       TEXT    NOT NULL DEFAULT '',
    plan             TEXT    NOT NULL DEFAULT '',
    chief_complaint  TEXT    NOT NULL DEFAULT '',
    hpi              TEXT    NOT NULL DEFAULT '',
    examination      TEXT    NOT NULL DEFAULT '',
    diagnosis        TEXT    NOT NULL DEFAULT '',
    icd10            TEXT    NOT NULL DEFAULT '',
    treatment_plan   TEXT    NOT NULL DEFAULT '',
    conclusion       TEXT    NOT NULL DEFAULT '',
    raw_json         TEXT    NOT NULL DEFAULT '{}',
    updated_at       TEXT    NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (claim_id) REFERENCES claims (claim_id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS invoices (
    invoice_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    claim_id         INTEGER NOT NULL,
    invoice_no       TEXT    NOT NULL DEFAULT '',
    invoice_date     TEXT    NOT NULL DEFAULT '',
    amount           TEXT    NOT NULL DEFAULT '',
    currency         TEXT    NOT NULL DEFAULT '',
    status           TEXT    NOT NULL DEFAULT '',
    description      TEXT    NOT NULL DEFAULT '',
    raw_json         TEXT    NOT NULL DEFAULT '{}',
    updated_at       TEXT    NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (claim_id) REFERENCES claims (claim_id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS invoice_items (
    item_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id       INTEGER NOT NULL,
    line_no          INTEGER NOT NULL DEFAULT 0,
    item             TEXT    NOT NULL DEFAULT '',
    quantity         TEXT    NOT NULL DEFAULT '',
    unit_price       TEXT    NOT NULL DEFAULT '',
    amount           TEXT    NOT NULL DEFAULT '',
    raw_json         TEXT    NOT NULL DEFAULT '{}',
    FOREIGN KEY (invoice_id) REFERENCES invoices (invoice_id) ON DELETE CASCADE
);

-- Approval / submission timeline shown in the dialog.
CREATE TABLE IF NOT EXISTS approval_history (
    event_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    claim_id         INTEGER NOT NULL,
    seq              INTEGER NOT NULL DEFAULT 0,
    event_time       TEXT    NOT NULL DEFAULT '',
    status           TEXT    NOT NULL DEFAULT '',
    reason           TEXT    NOT NULL DEFAULT '',
    actor            TEXT    NOT NULL DEFAULT '',
    FOREIGN KEY (claim_id) REFERENCES claims (claim_id) ON DELETE CASCADE
);

-- Equity-fund packages (ក្រុមគោលដៅ) attached to the claim.
CREATE TABLE IF NOT EXISTS equity_packages (
    package_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    claim_id         INTEGER NOT NULL,
    line_no          TEXT    NOT NULL DEFAULT '',
    name             TEXT    NOT NULL DEFAULT '',
    amount           TEXT    NOT NULL DEFAULT '',
    FOREIGN KEY (claim_id) REFERENCES claims (claim_id) ON DELETE CASCADE
);

-- Every label/value pair seen in the dialog, including unmapped ones.
CREATE TABLE IF NOT EXISTS modal_fields (
    field_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    claim_id         INTEGER NOT NULL,
    section          TEXT    NOT NULL DEFAULT '',
    label            TEXT    NOT NULL DEFAULT '',
    value            TEXT    NOT NULL DEFAULT '',
    FOREIGN KEY (claim_id) REFERENCES claims (claim_id) ON DELETE CASCADE
);

-- Every structured list rendered inside the dialog.
CREATE TABLE IF NOT EXISTS modal_tables (
    table_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    claim_id         INTEGER NOT NULL,
    section          TEXT    NOT NULL DEFAULT '',
    table_name       TEXT    NOT NULL DEFAULT '',
    headers_json     TEXT    NOT NULL DEFAULT '[]',
    rows_json        TEXT    NOT NULL DEFAULT '[]',
    row_count        INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (claim_id) REFERENCES claims (claim_id) ON DELETE CASCADE
);

-- Raw text of every card, so layout changes never lose data.
CREATE TABLE IF NOT EXISTS modal_cards (
    card_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    claim_id         INTEGER NOT NULL,
    seq              INTEGER NOT NULL DEFAULT 0,
    title            TEXT    NOT NULL DEFAULT '',
    text             TEXT    NOT NULL DEFAULT '',
    FOREIGN KEY (claim_id) REFERENCES claims (claim_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_patients_nssf      ON patients (nssf_id);
CREATE INDEX IF NOT EXISTS idx_claims_patient     ON claims (patient_id);
CREATE INDEX IF NOT EXISTS idx_claims_no          ON claims (claim_no);
CREATE INDEX IF NOT EXISTS idx_modal_fields_claim ON modal_fields (claim_id, section);
CREATE INDEX IF NOT EXISTS idx_modal_tables_claim ON modal_tables (claim_id, section);
CREATE INDEX IF NOT EXISTS idx_invoices_claim     ON invoices (claim_id);
CREATE INDEX IF NOT EXISTS idx_history_claim      ON approval_history (claim_id, seq);
CREATE INDEX IF NOT EXISTS idx_packages_claim     ON equity_packages (claim_id);
"""

# Columns added after the first schema version. init_db() ALTERs them in,
# so an older patients_data.db keeps working without a manual migration.
MIGRATIONS: dict[str, dict[str, str]] = {
    "patients": {
        "equity_card_id": "TEXT NOT NULL DEFAULT ''",
        "nationality": "TEXT NOT NULL DEFAULT ''",
        "address": "TEXT NOT NULL DEFAULT ''",
    },
    "claims": {
        "district": "TEXT NOT NULL DEFAULT ''",
        "facility_level": "TEXT NOT NULL DEFAULT ''",
        "payment_type": "TEXT NOT NULL DEFAULT ''",
        "error_type": "TEXT NOT NULL DEFAULT ''",
        "total_amount": "TEXT NOT NULL DEFAULT ''",
        "medical_amount": "TEXT NOT NULL DEFAULT ''",
        "non_medical_amount": "TEXT NOT NULL DEFAULT ''",
        "rejected_at": "TEXT NOT NULL DEFAULT ''",
        "rejected_by": "TEXT NOT NULL DEFAULT ''",
        "rejection_reason": "TEXT NOT NULL DEFAULT ''",
    },
    "clinical_records": {
        "department_out": "TEXT NOT NULL DEFAULT ''",
        "diagnosis_in": "TEXT NOT NULL DEFAULT ''",
        "diagnosis_out": "TEXT NOT NULL DEFAULT ''",
        "visit_type": "TEXT NOT NULL DEFAULT ''",
        "discharge_type": "TEXT NOT NULL DEFAULT ''",
        "disease_status": "TEXT NOT NULL DEFAULT ''",
        "service_unit": "TEXT NOT NULL DEFAULT ''",
        "appointment_date": "TEXT NOT NULL DEFAULT ''",
        "admission_reason": "TEXT NOT NULL DEFAULT ''",
    },
    "vitals": {},
    "soap_notes": {
        "evaluation": "TEXT NOT NULL DEFAULT ''",
        "current_medication": "TEXT NOT NULL DEFAULT ''",
        "past_medical_history": "TEXT NOT NULL DEFAULT ''",
        "past_surgical_history": "TEXT NOT NULL DEFAULT ''",
        "allergies": "TEXT NOT NULL DEFAULT ''",
        "immunizations": "TEXT NOT NULL DEFAULT ''",
        "family_history": "TEXT NOT NULL DEFAULT ''",
        "social_history": "TEXT NOT NULL DEFAULT ''",
        "symptoms": "TEXT NOT NULL DEFAULT ''",
        "examination_json": "TEXT NOT NULL DEFAULT '{}'",
        "soap_doctor": "TEXT NOT NULL DEFAULT ''",
        "soap_date": "TEXT NOT NULL DEFAULT ''",
    },
    "invoices": {
        "target_group": "TEXT NOT NULL DEFAULT ''",
        "target_group_amount": "TEXT NOT NULL DEFAULT ''",
    },
}


# ══════════════════════════════════════════════════════════════════════════
# BROWSER-SIDE EXTRACTOR
# ══════════════════════════════════════════════════════════════════════════
EXTRACT_JS = r"""
(() => {
  const clean = (s) =>
    String(s == null ? '' : s)
      .replace(/\u200b/g, '')
      .replace(/\u00a0/g, ' ')
      .replace(/[ \t]+/g, ' ')
      .trim();

  const multi = (s) =>
    String(s == null ? '' : s)
      .replace(/\u200b/g, '')
      .replace(/\u00a0/g, ' ')
      .replace(/\r/g, '')
      .split('\n')
      .map((line) => line.replace(/[ \t]+/g, ' ').trim())
      .filter((line) => line.length > 0)
      .join('\n')
      .trim();

  const visible = (el) => {
    if (!el || !el.isConnected) return false;
    if (!el.getClientRects().length) return false;
    const st = getComputedStyle(el);
    return st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0';
  };

  const NOISE = '.iconify, [class*="iconify"], svg, button, [role="button"]';

  const readText = (el) => {
    if (!el) return '';
    let clone;
    try {
      clone = el.cloneNode(true);
    } catch (err) {
      return '';
    }
    clone.querySelectorAll(NOISE).forEach((n) => n.remove());
    clone.querySelectorAll('br').forEach((n) => n.replaceWith('\n'));
    clone
      .querySelectorAll(
        'p, div, li, article, h1, h2, h3, h4, h5, h6, header, main, section, tr, fieldset, legend, label'
      )
      .forEach((n) => n.append('\n'));
    return multi(clone.textContent || '');
  };

  const dialogs = Array.from(document.querySelectorAll('[role="dialog"]')).filter(visible);
  const dialog = dialogs.length ? dialogs[dialogs.length - 1] : null;
  if (!dialog) return { ok: false, error: 'detail dialog not found' };

  const cardOf = (el) => {
    let node = el;
    while (node && node !== dialog) {
      const header = node.querySelector ? node.querySelector(':scope > header') : null;
      if (header) {
        const p = header.querySelector('p');
        const title = clean(readText(p || header).split('\n')[0]);
        if (title) return title;
      }
      if (node.tagName === 'FIELDSET') {
        const legend = node.querySelector(':scope > legend');
        if (legend) {
          const lines = readText(legend).split('\n').filter(Boolean);
          if (lines.length) return lines[lines.length - 1];
        }
      }
      node = node.parentElement;
    }
    return 'General';
  };

  const pairs = [];
  const seen = new Set();
  const add = (label, value, section, kind) => {
    label = clean(label)
      .replace(/^[\s\-•*]+/, '')
      .replace(/[\s:：៖]+$/, '')
      .trim();
    value = multi(value).replace(/^[:：៖\s]+/, '').trim();
    if (!label || !value) return;
    if (label.length > 120) return;
    if (/^(https?:|copy|ចម្លង)/i.test(label)) return;
    const key = section + '|' + label + '|' + value;
    if (seen.has(key)) return;
    seen.add(key);
    pairs.push({ section: section || 'General', label: label, value: value, kind: kind });
  };

  // 1. <strong>label</strong> ៖ value   (claim + patient cards)
  dialog.querySelectorAll('strong').forEach((st) => {
    if (!visible(st)) return;
    const label = clean(readText(st));
    if (!label) return;
    const parent = st.parentElement;
    if (!parent) return;
    let value = readText(parent);
    if (value.indexOf(label) === 0) value = value.slice(label.length);
    value = value.replace(/^[\s:：៖]+/, '').trim();
    add(label, value, cardOf(st), 'strong');
  });

  // 2. <h3>label៖</h3> + sibling value   (clinical, vitals, exams)
  dialog.querySelectorAll('h3').forEach((h) => {
    if (!visible(h)) return;
    const raw = readText(h);
    if (!raw || !/[:：៖]\s*$/.test(raw)) return;
    const label = raw.replace(/[:：៖]\s*$/, '').trim();
    let value = '';
    const sib = h.nextElementSibling;
    if (sib && visible(sib)) value = readText(sib);
    if (!value && h.parentElement) {
      value = readText(h.parentElement).replace(raw, '').trim();
    }
    add(label, value, cardOf(h), 'h3');
  });

  // 3. <label>label៖</label> + sibling / article   (SOAP, diagnoses, totals)
  dialog.querySelectorAll('label').forEach((lb) => {
    if (!visible(lb)) return;
    const raw = readText(lb);
    if (!raw) return;

    if (/Systolic/i.test(raw)) {
      const m = raw.match(
        /Systolic[^:：]*[:：]\s*([0-9]{2,3})[\s\S]*?Diastolic[^:：]*[:：]\s*([0-9]{2,3})/i
      );
      if (m) add('សម្ពាធឈាម', m[1] + '/' + m[2], cardOf(lb), 'bp');
      return;
    }

    const colon = /[:：៖]\s*$/.test(raw);
    const sib = lb.nextElementSibling;

    if (colon) {
      const label = raw.replace(/[:：៖]\s*$/, '').trim();
      let value = sib && visible(sib) ? readText(sib) : '';
      if (!value && lb.parentElement) {
        value = readText(lb.parentElement).replace(raw, '').trim();
      }
      add(label, value, cardOf(lb), 'label');
      return;
    }

    // Two <label> children inside a grid row: label + value.
    const parent = lb.parentElement;
    if (!parent || !/grid/.test(parent.className || '')) return;
    const labels = Array.from(parent.children).filter((c) => c.tagName === 'LABEL');
    if (labels.length === 2 && labels[0] === lb) {
      add(raw, readText(labels[1]), cardOf(lb), 'label-pair');
    }
  });

  // 4. <div>label៖ <span>value</span></div>   (invoice summary rows)
  const directText = (el) => {
    let out = '';
    el.childNodes.forEach((n) => {
      if (n.nodeType === 3) out += n.nodeValue;
    });
    return out;
  };
  dialog.querySelectorAll('div').forEach((el) => {
    if (!visible(el)) return;
    if (el.querySelector('strong, h1, h2, h3, h4, h5, h6, label, article')) return;
    const own = clean(directText(el));
    const m = own.match(/^([^:：៖]{1,120})[:：៖]\s*$/);
    if (!m) return;
    const value = readText(el).replace(/^[^:：៖]{1,120}[:：៖]\s*/, '').trim();
    add(m[1], value, cardOf(el), 'text');
  });

  // 5. Invoice line items: <div><h3>item</h3><div><p>x1</p><span>40,000 ៛</span></div></div>
  const items = [];
  const itemSeen = new Set();
  dialog.querySelectorAll('div').forEach((row) => {
    if (!visible(row)) return;
    const h3 = row.querySelector(':scope > h3');
    if (!h3 || !visible(h3)) return;
    const name = readText(h3);
    if (!name || /[:：៖]\s*$/.test(name)) return;
    const qtyEl = row.querySelector(':scope > div > p');
    const amtEl = row.querySelector(':scope > div > span');
    if (!qtyEl && !amtEl) return;
    const qty = qtyEl ? clean(qtyEl.innerText).replace(/^x\s*/i, '') : '';
    if (qty && !/^\d+(\.\d+)?$/.test(qty)) return;
    const amount = amtEl ? clean(amtEl.innerText) : '';
    if (!amount) return;
    const key = name + '|' + qty + '|' + amount;
    if (itemSeen.has(key)) return;
    itemSeen.add(key);
    items.push({ item: name, quantity: qty, amount: amount, section: cardOf(h3) });
  });

  // 6. Equity packages: three <h2> in a row (no, name, amount).
  const packages = [];
  dialog.querySelectorAll('div').forEach((row) => {
    const h2s = Array.from(row.querySelectorAll(':scope > div > h2'));
    if (h2s.length !== 3) return;
    const vals = h2s.map((h) => clean(h.innerText));
    if (!/^\d+$/.test(vals[0])) return;
    packages.push({ line_no: vals[0], name: vals[1], amount: vals[2], section: cardOf(row) });
  });

  // 7. Approval history: <ol><li> date / actor / status / reason
  const history = [];
  dialog.querySelectorAll('ol > li').forEach((li) => {
    if (!visible(li)) return;
    const statusEl = li.querySelector('h3');
    const status = statusEl ? clean(statusEl.innerText) : '';
    const spans = Array.from(li.querySelectorAll('div > span'))
      .map((s) => clean(s.innerText))
      .filter(Boolean);
    const eventTime = spans.find((s) => /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}/.test(s)) || '';
    const actor = spans.find((s) => s !== eventTime) || '';
    const reasonEl = li.querySelector('small');
    const reason = reasonEl ? clean(reasonEl.innerText) : '';
    if (!status && !eventTime && !reason) return;
    history.push({ event_time: eventTime, status: status, reason: reason, actor: actor });
  });

  // 8. Daily-visit meta: <h6>department</h6> + service + time
  const daily = [];
  dialog.querySelectorAll('h6').forEach((h6) => {
    if (!visible(h6)) return;
    let box = h6.closest('div.col-span-full');
    if (!box && h6.parentElement) box = h6.parentElement.parentElement;
    const spans = box
      ? Array.from(box.querySelectorAll('span.text-xs'))
          .map((s) => clean(s.innerText))
          .filter(Boolean)
      : [];
    daily.push({
      department: clean(h6.innerText),
      service: spans[0] || '',
      visit_time: spans[1] || '',
    });
  });

  // 9. Raw cards (title + full text) as a safety net.
  const cards = [];
  dialog.querySelectorAll('div, section, article').forEach((el) => {
    if (!visible(el)) return;
    const header = el.querySelector(':scope > header');
    if (!header) return;
    const p = header.querySelector('p');
    const title = clean(readText(p || header).split('\n')[0]);
    if (!title) return;
    cards.push({ title: title, text: readText(el) });
  });

  let targetGroup = '';
  cards.forEach((card) => {
    if (!targetGroup && card.title.indexOf('ក្រុមគោលដៅ') === 0) {
      targetGroup = card.title.replace(/^ក្រុមគោលដៅ\s*[:：៖]\s*/, '').trim();
    }
  });

  return {
    ok: true,
    url: location.href,
    title: document.title,
    text: readText(dialog),
    pairs: pairs,
    items: items,
    packages: packages,
    history: history,
    daily: daily,
    cards: cards,
    target_group: targetGroup,
  };
})()
"""


# ══════════════════════════════════════════════════════════════════════════
# TEXT HELPERS
# ══════════════════════════════════════════════════════════════════════════
def norm_text(value: Any) -> str:
    """Collapse horizontal whitespace, keep line breaks, drop zero-width chars."""
    if value is None:
        return ""
    text = str(value).replace("\u200b", "").replace("\xa0", " ")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def ff(v: str) -> str:
    """Khmer-free, case-insensitive form used for label matching."""
    return norm_text(v).casefold().strip(" *_:-–—៖")


def find_value(fields: dict[str, str], aliases: Sequence[str]) -> str:
    """Exact (normalised) match first, then substring match for long labels."""
    normalised = {ff(k): v for k, v in fields.items() if v}
    for alias in aliases:
        key = ff(alias)
        if key in normalised:
            return normalised[key]
    for alias in aliases:
        key = ff(alias)
        if len(key) < 4:
            continue
        for label, value in normalised.items():
            if key in label:
                return value
    return ""


def split_bp(value: str) -> tuple[str, str]:
    match = re.search(r"(\d{2,3})\s*/\s*(\d{2,3})", value or "")
    return (match.group(1), match.group(2)) if match else ("", "")


def detect_currency(amount: str) -> str:
    text = amount or ""
    if "៛" in text or "riel" in text.lower():
        return "KHR"
    if "$" in text or "usd" in text.lower():
        return "USD"
    return ""


def extract_icd(text: str) -> str:
    match = ICD_RE.search(text or "")
    return match.group(1) if match else ""


DISABLED_TOKEN = re.compile(r"(^|\s)(is-)?disabled($|\s)", re.IGNORECASE)


def is_disabled(locator: Locator) -> bool:
    """Disabled check that ignores Tailwind's ``disabled:`` utility classes."""
    try:
        if locator.is_disabled():
            return True
    except Exception:
        pass
    try:
        if (locator.get_attribute("aria-disabled") or "").lower() == "true":
            return True
    except Exception:
        pass
    try:
        for node in (locator, locator.locator("..")):
            classes = node.get_attribute("class") or ""
            if DISABLED_TOKEN.search(classes):
                return True
    except Exception:
        pass
    return False


def section_matches(name: str, hints: Sequence[str]) -> bool:
    target = (name or "").casefold()
    return any((hint or "").casefold() in target for hint in hints)


# ══════════════════════════════════════════════════════════════════════════
# DATABASE LAYER
# ══════════════════════════════════════════════════════════════════════════
def connect_db(db_path: Path | str = DB_PATH) -> sqlite3.Connection:
    path = Path(db_path).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path, timeout=20)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    conn.execute("PRAGMA busy_timeout = 5000")
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(SCHEMA)
    for table, columns in MIGRATIONS.items():
        existing = {row[1] for row in conn.execute(f"PRAGMA table_info({table})")}
        for column, declaration in columns.items():
            if column not in existing:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {declaration}")
    conn.commit()


def reset_db(conn: sqlite3.Connection) -> None:
    """Remove scraped data but keep the schema."""
    for table in (
        "modal_cards",
        "modal_tables",
        "modal_fields",
        "equity_packages",
        "approval_history",
        "invoice_items",
        "invoices",
        "soap_notes",
        "vitals",
        "clinical_records",
        "claims",
        "patients",
    ):
        conn.execute(f"DELETE FROM {table}")
    conn.commit()


def upsert(conn: sqlite3.Connection, table: str, data: dict[str, Any], key_columns: Sequence[str]) -> None:
    """INSERT ... ON CONFLICT DO UPDATE using only ``?`` placeholders."""
    if not data:
        return
    columns = list(data)
    placeholders = ", ".join("?" for _ in columns)
    updates = ", ".join(
        f"{column} = COALESCE(NULLIF(excluded.{column}, ''), {table}.{column})"
        for column in columns
        if column not in key_columns
    )
    sql = (
        f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({placeholders}) "
        f"ON CONFLICT({', '.join(key_columns)}) "
        + (f"DO UPDATE SET {updates}" if updates else "DO NOTHING")
    )
    conn.execute(sql, [data[column] for column in columns])


def make_patient_key(patient: dict[str, str]) -> str:
    if patient.get("nssf_id"):
        return f"nssf:{patient['nssf_id']}"
    if patient.get("health_card_id"):
        return f"card:{patient['health_card_id']}"
    basis = "|".join(
        [patient.get("patient_name", ""), patient.get("dob", ""), patient.get("phone", "")]
    )
    return "auto:" + hashlib.sha1(basis.encode("utf-8")).hexdigest()[:24]


def make_record_key(
    claim: dict[str, str],
    patient: dict[str, str],
    list_row: dict[str, Any],
) -> str:
    if claim.get("claim_no"):
        return f"claim:{claim['claim_no'].strip()}"
    basis = "|".join(
        [
            make_patient_key(patient),
            norm_text(list_row.get("__text__", "")),
            str(list_row.get("No", "")),
        ]
    )
    return "auto:" + hashlib.sha1(basis.encode("utf-8")).hexdigest()[:24]


def save_claim(conn: sqlite3.Connection, record: dict[str, Any]) -> int:
    """Upsert one patient + claim + clinical + vitals + SOAP + invoices."""
    patient = record["patient"]
    claim = record["claim"]
    clinical = record["clinical"]
    vitals = record["vitals"]
    soap = record["soap"]
    invoices = record.get("invoices", [])
    history = record.get("history", [])
    packages = record.get("packages", [])
    cards = record.get("cards", [])
    sections = record.get("sections", [])

    patient_key = make_patient_key(patient)
    record_key = make_record_key(claim, patient, record.get("list_row", {}))

    with conn:  # one transaction per patient: all tables or none
        upsert(
            conn,
            "patients",
            {
                "patient_key": patient_key,
                "nssf_id": patient.get("nssf_id", ""),
                "patient_name": patient.get("patient_name", ""),
                "gender": patient.get("gender", ""),
                "dob": patient.get("dob", ""),
                "age": patient.get("age", ""),
                "phone": patient.get("phone", ""),
                "health_card_id": patient.get("health_card_id", ""),
                "equity_card_id": patient.get("equity_card_id", ""),
                "nationality": patient.get("nationality", ""),
                "address": patient.get("address", ""),
                "card_type": patient.get("card_type", ""),
            },
            ["patient_key"],
        )
        patient_id = conn.execute(
            "SELECT patient_id FROM patients WHERE patient_key = ?", (patient_key,)
        ).fetchone()["patient_id"]

        upsert(
            conn,
            "claims",
            {
                "record_key": record_key,
                "claim_no": claim.get("claim_no", ""),
                "request_date": claim.get("request_date", ""),
                "facility": claim.get("facility", ""),
                "district": claim.get("district", ""),
                "facility_level": claim.get("facility_level", ""),
                "status": claim.get("status", ""),
                "reason": claim.get("reason", ""),
                "payment_type": claim.get("payment_type", ""),
                "error_type": claim.get("error_type", ""),
                "total_amount": claim.get("total_amount", ""),
                "medical_amount": claim.get("medical_amount", ""),
                "non_medical_amount": claim.get("non_medical_amount", ""),
                "rejected_at": claim.get("rejected_at", ""),
                "rejected_by": claim.get("rejected_by", ""),
                "rejection_reason": claim.get("rejection_reason", ""),
                "patient_id": patient_id,
                "source_page": record.get("source_page", 0),
                "source_row": record.get("source_row", 0),
                "list_row_json": json.dumps(record.get("list_row", {}), ensure_ascii=False),
                "raw_json": json.dumps(
                    {k: v for k, v in claim.items() if v}, ensure_ascii=False
                ),
            },
            ["record_key"],
        )
        claim_id = conn.execute(
            "SELECT claim_id FROM claims WHERE record_key = ?", (record_key,)
        ).fetchone()["claim_id"]

        upsert(
            conn,
            "clinical_records",
            {
                "claim_id": claim_id,
                "department": clinical.get("department", ""),
                "department_out": clinical.get("department_out", ""),
                "admission_date": clinical.get("admission_date", ""),
                "discharge_date": clinical.get("discharge_date", ""),
                "diagnosis": clinical.get("diagnosis", ""),
                "diagnosis_in": clinical.get("diagnosis_in", ""),
                "diagnosis_out": clinical.get("diagnosis_out", ""),
                "icd10": clinical.get("icd10", ""),
                "treatment_type": clinical.get("treatment_type", ""),
                "visit_type": clinical.get("visit_type", ""),
                "discharge_type": clinical.get("discharge_type", ""),
                "disease_status": clinical.get("disease_status", ""),
                "service_unit": clinical.get("service_unit", ""),
                "appointment_date": clinical.get("appointment_date", ""),
                "admission_reason": clinical.get("admission_reason", ""),
                "raw_json": json.dumps(clinical, ensure_ascii=False),
            },
            ["claim_id"],
        )

        upsert(
            conn,
            "vitals",
            {
                "claim_id": claim_id,
                "bp": vitals.get("bp", ""),
                "systolic": vitals.get("systolic", ""),
                "diastolic": vitals.get("diastolic", ""),
                "pulse": vitals.get("pulse", ""),
                "temp": vitals.get("temp", ""),
                "spo2": vitals.get("spo2", ""),
                "weight": vitals.get("weight", ""),
                "height": vitals.get("height", ""),
                "resp_rate": vitals.get("resp_rate", ""),
                "bmi": vitals.get("bmi", ""),
                "glucose": vitals.get("glucose", ""),
                "measured_at": vitals.get("measured_at", ""),
                "raw_json": json.dumps(vitals, ensure_ascii=False),
            },
            ["claim_id"],
        )

        upsert(
            conn,
            "soap_notes",
            {
                "claim_id": claim_id,
                "subjective": soap.get("subjective", ""),
                "objective": soap.get("objective", ""),
                "assessment": soap.get("assessment", ""),
                "evaluation": soap.get("evaluation", ""),
                "plan": soap.get("plan", ""),
                "chief_complaint": soap.get("chief_complaint", ""),
                "hpi": soap.get("hpi", ""),
                "examination": soap.get("examination", ""),
                "diagnosis": soap.get("diagnosis", ""),
                "icd10": soap.get("icd10", ""),
                "treatment_plan": soap.get("treatment_plan", ""),
                "conclusion": soap.get("conclusion", ""),
                "current_medication": soap.get("current_medication", ""),
                "past_medical_history": soap.get("past_medical_history", ""),
                "past_surgical_history": soap.get("past_surgical_history", ""),
                "allergies": soap.get("allergies", ""),
                "immunizations": soap.get("immunizations", ""),
                "family_history": soap.get("family_history", ""),
                "social_history": soap.get("social_history", ""),
                "symptoms": soap.get("symptoms", ""),
                "examination_json": json.dumps(
                    soap.get("examination_sections", {}), ensure_ascii=False
                ),
                "soap_doctor": soap.get("soap_doctor", ""),
                "soap_date": soap.get("soap_date", ""),
                "raw_json": json.dumps(soap, ensure_ascii=False),
            },
            ["claim_id"],
        )

        # Invoices and raw dialog content are replaced wholesale on every run,
        # so re-scraping never duplicates rows.
        conn.execute("DELETE FROM invoices WHERE claim_id = ?", (claim_id,))  # cascades items
        for invoice in invoices:
            cursor = conn.execute(
                """
                INSERT INTO invoices
                    (claim_id, invoice_no, invoice_date, amount, currency, status,
                     description, target_group, target_group_amount, raw_json)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    claim_id,
                    invoice.get("invoice_no", ""),
                    invoice.get("invoice_date", ""),
                    invoice.get("amount", ""),
                    invoice.get("currency", ""),
                    invoice.get("status", ""),
                    invoice.get("description", ""),
                    invoice.get("target_group", ""),
                    invoice.get("target_group_amount", ""),
                    json.dumps(invoice.get("raw", {}), ensure_ascii=False),
                ),
            )
            invoice_id = cursor.lastrowid
            for line_no, item in enumerate(invoice.get("items", []), start=1):
                conn.execute(
                    """
                    INSERT INTO invoice_items
                        (invoice_id, line_no, item, quantity, unit_price, amount, raw_json)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        invoice_id,
                        line_no,
                        item.get("item", ""),
                        item.get("quantity", ""),
                        item.get("unit_price", ""),
                        item.get("amount", ""),
                        json.dumps(item, ensure_ascii=False),
                    ),
                )

        conn.execute("DELETE FROM approval_history WHERE claim_id = ?", (claim_id,))
        if history:
            conn.executemany(
                """
                INSERT INTO approval_history
                    (claim_id, seq, event_time, status, reason, actor)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        claim_id,
                        index,
                        event.get("event_time", ""),
                        event.get("status", ""),
                        event.get("reason", ""),
                        event.get("actor", ""),
                    )
                    for index, event in enumerate(history)
                ],
            )

        conn.execute("DELETE FROM equity_packages WHERE claim_id = ?", (claim_id,))
        if packages:
            conn.executemany(
                """
                INSERT INTO equity_packages (claim_id, line_no, name, amount)
                VALUES (?, ?, ?, ?)
                """,
                [
                    (
                        claim_id,
                        package.get("line_no", ""),
                        package.get("name", ""),
                        package.get("amount", ""),
                    )
                    for package in packages
                ],
            )

        conn.execute("DELETE FROM modal_fields WHERE claim_id = ?", (claim_id,))
        field_rows = [
            (claim_id, section["name"], label, value)
            for section in sections
            for label, value in section["fields"].items()
            if value
        ]
        if field_rows:
            conn.executemany(
                "INSERT INTO modal_fields (claim_id, section, label, value) VALUES (?, ?, ?, ?)",
                field_rows,
            )

        conn.execute("DELETE FROM modal_tables WHERE claim_id = ?", (claim_id,))
        table_rows = [
            (
                claim_id,
                section["name"],
                table["name"],
                json.dumps(table.get("headers", []), ensure_ascii=False),
                json.dumps(table.get("rows", []), ensure_ascii=False),
                len(table.get("rows", [])),
            )
            for section in sections
            for table in section.get("tables", [])
            if table.get("rows") or table.get("headers")
        ]
        if table_rows:
            conn.executemany(
                """
                INSERT INTO modal_tables
                    (claim_id, section, table_name, headers_json, rows_json, row_count)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                table_rows,
            )

        conn.execute("DELETE FROM modal_cards WHERE claim_id = ?", (claim_id,))
        if cards:
            conn.executemany(
                "INSERT INTO modal_cards (claim_id, seq, title, text) VALUES (?, ?, ?, ?)",
                [
                    (claim_id, index, card.get("title", ""), card.get("text", ""))
                    for index, card in enumerate(cards)
                ],
            )

    return int(claim_id)


# ══════════════════════════════════════════════════════════════════════════
# MODAL EXTRACTION
# ══════════════════════════════════════════════════════════════════════════
def visible_modal(page: Page) -> Locator | None:
    for selector in MODAL_SELECTORS:
        locator = page.locator(selector)
        for index in range(locator.count()):
            candidate = locator.nth(index)
            try:
                if candidate.is_visible():
                    return candidate
            except Exception:
                continue
    return None


def wait_modal_idle(page: Page, modal: Locator, timeout_ms: int = 15000) -> None:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        try:
            busy = modal.locator(
                ".spinner-border:visible, .spinner-grow:visible, .loading:visible, "
                ".overlay:visible, .progress-bar-animated:visible"
            )
            if busy.count() == 0:
                break
        except Exception:
            break
        page.wait_for_timeout(250)
    try:
        page.wait_for_load_state("networkidle", timeout=4000)
    except Exception:
        pass


def wait_for_modal(page: Page, timeout_ms: int = 25000) -> Locator | None:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        modal = visible_modal(page)
        if modal is not None:
            wait_modal_idle(page, modal)
            return modal
        page.wait_for_timeout(200)
    return None


def wait_for_dialog_content(page: Page, timeout_ms: int = 20000) -> bool:
    """Wait until the PMRS dialog actually rendered its cards."""
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        try:
            page.wait_for_selector(
                "[role='dialog'] strong, [role='dialog'] h3, [role='dialog'] h6",
                state="visible",
                timeout=1500,
            )
            return True
        except Exception:
            page.wait_for_timeout(200)
    return False


def extract_modal(page: Page, timeout_ms: int = 5000) -> dict[str, Any]:
    """Read the whole detail dialog in one page-side pass.

    The Nuxt list re-renders on its own every few seconds and briefly unmounts
    the overlay, so a single evaluate can land in a gap. Retry until the
    dialog is back instead of failing the row.
    """
    deadline = time.monotonic() + timeout_ms / 1000
    raw: Any = None
    while time.monotonic() < deadline:
        raw = page.evaluate(EXTRACT_JS)
        if isinstance(raw, dict) and raw.get("ok"):
            return raw
        page.wait_for_timeout(350)
    error = raw.get("error") if isinstance(raw, dict) else "extraction returned no data"
    raise RuntimeError(str(error))


def dialog_html(page: Page) -> str:
    try:
        return page.evaluate(
            """
            () => {
              const list = Array.from(document.querySelectorAll('[role=dialog]'))
                .filter((el) => el.getClientRects().length);
              const dialog = list.length ? list[list.length - 1] : null;
              return dialog ? dialog.outerHTML : '';
            }
            """
        )
    except Exception:
        return ""


# ══════════════════════════════════════════════════════════════════════════
# RECORD MAPPING
# ══════════════════════════════════════════════════════════════════════════
def merge_sections(
    sections: Iterable[dict[str, Any]], hints: Sequence[str] | None = None
) -> dict[str, str]:
    merged: dict[str, str] = {}
    for section in sections:
        if hints and not section_matches(section["name"], hints):
            continue
        for label, value in section["fields"].items():
            if value and label not in merged:
                merged[label] = value
    return merged


def pick_from(
    sections: list[dict[str, Any]],
    hints: Sequence[str],
    aliases: Sequence[str],
    fallback: dict[str, str] | None = None,
) -> str:
    value = find_value(merge_sections(sections, hints), aliases)
    if not value and fallback:
        value = find_value(fallback, aliases)
    return value


def build_sections(raw: dict[str, Any]) -> list[dict[str, Any]]:
    """Group the page-side label/value pairs by dialog card."""
    order: list[str] = []
    by_name: dict[str, dict[str, Any]] = {}

    def bucket(name: str) -> dict[str, Any]:
        name = norm_text(name) or "General"
        if name not in by_name:
            by_name[name] = {"name": name, "fields": {}, "tables": []}
            order.append(name)
        return by_name[name]

    for pair in raw.get("pairs") or []:
        section = bucket(pair.get("section") or "General")
        label = norm_text(pair.get("label"))
        value = norm_text(pair.get("value"))
        if not label or not value:
            continue
        existing = section["fields"].get(label)
        if not existing:
            section["fields"][label] = value
        elif value not in existing:
            section["fields"][label] = f"{existing}\n{value}"

    history = raw.get("history") or []
    if history:
        bucket("ប្រវត្តិនៃការអនុម័ត")["tables"].append(
            {
                "name": "approval history",
                "headers": ["event_time", "status", "reason", "actor"],
                "rows": [
                    [
                        event.get("event_time", ""),
                        event.get("status", ""),
                        event.get("reason", ""),
                        event.get("actor", ""),
                    ]
                    for event in history
                ],
            }
        )

    items = raw.get("items") or []
    if items:
        bucket("Invoice")["tables"].append(
            {
                "name": "invoice items",
                "headers": ["item", "quantity", "amount"],
                "rows": [
                    [item.get("item", ""), item.get("quantity", ""), item.get("amount", "")]
                    for item in items
                ],
            }
        )

    packages = raw.get("packages") or []
    if packages:
        bucket("ក្រុមគោលដៅ")["tables"].append(
            {
                "name": "equity packages",
                "headers": ["line_no", "name", "amount"],
                "rows": [
                    [
                        package.get("line_no", ""),
                        package.get("name", ""),
                        package.get("amount", ""),
                    ]
                    for package in packages
                ],
            }
        )

    return [by_name[name] for name in order]


def build_invoices(sections: list[dict[str, Any]], raw: dict[str, Any]) -> list[dict[str, Any]]:
    fields = merge_sections(sections, INVOICE_SECTIONS)
    items = raw.get("items") or []
    packages = raw.get("packages") or []
    target_group = norm_text(raw.get("target_group"))
    amount = find_value(fields, INVOICE_AMOUNT)
    invoice_no = find_value(fields, INVOICE_NO)

    if not (invoice_no or amount or items):
        return []

    return [
        {
            "invoice_no": invoice_no,
            "invoice_date": find_value(fields, INVOICE_DATE),
            "amount": amount,
            "currency": detect_currency(amount),
            "status": "",
            "description": find_value(fields, INVOICE_DESCRIPTION),
            "target_group": target_group,
            "target_group_amount": packages[0].get("amount", "") if packages else "",
            "raw": {k: v for k, v in fields.items() if v},
            "items": [
                {
                    "item": item.get("item", ""),
                    "quantity": item.get("quantity", ""),
                    "unit_price": "",
                    "amount": item.get("amount", ""),
                }
                for item in items
            ],
        }
    ]


def rejection_event(history: list[dict[str, Any]]) -> dict[str, Any]:
    for event in history:
        status = (event.get("status") or "").casefold()
        if "បដិសេធ" in status or "reject" in status:
            return event
    return history[0] if history else {}


def build_record(
    list_row: dict[str, Any],
    raw: dict[str, Any],
    sections: list[dict[str, Any]],
    page_no: int,
    row_no: int,
) -> dict[str, Any]:
    claim_fields = merge_sections(sections, CLAIM_SECTIONS)
    patient_fields = merge_sections(sections, PATIENT_SECTIONS)
    clinical_fields = merge_sections(sections, CLINICAL_SECTIONS)
    vital_fields = merge_sections(sections, VITAL_SECTIONS)
    soap_fields = merge_sections(sections, SOAP_SECTIONS)
    history_fields = merge_sections(sections, HISTORY_SECTIONS)
    exam_fields = merge_sections(sections, EXAM_SECTIONS)

    history = raw.get("history") or []
    rejected = rejection_event(history)

    dob = find_value(patient_fields, DOB)
    age_match = AGE_RE.search(dob)
    age = age_match.group(1).strip() if age_match else find_value(patient_fields, AGE, )

    clinical = {
        "department": find_value(clinical_fields, DEPARTMENT_IN),
        "department_out": find_value(clinical_fields, DEPARTMENT_OUT),
        "admission_date": find_value(clinical_fields, ADMISSION_DATE),
        "discharge_date": find_value(clinical_fields, DISCHARGE_DATE),
        "diagnosis": find_value(clinical_fields, DIAGNOSIS_OUT)
        or find_value(clinical_fields, DIAGNOSIS_IN),
        "diagnosis_in": find_value(clinical_fields, DIAGNOSIS_IN),
        "diagnosis_out": find_value(clinical_fields, DIAGNOSIS_OUT),
        "icd10": extract_icd(
            find_value(clinical_fields, DIAGNOSIS_OUT)
            or find_value(clinical_fields, DIAGNOSIS_IN)
        ),
        "treatment_type": find_value(clinical_fields, VISIT_TYPE),
        "visit_type": find_value(clinical_fields, VISIT_TYPE),
        "discharge_type": find_value(clinical_fields, DISCHARGE_TYPE),
        "disease_status": find_value(clinical_fields, DISEASE_STATUS),
        "service_unit": find_value(clinical_fields, SERVICE_UNIT),
        "appointment_date": find_value(clinical_fields, APPOINTMENT_DATE),
        "admission_reason": find_value(clinical_fields, ADMISSION_REASON)
        or find_value(vital_fields, ADMISSION_REASON),
    }

    bp = find_value(vital_fields, BP)
    systolic, diastolic = split_bp(bp)

    # Some claims have no "initial assessment" card at all: the only vitals
    # then live in the ward/examination card (as Systolic/Diastolic rows).
    measure_fields = vital_fields
    if not (systolic or find_value(vital_fields, PULSE)):
        ward_fields = merge_sections(sections, WARD_SECTIONS)
        if ward_fields:
            measure_fields = {**ward_fields, **vital_fields}
            bp = find_value(measure_fields, BP)
            systolic, diastolic = split_bp(bp)
            if not systolic:
                sys_value = find_value(measure_fields, SYSTOLIC)
                dia_value = find_value(measure_fields, DIASTOLIC)
                if sys_value and dia_value:
                    bp = f"{sys_value}/{dia_value}"
                    systolic, diastolic = sys_value, dia_value

    vitals = {
        "bp": bp,
        "systolic": systolic,
        "diastolic": diastolic,
        "pulse": find_value(measure_fields, PULSE),
        "temp": find_value(measure_fields, TEMP),
        "spo2": find_value(measure_fields, SPO2),
        "weight": find_value(measure_fields, WEIGHT),
        "height": find_value(measure_fields, HEIGHT),
        "resp_rate": find_value(measure_fields, RESP_RATE),
        "bmi": "",
        "glucose": find_value(measure_fields, GLUCOSE),
        "measured_at": find_value(measure_fields, ("កាលបរិច្ឆេទ", "date", "time")),
    }

    examinations = {
        label: value
        for label, value in exam_fields.items()
        if label and value and not section_matches(label, SYMPTOMS)
    }
    symptoms = find_value(exam_fields, SYMPTOMS)

    assessment = find_value(soap_fields, ASSESSMENT)
    soap_doctor = ""
    soap_date = ""
    for card in raw.get("cards") or []:
        if (card.get("title") or "").casefold() != "soap":
            continue
        lines = norm_text(card.get("text")).splitlines()
        for line in lines:
            if not soap_doctor and line.startswith("វេជ្ជ"):
                soap_doctor = line
            if not soap_date and re.match(r"^\d{1,2}/\d{1,2}/\d{4}", line):
                soap_date = line
        break

    soap = {
        "subjective": find_value(soap_fields, SUBJECTIVE),
        "objective": find_value(soap_fields, OBJECTIVE),
        "assessment": assessment,
        "evaluation": find_value(soap_fields, EVALUATION),
        "plan": find_value(soap_fields, PLAN),
        "chief_complaint": find_value(history_fields, HPI),
        "hpi": find_value(history_fields, HPI),
        "examination": "\n".join(
            f"{label}: {value}" for label, value in examinations.items()
        ),
        "diagnosis": clinical["diagnosis"],
        "icd10": clinical["icd10"],
        "treatment_plan": find_value(soap_fields, PLAN),
        "conclusion": find_value(soap_fields, EVALUATION),
        "current_medication": find_value(history_fields, CURRENT_MEDICATION),
        "past_medical_history": find_value(history_fields, PAST_MEDICAL_HISTORY),
        "past_surgical_history": find_value(history_fields, PAST_SURGICAL_HISTORY),
        "allergies": find_value(history_fields, ALLERGIES),
        "immunizations": find_value(history_fields, IMMUNIZATIONS),
        "family_history": find_value(history_fields, FAMILY_HISTORY),
        "social_history": find_value(history_fields, SOCIAL_HISTORY),
        "symptoms": symptoms,
        "examination_sections": examinations,
        "soap_doctor": soap_doctor,
        "soap_date": soap_date,
    }

    claim = {
        "claim_no": find_value(claim_fields, CLAIM_NO),
        "request_date": find_value(claim_fields, REQUEST_DATE),
        "district": find_value(claim_fields, DISTRICT),
        "facility": find_value(claim_fields, FACILITY),
        "facility_level": find_value(claim_fields, FACILITY_LEVEL),
        "status": str(list_row.get("ស្ថានភាពបញ្ជូន", "") or ""),
        "reason": rejected.get("reason", "") or find_value(claim_fields, REASON),
        "payment_type": str(list_row.get("ប្រភេទនៃការបង់ប្រាក់", "") or ""),
        "error_type": str(list_row.get("ប្រភេទកំហុស", "") or ""),
        "total_amount": str(list_row.get("ទឹកប្រាក់សរុប", "") or ""),
        "medical_amount": str(list_row.get("ទឹកប្រាក់វេជ្ជសាស្ត្រ", "") or ""),
        "non_medical_amount": str(list_row.get("ទឹកប្រាក់មិនមែនវេជ្ជសាស្ត្រ", "") or ""),
        "rejected_at": rejected.get("event_time", ""),
        "rejected_by": rejected.get("actor", ""),
        "rejection_reason": rejected.get("reason", ""),
    }

    patient = {
        "nssf_id": find_value(patient_fields, ("nssf id", "nssf", "nssf number")),
        "patient_name": find_value(patient_fields, PATIENT_NAME),
        "gender": find_value(patient_fields, GENDER)
        or str(list_row.get("ភេទ", "") or ""),
        "dob": dob,
        "age": age or str(list_row.get("អាយុ", "") or ""),
        "phone": find_value(patient_fields, PHONE),
        "health_card_id": find_value(patient_fields, HEALTH_CARD_ID)
        or str(list_row.get("បណ្ណសុខភាព", "") or ""),
        "equity_card_id": find_value(patient_fields, EQUITY_CARD_ID),
        "nationality": find_value(patient_fields, NATIONALITY),
        "address": find_value(patient_fields, ADDRESS),
        "card_type": "",
    }

    invoices = build_invoices(sections, raw)

    return {
        "claim": claim,
        "patient": patient,
        "clinical": clinical,
        "vitals": vitals,
        "soap": soap,
        "invoices": invoices,
        "history": history,
        "packages": raw.get("packages") or [],
        "cards": raw.get("cards") or [],
        "sections": sections,
        "list_row": list_row,
        "source_page": page_no,
        "source_row": row_no,
    }


# ══════════════════════════════════════════════════════════════════════════
# TABLE / ROW NAVIGATION
# ══════════════════════════════════════════════════════════════════════════
def main_table(page: Page) -> Locator | None:
    best: Locator | None = None
    best_rows = -1
    try:
        for table in page.locator("table").all():
            if not table.is_visible():
                continue
            rows = table.locator("tbody tr").count()
            if rows > best_rows:
                best, best_rows = table, rows
    except Exception:
        return None
    return best


def table_headers(table: Locator) -> list[str]:
    """Return the richest <thead> row (PMRS may render extra empty rows)."""
    try:
        best: list[str] = []
        best_filled = 0
        for row in table.locator("thead tr").all():
            cells = [norm_text(x) for x in row.locator("th, td").all_inner_texts()]
            filled = sum(1 for cell in cells if cell)
            if filled > best_filled:
                best, best_filled = cells, filled
        if best_filled:
            return best

        first = table.locator("tr").first
        return [norm_text(x) for x in first.locator("th").all_inner_texts()]
    except Exception:
        return []


def extract_list_row(row: Locator, headers: list[str]) -> dict[str, Any]:
    cells = [norm_text(x) for x in row.locator(":scope > td, :scope > th").all_inner_texts()]
    mapped: dict[str, Any] = {}
    for index, value in enumerate(cells):
        if index < len(headers) and headers[index]:
            mapped[headers[index]] = value
    mapped["__cells__"] = cells
    mapped["__text__"] = norm_text(row.inner_text())
    return mapped


NUMERIC_CELL_RE = re.compile(r"^[\d\s.,/:+-]+$")
DATE_LIKE_RE = re.compile(r"\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}")
NUMBER_ALIASES = ("លេខ", "no", "number", "#", "nº")


def row_identity(list_row: dict[str, Any], index: int) -> tuple[str, str]:
    """Best-effort stable identity for one list row (row number + patient name)."""
    cells = [c for c in list_row.get("__cells__", []) if c]
    number = find_value(list_row, NUMBER_ALIASES)
    name = find_value(list_row, PATIENT_NAME)

    if not name:
        for cell in cells:
            if (
                not NUMERIC_CELL_RE.fullmatch(cell)
                and not DATE_LIKE_RE.search(cell)
                and not re.fullmatch(r"[A-Z]{2,}[-_/]?\d[\w\-/]*", cell)
                and len(cell) > 2
                and not re.fullmatch(
                    r"(?:បដិសេធ|អនុម័ត|រង់ចាំ|rejected|approved|pending)",
                    cell,
                    re.IGNORECASE,
                )
            ):
                name = cell
                break

    if not number:
        for cell in cells:
            if NUMERIC_CELL_RE.fullmatch(cell) and len(cell) <= 8:
                number = cell
                break

    if not number:
        basis = norm_text(list_row.get("__text__", "")) or str(index)
        number = "row-" + hashlib.sha1(basis.encode("utf-8")).hexdigest()[:10]

    return number, name


def rows_signature(table: Locator) -> str:
    try:
        rows = table.locator("tbody tr")
        count = rows.count()
        if count == 0:
            return "0"
        first = norm_text(rows.first.inner_text())[:200]
        last = norm_text(rows.nth(count - 1).inner_text())[:200]
        return f"{count}|{first}|{last}"
    except Exception:
        return ""


def find_view_button(row: Locator) -> Locator | None:
    """Locate the eye / view-details control inside one table row."""
    candidates = [
        "button:has(i[class*='eye' i])",
        "a:has(i[class*='eye' i])",
        "button:has(svg[class*='eye' i])",
        "a:has(svg[class*='eye' i])",
        "button:has(span[class*='eye' i])",
        "a:has(span[class*='eye' i])",
        "[title*='view' i]",
        "[aria-label*='view' i]",
        "[title*='detail' i]",
        "[aria-label*='detail' i]",
        "button:has(i[class*='detail' i])",
        "a:has(i[class*='detail' i])",
        "td:last-child button",
        "td:last-child a",
    ]
    for selector in candidates:
        try:
            locator = row.locator(selector)
            for index in range(locator.count()):
                candidate = locator.nth(index)
                if candidate.is_visible() and not is_disabled(candidate):
                    return candidate
        except Exception:
            continue

    # Last resort: any button/link whose markup mentions an eye or "view".
    try:
        for candidate in row.locator("button, a").all():
            if not candidate.is_visible():
                continue
            blob = " ".join(
                filter(
                    None,
                    [
                        (candidate.get_attribute("class") or "").lower(),
                        (candidate.get_attribute("title") or "").lower(),
                        (candidate.get_attribute("aria-label") or "").lower(),
                        (candidate.inner_html() or "").lower(),
                    ],
                )
            )
            if "eye" in blob or "view" in blob or "detail" in blob:
                return candidate
    except Exception:
        pass
    return None


def click_element(locator: Locator) -> None:
    try:
        locator.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass
    try:
        locator.click(timeout=8000)
        return
    except Exception:
        pass
    try:
        locator.dispatch_event("click")
        return
    except Exception:
        pass
    locator.evaluate("el => el.click()")


def close_modal(page: Page, modal: Locator | None = None) -> None:
    scope = modal or page
    clicked = False
    for selector in CLOSE_SELECTORS:
        try:
            candidates = scope.locator(selector)
            for index in range(candidates.count()):
                button = candidates.nth(index)
                if button.is_visible():
                    click_element(button)
                    clicked = True
                    break
        except Exception:
            continue
        if clicked:
            break

    # Fallbacks for custom icon-only close buttons.
    if not clicked:
        for selector in (
            ".modal-header button:visible",
            ".offcanvas-header button:visible",
            "[class*='close']:visible",
            "button[aria-label]:visible",
        ):
            try:
                candidates = scope.locator(selector)
                for index in range(min(candidates.count(), 6)):
                    button = candidates.nth(index)
                    if not button.is_visible():
                        continue
                    blob = " ".join(
                        filter(
                            None,
                            [
                                (button.get_attribute("class") or "").lower(),
                                (button.get_attribute("title") or "").lower(),
                                (button.get_attribute("aria-label") or "").lower(),
                            ],
                        )
                    )
                    text = norm_text(button.inner_text())
                    if any(word in blob for word in ("close", "dismiss", "cancel", "x")) or text in (
                        "×",
                        "✕",
                        "X",
                    ):
                        click_element(button)
                        clicked = True
                        break
            except Exception:
                continue
            if clicked:
                break

    if not clicked:
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass

    # Wait until no modal is visible any more.
    deadline = time.monotonic() + 15
    while time.monotonic() < deadline:
        remaining = visible_modal(page)
        if remaining is None:
            break
        page.wait_for_timeout(200)

    # If a custom overlay ignores both X and Escape, hide it in the DOM.
    if visible_modal(page) is not None:
        try:
            page.evaluate(
                """
                () => {
                    document.querySelectorAll('.modal, .offcanvas, .drawer, [role=dialog]')
                        .forEach(el => {
                            el.classList.remove('show', 'in', 'open');
                            el.style.removeProperty('display');
                            el.setAttribute('aria-hidden', 'true');
                        });
                }
                """
            )
        except Exception:
            pass

    # Clean up Bootstrap leftovers that would block the next click.
    try:
        page.evaluate(
            """
            () => {
                document.querySelectorAll('.modal-backdrop').forEach(el => el.remove());
                document.body.classList.remove('modal-open');
                document.body.style.removeProperty('overflow');
                document.body.style.removeProperty('padding-right');
            }
            """
        )
    except Exception:
        pass


def locate_row(
    page: Page, index: int, expected_no: str = "", expected_name: str = ""
) -> Locator | None:
    table = main_table(page)
    if table is None:
        return None
    rows = table.locator("tbody tr")
    count = rows.count()

    if index < count:
        row = rows.nth(index)
        text = norm_text(row.inner_text())
        no_ok = not expected_no or expected_no in text
        name_ok = not expected_name or expected_name in text
        if no_ok and name_ok:
            return row

    for candidate_index in range(count):
        row = rows.nth(candidate_index)
        text = norm_text(row.inner_text())
        if expected_name and expected_name in text and (not expected_no or expected_no in text):
            return row
    return rows.nth(index) if index < count else None


def open_detail(page: Page, row: Locator) -> bool:
    button = find_view_button(row)
    if button is None:
        return False
    click_element(button)
    if wait_for_modal(page) is None:
        return False
    return wait_for_dialog_content(page)


def find_next_button(page: Page) -> Locator | None:
    selectors = (
        "button[aria-label*='next' i]",
        "a[aria-label*='next' i]",
        "button[data-slot='next']",
        "li.paginate_button.next:not(.disabled) a",
        "a.paginate_button.next:not(.disabled)",
        ".dataTables_paginate .next:not(.disabled)",
        ".pagination a.next",
        "button:has-text('Next')",
        "a:has-text('Next')",
        "button:has-text('បន្ទាប់')",
        "a:has-text('បន្ទាប់')",
    )
    for selector in selectors:
        try:
            locator = page.locator(selector)
            for index in range(locator.count()):
                candidate = locator.nth(index)
                if candidate.is_visible() and not is_disabled(candidate):
                    return candidate
        except Exception:
            continue

    # The PMRS pager can also be a bare "›" glyph.
    for glyph in ("›", "»"):
        try:
            locator = page.get_by_text(glyph, exact=True)
            for index in range(locator.count()):
                candidate = locator.nth(index)
                if candidate.is_visible() and not is_disabled(candidate):
                    return candidate
        except Exception:
            continue
    return None


def wait_for_rows_change(page: Page, old_signature: str, timeout_ms: int = 25000) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        page.wait_for_timeout(300)
        table = main_table(page)
        if table is None:
            continue
        signature = rows_signature(table)
        if signature and signature != old_signature and signature != "0":
            page.wait_for_timeout(300)
            return True
    return False


def wait_for_table(page: Page, timeout_ms: int = 60000) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        try:
            page.wait_for_selector(ROW_SELECTOR, state="visible", timeout=2000)
            return True
        except Exception:
            page.wait_for_timeout(500)
    return False


# ══════════════════════════════════════════════════════════════════════════
# SCRAPE LOOP
# ══════════════════════════════════════════════════════════════════════════
def scrape(page: Page, conn: sqlite3.Connection, args: argparse.Namespace) -> dict[str, int]:
    stats = {"pages": 0, "rows": 0, "saved": 0, "failed": 0, "skipped": 0}
    debug_budget = max(0, args.dump_html)
    page_no = 1

    # --start-page N: click "next" N-1 times so a long scrape can be resumed
    # without re-visiting the pages that are already in the database.
    start_page = max(1, int(getattr(args, "start_page", 1) or 1))
    while page_no < start_page:
        if not wait_for_table(page):
            print("⏳ No table rows found — cannot skip to the start page.")
            return stats
        table = main_table(page)
        signature = rows_signature(table) if table is not None else ""
        next_button = find_next_button(page)
        if next_button is None:
            print(f"⚠️ Cannot skip to page {start_page}: no next-page control.")
            return stats
        print(f"⏭️  Skipping to page {page_no + 1}…")
        click_element(next_button)
        if not wait_for_rows_change(page, signature):
            print("⚠️ Table did not change while skipping pages — stopping.")
            return stats
        page_no += 1

    while True:
        if not wait_for_table(page):
            print("⏳ No table rows found. Log in, open the Rejected list and try again.")
            break

        table = main_table(page)
        if table is None:
            print("⚠️ Could not identify the claims table.")
            break

        stats["pages"] += 1
        print(
            f"\n{'═' * 70}\n📄 PAGE {page_no} — {table.locator('tbody tr').count()} row(s)"
            f"\n{'═' * 70}"
        )

        for index in range(table.locator("tbody tr").count()):
            if args.max_rows is not None and stats["rows"] >= args.max_rows:
                print(f"🛑 Reached --max-rows {args.max_rows}.")
                return stats

            # Re-resolve the table every row: closing the dialog may re-render
            # the list, which would invalidate cached locators/headers.
            table = main_table(page) or table
            rows = table.locator("tbody tr")
            row_count = rows.count()
            if index >= row_count:
                print(f"   ↷ Row {index + 1}: table shrank while iterating — skipped.")
                continue
            headers = table_headers(table)

            stats["rows"] += 1
            list_row = extract_list_row(rows.nth(index), headers)
            row_text = list_row.get("__text__", "")
            if not row_text or re.search(r"no data|no record|គ្មានទិន្នន័យ", row_text, re.IGNORECASE):
                print(f"   ↷ Row {index + 1}: empty/placeholder row — skipped.")
                stats["skipped"] += 1
                continue

            row_no, row_name = row_identity(list_row, index)
            print(f"\n--- Page {page_no} | Row {index + 1}/{row_count} | No={row_no} | {row_name} ---")

            saved = False
            last_error = ""

            for attempt in range(1, 4):
                if attempt > 1:
                    print(f"   ⟳ Retry {attempt}/3: reopening the dialog…")
                    close_modal(page)
                    page.wait_for_timeout(700)

                opened = False
                for open_try in (1, 2):
                    try:
                        row = locate_row(page, index, expected_no=row_no, expected_name=row_name)
                        if row is None:
                            raise RuntimeError("row disappeared (table re-rendered)")
                        opened = open_detail(page, row)
                        if opened:
                            break
                        print(f"   ⚠️ Dialog did not open (attempt {open_try}/2).")
                    except Exception as exc:
                        print(f"   ⚠️ Open attempt {open_try}/2 failed: {exc}")
                        close_modal(page)
                        page.wait_for_timeout(700)

                if not opened:
                    last_error = "dialog did not open"
                    continue

                try:
                    raw = extract_modal(page)
                    sections = build_sections(raw)
                    field_count = sum(len(section["fields"]) for section in sections)
                    item_count = len(raw.get("items") or [])
                    history_count = len(raw.get("history") or [])

                    # A dialog that has just been painted can show its card
                    # frames one tick before the values arrive.
                    for _ in range(3):
                        if field_count or item_count or history_count:
                            break
                        page.wait_for_timeout(900)
                        if not wait_for_dialog_content(page, timeout_ms=4000):
                            continue
                        raw = extract_modal(page)
                        sections = build_sections(raw)
                        field_count = sum(len(section["fields"]) for section in sections)
                        item_count = len(raw.get("items") or [])
                        history_count = len(raw.get("history") or [])

                    if field_count == 0 and item_count == 0 and history_count == 0:
                        last_error = "dialog yielded no fields"
                        continue

                    record = build_record(list_row, raw, sections, page_no, index + 1)

                    if debug_budget > 0:
                        DEBUG_DIR.mkdir(exist_ok=True)
                        stem = DEBUG_DIR / f"page{page_no}_row{index + 1}"
                        stem.with_suffix(".html").write_text(dialog_html(page), encoding="utf-8")
                        stem.with_suffix(".json").write_text(
                            json.dumps(
                                {"raw": raw, "record": record},
                                ensure_ascii=False,
                                indent=2,
                            ),
                            encoding="utf-8",
                        )
                        debug_budget -= 1

                    claim_id = save_claim(conn, record)
                    stats["saved"] += 1
                    saved = True
                    print(
                        f"   ✅ Saved claim_id={claim_id} | claim_no={record['claim']['claim_no'] or '?'} "
                        f"| fields={field_count} | items={item_count} | events={history_count}"
                    )
                    break
                except Exception as exc:
                    last_error = str(exc)
                    print(f"   ⚠️ Extraction attempt {attempt}/3 failed: {exc}")

            if not saved:
                stats["failed"] += 1
                DEBUG_DIR.mkdir(exist_ok=True)
                stem = DEBUG_DIR / f"failed_page{page_no}_row{index + 1}"
                stem.with_suffix(".html").write_text(page.content()[:200000], encoding="utf-8")
                print(f"   ❌ Row not saved: {last_error} (dump: {stem.name})")

            close_modal(page)
            page.wait_for_timeout(400)

        if args.max_pages is not None and page_no >= args.max_pages:
            print(f"🛑 Reached --max-pages {args.max_pages}.")
            break

        next_button = find_next_button(page)
        if next_button is None:
            print("🏁 No next-page control — finished.")
            break

        signature = rows_signature(table)
        print("➡️  Next page…")
        click_element(next_button)
        if not wait_for_rows_change(page, signature):
            print("⚠️ Table did not change after clicking next — stopping.")
            break
        page_no += 1

    return stats


# ══════════════════════════════════════════════════════════════════════════
# OPTIONAL FLAT EXPORT
# ══════════════════════════════════════════════════════════════════════════
def export_flat(conn: sqlite3.Connection, out_dir: Path) -> None:
    try:
        import pandas as pd
    except Exception as exc:
        print(f"⚠️ pandas not available, skipping flat export: {exc}")
        return

    query = """
        SELECT
            c.claim_no,
            c.request_date,
            c.facility,
            c.district,
            c.facility_level,
            c.status,
            c.error_type,
            c.rejection_reason,
            c.rejected_at,
            c.rejected_by,
            c.payment_type,
            c.total_amount,
            p.health_card_id,
            p.equity_card_id,
            p.patient_name,
            p.gender,
            p.dob,
            p.age,
            p.phone,
            p.nationality,
            p.address,
            cr.department,
            cr.department_out,
            cr.admission_date,
            cr.discharge_date,
            cr.diagnosis_in,
            cr.diagnosis_out,
            cr.icd10,
            cr.visit_type,
            cr.discharge_type,
            cr.disease_status,
            cr.service_unit,
            cr.appointment_date,
            cr.admission_reason,
            v.bp,
            v.pulse,
            v.temp,
            v.spo2,
            v.weight,
            v.height,
            v.resp_rate,
            v.glucose,
            s.subjective,
            s.objective,
            s.assessment,
            s.evaluation,
            s.plan,
            s.hpi,
            s.current_medication,
            s.past_medical_history,
            s.past_surgical_history,
            s.allergies,
            s.family_history,
            s.social_history,
            s.symptoms,
            s.soap_doctor,
            s.soap_date,
            s.examination_json,
            i.invoice_no,
            i.invoice_date,
            i.amount          AS invoice_amount,
            i.target_group,
            i.target_group_amount
        FROM claims c
        LEFT JOIN patients p          ON p.patient_id = c.patient_id
        LEFT JOIN clinical_records cr ON cr.claim_id  = c.claim_id
        LEFT JOIN vitals v            ON v.claim_id   = c.claim_id
        LEFT JOIN soap_notes s        ON s.claim_id   = c.claim_id
        LEFT JOIN invoices i          ON i.claim_id   = c.claim_id
        ORDER BY c.claim_id
    """
    frame = pd.read_sql_query(query, conn)
    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx = out_dir / "rejected_claims_full.xlsx"
    csv = out_dir / "rejected_claims_full.csv"
    frame.to_excel(xlsx, index=False)
    frame.to_csv(csv, index=False, encoding="utf-8-sig")
    print(f"\n📄 Flat export: {len(frame)} row(s)\n   {xlsx}\n   {csv}")


# ══════════════════════════════════════════════════════════════════════════
# READY / LOGIN WAIT
# ══════════════════════════════════════════════════════════════════════════
def table_ready(page: Page) -> bool:
    """True as soon as any row of a visible table is on screen."""
    try:
        rows = page.locator(ROW_SELECTOR)
        for index in range(min(rows.count(), 5)):
            if rows.nth(index).is_visible():
                return True
    except Exception:
        pass
    return False


def wait_until_ready(page: Page, args: argparse.Namespace) -> bool:
    """Wait for the user to log in and reach the claims list (no terminal input).

    Keeps polling, so the scraper starts by itself the moment the table
    appears — whether the stored profile was already logged in or not.
    """
    if args.no_pause:
        return True

    print("\n" + "═" * 74)
    print("A Chromium window is open on PMRS.")
    print("Log in, open the Rejected list and apply your filters.")
    print("Scraping starts automatically when the claims table appears")
    print(f"(waiting up to {args.login_timeout}s).")
    print("═" * 74)

    deadline = time.monotonic() + max(30, args.login_timeout)
    last_note = 0.0
    while time.monotonic() < deadline:
        if table_ready(page):
            print("✅ Claims table detected — starting scrape.\n")
            return True
        now = time.monotonic()
        if now - last_note >= 15:
            last_note = now
            if "/login" in page.url:
                print("⏳ Waiting for login… (browser window is open)")
            else:
                print(f"⏳ Waiting for the claims table… ({page.url})")
        page.wait_for_timeout(1000)

    print("⏰ Timed out waiting for login / claims table.")
    return False


# ══════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════
def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scrape the PMRS rejected-claims detail dialog into SQLite."
    )
    parser.add_argument("--url", default=BASE_URL, help="report URL (default: %(default)s)")
    parser.add_argument("--db", default=str(DB_PATH), help="SQLite file (default: %(default)s)")
    parser.add_argument("--profile", default=str(PROFILE_DIR), help="persistent browser profile")
    parser.add_argument("--max-rows", type=int, default=None, help="stop after N rows (testing)")
    parser.add_argument("--max-pages", type=int, default=None, help="stop after N pages (testing)")
    parser.add_argument(
        "--start-page",
        type=int,
        default=1,
        help="skip ahead to page N before scraping (resume a long run)",
    )
    parser.add_argument("--dump-html", type=int, default=2, help="save N dialog dumps (0 = off)")
    parser.add_argument("--reset", action="store_true", help="delete existing scraped rows first")
    parser.add_argument("--no-pause", action="store_true", help="start scraping immediately")
    parser.add_argument(
        "--login-timeout",
        type=int,
        default=900,
        help="seconds to wait for login + claims table when running without a terminal prompt",
    )
    parser.add_argument("--no-excel", action="store_true", help="skip the flat Excel/CSV export")
    parser.add_argument("--headless", action="store_true", help="run without a visible window")
    parser.add_argument("--slow-mo", type=int, default=0, help="Playwright slow_mo in ms")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)

    conn = connect_db(args.db)
    init_db(conn)
    print(f"🗄️ SQLite: {Path(args.db).resolve()}")

    if args.reset:
        reset_db(conn)
        print("🧹 Existing scraped rows deleted (--reset).")

    profile = Path(args.profile).expanduser().resolve()
    profile.mkdir(parents=True, exist_ok=True)

    if args.dump_html > 0:
        DEBUG_DIR.mkdir(exist_ok=True)

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile),
            headless=args.headless,
            slow_mo=args.slow_mo,
            viewport={"width": 1500, "height": 950},
            accept_downloads=True,
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(15000)

        try:
            page.goto(args.url, wait_until="domcontentloaded")

            if not wait_until_ready(page, args):
                print("No data scraped.")
                return 2

            stats = scrape(page, conn, args)

            if not args.no_excel and stats["saved"]:
                try:
                    export_flat(conn, PROJECT_DIR)
                except Exception as exc:
                    print(f"⚠️ Flat export failed (database is still fine): {exc}")

            print(
                f"\n🎉 Done. pages={stats['pages']} rows={stats['rows']} "
                f"saved={stats['saved']} failed={stats['failed']} skipped={stats['skipped']}"
            )
            for table in (
                "patients",
                "claims",
                "approval_history",
                "clinical_records",
                "vitals",
                "soap_notes",
                "invoices",
                "invoice_items",
                "equity_packages",
                "modal_fields",
                "modal_tables",
                "modal_cards",
            ):
                count = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
                print(f"   {table:<16} {count:>6}")
            return 0 if stats["failed"] == 0 else 1
        finally:
            try:
                context.close()
            except Exception:
                pass
            conn.commit()
            conn.close()


if __name__ == "__main__":
    sys.exit(main())
