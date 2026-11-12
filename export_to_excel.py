#!/usr/bin/env python3
"""
══════════════════════════════════════════════════════════════════════════
PMRS REJECTED CLAIMS - FLAT TEAM REVIEW EXPORT
══════════════════════════════════════════════════════════════════════════

Reads ``patients_data.db`` (the normalised scraper output) and flattens the
six related tables into one wide, human-readable sheet:

    patients  ← claims ← clinical_records
                        ← vitals
                        ← soap_notes
                        ← invoices (+ invoice_items count)

Every relation is joined with LEFT JOIN, so a claim is never dropped just
because one of its detail tables is empty. One row = one claim.

Usage
-----
    python export_to_excel.py                     # patients_data.db -> team_review_report.xlsx
    python export_to_excel.py --db other.db
    python export_to_excel.py --out review_2026_10.xlsx
    python export_to_excel.py --no-format          # raw pandas output, no styling

The workbook gets a styled header, frozen panes, an autofilter, sensible
column widths, wrapped long-text columns and a landscape print layout.
"""

from __future__ import annotations

import argparse
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DB = BASE_DIR / "patients_data.db"
DEFAULT_OUT = BASE_DIR / "team_review_report.xlsx"
SHEET_NAME = "Claims Review"

# ── look & feel ───────────────────────────────────────────────────────────
HEADER_BG = "14532D"       # dark green, same family as the extension theme
HEADER_FG = "FFFFFF"
BAND_BG = "F2F8F3"         # very light green for even rows
BORDER_COLOR = "C9D6CC"
TITLE_FONT = "Calibri"

# Columns that hold paragraphs — wrapped and given a wider cap.
LONG_TEXT_COLUMNS = {
    "Address",
    "Rejection Reason",
    "Admission Reason",
    "Diagnosis (In)",
    "Diagnosis (Out)",
    "Examination",
    "SOAP — Subjective",
    "SOAP — Objective",
    "SOAP — Assessment",
    "SOAP — Evaluation",
    "SOAP — Plan",
    "History of Present Illness",
    "Current Medication",
    "Past Medical History",
    "Past Surgical History",
    "Allergies",
    "Immunizations",
    "Family History",
    "Social History",
    "Target Group",
}

RIGHT_ALIGNED = {
    "Systolic",
    "Diastolic",
    "Pulse (/min)",
    "Temperature (°C)",
    "SpO2 (%)",
    "Weight (kg)",
    "Height (cm)",
    "Respiratory Rate (/min)",
    "Glucose",
    "Invoice Amount (KHR)",
    "Invoice Items",
    "Source Page",
    "Source Row",
}


# ══════════════════════════════════════════════════════════════════════════
# SQL: one flat row per claim
# ══════════════════════════════════════════════════════════════════════════
QUERY = """
WITH inv AS (
    -- Aggregate first: several invoices per claim must not multiply rows.
    SELECT  claim_id,
            COUNT(*)                 AS invoice_count,
            MAX(invoice_no)          AS invoice_no,
            MAX(invoice_date)        AS invoice_date,
            MAX(amount)              AS invoice_amount,
            MAX(currency)            AS currency,
            MAX(target_group)        AS target_group,
            MAX(target_group_amount) AS target_group_amount
    FROM    invoices
    GROUP BY claim_id
),
items AS (
    SELECT  v.claim_id,
            COUNT(*) AS item_count
    FROM    invoice_items i
    JOIN    invoices v ON v.invoice_id = i.invoice_id
    GROUP BY v.claim_id
)
SELECT
    -- ── patient ──────────────────────────────────────────────────────────
    COALESCE(p.nssf_id,        '') AS "NSSF ID",
    COALESCE(p.health_card_id, '') AS "Health Card No",
    COALESCE(p.equity_card_id, '') AS "Equity Card No",
    COALESCE(p.patient_name,   '') AS "Patient Name",
    COALESCE(p.gender,         '') AS "Gender",
    COALESCE(p.dob,            '') AS "Date of Birth",
    COALESCE(p.age,            '') AS "Age",
    COALESCE(p.phone,          '') AS "Phone",
    COALESCE(p.nationality,    '') AS "Nationality",
    COALESCE(p.address,        '') AS "Address",

    -- ── claim ────────────────────────────────────────────────────────────
    COALESCE(c.claim_no,        '') AS "Claim No",
    COALESCE(c.request_date,    '') AS "Request Date",
    COALESCE(c.district,        '') AS "District",
    COALESCE(c.facility,        '') AS "Facility",
    COALESCE(c.facility_level,  '') AS "Facility Level",
    COALESCE(c.status,          '') AS "Claim Status",
    COALESCE(c.rejection_reason,'') AS "Rejection Reason",
    COALESCE(c.rejected_at,     '') AS "Rejected At",
    COALESCE(c.rejected_by,     '') AS "Rejected By",
    COALESCE(c.payment_type,    '') AS "Payment Type",
    COALESCE(c.error_type,      '') AS "Error Type",
    COALESCE(c.total_amount,    '') AS "Claim Total Amount",

    -- ── clinical visit ───────────────────────────────────────────────────
    COALESCE(cr.department,       '') AS "Department (In)",
    COALESCE(cr.department_out,   '') AS "Department (Out)",
    COALESCE(cr.admission_date,   '') AS "Admission Date",
    COALESCE(cr.discharge_date,   '') AS "Discharge Date",
    COALESCE(cr.diagnosis_in,     '') AS "Diagnosis (In)",
    COALESCE(cr.diagnosis_out,    '') AS "Diagnosis (Out)",
    COALESCE(cr.icd10,            '') AS "ICD-10",
    COALESCE(cr.visit_type,       '') AS "Visit Type",
    COALESCE(cr.discharge_type,   '') AS "Discharge Type",
    COALESCE(cr.disease_status,   '') AS "Disease Status",
    COALESCE(cr.service_unit,     '') AS "Service Unit",
    COALESCE(cr.appointment_date, '') AS "Appointment Date",
    COALESCE(cr.admission_reason, '') AS "Admission Reason",

    -- ── vitals ───────────────────────────────────────────────────────────
    COALESCE(v.bp,        '') AS "BP (mmHg)",
    COALESCE(v.systolic,  '') AS "Systolic",
    COALESCE(v.diastolic, '') AS "Diastolic",
    COALESCE(v.pulse,     '') AS "Pulse (/min)",
    COALESCE(v.temp,      '') AS "Temperature (°C)",
    COALESCE(v.spo2,      '') AS "SpO2 (%)",
    COALESCE(v.weight,    '') AS "Weight (kg)",
    COALESCE(v.height,    '') AS "Height (cm)",
    COALESCE(v.resp_rate, '') AS "Respiratory Rate (/min)",
    COALESCE(v.glucose,   '') AS "Glucose",

    -- ── SOAP note + history ──────────────────────────────────────────────
    COALESCE(s.subjective,            '') AS "SOAP — Subjective",
    COALESCE(s.objective,             '') AS "SOAP — Objective",
    COALESCE(s.assessment,            '') AS "SOAP — Assessment",
    COALESCE(s.evaluation,            '') AS "SOAP — Evaluation",
    COALESCE(s.plan,                  '') AS "SOAP — Plan",
    COALESCE(s.hpi,                   '') AS "History of Present Illness",
    COALESCE(s.current_medication,    '') AS "Current Medication",
    COALESCE(s.past_medical_history,  '') AS "Past Medical History",
    COALESCE(s.past_surgical_history, '') AS "Past Surgical History",
    COALESCE(s.allergies,             '') AS "Allergies",
    COALESCE(s.immunizations,         '') AS "Immunizations",
    COALESCE(s.family_history,        '') AS "Family History",
    COALESCE(s.social_history,        '') AS "Social History",
    COALESCE(s.examination,           '') AS "Examination",
    COALESCE(s.soap_doctor,           '') AS "SOAP Doctor",
    COALESCE(s.soap_date,             '') AS "SOAP Date",

    -- ── billing ──────────────────────────────────────────────────────────
    COALESCE(i.invoice_no,          '') AS "Invoice No",
    COALESCE(i.invoice_date,        '') AS "Invoice Date",
    COALESCE(i.invoice_amount,      '') AS "Invoice Amount",
    COALESCE(i.currency,            '') AS "Currency",
    COALESCE(i.target_group,        '') AS "Target Group",
    COALESCE(i.target_group_amount, '') AS "Target Group Amount",
    COALESCE(it.item_count,          0) AS "Invoice Items",

    -- ── traceability ─────────────────────────────────────────────────────
    c.source_page                      AS "Source Page",
    c.source_row                       AS "Source Row",
    c.updated_at                       AS "Last Updated"
FROM        claims            c
LEFT JOIN   patients          p  ON p.patient_id = c.patient_id
LEFT JOIN   clinical_records  cr ON cr.claim_id  = c.claim_id
LEFT JOIN   vitals            v  ON v.claim_id   = c.claim_id
LEFT JOIN   soap_notes        s  ON s.claim_id   = c.claim_id
LEFT JOIN   inv               i  ON i.claim_id   = c.claim_id
LEFT JOIN   items             it ON it.claim_id  = c.claim_id
ORDER BY    c.claim_id
"""


# ══════════════════════════════════════════════════════════════════════════
# DATA
# ══════════════════════════════════════════════════════════════════════════
def connect(db_path: Path) -> sqlite3.Connection:
    if not db_path.exists():
        raise SystemExit(
            f"❌ Database not found: {db_path}\n"
            "   Run scrape_rejected_claims.py first, or pass --db <path>."
        )
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def table_names(conn: sqlite3.Connection) -> set[str]:
    rows = conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    return {row[0] for row in rows}


def required_tables() -> set[str]:
    return {"patients", "claims", "clinical_records", "vitals", "soap_notes", "invoices"}


def load_frame(conn: sqlite3.Connection) -> pd.DataFrame:
    missing = required_tables() - table_names(conn)
    if missing:
        raise SystemExit(
            "❌ Missing table(s): " + ", ".join(sorted(missing)) + "\n"
            "   Run scrape_rejected_claims.py once so it creates/migrates the schema."
        )

    try:
        frame = pd.read_sql_query(QUERY, conn)
    except sqlite3.OperationalError as exc:
        raise SystemExit(
            f"❌ Query failed: {exc}\n"
            "   The database schema looks older than the scraper expects.\n"
            "   Run scrape_rejected_claims.py once to apply its migrations."
        ) from exc

    # "40,000 ៛" -> 40000, so the team can sum/sort the invoice column.
    if "Invoice Amount" in frame.columns:
        digits = (
            frame["Invoice Amount"]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True)
        )
        numeric = pd.to_numeric(digits, errors="coerce")
        position = list(frame.columns).index("Invoice Amount") + 1
        frame.insert(position, "Invoice Amount (KHR)", numeric)

    return frame


# ══════════════════════════════════════════════════════════════════════════
# EXCEL FORMATTING
# ══════════════════════════════════════════════════════════════════════════
def _display_width(value: object) -> int:
    """Rough character width, tolerant of long paragraphs."""
    if value is None:
        return 0
    text = str(value)
    if not text:
        return 0
    longest_line = max((len(line) for line in text.splitlines()), default=0)
    return min(longest_line, 60)


def style_worksheet(ws, frame: pd.DataFrame) -> None:
    header_font = Font(name=TITLE_FONT, bold=True, color=HEADER_FG, size=11)
    header_fill = PatternFill("solid", fgColor=HEADER_BG)
    thin = Side(style="thin", color=BORDER_COLOR)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    band_fill = PatternFill("solid", fgColor=BAND_BG)

    columns = list(frame.columns)
    last_col = get_column_letter(len(columns))
    last_row = max(len(frame) + 1, 2)

    # Header row
    ws.row_dimensions[1].height = 34
    for index, name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=index)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    # Column widths: max(header, content) with a per-type cap.
    for index, name in enumerate(columns, start=1):
        header_width = max(len(part) for part in name.split()) + 2
        content_width = max(
            (_display_width(value) for value in frame[name].head(300)), default=0
        )
        cap = 52 if name in LONG_TEXT_COLUMNS else 26
        floor = 12 if name not in LONG_TEXT_COLUMNS else 22
        width = max(floor, min(cap, max(header_width, content_width + 2)))
        ws.column_dimensions[get_column_letter(index)].width = width

    # Body alignment / wrapping for the columns that need it.
    wrapped = {i for i, n in enumerate(columns, start=1) if n in LONG_TEXT_COLUMNS}
    right = {i for i, n in enumerate(columns, start=1) if n in RIGHT_ALIGNED}
    centered = {i for i, n in enumerate(columns, start=1) if n in
                {"NSSF ID", "Health Card No", "Equity Card No", "Gender", "Age",
                 "Facility Level", "Claim Status", "ICD-10", "Currency",
                 "Invoice No", "Source Page", "Source Row"}}

    if wrapped or right or centered:
        for row_index in range(2, len(frame) + 2):
            for col_index in wrapped | right | centered:
                cell = ws.cell(row=row_index, column=col_index)
                horizontal = "right" if col_index in right else (
                    "center" if col_index in centered else "left"
                )
                cell.alignment = Alignment(
                    horizontal=horizontal,
                    vertical="top",
                    wrap_text=col_index in wrapped,
                )

    # Zebra striping via conditional formatting (cheap on 10k+ rows).
    from openpyxl.formatting.rule import FormulaRule

    ws.conditional_formatting.add(
        f"A2:{last_col}{last_row}",
        FormulaRule(formula=["MOD(ROW(),2)=0"], fill=band_fill, stopIfTrue=False),
    )

    # Amount columns as real numbers.
    for name in ("Invoice Amount (KHR)", "Source Page", "Source Row", "Invoice Items"):
        if name not in columns:
            continue
        col_index = columns.index(name) + 1
        for row_index in range(2, len(frame) + 2):
            ws.cell(row=row_index, column=col_index).number_format = "#,##0"

    # Usability & print setup.
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{last_col}{last_row}"
    ws.sheet_view.zoomScale = 90
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


# ══════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════
def export(
    db_path: Path = DEFAULT_DB,
    out_path: Path = DEFAULT_OUT,
    styled: bool = True,
) -> tuple[Path, pd.DataFrame]:
    conn = connect(db_path)
    try:
        frame = load_frame(conn)
    finally:
        conn.close()

    if frame.empty:
        print("⚠️  The database has no claims yet - writing a header-only sheet.")

    out_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        workbook = writer.book
        worksheet = writer.sheets[SHEET_NAME]

        if styled:
            style_worksheet(worksheet, frame)

        workbook.properties.title = "PMRS rejected claims - team review"
        workbook.properties.subject = "Flat export of patients_data.db"
        workbook.properties.creator = "PMRS scraper"
        workbook.properties.created = datetime.now()
        workbook.properties.description = (
            f"{len(frame)} claim(s) exported from {db_path.name}"
        )

    return out_path, frame


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Flatten patients_data.db into a styled team-review Excel sheet."
    )
    parser.add_argument("--db", default=str(DEFAULT_DB), help="SQLite file (default: %(default)s)")
    parser.add_argument("--out", default=str(DEFAULT_OUT), help="Excel file (default: %(default)s)")
    parser.add_argument(
        "--no-format",
        action="store_true",
        help="skip header colours / widths (plain pandas output)",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    db_path = Path(args.db).expanduser().resolve()
    out_path = Path(args.out).expanduser().resolve()

    print(f"🗄️  Reading : {db_path}")
    result, frame = export(db_path, out_path, styled=not args.no_format)

    size_kb = result.stat().st_size / 1024
    print(f"📄 Written : {result}")
    print(f"   {len(frame):,} claim row(s) · {len(frame.columns)} column(s) · {size_kb:,.0f} KB")
    if not frame.empty:
        missing = int(frame["Claim No"].astype(str).str.strip().eq("").sum())
        if missing:
            print(f"   ⚠️  {missing} row(s) have no claim number.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
