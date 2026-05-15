---
name: idx-excel-quarterly
description: >
  Extract IDX quarterly financial report PDFs into an Excel workbook containing main financial
  statements only (Balance Sheet, Income Statement, Cash Flow, Changes in Equity, Key Ratios).
  No notes extraction. Supports batch mode: up to 4 PDFs (1 year of quarters) per run.
  Handles both creating a new workbook and appending periods to an existing one.
  Use this skill when the user provides quarterly (Q1/Q2/Q3) or interim financial statement PDFs.
---

# IDX Excel Extractor — Quarterly Statements

Extract main financial statements from quarterly/interim report PDFs.
Supports batch processing of up to 4 quarterly PDFs per run (recommended: 1 year at a time).
Produces a leaner workbook than the annual extractor: no notes sheets, no narrative sections.

---

## Role

You are an expert financial analyst and certified accountant with deep experience in:
- Indonesian capital markets (IDX/BEI) and OJK regulations
- PSAK (Indonesian Financial Accounting Standards) and IFRS
- Reading and interpreting Annual Reports and Financial Statements of Indonesian listed companies

---

## Required Input from User

| Parameter | Example |
|---|---|
| PDF file(s) | Single: `report/WIFI_Financial Report_2025Q2.pdf` |
| | Batch: `report/WIFI_Financial Report_2024Q1.pdf, ...Q2.pdf, ...Q3.pdf, ...Q4.pdf` |
| Ticker / company name | `WIFI – PT Solusi Sinergi Digital Tbk` |
| Period(s) | `2025Q2` or `2024Q1, 2024Q2, 2024Q3, 2024Q4` |
| Existing Excel (if appending) | `WIFI_Quarterly_Statements.xlsx` (or omit for new) |
| Output filename | `WIFI_Quarterly_Statements.xlsx` |

If any parameter is missing, ask the user before proceeding.

---

## Batch Processing Rules

### Recommended batch size: up to 4 PDFs per run (= 1 year of quarters)

| Batch size | Feasibility | Notes |
|---|---|---|
| 1 PDF | Always fine | Single quarter |
| 2–4 PDFs | Recommended max | One year of quarters, comfortable within context |
| 5–8 PDFs | Risky | May hit context limits; split into 2 runs |
| 9+ PDFs | Do not attempt | Split into multiple runs of 4 |

### Processing order within a batch

```
1. Sort all input PDFs chronologically (oldest first)
2. Process the FIRST PDF fully:
   → Probe structure, extract all statements, build/update Excel
3. Process each subsequent PDF:
   → Probe, extract, append columns to the SAME workbook
4. Save once at the end (not after each PDF)
```

**Why oldest first:** When the first PDF creates the row structure (new workbook mode),
starting with the oldest ensures that newer quarters append to an established structure.
Older reports sometimes have fewer line items, so rows added by newer reports leave
older period columns empty (with comment) — which is the correct behavior.

### Deduplication

Each quarterly PDF contains TWO columns: current period + comparative period.
Across a batch, the same period may appear multiple times as a comparative column
in different PDFs. Example:

```
PDF 2024Q1 → columns: 2024Q1, 2023Q4 (comparative BS) / 2023Q1 (comparative IS)
PDF 2024Q2 → columns: 2024Q2, 2023Q4 (comparative BS) / 2023Q2 (comparative IS)
PDF 2024Q3 → columns: 2024Q3, 2023Q4 (comparative BS) / 2023Q3 (comparative IS)
```

Here `2023Q4` appears as a BS comparative in all three PDFs.

**Rule: first write wins.** If a period column already exists in the workbook
(from an earlier PDF in the batch or a previous run), SKIP it. Do not overwrite.
Log it as "skipped (already exists)" in the terminal output.

---

## Key Differences: Quarterly vs Annual

| Aspect | Annual (Run 1 / Append) | Quarterly (this file) |
|---|---|---|
| Sheets produced | BS, IS, CF, Equity, Notes 1–N, Ratios | BS, IS, CF, Equity, Ratios ONLY |
| Notes extraction | Full (all notes) | NONE |
| Column headers | Year only: `2025` | Period: `2025Q2` or `Jun 2025` |
| Comparative columns in PDF | Prior year | Prior year same quarter + prior year-end (BS) |
| Typical PDF length | 40–80+ pages | 10–30 pages |
| Statement page locations | Varies widely by era | Generally more compact and consistent |

---

## General Rules

1. **Bilingual labels** — all labels use `Indonesian Term / English Term` format
2. **Data integrity** — NEVER fabricate or interpolate numbers; if not found in source, leave cell empty with cell comment explaining why
3. **Consolidated priority** — always use consolidated statements if both exist
4. **Indonesian number format** — dot (`.`) = thousands separator, parentheses = negative
5. **Chronological column order** — oldest period on LEFT, newest on RIGHT
6. **Period header format** — use `YYYYQN` in column headers (e.g. `2025Q1`), PLUS a sub-header row on each sheet explaining the reporting basis (see Period Labeling below)
7. **Existing rows are immutable** (when appending) — do NOT rename, reorder, or delete existing rows
8. **New rows are additive** (when appending) — if new PDF has a line item not in the workbook, INSERT at the logical position; leave existing period columns empty + comment
9. **Batch limit** — process up to 4 quarterly PDFs per run; if more provided, process first 4, instruct user to run again with remaining + output Excel
10. **Empty cell, not N/A string** — NEVER write string "N/A" into data cells — it breaks Excel formulas. Leave cell empty + cell comment. Use `0` only when report explicitly shows zero.
11. **Backup before mutation** (append mode) — save backup as `[FILENAME]_backup.xlsx` before any column/row insertion

---

## Period Labeling & Reporting Basis

This is the most critical quarterly-specific concept. BS is point-in-time; IS and CF are YTD cumulative. A single `2025Q2` header means different things on different sheets.

### Column headers

Use `YYYYQN` as the primary column header across all sheets for sortability.

### Sub-header row (Row 2 on each sheet)

Add a second header row below the period header that states the reporting basis:

| Sheet | Sub-header format | Example for 2025Q2 |
|---|---|---|
| Balance Sheet | `Per [DATE]` | `Per 30 Jun 2025` |
| Income Statement | `[N] bulan s.d. [DATE] / [N] months ended [DATE]` | `6 bulan s.d. 30 Jun 2025 / 6 months ended 30 Jun 2025` |
| Cash Flow | `[N] bulan s.d. [DATE] / [N] months ended [DATE]` | `6 bulan s.d. 30 Jun 2025 / 6 months ended 30 Jun 2025` |
| Equity | `[N] bulan s.d. [DATE] / [N] months ended [DATE]` | `6 bulan s.d. 30 Jun 2025 / 6 months ended 30 Jun 2025` |

This removes all ambiguity about what each column represents.

### Comparative period handling per statement type

Each quarterly PDF shows TWO columns. The comparative period DIFFERS by statement:

| Statement | Current column | Comparative column | Why different |
|---|---|---|---|
| Balance Sheet | Point-in-time: current quarter-end | Point-in-time: prior YEAR-END (31 Dec) | BS compares balance dates |
| Income Statement | YTD: N months current year | YTD: same N months prior year | IS compares like-for-like periods |
| Cash Flow | YTD: N months current year | YTD: same N months prior year | CF compares like-for-like periods |
| Equity | Current period movements | Prior year movements | Shows full movement reconciliation |

**Example for Q2 2025 PDF:**
```
BS columns:   2025Q2 (per 30 Jun 2025)     vs  2024Q4 (per 31 Dec 2024)
IS columns:   2025Q2 (6mo to 30 Jun 2025)  vs  2024Q2 (6mo to 30 Jun 2024)
CF columns:   2025Q2 (6mo to 30 Jun 2025)  vs  2024Q2 (6mo to 30 Jun 2024)
```

**CRITICAL:** Read the actual column headers from the PDF to determine comparative periods.
Do NOT assume. The PDF header will show dates like:
```
"30 Juni 2025 / June 30, 2025    31 Desember 2024 / December 31, 2024"
```

When inserting comparative columns, check PER SHEET whether that period already exists,
because BS and IS comparatives may be different periods from the same PDF.

---

## Quarterly PDF Structure

Quarterly reports (Laporan Keuangan Interim) are compact:

```
Page 1–2  : Cover, table of contents, or review report
Page 3–4  : Balance Sheet (Laporan Posisi Keuangan)
Page 5–6  : Income Statement (Laporan Laba Rugi)
Page 7    : Changes in Equity (Laporan Perubahan Ekuitas)
Page 8    : Cash Flow Statement (Laporan Arus Kas)
Page 9+   : Selected Notes (SKIP — not extracted in quarterly mode)
```

FS pages are near the front (unlike Annual Reports where they're in the back).
No need to search for an auditor's report — quarterly FS are typically unaudited/reviewed.

---

## Q4 / Full-Year Detection (Decision Tree)

When a PDF is detected as Q4 (Desember / 12 months), determine what type it is:

```
Is the period "31 Desember" or "12 months"?
  │
  ├─ NO → Normal quarterly (Q1/Q2/Q3). Proceed normally.
  │
  └─ YES → This is a Q4 / full-year document. What kind?
       │
       ├─ Does the PDF have 30+ pages with narrative sections
       │  (company profile, MD&A, governance)?
       │    └─ YES → This is an ANNUAL REPORT (Laporan Tahunan).
       │             STOP. Tell user:
       │             "This is an Annual Report. Use idx-excel-run1-template.md
       │              or idx-excel-append-year.md for full extraction with notes."
       │
       ├─ Does the PDF have <30 pages, mostly FS + notes,
       │  and an auditor's report on pages 1–3?
       │    └─ YES → This is an AUDITED full-year Financial Report.
       │             WARN: "This is an audited full-year FS, not an interim report.
       │             Annual extractor will produce a more complete result with notes.
       │             Proceed with quarterly extractor (main statements only)?"
       │             If user confirms → proceed. Label columns as YYYY (not YYYYQ4).
       │
       └─ Does the PDF have <30 pages, FS + limited notes,
          and a review report (or no auditor's report)?
            └─ YES → This is an UNAUDITED Q4 interim report.
                     WARN: "This Q4 report appears unaudited. Values may differ
                     from the audited annual figures."
                     Proceed. Label columns as YYYYQ4.
```

---

## Execution Sequence

### Overview: Batch Loop Structure

```
Phase 0  : Mode Detection + Workbook Setup          ← ONCE
  ┌─── FOR EACH PDF (sorted chronologically, oldest first) ───┐
  │ Phase 1  : Structure Probe + Period Detection (this PDF)   │
  │ Phase 2  : Main Statement Extraction (this PDF)            │
  │ Phase 3  : Row Matching + Column Insertion (this PDF)      │
  │ Phase 4  : Write to Workbook (this PDF's data)             │
  │            → deduplicate: skip periods already in workbook │
  └────────────────────────────────────────────────────────────┘
Phase 5  : Key Ratios (all periods at once)          ← ONCE
Phase 6  : Validation & Save                         ← ONCE
```

### Mode Detection

```
1. Check if an existing Excel file was provided
   → YES: Append Mode (go to Phase 0A)
   → NO:  New Workbook Mode (go to Phase 0B)
2. Sort all input PDFs chronologically (oldest first)
3. If more than 4 PDFs provided:
   → Process only the first 4
   → Print: "Processing 4 of [N] PDFs. Run again with the output Excel
     and remaining PDFs to continue."
```

### Phase 0A — Read Existing Workbook & Backup (Append Mode only)

```
1. BACKUP: Copy existing Excel to [FILENAME]_backup.xlsx

2. Load existing Excel with openpyxl (NOT data_only — preserve formulas)
3. For each sheet, record:
   a. Sheet name
   b. All labels in Column A
   c. All period headers and their column positions
   d. Merged cell ranges (column insertion can break these)
4. Build a set of existing_periods for deduplication (check PER SHEET,
   because BS may have 2024Q4 while IS does not)
5. Print existing structure to terminal
```

### Phase 0B — Setup New Workbook (New Workbook Mode only)

```
1. Create a new openpyxl Workbook
2. Plan sheet structure:
   Sheet 1: Balance Sheet
   Sheet 2: Income Statement
   Sheet 3: Cash Flow Statement
   Sheet 4: Changes in Equity
   Sheet 5: Key Ratios Summary
3. No notes sheets — quarterly mode skips notes entirely
4. Initialize existing_periods = empty set
```

---

### ═══ BEGIN LOOP: for each PDF (oldest first) ═══

### Phase 1 — Structure Probe (per PDF)

```
1. Open the quarterly PDF with pdfplumber
2. Count total pages
3. For each page:
   a. Extract text
   b. Print: page number, first 120 chars (trimmed)
4. Identify pages containing:
   - Balance Sheet        → "LAPORAN POSISI KEUANGAN" or "NERACA"
   - Income Statement     → "LAPORAN LABA RUGI"
   - Cash Flow Statement  → "LAPORAN ARUS KAS"
   - Changes in Equity    → "LAPORAN PERUBAHAN EKUITAS"
5. Record page ranges for each statement
6. Identify the period from the report header:
   → Look for date patterns: "30 Juni 2025", "31 Maret 2024", "30 September 2025"
   → Map to quarter: Maret=Q1, Juni=Q2, September=Q3, Desember=Q4
7. Identify the comparative period from the PDF column headers
8. Print structure map to terminal
```

**Period detection keywords:**

| Indonesian | English | Quarter |
|---|---|---|
| 31 Maret / 31 March | Three months | Q1 |
| 30 Juni / 30 June | Six months | Q2 |
| 30 September / 30 September | Nine months | Q3 |
| 31 Desember / 31 December | Twelve months | Q4 (annual) |

If the PDF is a Q4 / full-year report, warn the user:
"This appears to be a full-year (Q4) report. Consider using idx-excel-run1-template.md
or idx-excel-append-year.md instead for full annual extraction including notes."
Proceed anyway if user confirms.

### Phase 2 — Main Statement Extraction (per PDF)

For each statement (BS, IS, CF, Equity):

```
1. Extract raw text from identified pages
2. Run parse_financial_lines() on raw text
3. Print ALL parsed rows to terminal
4. Identify TWO sets of numbers per row (current period + comparative period)
5. Deduplication check:
   → If current_period is in existing_periods → skip current column
   → If comparative_period is in existing_periods → skip comparative column
   → If BOTH already exist → skip this PDF entirely, print warning
6. Validate BS: Total Assets = Total Liabilities + Total Equity (both columns)
7. Add newly written periods to existing_periods set
```

#### Handling Two Columns of Numbers

Quarterly PDFs have two value columns. When parsing:

```python
# Each parsed row will typically have 2 numbers
# Row format in PDF: "Label    123.456.789    98.765.432"
#                     ^^^^^^^^ ^^^^^^^^^^^    ^^^^^^^^^^
#                     label    current_period  comparative

# parse_financial_lines() should return:
# (label, [value_current, value_comparative])
```

Map them to the correct periods:
```
current_value     → column header = detected quarter (e.g. "2025Q2")
comparative_value → column header = comparative period:
  - For BS: prior year-end (e.g. "2024Q4")
  - For IS/CF: same quarter prior year (e.g. "2024Q2")
```

Verify the comparative period by reading the column headers in the PDF.
Do NOT assume — the PDF header will say something like:
```
"30 Juni 2025 / June 30, 2025    31 Desember 2024 / December 31, 2024"
```

#### Number Parsing (same as annual)

```python
import re

_NUM_RE = re.compile(
    r'\([\d]{1,3}(?:\.[\d]{3})+(?:\s*\.[\d]{3})*(?:\s*)?\)'
    r'|'
    r'[\d]{1,3}(?:\.[\d]{3})+(?:\s*\.[\d]{3})*'
)

def parse_id_number(s: str) -> int | None:
    s = s.strip()
    if not s or s == '-' or s == '—':
        return None
    neg = s.startswith('(') and s.endswith(')')
    if neg:
        s = s[1:-1]
    s = s.replace(' ', '').replace('.', '')
    try:
        val = int(s)
        return -val if neg else val
    except ValueError:
        return None
```

#### Known Pitfalls (same as annual — all still apply)

| Problem | Solution |
|---|---|
| "Total Aset" matches "Total Aset Lancar" | Exclusion list: `excl: ["lancar", "tidak"]` |
| Two numbers per row instead of one | Parse both; map to correct period columns |
| Comparative column may differ BS vs IS | Read period from PDF header, don't assume |
| Labels split across lines | Match fragments, same as annual |
| Space artifacts in numbers | Regex handles with `(?:\s*\.[\d]{3})*` |
| Total Liabilities unreliable from direct match | Fallback: Total Assets − Total Equity |

### Phase 3 — Row Matching & Column Insertion (per PDF)

Applies in BOTH modes — after the first PDF in a batch, even a new workbook has existing rows.

```
1. For each sheet, get existing labels from Column A
2. Determine insert position for each new period column:
   → Parse all existing period headers into sortable values
   → Insert at correct chronological position (oldest left, newest right)
   → Check PER SHEET: BS and IS may have different existing periods
3. For each parsed row from the current PDF:
   a. Normalize label (lowercase, strip, remove punctuation)
   b. Try exact match against existing Column A labels
   c. Try fuzzy match for known variations:
      - "Beban Pokok Penjualan" ↔ "Harga Pokok Penjualan"
      - "Kas dan Bank" ↔ "Kas dan Setara Kas"
   d. No match → insert new row, leave existing period columns empty + comment
4. For existing labels with no match in this PDF → leave new period column empty + comment
```

### Phase 4 — Write to Workbook (per PDF)

#### Formatting Rules (same as annual)

| Element | Format |
|---|---|
| Font | Arial 10pt |
| Column A width | 55 characters |
| Data column width | 18 characters |
| Period header | Bold, centered, text format |
| Currency header | `IDR – full amount` (or as in report) |
| Subtotals | Bold, thin top border |
| Grand totals | Bold, double top border |
| Section headers | Merged, bold, light gray fill (#F2F2F2) |
| Negative numbers | `#,##0_);[Red](#,##0)` |
| Period values | Text format (avoid "2,025") |

#### For New Workbook:

```
Sheet 1 — Balance Sheet:
  Row 1: "Laporan Posisi Keuangan / Balance Sheet"
  Row 2: Reporting basis sub-headers — e.g. B2: "Per 30 Jun 2025", C2: "Per 31 Des 2024"
  Row 3: Column headers — A: "Keterangan / Description", B: [2025Q2], C: [2024Q4]
  Row 4: Currency/unit note — "Dalam Rupiah penuh / In full IDR"
  Row 5: Empty separator
  Row 6+: All line items from PDF

Sheet 2 — Income Statement:
  Row 1: "Laporan Laba Rugi / Income Statement"
  Row 2: Reporting basis — e.g. B2: "6 bulan s.d. 30 Jun 2025", C2: "6 bulan s.d. 30 Jun 2024"
  Row 3+: Same structure as BS

Sheet 3 — Cash Flow Statement:
  Same structure, CF-specific reporting basis in Row 2

Sheet 4 — Changes in Equity:
  Same structure

Sheet 5 — Key Ratios Summary:
  Row 2 sub-headers: note which ratios use BS (point-in-time) vs IS/CF (YTD)
```

#### For Append Mode:

```
1. Insert new column(s) at correct chronological position
2. Populate values matched to existing row labels
3. Insert new rows for items not previously seen
4. Skip periods that already exist in the workbook
```

### ═══ END LOOP ═══

---

### Phase 5 — Key Ratios (ONCE, after all PDFs processed)

Calculate these ratios for each period extracted.

**CRITICAL — quarterly ratio interpretation rules:**

Quarterly ratios mix two measurement bases:
- **BS-based ratios** (liquidity, leverage) use point-in-time values → directly comparable across quarters
- **IS/CF-based ratios** (profitability, cash flow) use YTD cumulative values → NOT directly comparable across different quarters (Q2 YTD covers 6 months, Q3 YTD covers 9 months)

**DO NOT annualize.** Annualizing (multiplying Q2 by 2, Q3 by 4/3) introduces assumptions.
Report YTD figures as-is. Add a footnote explaining the basis.

**Liquidity (BS-based — directly comparable):**
- Current Ratio = Current Assets / Current Liabilities
- Quick Ratio = (Current Assets − Inventory) / Current Liabilities
- Cash Ratio = Cash & Equivalents / Current Liabilities

**Leverage (BS-based — directly comparable):**
- Debt-to-Equity = Total Liabilities / Total Equity
- Interest Coverage = EBIT / Finance Costs — **CAUTION: EBIT is YTD; add footnote**

**Profitability (IS-based — YTD, add footnote per ratio):**
- Gross Margin = Gross Profit / Revenue — footnote: "YTD basis"
- Operating Margin = Operating Profit / Revenue — footnote: "YTD basis"
- Net Margin = Net Profit / Revenue — footnote: "YTD basis"
- ROE = Net Profit / Total Equity — footnote: "YTD profit ÷ period-end equity; not annualized"
- ROA = Net Profit / Total Assets — footnote: "YTD profit ÷ period-end assets; not annualized"

**Cash Flow (CF-based — YTD):**
- Free Cash Flow = Operating CF − Capex — footnote: "YTD basis"
- FCF Margin = FCF / Revenue — footnote: "YTD basis"

**Footnote rows to add at bottom of Ratios sheet:**
```
Row N+1: (italic) "Profitability and cash flow ratios use YTD figures as reported.
         Q1 = 3 months, Q2 = 6 months, Q3 = 9 months, Q4 = 12 months.
         These are NOT directly comparable across different quarters."
Row N+2: (italic) "ROE and ROA use period-end BS values, not averages. Not annualized."
Row N+3: (italic) "Liquidity and leverage ratios use point-in-time BS values
         and are directly comparable across quarters."
```

Format: percentages as `0.0%`, multiples as `0.0x`, leave empty if not computable.

Use Excel formulas referencing main statement sheet cells where possible.

### Phase 6 — Validation & Save

```
STEP 1 — POST-INSERTION INTEGRITY CHECK (append mode only):

   a. Verify no #REF! errors across all sheets
   b. Verify merged cells still intact
   c. Spot-check 3 existing values haven't changed (Total Assets, Revenue, Net Profit)
   d. Verify column headers still chronologically ordered
   e. If any check fails → restore from backup, report failure, STOP

STEP 2 — FINANCIAL VALIDATION (for each newly added period):

   BALANCE SHEET:
   a. Total Assets = Total Liabilities + Total Equity
   b. Total Current Assets + Total Non-Current Assets = Total Assets

   INCOME STATEMENT:
   c. Revenue − COGS = Gross Profit

   CASH FLOW:
   d. Operating CF + Investing CF + Financing CF = Net Change in Cash (± FX effect)
   e. Cash Beginning + Net Change = Cash End

   CROSS-STATEMENT:
   f. Cash End in CF ≈ Cash & Equivalents in BS (same period)

   For each failure → add footnote row. Do NOT silently adjust.

STEP 3 — SAVE & RECALC:
   a. Save workbook
   b. Run recalc: python /mnt/skills/public/xlsx/scripts/recalc.py [output_file]
   c. Check for formula errors; fix any found
   d. Print terminal summary
```

### Anti-Hallucination Rules

```
1. NEVER fill a cell with a calculated or estimated value without a footnote.

2. FALLBACK CALCULATIONS allowed ONLY for:
   - Total Liabilities = Total Assets − Total Equity (when direct match fails)
   Every fallback MUST have an italic footnote row.

3. COMPARATIVE PERIOD values: extract them ONLY if they don't already exist in the
   workbook. Check PER SHEET (BS comparative may differ from IS comparative).

4. CONSOLIDATED vs STANDALONE: verify statement title says "KONSOLIDASIAN".
   If not, check if the company has subsidiaries. Log the finding.

5. ZERO vs EMPTY: report shows 0 or dash → put 0. Not found → leave empty + comment.

6. DO NOT ANNUALIZE: never multiply quarterly IS/CF figures to project annual values.

7. WORKBOOK DATA IS NOT A SOURCE: every value in a new column must come from the PDF.
   Never copy from existing columns.

8. YTD AWARENESS: when two IS/CF periods from different quarters appear side by side,
   they cover different durations. Q1=3mo, Q2=6mo, Q3=9mo. Do not compare them as
   if they represent equal time periods. The sub-header row makes this explicit.
```

---

## Terminal Output

```
=== IDX Quarterly Excel Complete ===
Company      : [TICKER] – [COMPANY NAME]
Mode         : [NEW WORKBOOK / APPEND]
PDFs Processed: [N] of [TOTAL]
Output File  : [XLSX PATH]
Backup File  : [PATH or N/A — new workbook mode]
All Periods  : [LIST ALL PERIODS IN WORKBOOK, left to right]

Sheets:
  1. Balance Sheet         — periods: [list with reporting basis]
  2. Income Statement      — periods: [list with reporting basis]
  3. Cash Flow Statement   — periods: [list with reporting basis]
  4. Changes in Equity     — periods: [list]
  5. Key Ratios Summary    — periods: [list]

Per-PDF Results:
  PDF 1: [FILENAME]
    Detected period  : [2024Q1] (3 months to 31 Mar 2024)
    Q4 check         : [N/A — not Q4 / WARNED — Q4 detected, user confirmed]
    Current column   : [2024Q1] — written
    BS comparative   : [2023Q4] — written / SKIPPED (exists)
    IS/CF comparative: [2023Q1] — written / SKIPPED (exists)
  PDF 2: [FILENAME]
    ...

Post-Insertion Integrity (append mode):
  #REF! errors       : [NONE / count]
  Existing data check : [PASS / FAIL]

Financial Validation:
  [PERIOD]: BS identity [PASS/FAIL] | IS gross profit [PASS/FAIL] | CF net change [PASS/FAIL]
  ...

Fallback Calculations:
  [LIST any computed values]

Row Changes:
  [N] new rows inserted across all sheets
  [N] cells left empty (data not available)

Remaining PDFs (if batch exceeded 4):
  [LIST UNPROCESSED FILES]

Warnings:
  [Q4 detections, parsing issues, duplicates, etc.]
==========================================
```

---

## Combining Quarterly and Annual Data

The quarterly workbook (`WIFI_Quarterly_Statements.xlsx`) is intentionally SEPARATE
from the annual workbook (`WIFI_Financial_Statements.xlsx`).

Reasons:
- Annual has notes sheets; quarterly does not
- Annual figures may differ from Q4 quarterly (audit adjustments)
- Mixing granularities in one file creates confusing column headers

If the user wants a combined view, suggest creating a separate summary workbook
that pulls key metrics from both files — but that is a different task.

---

## Debugging Tips

1. **Two numbers per row** — the most common quarterly-specific issue; always verify you're capturing both columns
2. **Period detection** — read the PDF header carefully; don't assume Q2 2025 comparative is always Q2 2024
3. **YTD vs single-quarter** — IS and CF are cumulative; BS is point-in-time. The sub-header row must reflect this.
4. **Q4 decision tree** — follow the full decision tree above; don't just "warn and proceed" blindly
5. **Comparative differs by statement** — BS comparative may be 2024Q4 while IS comparative is 2024Q2, from the same PDF. Check PER SHEET when deduplicating.
6. **Fragment matching** — same PDF rendering issues as annual reports apply
7. **Batch column drift** — after inserting multiple columns in one session, verify that all column references in the Ratios sheet still point correctly
8. **Sort PDFs before processing** — oldest first ensures stable row structure
9. **Save once at the end** — do not save the workbook after each PDF in the loop
10. **Never annualize** — do not multiply partial-year IS/CF figures. Report as-is with basis footnotes.
11. **Backup is your safety net** — if post-insertion integrity fails, restore from backup
