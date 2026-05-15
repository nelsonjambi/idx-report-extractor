---
name: idx-excel-run1-template
description: >
  Extract audited financial statements from an IDX-listed company's Annual Report (Laporan Tahunan)
  into a comprehensive Excel workbook. Annual Reports contain audited FS in the back portion,
  typically starting 40–60% into the PDF. This skill locates the audited FS section,
  extracts all statements and notes, and builds a single-year template Excel for future year appending.
  Do NOT use for standalone Financial Reports (Laporan Keuangan Q4) — those are often unaudited.
---

# IDX Excel Extractor — Run 1: Single-Year Template

Build a complete Excel workbook from the **audited financial statements** found inside
ONE Annual Report (Laporan Tahunan) PDF — the most recent fiscal year.
This workbook becomes the template; future runs will append earlier years as new columns.

**IMPORTANT — source file distinction:**

| Document type | Indonesian | Contents | Use this skill? |
|---|---|---|---|
| **Annual Report** | Laporan Tahunan | Narrative (100+ pages) + audited FS in the back | ✅ YES |
| Financial Report Q4 | Laporan Keuangan Q4 | FS + notes only, often unaudited | ❌ NO — use quarterly extractor |
| Financial Report Q1–Q3 | Laporan Keuangan Interim | FS + limited notes, unaudited | ❌ NO — use quarterly extractor |

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
| PDF file path | `report/WIFI_Annual_Report_2025.pdf` |
| Ticker / company name | `WIFI – PT Solusi Sinergi Digital Tbk` |
| Fiscal year | `2025` |
| Output filename | `WIFI_Financial_Statements.xlsx` |

If any parameter is missing, ask the user before proceeding.

If the filename suggests a standalone Financial Report (e.g. `Laporan_Keuangan_Q4`,
`Financial_Report_2025Q4`), warn the user:
"This appears to be a standalone Financial Report, not an Annual Report.
Standalone Q4 reports are often unaudited. Consider using idx-excel-quarterly.md instead.
Proceed anyway?"

---

## Annual Report PDF Structure

Annual Reports (Laporan Tahunan) of IDX-listed companies follow a roughly standard layout.
The audited financial statements are in the BACK portion, not the front.

```
Typical Annual Report structure (100–250+ pages):

Pages 1–5      : Cover, Table of Contents
Pages 5–15     : Financial Highlights / Ikhtisar Keuangan
Pages 15–30    : Chairman & Directors' Reports / Sambutan
Pages 30–60    : Company Profile, Business Overview
Pages 60–80    : Management Discussion & Analysis
Pages 80–100   : Corporate Governance
Pages 100–120  : CSR / Sustainability
Pages 120–130  : Independent Auditor's Report  ← LOOK FOR THIS
Pages 130–170  : Audited Financial Statements  ← THIS IS WHAT WE EXTRACT
  ├── Balance Sheet (Laporan Posisi Keuangan)
  ├── Income Statement (Laporan Laba Rugi)
  ├── Changes in Equity (Laporan Perubahan Ekuitas)
  ├── Cash Flow Statement (Laporan Arus Kas)
  └── Notes to Financial Statements (Catatan atas Laporan Keuangan)
Pages 170+     : Appendices, proxy statement
```

**Page numbers are approximate.** Small companies may have 80-page reports;
large conglomerates (ASII, BBCA) can exceed 500 pages. The key is to FIND
the auditor's report, which immediately precedes the financial statements.

### How to locate the audited FS section

Search for these markers (in order of reliability):

1. **Auditor's Report** — the most reliable anchor point:
   - "LAPORAN AUDITOR INDEPENDEN" / "INDEPENDENT AUDITOR'S REPORT"
   - "Laporan Akuntan Independen"
   - The FS section starts 1–3 pages AFTER the auditor's signature page

2. **First financial statement title** — after the auditor's report:
   - "LAPORAN POSISI KEUANGAN KONSOLIDASIAN" (Consolidated Balance Sheet)
   - Key word: "KONSOLIDASIAN" (Consolidated) — this distinguishes audited
     consolidated FS from the financial highlights summary near the front

3. **Standalone vs Consolidated** — the financial highlights section (pages 5–15)
   often has a condensed BS/IS. Do NOT extract from there. Always look for
   "KONSOLIDASIAN" / "CONSOLIDATED" in the statement title.

---

## General Rules (apply throughout without exception)

1. **Bilingual labels** — all labels use `Indonesian Term / English Term` format
2. **Data integrity** — NEVER fabricate or interpolate numbers; use `N/A` or `—` if not found in source
3. **Consolidated priority** — if both consolidated and standalone statements exist, always use consolidated
4. **Indonesian number format** — thousands separator is dot (`.`), negative in parentheses `(123.456.789)`
5. **Single-year layout** — Column A = labels, Column B = the single fiscal year's values
6. **Scanned PDFs** — if a page cannot be parsed as text, attempt OCR; flag values with `[OCR – verify]`
7. **Audited FS only** — extract ONLY from the audited FS section (after the auditor's report), never from financial highlights or summary tables near the front of the Annual Report
8. **Source pages** — record which PDF pages were used for each statement; include in terminal output

---

## Execution Sequence

Follow this exact order. Each phase completes fully before starting the next.

### Phase 0 — Environment Setup

```bash
pip install pdfplumber openpyxl --break-system-packages
```

### Phase 1 — Structure Probe

Before writing any extraction code, probe the PDF to locate the audited FS section.
Annual Reports have financial highlights near the front AND audited FS in the back.
You must find the RIGHT section.

```
1. Open the PDF with pdfplumber
2. Count total pages (expect 80–300+)
3. For EVERY page, extract text and store: pages[page_num] = first 200 chars

4. STEP A — Find the Auditor's Report (the anchor point):
   Scan ALL pages for:
   - "LAPORAN AUDITOR INDEPENDEN" (case-insensitive)
   - "INDEPENDENT AUDITOR" (case-insensitive)
   - "LAPORAN AKUNTAN INDEPENDEN" (case-insensitive)
   Record: auditor_report_page = first match
   → If NOT found: warn user "Auditor's report not found — FS may be unaudited or
     this may not be an Annual Report." Continue, but flag all output as [unverified].

5. STEP B — Find financial statements AFTER the auditor's report:
   Scan pages from (auditor_report_page + 1) onward for:
   - Balance Sheet   → "LAPORAN POSISI KEUANGAN" containing "KONSOLIDASIAN"
   - Income Statement → "LAPORAN LABA RUGI" containing "KONSOLIDASIAN"
   - Cash Flow       → "LAPORAN ARUS KAS" (may be shortened across lines)
   - Equity          → "LAPORAN PERUBAHAN EKUITAS"

   CRITICAL: Do NOT match occurrences BEFORE the auditor's report.
   Those are the financial highlights / summary section, not the audited FS.

   If "KONSOLIDASIAN" is not found on BS/IS pages, check if the company
   has no subsidiaries (standalone FS is acceptable in that case). Log the finding.

6. STEP C — Determine FS page ranges:
   The statements always appear in a fixed order:
     BS → IS → Equity → CF (or BS → IS → CF → Equity, varies by company)
   Each statement ends where the next one begins.
   Record: fs_pages = {
     auditor: auditor_report_page,
     bs: [start, end],
     is: [start, end],
     cf: [start, end],
     eq: [start, end]
   }

7. STEP D — Find notes section start:
   Notes start AFTER the last main statement.
   Scan from (last_statement_end + 1) for lines matching:
     ^(\d{1,2})\.\s+([A-Z]{3,}...)
   Skip lines with "(lanjutan)" or "(continued)".
   Record: notes_min_page

8. Print the full structure map to terminal:
   - Total pages
   - Auditor's report: page [N]
   - BS: pages [N]–[N]
   - IS: pages [N]–[N]
   - CF: pages [N]–[N]
   - Equity: pages [N]–[N]
   - Notes start: page [N]
   - Estimated notes end: page [N] (last page of PDF or before appendices)
```

**CRITICAL:** The header "CATATAN ATAS LAPORAN KEUANGAN / NOTES TO THE CONSOLIDATED
FINANCIAL STATEMENTS" appears on EVERY page of the FS section — do NOT use it as a
notes section start marker.

#### Guarding against Financial Highlights false matches

The financial highlights section (typically pages 5–15) often contains:
- "Laporan Posisi Keuangan" as a table header
- Summary BS/IS numbers in a condensed format

These will match the same search strings. The auditor's report anchor prevents this:
only match statements on pages AFTER the auditor's report. If no auditor's report is
found, use a secondary heuristic: the audited FS section is usually in the last 40%
of the PDF. Require statement pages to be past the 50% mark of total pages.

### Phase 2 — Main Statement Extraction

For each main statement (BS, IS, CF, Equity):

```
1. Extract raw text from the identified pages
2. Run parse_financial_lines() on the raw text
3. Print ALL parsed rows (label + values) to terminal
4. Build matchers for the target line items
5. Validate: Total Assets = Total Liabilities + Total Equity (BS identity check)
6. If validation fails, debug before proceeding
```

#### Number Parsing Rules

```python
import re

_NUM_RE = re.compile(
    r'\([\d]{1,3}(?:\.[\d]{3})+(?:\s*\.[\d]{3})*(?:\s*)?\)'  # negative (parentheses)
    r'|'
    r'[\d]{1,3}(?:\.[\d]{3})+(?:\s*\.[\d]{3})*'               # positive
)

def parse_id_number(s: str) -> int | None:
    """Parse Indonesian-format number: dots as thousands, parens as negative."""
    s = s.strip()
    if not s or s == '-' or s == '—':
        return None
    neg = s.startswith('(') and s.endswith(')')
    if neg:
        s = s[1:-1]
    s = s.replace(' ', '').replace('.', '')  # remove space artifacts and dots
    try:
        val = int(s)
        return -val if neg else val
    except ValueError:
        return None
```

#### Known Pitfalls

| Problem | Solution |
|---|---|
| "Total Aset" matches "Total Aset Lancar" | Use exclusion list: `excl: ["lancar", "tidak"]` |
| Early reports use "Kas dan bank" not "Kas dan setara kas" | Include both patterns in matcher |
| Trade receivables header has no numbers; data under "Pihak ketiga" | Match "pihak ketiga" too |
| "LABA SEBELUM PAJAK PENGHASILAN" split across lines | Match fragment "penghasilan" with exclusions |
| Statement title split across PDF lines | Use shortened search strings with secondary validation |
| Space artifacts in numbers: `97.27 8.291.602` | Regex handles with `(?:\s*\.[\d]{3})*` |
| Total Liabilities sometimes wrong from direct match | Fallback: compute as Total Assets − Total Equity |

### Phase 3 — Notes Index Building

Scan pages from `notes_min_page` to end of PDF:

```
1. For each page from notes_min_page onward:
   a. Extract text
   b. Look for lines matching: ^(\d{1,2})\.\s+([A-Z][A-Z\s/\(\)\-]{3,})
   c. SKIP lines containing "(lanjutan)" or "(continued)" — these are continuation headers
   d. Record: note_index[note_num] = {title_id, title_en, start_page}
2. Compute end_page for each note = start_page of next note − 1
   (last note runs to end of PDF or to a clearly different section)
3. Print the notes index to terminal for verification
```

**False positive guards:**
- Require title text after the number to be ALL CAPS (≥3 characters)
- Require page ≥ `notes_min_page`
- Numbered lists inside note text can match — filter by requiring the match to start a major section

### Phase 4 — Notes Content Extraction

For each note in `note_index`, classify and extract:

#### Type A: Quantitative / Tabular Notes
Examples: Cash, Trade Receivables, Inventories, Fixed Assets, Debt

```
1. Extract text from note pages using pdfplumber extract_text()
2. Re-use parse_financial_lines() — same bilingual number format
3. For Fixed Assets movement tables (opening/additions/disposals/closing):
   → use pdfplumber extract_tables() instead of extract_text()
4. Store as: {note_num: {title, type: "table", rows: [(label, value), ...]}}
```

#### Type B: Narrative / Qualitative Notes
Examples: General (Umum), Accounting Policies, Commitments & Contingencies

```
1. Extract text from note pages
2. Build a summary table: Topic | Detail | Reference
3. Store as: {note_num: {title, type: "narrative", summary: [...]}}
```

#### Type C: Mixed Notes
Some notes have both tables and narrative (e.g. Related Party Transactions)

```
1. Extract tables first (Type A approach)
2. Extract remaining narrative (Type B approach)
3. Store both
```

### Phase 5 — Excel Workbook Construction

Build the workbook using openpyxl. Follow these formatting rules strictly:

#### Workbook-Level Formatting

| Element | Format |
|---|---|
| Font | Arial 10pt throughout |
| Column A width | 55 characters (bilingual labels are long) |
| Data column width | 20 characters |
| Year header | Bold, centered |
| Currency header | `IDR – full amount` (or as stated in report) |
| Subtotals | Bold, thin top border |
| Grand totals | Bold, double top border |
| Section headers | Merged across columns, bold, light gray fill (#F2F2F2) |
| Empty separator rows | Between sections |
| Negative numbers | Format: `#,##0_);[Red](#,##0)` |
| Footnote rows | Italic, 9pt, indented |
| Year values | Format as text string (not number) to avoid "2,025" |

#### Sheet Construction Order

```
Sheet 1  : Balance Sheet
Sheet 2  : Income Statement
Sheet 3  : Cash Flow Statement
Sheet 4  : Changes in Equity
Sheet 5  : Note 1 - [Title]
Sheet 6  : Note 2 - [Title]
...       (one sheet per note found in notes_index)
Sheet N  : Segment Information (if segment note exists)
Sheet N+1: Key Ratios Summary
```

**Sheet naming:** `Note [N] - [Short English Title]`
- Max 31 characters (Excel limit)
- If title is too long, abbreviate sensibly

#### Sheet 1–4: Main Financial Statements

For each statement, extract EVERY line item as it appears in the PDF.
Do not aggregate, omit, or rename sub-line items.

Structure for Balance Sheet:
```
Row 1: Title row — "Laporan Posisi Keuangan / Balance Sheet"
Row 2: Currency/unit note — e.g. "Dalam Rupiah penuh / In full IDR"
Row 3: Column headers — A: blank, B: [YEAR]
Row 4: Empty separator
Row 5+: Data rows following the structure in the PDF
```

Structure for Income Statement, Cash Flow, Equity: follow same pattern,
extracting every line item exactly as it appears in the PDF with bilingual labels.

#### Sheet 5+: Notes Sheets

For **tabular notes** (Type A):
```
Row 1: Note title (bold, merged) — "Catatan [N] - [ID Title] / Note [N] - [EN Title]"
Row 2: Empty
Row 3: Column headers — A: "Keterangan / Description", B: [YEAR]
Row 4+: Data rows from parsed tables
```

If a note contains multiple distinct tables, separate them with:
- One blank row
- A bold italic label: `*Table A: [Table Name / Nama Tabel]*`

For **narrative notes** (Type B):
```
Row 1: Note title (bold, merged)
Row 2: Empty
Row 3: Summary table headers — A: "Topik / Topic", B: "Detail", C: "Referensi / Reference"
Row 4+: Summary rows
```

For **mixed notes** (Type C):
- Narrative summary section first (Type B format)
- Blank separator row
- Then tabular data (Type A format)

#### Key Ratios Summary Sheet

Calculate these ratios from the extracted single-year data:

**Liquidity:**
- Current Ratio = Current Assets / Current Liabilities
- Quick Ratio = (Current Assets − Inventory) / Current Liabilities
- Cash Ratio = Cash & Equivalents / Current Liabilities

**Leverage:**
- Debt-to-Equity = Total Liabilities / Total Equity
- Interest Coverage = EBIT / Finance Costs

**Profitability:**
- Gross Margin = Gross Profit / Revenue
- Operating Margin = Operating Profit / Revenue
- Net Margin = Net Profit / Revenue
- ROE = Net Profit / Total Equity (single-year, no average)
- ROA = Net Profit / Total Assets (single-year, no average)

**Efficiency:**
- Asset Turnover = Revenue / Total Assets
- Receivables Days = Trade Receivables / Revenue × 365
- Payables Days = Trade Payables / COGS × 365

**Cash Flow:**
- Free Cash Flow = Operating CF − Capex
- FCF Margin = FCF / Revenue

Format:
- Percentages as `0.0%`
- Multiples as `0.0x`
- Days as `0` (integer)
- Mark `N/A` for any ratio that cannot be computed

Use Excel formulas referencing cells in the main statement sheets where possible.
For ratios that depend on values only available as hardcoded (e.g. trade receivables
extracted from notes), use hardcoded values and add a comment noting the source.

### Phase 6 — Validation & Save

```
1. Run ALL validation checks (see below). For each failure, add a footnote row
   in the relevant sheet noting the discrepancy. Do NOT silently adjust numbers.

   BALANCE SHEET:
   a. Total Assets = Total Liabilities + Total Equity
   b. Total Current Assets + Total Non-Current Assets = Total Assets
   c. Total Current Liabilities + Total Non-Current Liabilities = Total Liabilities

   INCOME STATEMENT:
   d. Revenue − COGS = Gross Profit
   e. Gross Profit − Operating Expenses = Operating Profit (if directly extractable)
   f. Profit Before Tax − Tax Expense = Net Profit (approximate — deferred tax may cause small difference)

   CASH FLOW:
   g. Operating CF + Investing CF + Financing CF + FX effect (if any) = Net Change in Cash
   h. Cash Beginning + Net Change = Cash End

   CROSS-STATEMENT:
   i. Net Profit in IS ≈ Net Profit in Equity Statement (may differ by OCI)
   j. Cash End in CF ≈ Cash & Equivalents in BS

2. Save workbook to output path
3. Run recalc: python /mnt/skills/public/xlsx/scripts/recalc.py [output_file]
4. Check for formula errors; fix any found
5. Print terminal summary (see below)
```

### Anti-Hallucination Rules

Apply these throughout the entire extraction process:

```
1. NEVER fill a cell with a calculated or estimated value without an explicit footnote
   stating the calculation method. If a number is not directly readable from the PDF → N/A.

2. FALLBACK CALCULATIONS are allowed ONLY for:
   - Total Liabilities = Total Assets − Total Equity (when direct match fails)
   Every fallback MUST have an italic footnote row:
   "Computed as [formula]. Direct extraction failed because [reason]."

3. COMPARATIVE YEAR values from the PDF: do NOT extract these in Run 1.
   Run 1 is single-year only. If the PDF shows a comparative prior-year column,
   ignore it — that year will be added properly via idx-excel-append-year.md.

4. CONSOLIDATED vs STANDALONE: check the page header for "KONSOLIDASIAN" / "CONSOLIDATED".
   If a page does NOT say consolidated, skip it — it may be a standalone parent statement.
   Log which type was used in the terminal output.

5. OCR values: if any page required OCR (text extraction returned empty/garbled),
   mark EVERY value from that page with a cell comment: "[OCR – verify against source]".

6. ZERO vs N/A: a genuine zero (the report explicitly shows 0 or dash for a line item)
   is different from N/A (the line item was not found). Use 0 for the former, N/A for the latter.
   Never treat a parsing failure as zero.

7. FINANCIAL HIGHLIGHTS ≠ AUDITED FS: the financial highlights section (Ikhtisar Keuangan,
   typically pages 5–15) often shows summary BS/IS numbers. These may be rounded, condensed,
   or restated differently. NEVER extract from the highlights section. Extract ONLY from pages
   identified AFTER the auditor's report in Phase 1.

8. AUDITOR'S REPORT ABSENT: if no auditor's report is found in the PDF, add a prominent
   footnote on the Cover/first sheet: "WARNING: Auditor's report not found in source PDF.
   Financial statement data may be unaudited. Verify against official filings."
```

---

## Terminal Output

```
=== IDX Excel Template — Run 1 Complete ===
Company     : [TICKER] – [COMPANY NAME]
Year        : [FISCAL YEAR]
Source File : [PDF PATH] ([TOTAL PAGES] pages)
Output File : [XLSX PATH]

Source Location in PDF:
  Auditor's Report : page [N]
  Audit Status     : [AUDITED / UNVERIFIED — auditor's report not found]
  Statement Type   : [CONSOLIDATED / STANDALONE]
  Balance Sheet    : pages [N]–[N]
  Income Statement : pages [N]–[N]
  Cash Flow        : pages [N]–[N]
  Equity           : pages [N]–[N]
  Notes            : pages [N]–[N]

Sheets created:
  1. Balance Sheet
  2. Income Statement
  3. Cash Flow Statement
  4. Changes in Equity
  [5–N. Note sheets listed]
  N+1. Key Ratios Summary

Validation:
  BS: Total Assets = Liabilities + Equity    : [PASS / FAIL — details]
  BS: Current + Non-Current = Total          : [PASS / FAIL]
  IS: Revenue − COGS = Gross Profit          : [PASS / FAIL]
  CF: Operating + Investing + Financing = ΔCash : [PASS / FAIL]
  CF: Beginning + ΔCash = Ending             : [PASS / FAIL]
  Cross: IS Net Profit ≈ Equity Net Profit   : [PASS / FAIL]
  Cross: CF Cash End ≈ BS Cash               : [PASS / FAIL]
  Formula Errors                             : [count or NONE]

Fallback Calculations Used:
  [LIST any values computed via fallback, e.g. "Total Liabilities = Assets − Equity"]

Notes Index Found:
  [List of note numbers and titles]

Warnings / Missing Data:
  [N/A values, OCR flags, parsing issues, pages skipped]
=============================================
```

---

## Design for Future Runs

This template workbook is designed so that Run 2+ can:
1. Load the existing workbook
2. Insert a new column (to the LEFT of the existing year column, since older years go left)
3. Populate with the next year's data using the same extraction logic
4. Handle note number changes by matching on topic/title rather than note number
5. Add footnote rows for any reconciliation differences across years

The single-year template establishes:
- The complete set of sheets (no new sheets needed unless a new note appears)
- The row structure and labels in Column A
- The formatting standards

Future runs only need to add data columns — not rebuild the structure.

---

## Debugging Tips

1. **Always probe first** — run Phase 1 structure probe before writing any extraction code
2. **Print parsed rows** — dump all `parse_financial_lines()` output before building matchers
3. **BS identity gate** — validate Total Assets = Total Liabilities + Total Equity before accepting data
4. **Notes detection** — check for `(lanjutan)/(continued)` false positives
5. **Fixed Assets** — use `extract_tables()` not `extract_text()` for movement tables
6. **Fragment matching** — PDF often splits labels across lines; match the fragment with the number, not the full label
7. **Space artifacts** — large numbers may have spaces mid-number from PDF rendering
