---
name: idx-excel-append-year
description: >
  Append ONE fiscal year's data (older OR newer) to an existing IDX Financial Statements Excel workbook.
  Source PDF must be an Annual Report (Laporan Tahunan) containing audited financial statements.
  Locates the audited FS section (after the auditor's report), extracts all statements and notes,
  and inserts new columns at the correct chronological position (oldest left, newest right).
  Handles note numbering shifts across years by matching on topic/title, not note number.
  Do NOT use for standalone Financial Reports (Laporan Keuangan Q4) — use quarterly extractor instead.
---

# IDX Excel Extractor — Append Year

Add one fiscal year (older or newer) to an existing Excel workbook.
Source PDF must be an **Annual Report (Laporan Tahunan)** with audited FS in the back section.

The existing workbook defines the sheet structure and row labels.
This run adds a data column and, where necessary, new rows for line items not in the template.

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
| Existing Excel file | `WIFI_Financial_Statements.xlsx` |
| New PDF file | `report/WIFI_Annual_Report_2024.pdf` |
| Ticker / company name | `WIFI – PT Solusi Sinergi Digital Tbk` |
| Fiscal year being added | `2024` |

If any parameter is missing, ask the user before proceeding.

If the PDF filename suggests a standalone Financial Report (e.g. `Laporan_Keuangan_Q4`,
`Financial_Report_2024Q4`), warn the user:
"This appears to be a standalone Financial Report, not an Annual Report.
Standalone Q4 reports are often unaudited. Consider using idx-excel-quarterly.md instead.
Proceed anyway?"

---

## General Rules

1. **Bilingual labels** — all labels use `Indonesian Term / English Term` format
2. **Data integrity** — NEVER fabricate or interpolate numbers; if not found in source, leave cell empty with a cell comment: `"Not found in [YEAR] report"`
3. **Consolidated priority** — always use consolidated statements if both exist
4. **Indonesian number format** — dot (`.`) = thousands separator, parentheses = negative
5. **Column insertion** — maintain chronological order: oldest year on LEFT, newest on RIGHT. If the new year is older than all existing years → insert to the LEFT. If the new year is newer → insert to the RIGHT. If it falls between existing years → insert at the correct position.
6. **Existing rows are immutable** — do NOT rename, reorder, or delete any existing row in Column A. Existing data in other year columns must NEVER be overwritten or removed.
7. **New rows are additive** — if the new PDF contains a line item not present in the existing workbook, INSERT a new row at the logical position within the correct section. Leave existing year columns empty for this row (with cell comment: `"Item not in [YEAR] report"`). This is not a contradiction with rule 6: existing rows are untouched, new rows are appended.
8. **Audited FS only** — extract ONLY from the audited FS section (after the auditor's report in the Annual Report PDF), never from financial highlights or summary tables
9. **Backup before mutation** — before inserting any columns or rows, save a backup copy of the existing workbook as `[FILENAME]_backup_before_[YEAR].xlsx`
10. **Empty cell, not N/A string** — when data is unavailable, leave the cell empty and add a cell comment explaining why. NEVER write the string "N/A" into a data cell — it breaks Excel formulas (`SUM`, `AVERAGE`, `IF`). Use empty cells for missing data and `0` only when the report explicitly shows zero.

---

## Missing Data Handling (replaces N/A throughout)

| Situation | Cell value | Cell comment |
|---|---|---|
| Line item exists in workbook but not found in new PDF | Empty | `"Not found in [YEAR] report"` |
| Line item exists in new PDF but not in workbook | Value in new column, empty in existing columns | `"Item not in [TEMPLATE_YEAR] report"` |
| Note sheet exists but note not found in new year | Entire column empty | `"Note not present in [YEAR] report"` |
| Parsing failed but line item likely exists | Empty | `"Extraction failed — verify against source p.[PAGE]"` |
| Report explicitly shows zero or dash | `0` | No comment needed |

---

## Execution Sequence

### Phase 0 — Read Existing Workbook & Backup

```
1. BACKUP: Copy the existing Excel to [FILENAME]_backup_before_[YEAR].xlsx
   → This is the safety net. If anything goes wrong, the user can revert.

2. Load the existing Excel with openpyxl (NOT data_only — preserve formulas)

3. For each sheet, record:
   a. Sheet name
   b. All labels in Column A (the row structure)
   c. Which columns contain year data and what years they are
   d. The rightmost data column letter/index
   e. Any merged cells (record their ranges — column insertion can break these)
   f. Any conditional formatting rules (record their ranges)

4. Determine where to insert the new column:
   → Parse all existing year headers into integers
   → Sort them to know the chronological range
   → If new_year < min(existing_years): insert BEFORE the leftmost year column
   → If new_year > max(existing_years): insert AFTER the rightmost year column
   → If new_year is between existing years: insert at the correct chronological position
   → If new_year already exists: STOP — warn user "Year [YEAR] already exists in workbook"
   → Record: insert_position (column index) and direction ("left" / "right" / "between")

5. Print the existing structure to terminal for verification:
   - Sheets: [list]
   - Years present: [list, left to right]
   - Insert position: column [LETTER], [direction]
   - Merged cell ranges found: [count]
```

### Phase 1 — Structure Probe (New PDF)

Locate the audited FS section inside the Annual Report PDF.
Same approach as Run 1: find the auditor's report first, then FS after it.

```
1. Open PDF with pdfplumber, count total pages (expect 80–300+)

2. For EVERY page, extract text and store: pages[page_num] = first 200 chars

3. STEP A — Find the Auditor's Report:
   Scan ALL pages for:
   - "LAPORAN AUDITOR INDEPENDEN" (case-insensitive)
   - "INDEPENDENT AUDITOR" (case-insensitive)
   - "LAPORAN AKUNTAN INDEPENDEN" (case-insensitive)
   Record: auditor_report_page = first match
   → If NOT found: warn user — FS may be unaudited. Flag all output as [unverified].

4. STEP B — Find financial statements AFTER the auditor's report:
   Scan pages from (auditor_report_page + 1) onward for:
   - Balance Sheet    → "LAPORAN POSISI KEUANGAN" containing "KONSOLIDASIAN"
   - Income Statement → "LAPORAN LABA RUGI" containing "KONSOLIDASIAN"
   - Cash Flow        → "LAPORAN ARUS KAS"
   - Equity           → "LAPORAN PERUBAHAN EKUITAS"

   CRITICAL: Do NOT match occurrences BEFORE the auditor's report.
   Those are financial highlights / summary tables, not audited FS.

5. STEP C — Determine page ranges (each statement ends where the next begins)
   Record: fs_pages = {auditor: page, bs: [s,e], is: [s,e], cf: [s,e], eq: [s,e]}

6. STEP D — Find notes section start:
   Scan from (last_statement_end + 1) for lines matching:
     ^(\d{1,2})\.\s+([A-Z]{3,}...)
   Skip lines with "(lanjutan)" or "(continued)".
   Record: notes_min_page

7. Print the full structure map to terminal
```

**Fallback if no auditor's report found:** require statement pages to be past the 50%
mark of total pages — the audited FS section is always in the back half.

### Phase 2 — Main Statement Extraction

Extract BS, IS, CF, Equity from the new PDF using the same parsing logic as Run 1.

```
1. Extract raw text from identified pages (ONLY from pages after auditor's report)
2. Run parse_financial_lines()
3. Print all parsed rows
4. Match against the EXISTING row labels in the workbook (see Row Matching below)
5. Run validation checks (see Phase 6)
```

#### Row Matching Strategy

The existing Excel has row labels in Column A. For each label:
```
1. Normalize: lowercase, strip spaces, remove punctuation
2. Try exact match against parsed lines from the new PDF
3. If no exact match, try fuzzy match:
   - "Beban Pokok Penjualan" vs "Harga Pokok Penjualan" → same concept (COGS)
   - "Kas dan Bank" vs "Kas dan Setara Kas" → same concept (Cash)
4. If still no match → leave cell empty, add cell comment: "Not found in [YEAR] report"
```

**New line items in the new PDF** (not in the existing workbook):
```
→ INSERT a new row at the logical position within the correct section
→ Add the bilingual label in Column A
→ Put the value in the new year column
→ Leave existing year column(s) empty with comment: "Item not in [TEMPLATE_YEAR] report"
→ Add an italic footnote row below: "Row added from [YEAR] report — not present in earlier years"
```

This is consistent with General Rules 6 & 7: existing rows stay untouched, new rows are additive.

### Phase 3 — Note Index Sheet & Note Matching

This is the most critical phase. Note numbers shift between years.
A permanent **Note Index sheet** tracks all note titles across every appended year.

#### Step 3a: Build the new year's note index

```
1. Scan pages from notes_min_page onward
2. Find note headers: ^(\d{1,2})\.\s+([A-Z]{3,}...)
3. Skip "(lanjutan)" / "(continued)" lines
4. Record: new_notes[note_num] = {title_id, title_en, start_page, end_page}
```

#### Step 3b: Match new notes to existing sheets

**DO NOT match by note number. Match by TOPIC.**

```python
# Pseudocode for note matching
for new_num, new_info in new_notes.items():
    matched_sheet = None
    
    # Strategy 1: Exact title match (case-insensitive)
    for sheet_name in existing_sheets:
        if is_note_sheet(sheet_name):
            existing_title = extract_title_from_sheet(sheet_name)
            if normalize(new_info.title_id) == normalize(existing_title):
                matched_sheet = sheet_name
                break
            if normalize(new_info.title_en) == normalize(existing_title):
                matched_sheet = sheet_name
                break
    
    # Strategy 2: Keyword overlap match
    if not matched_sheet:
        for sheet_name in existing_sheets:
            if is_note_sheet(sheet_name):
                similarity = keyword_overlap(new_info.title, sheet_name)
                if similarity > 0.6:
                    matched_sheet = sheet_name
                    best_score = similarity
    
    # Strategy 3: No match → this is a new note not in the template
    if not matched_sheet:
        create_new_sheet(new_num, new_info)
```

#### Known Note Number Shifts (WIFI example — illustrative only)

| Topic | 2020 Note # | 2021 Note # | 2022 Note # | 2023 Note # | 2024 Note # | 2025 Note # |
|-------|-------------|-------------|-------------|-------------|-------------|-------------|
| UMUM / GENERAL | 1 | 1 | 1 | ? | ? | ? |
| KAS / CASH | 4 | 4 | 4 | 4 | 4 | 4 |
| PIUTANG USAHA | 5 | 5 | 5 | 5 | 5 | 5 |
| PERSEDIAAN | — | — | 6 | ? | ? | 6 |
| PIUTANG LAIN-LAIN | 6 | 6 | — | ? | ? | — |
| ASET TETAP | 10 | 10 | — | — | — | 10 |
| ASET TAKBERWUJUD | — | — | 10 | — | — | — |

This table is illustrative — always match by title content, never assume fixed mappings.

#### Step 3c: Add renumbering footnotes

When a note matches an existing sheet but has a DIFFERENT note number:

```
Add an italic footnote row at the bottom of the sheet:
"Catatan: Pada tahun [YEAR], item ini bernomor Catatan [N] /
 Note: In [YEAR], this item was numbered Note [N]"
```

#### Step 3d: Maintain the Note Index sheet

The **Note Index** sheet is a permanent cross-year reference. It lives between
'Changes in Equity' and 'Note 1'. Each row represents one note sheet in the workbook.
Each year adds **one** column showing the note title used in that year.

**Layout:**

| Column | Content |
|---|---|
| A | Note Number (e.g. `Note 1`, `Note 6`, `Note 25`) |
| B+ | One column per fiscal year (header = year, value = note title in that year) |

Year columns follow the same chronological convention as data sheets:
**oldest year on the LEFT, newest year on the RIGHT.** When appending an older year,
insert the column at the LEFT (before existing year columns). When appending a
newer year, insert at the RIGHT.

**On Run 1 (when the workbook is first created):**

```
1. Create sheet named "Note Index"
2. Position it immediately after 'Changes in Equity' and before 'Note 1'
3. Row 1 (merged A1:last_col): "Note Title Index — Cross-Year Reference"
4. Row 2: Column headers
   A: "Note Number"
   B: "[YEAR]"   (e.g. "2025")
5. Rows 3+: one row per note sheet in workbook order
   - A: note number label (e.g. "Note 1")
   - B: note title for that year (use English title from the PDF, or the workbook
        sheet's English title if it matches the PDF)
6. Bold row 2 headers; freeze top 2 rows
```

**On each append run:**

```
1. Locate the existing Note Index sheet
2. Determine insert position based on chronological order of the new year vs existing columns:
   → If new_year < min(existing_years): INSERT new column at position B (left of all years)
   → If new_year > max(existing_years): APPEND new column to the right
   → If new_year is between: INSERT at correct chronological position
3. Set the header cell to the new year (e.g. "2024")
4. For each existing row (one per note sheet):
   a. Look up whether that sheet was matched to a note in the new year (from Step 3b)
   b. If matched: write the English title from the new year PDF
   c. If NOT matched: leave the cell empty
      Add cell comment: "Note not present in [NEW_YEAR] report"
5. For new notes in the new year that have NO matching sheet in the template:
   a. Create a new note sheet (handled in Step 3b)
   b. Add a new row at the bottom of the index for that sheet
   c. Fill the new year's column with the title
   d. Leave all existing year columns empty for that row
6. Extend the Row 1 merged range to cover all year columns
```

**Index sheet layout example (after appending 2024 to a 2025 workbook):**

```
| Note Number | 2024              | 2025                |
|-------------|-------------------|---------------------|
| Note 1      | GENERAL           | GENERAL             |
| Note 4      | CASH              | CASH                |
| Note 6      |                   | INVENTORIES         |
| Note 7      | PREPAID EXPENSES  | PREPAID EXPENSES    |
| Note 25     |                   | SHARIA BONDS        |
```

Empty cells mean the note is absent in that year. The "Note Number" in column A
uses the workbook's note sheet numbering (which is fixed once the first run
creates the sheets). Cross-year note-number shifts are captured by the
per-sheet renumbering footnote (Step 3c), not by the index.

**Index sheet layout example (after later appending 2026):**

```
| Note Number | 2024              | 2025                | 2026                |
|-------------|-------------------|---------------------|---------------------|
| Note 1      | GENERAL           | GENERAL             | GENERAL             |
| Note 6      |                   | INVENTORIES         | INVENTORIES         |
| Note 7      | PREPAID EXPENSES  | PREPAID EXPENSES    | INVENTORIES         |
```

If a note number ends up with a different title in a later year (e.g. the
report restructured its notes), the index makes the divergence visible at a glance.

### Phase 4 — Notes Content Extraction & Insertion

**Every matched note sheet must have its sub-line values populated for the new year.**
Only leaving the year column header filled (with no sub-line data) is NOT acceptable —
it defeats the purpose of a multi-year workbook.

For each matched note (identified in Phase 3b):

```
1. Identify the page range for this note in the new PDF (start_page to end_page)
2. Extract full text from those pages using pdfplumber
3. Parse into (label, value) pairs — see Parsing Strategy below
4. Match each parsed row to existing Column A labels in the sheet — see Row Matching
5. Insert values into the new year column for matched rows
6. Insert new rows for PDF lines with no match in the sheet — see New Row Insertion
7. Leave empty + comment for sheet rows with no match in PDF
8. Add the year header in the correct header row (typically row 3, same row as other years)
```

#### Phase 4 Parsing Strategy

Indonesian financial note tables typically have one of two layouts:

**Layout A — Two-column table (label | value):**
```
Biaya dibayar dimuka - pihak ketiga        5.000.000.000
Biaya sewa                                 3.500.000.000
Biaya asuransi                             1.500.000.000
Jumlah / Total                             5.000.000.000
```
→ Extract each label and its rightmost numeric value on the same line.

**Layout B — Multi-column table (label | year_1 | year_2):**
```
                             2024              2023
Biaya sewa                 3.500.000.000    3.100.000.000
Biaya asuransi             1.500.000.000    1.200.000.000
```
→ Identify which column corresponds to the year being extracted.
→ Extract ONLY the column for the target year. Ignore comparative year columns.

**Parsing pseudocode:**
```python
def parse_note_table(pages_text, target_year):
    rows = []
    for line in pages_text.split('\n'):
        # Skip decorative lines, empty lines, page headers
        if is_empty_or_header(line):
            continue
        nums = extract_all_numbers(line)  # parse_id_number() on each token
        label = extract_label(line)       # everything before first numeric token
        if not nums or not label:
            continue
        if len(nums) == 1:
            # Layout A: single value
            rows.append((label.strip(), nums[0]))
        elif len(nums) >= 2:
            # Layout B: multi-column — use target_year column index
            col_idx = resolve_year_column(pages_text, target_year)
            rows.append((label.strip(), nums[col_idx]))
    return rows
```

**IMPORTANT:** If the note is purely narrative (no table), leave the note sheet column
populated only with the year header and add a sheet-level comment:
`"Note [N] ([YEAR]): Narrative only — no tabular data to extract."`

#### Phase 4 Row Matching

For each existing row in Column A of the note sheet:

```python
def match_label(sheet_label, pdf_rows):
    """
    Returns (pdf_row_index, match_type) or None if not matched.
    match_type: "exact" | "fuzzy" | "keyword"
    """
    norm_sheet = normalize(sheet_label)  # lowercase, strip, remove punctuation
    
    # Pass 1: Exact match (after normalization)
    for i, (pdf_label, val) in enumerate(pdf_rows):
        if normalize(pdf_label) == norm_sheet:
            return (i, "exact")
    
    # Pass 2: One label is a substring of the other (handles truncation)
    for i, (pdf_label, val) in enumerate(pdf_rows):
        norm_pdf = normalize(pdf_label)
        if norm_sheet in norm_pdf or norm_pdf in norm_sheet:
            if len(min(norm_sheet, norm_pdf, key=len)) >= 8:  # avoid short spurious matches
                return (i, "fuzzy")
    
    # Pass 3: Keyword overlap (at least 60% of significant words match)
    for i, (pdf_label, val) in enumerate(pdf_rows):
        if keyword_overlap(norm_sheet, normalize(pdf_label)) >= 0.6:
            return (i, "keyword")
    
    return None  # no match found

def normalize(s):
    s = s.lower().strip()
    s = re.sub(r'[^\w\s]', '', s)  # remove punctuation
    s = re.sub(r'\s+', ' ', s)
    return s

def keyword_overlap(a, b):
    STOP = {'dan', 'atau', 'yang', 'untuk', 'ke', 'dari', 'dengan', 'dalam',
            'dan', 'the', 'and', 'or', 'of', 'to', 'from', 'for', 'in', 'on'}
    words_a = set(a.split()) - STOP
    words_b = set(b.split()) - STOP
    if not words_a or not words_b:
        return 0
    return len(words_a & words_b) / max(len(words_a), len(words_b))
```

**When a match is found:**
- Insert the PDF value into the new year's column for that row
- If match_type is "fuzzy" or "keyword": add cell comment with original PDF label:
  `"Matched from: '[PDF_LABEL]' in [YEAR] report"`
- Mark the PDF row as used (consumed), so it is not matched again

**When NO match is found for a sheet row:**
- Leave the new year cell empty
- Add cell comment: `"Not found in [YEAR] report — verify against PDF p.[PAGE_RANGE]"`

#### New Row Insertion (Phase 4)

For each PDF row that was NOT consumed by Phase 4 Row Matching:
→ This is a line item that exists in the new year but has NO matching row in the sheet.
→ INSERT a new row at the logical position.

```
Insertion rules:
1. Find the section it belongs to (totals / sub-items / headers)
2. Place new sub-items ABOVE their section total row
3. Place new header/section rows at the section boundary
4. Preserve vertical ordering from the PDF where possible

New row format:
   Column A: bilingual label "[Indonesian] / [English]"
              If only Indonesian is available, append " / [translate or leave blank]"
   New year column: the extracted value
   All existing year column(s): leave empty
   Comment on each empty existing-year cell: "Item not in [TEMPLATE_YEAR] report"

Add an italic footnote row immediately below the new row:
   "Row added from [NEW_YEAR] report — not present in [TEMPLATE_YEAR] data"
```

#### Table Reconciliation Reference

| Situation | Action |
|---|---|
| Same label, exact match | Insert value in new year column |
| Label wording slightly different | Match by meaning; comment with original wording |
| Sub-row in new PDF but not in sheet | Insert new row; leave existing year columns empty + comment |
| Sub-row in sheet but not in new PDF | Leave new year column empty + comment |
| Note table structure completely different | Create a separate sub-table below with year label |
| Note in new PDF but NO matching sheet | Create new sheet; leave template year column empty + comment |
| Sheet exists but note NOT found in new PDF | Leave entire new year column empty + sheet-level comment |
| Note is purely narrative (no table) | Year header only + sheet-level narrative comment |

### Phase 5 — Key Ratios Update

```
1. Go to the Key Ratios Summary sheet
2. Insert a new column at the correct chronological position (same logic as Phase 0)
3. Calculate all ratios for the new year using extracted values
4. Use Excel formulas referencing cells in the main statement sheets where possible
5. For ratios needing values from notes (e.g. trade receivables), use hardcoded values
```

For ratios requiring averages (ROE, ROA, Asset Turnover):
- Now that two years exist, use `(Year N + Year N-1) / 2` for averages
- Update the template year's ratio formulas too if they were single-year approximations
- Add a footnote: "Averages computed using two available years"

### Phase 6 — Post-Insertion Integrity Check, Validation & Save

```
STEP 1 — POST-INSERTION INTEGRITY CHECK (run immediately after all column/row insertions):

   a. Verify no #REF! errors in any cell across all sheets
   b. Verify merged cell ranges are still intact (compare to Phase 0 recorded ranges)
   c. Verify existing year columns still contain their original data
      (spot-check 5 key values: Total Assets, Revenue, Net Profit, Operating CF, Total Equity
       against the values recorded in Phase 0)
   d. Verify column headers still show correct years in correct order
   e. If any check fails → restore from backup, report the failure, STOP

STEP 2 — FINANCIAL VALIDATION (for the newly added year):

   BALANCE SHEET:
   a. Total Assets = Total Liabilities + Total Equity
   b. Total Current Assets + Total Non-Current Assets = Total Assets
   c. Total Current Liabilities + Total Non-Current Liabilities = Total Liabilities

   INCOME STATEMENT:
   d. Revenue − COGS = Gross Profit
   e. Gross Profit − Operating Expenses = Operating Profit (if extractable)

   CASH FLOW:
   f. Operating CF + Investing CF + Financing CF + FX effect = Net Change in Cash
   g. Cash Beginning + Net Change = Cash End

   CROSS-STATEMENT:
   h. Net Profit in IS ≈ Net Profit in Equity Statement
   i. Cash End in CF ≈ Cash & Equivalents in BS

   For each failure → add a footnote row in the relevant sheet. Do NOT silently adjust.

STEP 3 — SAVE & RECALC:
   a. Save workbook
   b. Run recalc: python /mnt/skills/public/xlsx/scripts/recalc.py [output_file]
   c. Check for formula errors; fix any found
   d. Print terminal summary
```

### Anti-Hallucination Rules

```
1. NEVER fill a cell with a calculated or estimated value without an explicit footnote.
   If a number is not directly readable from the PDF → leave cell empty + comment.

2. FALLBACK CALCULATIONS allowed ONLY for:
   - Total Liabilities = Total Assets − Total Equity (when direct match fails)
   Every fallback MUST have an italic footnote row:
   "Computed as [formula]. Direct extraction failed because [reason]."

3. COMPARATIVE YEAR values in the PDF: the Annual Report's audited FS shows
   the current year AND a comparative prior year. Extract ONLY the year being
   appended. If the comparative year is already in the workbook → ignore it.
   If the comparative year is NOT in the workbook, still ignore it — it will
   be added in its own run from its own Annual Report.

4. CONSOLIDATED vs STANDALONE: verify the statement title says "KONSOLIDASIAN".
   If it doesn't, check if the company has subsidiaries. Log the finding.

5. OCR values: mark with cell comment "[OCR – verify against source p.[PAGE]]".

6. ZERO vs EMPTY: a genuine zero (report shows 0 or dash) → put 0.
   Parsing failure or item not found → leave empty + comment. Never confuse the two.

7. NEVER extract from Financial Highlights (Ikhtisar Keuangan) near the front of
   the Annual Report. Only from the audited FS section after the auditor's report.

8. WORKBOOK DATA IS NOT A SOURCE: when matching rows, use the existing workbook
   labels for matching only. Never copy values from existing columns into the new
   column. Every value in the new column must come from the new PDF.
```

---

## Terminal Output

```
=== IDX Excel — Year Append Complete ===
Company      : [TICKER] – [COMPANY NAME]
Year Added   : [FISCAL YEAR]
Source PDF   : [PDF PATH] ([TOTAL PAGES] pages)
Output File  : [XLSX PATH]
Backup File  : [BACKUP PATH]
Years in File: [LIST ALL YEARS NOW IN THE WORKBOOK, left to right]

Source Location in PDF:
  Auditor's Report : page [N]
  Audit Status     : [AUDITED / UNVERIFIED]
  Statement Type   : [CONSOLIDATED / STANDALONE]
  BS: pages [N]–[N]  |  IS: pages [N]–[N]  |  CF: pages [N]–[N]  |  EQ: pages [N]–[N]

Column Inserted: [COLUMN LETTER] ([left of / right of / between] existing data)

Post-Insertion Integrity:
  #REF! errors       : [NONE / count]
  Merged cells intact : [YES / NO — details]
  Existing data check : [PASS / FAIL — spot-check details]

Financial Validation (new year):
  BS: Assets = Liabilities + Equity        : [PASS / FAIL]
  IS: Revenue − COGS = Gross Profit        : [PASS / FAIL]
  CF: Operating + Investing + Financing = Δ : [PASS / FAIL]
  Cross: IS Net Profit ≈ Equity Net Profit : [PASS / FAIL]
  Cross: CF Cash End ≈ BS Cash             : [PASS / FAIL]

Note Index Sheet:
  Status       : [CREATED / UPDATED]
  Column added : "[NEW_YEAR]" inserted [LEFT OF / RIGHT OF / BETWEEN] existing year columns
  Total rows   : [N] note sheets tracked
  Titles filled: [M] / [N] (cells empty where note is absent in [NEW_YEAR])

Note Matching Results:
  [NEW_NOTE_NUM]. [TITLE] → matched to sheet "[SHEET_NAME]"
  [NEW_NOTE_NUM]. [TITLE] → matched to sheet "[SHEET_NAME]" (renumbered from Note [X])
  [NEW_NOTE_NUM]. [TITLE] → NEW SHEET CREATED (not in template)
  Sheet "[SHEET_NAME]" → note NOT FOUND in [YEAR] (column left empty)

Note Sub-line Extraction Results (per matched note):
  Note [N] "[SHEET_NAME]": [M] rows matched, [K] rows inserted (new in [YEAR]), [J] rows empty (not in [YEAR])
  Note [N] "[SHEET_NAME]": narrative only — no tabular data
  Note [N] "[SHEET_NAME]": [M] rows matched (all exact)

Row Changes:
  [N] new rows inserted across all note sheets
  [N] cells left empty (data not available in this year)
  [N] cells matched via fuzzy/keyword (see per-cell comments)

Fallback Calculations Used:
  [LIST any values computed via fallback]

Warnings:
  [Label mismatches, OCR flags, reconciliation notes]
================================================
```

---

## Number Parsing (same as Run 1)

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

---

## Debugging Tips

1. **Print existing structure first** — before extracting anything, understand what's already in the Excel
2. **Match by topic, never by note number** — this is the #1 source of errors across years
3. **Watch for shifted columns** — after inserting a column, ALL column references shift right by 1
4. **Test formulas after insertion** — run recalc and check for #REF! errors immediately
5. **Compare parsed rows side by side** — print both the existing labels and the new PDF's parsed labels to see mismatches
6. **Fragment matching** — PDF splits labels across lines; the fragment with the number may not match the full label in the existing Excel
7. **Don't force a match** — if a note genuinely doesn't exist in the older year, empty + comment is the correct answer
8. **Backup is your safety net** — if post-insertion integrity fails, restore from backup rather than trying to fix a corrupted workbook
9. **Never use existing workbook values as source** — every number in the new column must come from the new PDF
10. **Find the auditor's report first** — don't start extracting until you've confirmed which pages are the audited FS
11. **Note sub-lines are required** — a note column with only the year header and empty sub-lines is incomplete; every matched note must have its tabular rows populated
12. **Multi-column note tables are common** — the note table often shows two years side-by-side; always identify which column is the target year before extracting values
13. **Note Index sheet must be updated on every run** — seeded on Run 1 with the first year's note data; extended on each append run with new year column group
14. **Consumed tracking for row matching** — mark each PDF row as used after it matches a sheet row; unmatched PDF rows become new inserted rows; this prevents double-counting
