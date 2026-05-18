---
name: idx-excel-run1-template-rev
description: >
  Extract audited financial statements from an IDX-listed company's Annual
  Report (Laporan Tahunan) into a comprehensive Excel workbook. Annual Reports
  contain audited FS in the back portion, typically starting 40–60% into the
  PDF. This skill locates the audited FS section, extracts all statements and
  notes, and builds a single-year template Excel that future runs can append
  to via idx-excel-append-year-rev.md. Do NOT use for standalone Financial
  Reports (Laporan Keuangan Q4) — those are often unaudited.
---

# IDX Excel Extractor — Run 1: Single-Year Template (revised)

Build a complete Excel workbook from the **audited financial statements**
found inside ONE Annual Report (Laporan Tahunan) PDF — the most recent
fiscal year. This workbook becomes the template; future runs append earlier
or later years as new columns.

This document is a revision of `idx-excel-run1-template.md`, informed by
problems the FY2024 append run uncovered in the FY2025 Run 1 output. The
changes are flagged inline with **[FIXED]** or **[NEW]**.

**Important — source file distinction:**

| Document type | Indonesian | Contents | Use this skill? |
|---|---|---|---|
| **Annual Report** | Laporan Tahunan | Narrative (100+ pages) + audited FS in the back | ✅ YES |
| Financial Report Q4 | Laporan Keuangan Q4 | FS + notes only, often unaudited | ❌ NO — quarterly extractor |
| Financial Report Q1–Q3 | Laporan Keuangan Interim | FS + limited notes, unaudited | ❌ NO — quarterly extractor |

---

## Role

Expert financial analyst, certified accountant, IDX/BEI capital markets,
PSAK / IFRS. Reads and interprets Annual Reports of Indonesian listed
companies.

---

## Required Input

| Parameter | Example |
|---|---|
| PDF file path | `reports/WIFI_Annual_Report_2025.pdf` |
| Ticker / company name | `WIFI – PT Solusi Sinergi Digital Tbk` |
| Fiscal year | `2025` |
| Output filename | `WIFI_Financial_Statements.xlsx` |

If any parameter is missing, ask before proceeding. If the filename
suggests a standalone Financial Report (e.g. `Laporan_Keuangan_Q4`),
warn the user and offer the quarterly extractor.

---

## Annual Report PDF structure

Audited FS are in the BACK portion of the report, not the front.

```
Typical layout (100–250+ pages):

  pp.   1–  5  : Cover, Table of Contents
  pp.   5– 15  : Financial Highlights / Ikhtisar Keuangan
  pp.  15– 30  : Chairman & Directors' Reports / Sambutan
  pp.  30– 60  : Company Profile, Business Overview
  pp.  60– 80  : Management Discussion & Analysis
  pp.  80–100  : Corporate Governance
  pp. 100–120  : CSR / Sustainability
  pp. 120–130  : Independent Auditor's Report      ← anchor
  pp. 130–170  : Audited Financial Statements      ← extract this
    ├── Balance Sheet (Laporan Posisi Keuangan)
    ├── Income Statement (Laporan Laba Rugi)
    ├── Changes in Equity (Laporan Perubahan Ekuitas)
    ├── Cash Flow Statement (Laporan Arus Kas)
    └── Notes to Financial Statements (Catatan atas Laporan Keuangan)
  pp. 170+     : Appendices, proxy
```

Page numbers are approximate; the auditor's report is the only reliable
anchor. Locate it first.

### How to locate the audited FS section

1. **Auditor's report** — `LAPORAN AUDITOR INDEPENDEN` /
   `INDEPENDENT AUDITOR'S REPORT` / `LAPORAN AKUNTAN INDEPENDEN`.
   **[FIXED]** The keyword can match three pages: the FS cover page that
   references it, the Table of Contents that lists it, and the actual
   report. To distinguish: the **real** auditor's report page also
   contains the audit firm name and the business-license number
   (`KEP-…/KM.1/…`) within the first ~150 characters of extracted text.
   Require those markers in addition to the keyword.

   **[NEW]** If the auditor's report pages are scanned images (no text
   layer), the keyword scan won't find them. Fall back to the FS cover
   page's Table of Contents, which always lists the auditor's report
   with a page number. Pull that page number directly. Cross-check by
   verifying that the listed page begins with the audit firm letterhead
   when OCR is applied.

2. **First financial statement title** — after the auditor's report:
   `LAPORAN POSISI KEUANGAN KONSOLIDASIAN`. Require both the title and
   `KONSOLIDASIAN` AND `(Disajikan dalam Rupiah` on the same page.

3. **Standalone vs Consolidated** — the financial highlights section
   (pp.5–15) often has a condensed BS/IS that matches the same
   keywords. The `(Disajikan dalam Rupiah` discriminator catches this
   because the highlights tables don't carry it.

---

## General rules

1. **Bilingual labels** — Column A: `Indonesian Term / English Term`.
2. **Data integrity** — never fabricate.
3. **Consolidated priority** — always use consolidated if both exist.
4. **Indonesian number format** — dot = thousands, parens = negative.
5. **Single-year layout** — Column A = labels, Column B = the fiscal
   year. Add `NEW_ROW_GUARD` empty rows between sections so future
   appends have room to insert without renumbering.
6. **Scanned PDFs** — flag values with cell comment `[OCR – verify]`.
7. **Audited FS only** — pages AFTER the auditor's report.
8. **Empty cell, not "N/A" string** — missing data is an empty cell + a
   cell comment. `0` only when the report explicitly shows zero or a
   dash. **[FIXED]** This replaces the original "use N/A or —"
   guidance, which broke Excel formulas downstream.
9. **[NEW] Source pages recorded per row** — every populated cell
   carries a cell comment with the source PDF page when the value is
   not visually obvious (e.g., sub-totals computed across multi-page
   tables). Future readers need traceability.
10. **[NEW] Number normalisation pass** — every page of extracted text
    runs through `normalize_numbers()` (see Implementation in
    `idx-excel-append-year-rev.md`) before tokenisation. PDFs in this
    family insert spaces inside numbers; without the pre-pass, rows
    silently drop.

---

## Execution Sequence

### Phase 0 — Environment

```bash
pip install cffi pdfplumber openpyxl --break-system-packages
```

**[FIXED]** Install `cffi` first and explicitly. On many systems
`cffi` is present as a system package but the pip-installed
`cryptography` (a pdfminer dependency) requires a newer version that
needs recompilation, and `pdfplumber` fails to import with a confusing
`_cffi_backend` error.

### Phase 1 — Structure probe

Before writing any extraction code, probe the PDF to locate the audited
FS section.

```
1. Open PDF with pdfplumber, count pages.
2. For every page, extract text; store the first 200 chars.

3. STEP A — Find the auditor's report:
   a. Scan all pages for the keyword set.
   b. For each candidate, require:
      • Audit firm name pattern (e.g., uppercase "& REKAN", "& PARTNERS",
        "& CO") within the first 200 chars OR
      • Business license pattern `KEP-\d+/KM\.\d+/\d+` within the first
        200 chars.
   c. Reject pages whose top contains "Daftar Isi", "Tata Kelola",
      "Manajemen", "MD&A" — these are TOC / governance references.
   d. The real auditor's report is the first candidate that passes (b)
      and (c).
   e. [NEW] If no candidate passes, check the FS cover page (typically
      one page before the TOC) for a TOC entry "Laporan auditor
      independen" with a page number. Use that page number directly,
      and flag the run as "FS cover page used to locate auditor's
      report — auditor's report itself is scanned/non-text".

4. STEP B — Find financial statements AFTER the auditor's report:
   • BS  → LAPORAN POSISI KEUANGAN + KONSOLIDASIAN + (Disajikan dalam Rupiah
   • IS  → LAPORAN LABA RUGI + KONSOLIDASIAN
   • CF  → LAPORAN ARUS KAS
   • EQ  → LAPORAN PERUBAHAN EKUITAS
   Only consider pages with index > auditor_report_page.

5. STEP C — Determine page ranges. Each statement ends where the next
   begins.

6. STEP D — Find notes section start. Walk pages after the last main
   statement looking for `^(\d{1,2})\.\s+([A-Z]{3,}…)`, skipping
   lines containing "(lanjutan)" / "(continued)". Record the lowest
   matching page.

7. Print the structure map to terminal.
```

**Fallback if no auditor's report and no FS cover TOC entry can be
identified:** require statement pages to be past the 55% mark of total
pages (not 50% — narrow margin, especially for shorter ARs). Flag every
extracted value as `[unverified]`.

#### Guarding against financial-highlights false matches

The Financial Highlights section (Ikhtisar Keuangan, typically pp.5–15)
often contains "Laporan Posisi Keuangan" as a table header and
condensed BS/IS numbers. The auditor's report anchor blocks them.
**[FIXED]** Adding `(Disajikan dalam Rupiah` as a co-required marker
makes the guard work even when the auditor's report cannot be located
with confidence.

### Phase 2 — Main Statement Extraction

For each statement (BS, IS, CF, Equity):

1. Extract raw text from the identified pages.
2. Run `normalize_numbers()` on the text. **[NEW]**
3. Run `parse_financial_lines()` on the normalised text.
4. Print ALL parsed rows (label + values) to the terminal.
5. Build matchers for the target line items.
6. Validate: `Total Assets = Total Liabilities + Total Equity`. If it
   fails, debug before proceeding.

#### Number parsing

```python
NUM_TOKEN = re.compile(
    r'\(\s*\d{1,3}(?:\.\d{3})+(?:\s*\.\d{3})*(?:\s*)?\)'
    r'|'
    r'\d{1,3}(?:\.\d{3})+(?:\s*\.\d{3})*'
)

def parse_id_number(s):
    s = s.strip()
    if not s or s in ('-', '—'): return None
    neg = s.startswith('(') and s.endswith(')')
    if neg: s = s[1:-1]
    s = s.replace(' ', '').replace('.', '')
    try: v = int(s)
    except ValueError: return None
    return -v if neg else v
```

#### Known pitfalls

| Problem | Mitigation |
|---|---|
| "Total Aset" matches "Total Aset Lancar" | Exclusion list: `excl: ['lancar', 'tidak']` |
| Early reports use "Kas dan bank" not "Kas dan setara kas" | Include both patterns |
| Trade receivables header has no numbers; data under "Pihak ketiga" | Match `pihak ketiga` too |
| `LABA SEBELUM PAJAK PENGHASILAN` split across lines | Fragment match `penghasilan` with exclusions |
| Statement title split across PDF lines | Shortened search strings + secondary validation |
| **[FIXED]** Space artifacts mid-number: `97.27 8.291.602` | Run `normalize_numbers()` BEFORE regex extraction (the old `(?:\s*\.\d{3})*` band-aid covered only some cases) |
| **[NEW]** First letter of a label detaches: `P eriklanan`, `T otal` | Filter tokens of length ≤ 2 when computing keyword overlap; do not try to un-fragment |
| Total Liabilities sometimes wrong from direct match | Fallback: `TL = TA − TE` with italic footnote |
| **[NEW]** Equity pages can be mirrored | If the page's header text reads backwards, reverse every numeric token's characters; cross-validate closing balance against Total Equity |

### Phase 3 — Notes index building

Scan pages from `notes_min_page` to end of PDF:

1. For each page, extract text and apply `normalize_numbers()`.
2. Look for note headers: `^\s*(\d{1,2})\.\s+([A-Z][A-Z\s/\(\)\-]{3,})`.
3. Skip lines containing `(lanjutan)` or `(continued)`.
4. Record `note_index[N] = {title_id, title_en, start_page}`.
5. Compute `end_page[N] = start_page[N+1] − 1`. **[FIXED]** If two
   adjacent notes share a start page (e.g., Notes 10 and 11 both start
   on p.252), clamp `end_page` to `max(start_page, computed_end)` —
   never `start − 1`, which produces an invalid range. The page-range
   filter is liberal anyway because the section-header filter (see
   below) does the real boundary work.
6. **[NEW]** Print the index sorted by note number; for each note,
   include the title text and the page range. A human-readable index
   makes Phase 4 debugging straightforward.

False-positive guards:

- Require title text after the number to be ≥ 3 uppercase characters.
- Require page ≥ `notes_min_page`.

### Phase 4 — Notes content extraction

**[FIXED]** The original skill said "extract every line item from each
note as it appears in the PDF". That advice produced workbooks where
the same content appeared in multiple adjacent note sheets, because the
underlying note headers appeared on the same page and the extractor
copied every parsed row into every sheet whose page range overlapped.
The fix is the **section-header filter**: tracking which note section
each line belongs to as you walk the page.

#### Section-header-aware extraction

```python
NOTE_HEADER = re.compile(r'^\s*(\d{1,2})\.\s+[A-Z][A-Z]')

def lines_for_note(pdf, page_range, target_note):
    out = []
    for pn in range(page_range[0], page_range[1] + 1):
        txt = normalize_numbers(pdf.pages[pn-1].extract_text() or '')
        in_target = False
        for line in txt.split('\n'):
            stripped = line.strip()
            m = NOTE_HEADER.match(stripped)
            if m and 'lanjutan' not in stripped.lower() and 'continued' not in stripped.lower():
                in_target = (int(m.group(1)) == target_note)
                continue
            if in_target:
                out.append((pn, line))
    return out
```

Always pass `target_note=N` when reading lines for Note N. The Note N
sheet then contains ONLY lines that fall inside the `N. TITLE … (N+1). TITLE`
window on each page.

#### Type A — Quantitative / Tabular notes

Examples: Cash, Trade Receivables, Inventories, Fixed Assets, Debt.

1. Get lines via `lines_for_note(pdf, page_range[N], N)`.
2. Run `parse_financial_lines()` on those lines.
3. **[FIXED]** For Fixed-Assets and Intangible-Assets rollforward
   tables, the LAST numeric column on each line is the closing balance
   (Saldo Akhir = the year we want). For two-column tables (`2024
   2023`), the FIRST numeric column is the target year. Encode the
   layout per-note in metadata (`ROLLFORWARD_NOTES`).
4. Store `{N: {title, type: 'table', rows: [(label, value), ...]}}`.

#### Type B — Narrative / Qualitative notes

Examples: General (Umum), Accounting Policies, Commitments.

1. Lines via `lines_for_note`.
2. Build a topic / detail / reference summary.

#### Type C — Mixed notes

Tables first (Type A), then narrative (Type B).

#### [NEW] Per-sheet topic-range metadata

For every note sheet created, write a hidden cell or workbook-level
metadata key `TOPIC_ROW_RANGE = (low_row, high_row)` so future append
runs know which rows on this sheet truly belong to the sheet's topic.
The append skill reads this and refuses to write outside the range.

A simple way to persist this is a hidden sheet `_TopicRanges` with two
columns (`Sheet`, `Range`) that the append skill loads on startup.

### Phase 5 — Excel workbook construction

Build with openpyxl. Strict formatting rules:

| Element | Format |
|---|---|
| Font | Arial 10pt |
| Column A width | 55 |
| Data column width | 20 |
| Year header | Bold, centred |
| Currency header | `IDR – full amount` (or as stated) |
| Subtotals | Bold, thin top border |
| Grand totals | Bold, double top border |
| Section headers | Merged, bold, light gray fill (#F2F2F2) |
| Empty separator rows | Between sections |
| Negative numbers | `#,##0_);[Red](#,##0)` |
| Footnote rows | Italic, 9pt, indented |
| Year values | Text string (not number) — avoids "2,025" |

#### Sheet construction order

```
Sheet 1   : Balance Sheet
Sheet 2   : Income Statement
Sheet 3   : Cash Flow Statement
Sheet 4   : Changes in Equity
Sheet 5   : Note Index           ← [NEW] reserve this slot in Run 1
Sheet 6+  : Note 1, Note 2, …
Final     : Key Ratios Summary
Hidden    : _TopicRanges         ← [NEW]
```

**Sheet naming:** `Note <N> - <Short English Title>` (max 31 chars).

**Note Index belongs in Sheet 5.** Even though Run 1 is single-year,
this sheet is reserved here so future appends only add a column to it —
they never insert it. Full construction steps are in the dedicated
section below.

#### [NEW] Sheet 5 — Note Index (reserved by Run 1)

The Note Index is a single permanent sheet that holds the cross-year
reference of note titles. Run 1 **must** create it even though only one
year of data exists yet — every future append simply adds one column to
it instead of restructuring it. Failing to reserve this sheet in Run 1
forces append runs to insert it later, which shifts every subsequent
sheet's index and breaks any downstream tooling that references sheets
positionally.

**Position.** Immediately AFTER `Changes in Equity` and BEFORE
`Note 1`. In openpyxl: create the sheet at the end with
`wb.create_sheet('Note Index')`, then move it into position with
`wb.move_sheet('Note Index', offset=...)` so its index equals
`wb.sheetnames.index('Changes in Equity') + 1`.

**Layout.**

| Row | Col A | Col B | Notes |
|---|---|---|---|
| 1 | `Note Title Index — Cross-Year Reference` | (merged into A1:B1) | Bold, size 12, centred |
| 2 | `Note Number` | `<YEAR>` | Bold, light-gray fill `#DDDDDD` |
| 3 | `Note 1` | English title of Note 1 in this year | English title in col B; Indonesian title attached as cell comment |
| 4 | `Note 2` | English title of Note 2 in this year | … |
| …  | one row per note sheet | … | … |

Append-year runs will insert a new column at position B (for an older
year added to the left) or to the right (for a newer year added to the
right), re-merging row 1 across the widened range and re-bolding row 2.
The Run 1 layout MUST be consistent with that contract: row 1 merged,
row 2 headers, data starting row 3, one row per note sheet in workbook
order.

**Construction steps.**

```python
from openpyxl.comments import Comment
from openpyxl.styles import Font, Alignment, PatternFill

# 1. Build the per-note title list in workbook order.
#    NOTE_MATCH below is the same ordered list the note sheets were
#    created from in Phase 4/5. Each entry: (sheet_name, title_id_in_year,
#    title_en_in_year). For Run 1, every sheet has data; for later
#    append runs, a sheet may be missing from this year — record None.
NOTE_MATCH = [
    ('Note 1 - GENERAL',                 'UMUM',                            'GENERAL'),
    ('Note 2 - SIGNIFICANT ACCOUNTING',  'INFORMASI KEBIJAKAN AKUNTANSI…',  'MATERIAL ACCOUNTING POLICY INFORMATION'),
    # … one entry per note sheet, in the order the sheets sit in the workbook
]

# 2. Create and position the sheet.
if 'Note Index' in wb.sheetnames:
    del wb['Note Index']
idx = wb.create_sheet('Note Index')
ce_idx = wb.sheetnames.index('Changes in Equity')
wb.move_sheet(idx, offset=-(len(wb.sheetnames) - 1 - ce_idx - 1))

# 3. Row 1 — title (merged across A1:B1).
idx['A1'] = 'Note Title Index — Cross-Year Reference'
idx['A1'].font = Font(bold=True, size=12)
idx['A1'].alignment = Alignment(horizontal='center')
idx.merge_cells('A1:B1')

# 4. Row 2 — headers.
idx['A2'] = 'Note Number'
idx['B2'] = str(YEAR)                                # e.g. '2025'
for c in ('A2', 'B2'):
    idx[c].font = Font(bold=True)
    idx[c].fill = PatternFill('solid', fgColor='DDDDDD')

# 5. Rows 3+ — one row per note sheet.
for i, (sheet_name, title_id, title_en) in enumerate(NOTE_MATCH):
    r = 3 + i
    idx.cell(row=r, column=1, value=sheet_name.split(' - ')[0])  # "Note 1"
    cell = idx.cell(row=r, column=2, value=title_en)
    cell.comment = Comment(f'Indonesian title: {title_id}', 'Run1')

# 6. Column widths + freeze panes.
idx.column_dimensions['A'].width = 14
idx.column_dimensions['B'].width = 55
idx.freeze_panes = 'A3'
```

**Why English in column B even in Run 1.** Putting Indonesian here in
Run 1 forces a subsequent append run to either flip the column language
(creating an asymmetric `Indonesian for Year-N | English for Year-N+1`
table — visually inconsistent and confusing for readers) or do a
breaking schema migration to swap them. English in every year column
with Indonesian in cell comments is the contract that future runs will
extend; setting it now means no later cleanup.

**Validation.** After construction, verify:

- Row 1 is merged across A1:B1 with the exact title text.
- Row 2 headers are present and bold.
- Number of rows from row 3 to the last populated row equals the number
  of note sheets in the workbook (one-to-one).
- Every row 3+ has both column A (`Note N`) and column B (English title)
  populated, and column B has a non-empty cell comment containing the
  Indonesian title.
- Sheet position: `wb.sheetnames.index('Note Index') ==
  wb.sheetnames.index('Changes in Equity') + 1`.

#### Sheet 6+ : Note sheets

**[FIXED]** Each note sheet contains ONLY the rows extracted for THAT
note via the section-header filter. Never copy parsed rows into
multiple sheets even if their PDF page range overlaps. Verify by
spot-checking: pick two sheets that started on the same PDF page (e.g.,
Notes 10 and 11 if they share p.252) and confirm row labels differ.

**[FIXED]** Drop tokens whose normalised length ≤ 2 from Column-A
labels before writing them. Single-letter prefixes like `P` in
`P eriklanan` or `T` in `T otal` carry no information and only make
labels harder to read AND harder to match in future appends. The clean
label is what matters.

For **tabular notes** (Type A):

```
Row 1: Note title (bold, merged)
Row 2: empty
Row 3: A: "Keterangan / Description", B: <YEAR>
Row 4+: data rows
```

If a note contains multiple distinct tables, separate with one blank
row and an italic header `*Table A: <table name>*`.

For **narrative notes** (Type B):

```
Row 1: Note title (bold, merged)
Row 2: empty
Row 3: A: "Topik / Topic", B: "Detail", C: "Referensi / Reference"
Row 4+: summary rows
```

#### Key Ratios Summary

Compute the standard ratios from the extracted single-year data.

**[FIXED]** Year header on this sheet is NOT necessarily on row 3 —
detect the row dynamically by locating the existing year label. Future
appends will rely on this; hardcoding row 3 here misaligns the
appended column. Use a consistent row (row 4 is recommended for this
sheet because it leaves room for the title block).

Liquidity, leverage, profitability, efficiency, cash-flow ratios as in
the original skill. Format percentages `0.0%`, multiples `0.0x`, days
`0`. Use formulas referencing main statement sheets where possible.

### Phase 6 — Validation & save

1. Validation suite (BS identity, sub-totals, IS reconciliation, CF
   reconciliation, IS↔Equity, CF↔BS cash). For each failure, add a
   footnote row in the relevant sheet. Do NOT silently adjust.
2. **[NEW] Per-sheet topic-range verification.** For each note sheet,
   confirm all rows fall within a sensible range. If a sheet has rows
   that look like they belong to the previous or next note's topic,
   investigate before saving — it usually means the section-header
   filter failed.
3. **[NEW] Per-sheet uniqueness check.** For each pair of adjacent note
   sheets, compute the set of non-empty Column-B values; the overlap
   should be small or zero. A high overlap (e.g., the same totals on
   both sheets) means the section filter is not working.
4. Save workbook.
5. Run recalc.
6. Print terminal summary.

---

## Anti-hallucination rules

1. Never fill a cell with a calculated value without an italic footnote.
2. Fallback calculations allowed only for `TL = TA − TE`.
3. Comparative-year values from the PDF: **do not extract them in Run
   1.** Run 1 is single-year. The comparative year will be added later
   via the append skill from its own AR (lesson C2 in `lesson-rev.md`
   — the comparative column of a later year's AR is unreliable as a
   source).
4. Consolidated vs Standalone: check `KONSOLIDASIAN` on every page; if
   absent, verify the company has no subsidiaries.
5. OCR values: cell-comment every one.
6. Zero vs empty: `0` only for explicit zero/dash; empty + comment for
   parsing failure.
7. Financial Highlights ≠ Audited FS.
8. If no auditor's report can be found at all, add a prominent
   footnote on the first sheet warning the user.

---

## Terminal output

```
=== IDX Excel Template — Run 1 Complete ===
Company     : <TICKER> – <COMPANY>
Year        : <YEAR>
Source File : <PDF> (<N> pages)
Output File : <XLSX>

Source location in PDF:
  Auditor's Report : page <N>  (firm: <NAME>, license <KEP-...>)
  Audit Status     : <AUDITED | UNVERIFIED (no auditor's report found)>
  Statement Type   : <CONSOLIDATED | STANDALONE>
  BS / IS / CF / EQ: <page ranges>
  Notes            : pages <N>–<N>

Sheets created:
  1. Balance Sheet
  2. Income Statement
  3. Cash Flow Statement
  4. Changes in Equity
  5. Note Index
  6–<M>. Note 1, Note 2, …
  <M+1>. Key Ratios Summary
  Hidden: _TopicRanges

Validation:
  BS Assets = Liab + Equity         : <PASS | FAIL>
  BS Current + Non-Current = Total  : <PASS | FAIL>
  IS Revenue − COGS = Gross Profit  : <PASS | FAIL>
  CF Op + Inv + Fin = ΔCash         : <PASS | FAIL>
  CF Begin + ΔCash = End            : <PASS | FAIL>
  Cross IS NP ≈ Eq NP               : <PASS | FAIL>
  Cross CF End ≈ BS Cash            : <PASS | FAIL>
  Per-sheet topic uniqueness        : <PASS | FAIL>  ← [NEW]

Fallback calculations used:
  <list, if any>

Notes index found:
  <table of note numbers, titles, page ranges>

Warnings / missing data:
  <N/A values, OCR flags, parsing issues, pages skipped>
=============================================
```

---

## Design for future runs

This template is designed so that an append run can:

1. Load the existing workbook.
2. Read the `_TopicRanges` hidden sheet to know which rows on each note
   sheet are in-topic.
3. Insert a new column at the correct chronological position.
4. Populate using the two-layer extractor (hardcoded then systematic).
5. Match notes by topic title rather than number.
6. Add footnote rows for any reconciliation differences and append
   year-only rows under an italic "Baris berikut hanya muncul di
   Laporan Tahunan <YEAR>" footnote.

The single-year template establishes:

- The complete set of sheets (no new sheets needed unless a new note
  appears).
- The row structure and labels in Column A — **clean**, with
  single-letter fragments dropped.
- Per-sheet topic ranges so future runs can guard against cross-note
  contamination.
- The Note Index sheet structure with English titles in every year
  column.
- The formatting standards.

Future runs only add data columns and append year-only rows — they do
not rebuild the structure.

---

## Debugging tips

1. **Always probe first** — Phase 1 structure probe before any extraction.
2. **Print parsed rows** — dump every `parse_financial_lines()` output
   before building matchers.
3. **BS identity gate** — validate `TA = TL + TE` before accepting data.
4. **Notes detection** — `(lanjutan) / (continued)` false-positive guard.
5. **Section-header filter** — verify each note sheet has lines from
   only one note's section.
6. **Per-sheet uniqueness** — compute non-empty value overlap between
   adjacent sheets; high overlap means a section-filter bug.
7. **Fixed-Assets table** — `extract_tables()` can help when
   `extract_text()` is ambiguous; cross-validate row totals.
8. **Number normalisation** — every page through `normalize_numbers()`
   before tokenisation.
9. **Equity mirroring** — if header reads backwards, reverse the tokens.
10. **Persist topic ranges** — write them to `_TopicRanges` even for
    sheets that look clean; future appends rely on the metadata.

---

## Cross-reference

The fixes in this revision come directly from issues uncovered while
running the FY2024 append against the FY2025 Run 1 output:

- **[FIXED]** Auditor's report disambiguation (lesson C-9, C-1 in
  `lesson-rev.md`).
- **[FIXED]** Financial-highlights guard via `(Disajikan dalam Rupiah`
  co-marker.
- **[FIXED]** Negative-number / N/A handling — empty cell + comment
  rather than `N/A` string.
- **[FIXED]** Multi-note section filter — replaces "extract every
  parsed line into each note's sheet" (the root cause of the FY2025
  Notes 31/32 and 33–36 content cloning).
- **[FIXED]** Single-letter fragment cleanup in labels.
- **[FIXED]** Rollforward last-column vs two-column first-column
  extraction rule.
- **[FIXED]** Key Ratios year-header row detection.
- **[NEW]** Number normalisation pre-pass.
- **[NEW]** Note Index sheet reserved at Run 1 with English titles.
- **[NEW]** `_TopicRanges` hidden metadata sheet for future appends.
- **[NEW]** Audit-firm letterhead requirement for auditor's report
  identification.
- **[NEW]** FS-cover TOC fallback for scanned auditor pages.
- **[NEW]** Per-sheet topic uniqueness check in validation.
