---
name: idx-excel-append-year-rev
description: >
  Append ONE fiscal year's data (older OR newer) to an existing IDX Financial
  Statements Excel workbook produced by idx-excel-run1-template-rev.md or its
  predecessor. Source PDF must be an Annual Report (Laporan Tahunan) with
  audited FS in the back section. The skill matches notes by topic title (not
  number), populates sub-line tables row-by-row with a two-layer extractor
  (verified-hardcoded first, systematic PDF-driven second), and appends
  year-only rows as new bottom-of-sheet entries. Do NOT use for standalone
  Financial Reports (Laporan Keuangan Q4) — use the quarterly extractor.
---

# IDX Excel Extractor — Append Year (revised)

This skill adds one fiscal year (older or newer) to an existing workbook.
The previous run's workbook defines the sheet structure and Column-A row
labels; this skill adds a new data column and, where the PDF has line items
not in the template, new rows under a footnote separator.

It supersedes `idx-excel-append-year.md`. The major changes from that
version are:

- A **two-layer extractor**: hardcoded verified totals first, systematic
  PDF parsing second.
- A **`TOPIC_ROW_RANGES`** whitelist that guards against Run 1 having
  cloned content into the wrong note sheets.
- A **section-header filter** so multi-note pages don't bleed across
  note boundaries.
- A **`normalize_numbers()`** pre-pass for PDF whitespace artifacts inside
  numbers.
- A **`_GENERIC_LABELS`** skip-list that forces ambiguous "Total Total"
  rows to be hardcoded rather than auto-matched.
- An **asymmetric keyword-overlap** denominator so short labels
  (`Tanah / Land`) match correctly.
- A **magnitude sanity check** during matching, and **value-based dedup**
  during append.
- A Note Index sheet whose year columns are **all English** with the
  original Indonesian title kept in cell comments.

---

## Role

Expert financial analyst, certified accountant, IDX/BEI capital markets,
PSAK / IFRS. Reads and interprets Annual Reports of Indonesian listed
companies and reconciles them across years.

---

## Required Input

| Parameter | Example |
|---|---|
| Existing Excel file | `WIFI_Financial_Statements.xlsx` |
| New PDF file | `reports/WIFI_Annual_Report_2024.pdf` |
| Ticker / company name | `WIFI – PT Solusi Sinergi Digital Tbk` |
| Fiscal year being added | `2024` |

If any parameter is missing, ask the user before proceeding. If the PDF
filename suggests a standalone Q4 Financial Report, warn the user and
offer the quarterly extractor instead.

---

## General Rules

1. **Bilingual labels** — Column A: `Indonesian Term / English Term`.
2. **Data integrity** — never fabricate. If not found in source, leave the
   cell empty and add a cell comment.
3. **Consolidated priority** — always use consolidated statements if
   both exist.
4. **Indonesian number format** — dot = thousands, parens = negative.
5. **Column insertion** — chronological order, oldest left, newest right.
   If the new year is older than all existing years, insert at the LEFT.
   If newer, insert at the RIGHT. If it falls between, insert at the
   correct position. If the year already exists, STOP and ask.
6. **Existing rows are immutable** — do NOT rename, reorder, or delete
   any existing row in Column A. Data in other year columns is never
   overwritten.
7. **New rows are additive** — items in the new PDF without a workbook
   row become NEW rows appended at the bottom of the relevant sheet,
   under an italic footnote separator. Existing year columns remain
   empty for those rows.
8. **Audited FS only** — extract only from pages AFTER the auditor's
   report; never from financial highlights.
9. **Backup before mutation** — save
   `<FILENAME>_backup_before_<YEAR>.xlsx` before writing.
10. **Empty cell, not "N/A" string** — missing data is an empty cell + a
    cell comment. `0` only when the report explicitly shows zero or a
    dash.

---

## Missing-Data Handling

| Situation | Cell value | Cell comment |
|---|---|---|
| Workbook row exists, not found in new PDF | Empty | `"Not found in [YEAR] report"` |
| PDF row exists, not in workbook | Value in new row; existing-year cells empty | On existing-year cell: `"Item not in [TEMPLATE_YEAR] report"` |
| Whole note absent in new year | Year column empty | Sheet-level comment on year-header cell |
| Match was fuzzy/substring/keyword | Value populated | `"Matched from PDF p.<N>: '<PDF label>' (<match type>)"` |
| Sign was flipped from PDF presentation | Value with workbook sign | Hardcoded comment explaining the flip |
| Report explicitly shows zero or dash | `0` | No comment needed |

---

## Execution Sequence

### Phase 0 — Setup and integrity check

1. `pip install cffi pdfplumber openpyxl --break-system-packages` (cffi
   first — see lesson L11 below).
2. Print the existing workbook's state: sheet names, year headers in
   each main sheet's header row, the first three populated values per
   main sheet. Compare to the README's "Years in File" claim. If they
   disagree, STOP and surface to the user. (This is the documentation-
   drift guard.)
3. Determine the insert column based on the new year's chronological
   position. Record `insert_position` and direction (`left` / `right` /
   `between`).
4. Copy the workbook to `<FILENAME>_backup_before_<YEAR>.xlsx`.

### Phase 1 — Structure probe on the new PDF

Find the audited FS section AFTER the auditor's report.

1. Open the new PDF with pdfplumber, count pages.
2. Locate the auditor's report by scanning for
   `LAPORAN AUDITOR INDEPENDEN` / `INDEPENDENT AUDITOR` /
   `LAPORAN AKUNTAN INDEPENDEN`. Require the matching page to also
   contain the **audit firm letterhead** (firm name, business license
   number "KEP-…"). This eliminates Table-of-Contents and MD&A
   references that match the same keyword but are not the report itself.
3. Find the main statements *after* the auditor's report page:
   - Balance Sheet: `LAPORAN POSISI KEUANGAN` + `KONSOLIDASIAN` on the
     same page.
   - Income Statement: `LAPORAN LABA RUGI` + `KONSOLIDASIAN`.
   - Cash Flow: `LAPORAN ARUS KAS`.
   - Equity: `LAPORAN PERUBAHAN EKUITAS`.
   - Each must include `(Disajikan dalam Rupiah` somewhere on the same
     page (rules out financial-highlights summary tables).
4. Determine page ranges for BS, IS, CF, Equity.
5. Walk all pages after the equity statement and detect note headers:
   `^\s*(\d{1,2})\.\s+[A-Z]{3,}` (excluding `lanjutan` / `continued`).
   Record `notes[N] = (start_page, end_page)` where `end_page` =
   `next note's start − 1`, clamped to `start_page` if two notes share a
   start page.
6. **If the auditor's report cannot be confidently identified** (e.g.,
   the actual report pages are scanned images with no text layer):
   fall back to requiring statement pages to be past the 55% mark of
   total pages, and flag every extracted value as `[unverified]`.

### Phase 2 — Main statement extraction

Apply `normalize_numbers()` to every page of extracted text before
parsing (see Implementation below).

For each of BS, IS, CF, Equity:

1. Extract raw text from the identified pages only.
2. Parse into `(label, value)` candidate rows.
3. Match each existing Column-A label to a parsed row via the matching
   pipeline (Phase 3 below).
4. Insert values into the new year's column.
5. Append rows for PDF lines with no workbook match, under an italic
   footnote (`Catatan: Baris berikut hanya muncul di Laporan Tahunan
   <YEAR>`).

**Equity statement requires special care** (lesson C2/C3). Before
trusting the extracted opening balance, cross-check against the prior
year's AR if available. If the equity pages appear mirrored (header text
reads backwards), reverse each numeric token's character order.

### Phase 3 — Note matching by topic

#### Step 3a — Build the new year's note index

Scan from `notes_min_page` onward; build
`new_notes[n] = {title_id, title_en, start_page, end_page}`.

#### Step 3b — Match new notes to existing sheets

For each `(n, info)`:

1. **Exact title match** (case-insensitive normalisation) on either
   `title_id` or `title_en` against existing note-sheet titles.
2. **Substring match** if exact fails (require ≥ 8 character common
   substring).
3. **Keyword overlap** with asymmetric denominator (see Implementation):
   if both sides ≤ 2 significant keywords, score is `intersect /
   min(|a|, |b|)`; otherwise `intersect / max(|a|, |b|)`. Threshold 0.6.
4. **No match** → record as a new note for which a new sheet must be
   created.

Match by topic, not number. Note numbers shift ±1 to ±6 between adjacent
fiscal years (lesson C4).

#### Step 3c — Update the Note Index sheet

Layout: column A = workbook note number (`Note 1`, `Note 2`, ...); one
column per year. **Every year column uses English titles**; the
Indonesian title for that year is attached as a cell comment of the
form: `Indonesian title: <ID title>` (plus, if the year's note number
differs from the workbook number, `In <YEAR> this topic was Note <N>.`).

Position: between `Changes in Equity` and `Note 1`. Title row 1
(merged); header row 2; data rows 3+ (frozen at A3).

### Phase 4 — Note content extraction (two-layer)

Every renumbered note sheet must have its 2024 column populated with
sub-line values where the underlying PDF table contains them.

#### Layer 1 — Hardcoded verified totals

Maintain a `NOTE_<YEAR>_DATA` dict keyed by `(sheet_name, row_num)` of
values that have been manually verified against the source PDF AND
cross-reference to BS / IS / CF. Apply this layer FIRST — it protects
against the matcher's mistakes on ambiguous labels and is the place to
encode sign-convention overrides.

Hardcode at minimum, for every applicable note:
- Each total / sub-total row (these usually have generic Column-A labels
  like `Total Total` that auto-matching cannot resolve).
- Each row whose PDF sign disagrees with the workbook column's IS
  convention (current tax, deferred tax, finance costs, certain CF
  items — see lesson L9 and C5).
- The cross-reference rows that anchor BS/IS validation.

Comment every hardcoded value with the source page and the reason if
non-obvious.

#### Layer 2 — Systematic PDF-driven fill

For each renumbered note whose 2024 page range is known:

```
for sheet_name, n_year, _id, _en in NOTE_MATCH:
    if n_year is None or sheet_name not in wb: continue
    if n_year in NARRATIVE_ONLY_NOTES: continue
    start, end = NOTE_PAGE_RANGES[n_year]
    pdf_rows = extract_note_rows(pdf, start, end,
                                 rollforward=n_year in ROLLFORWARD_NOTES,
                                 target_note=n_year)
    s = wb[sheet_name]
    rows_to_check = topic_rows_for(sheet_name, s.max_row)
    for r in rows_to_check:
        a = s.cell(row=r, column=1).value
        c = s.cell(row=r, column=3).value          # 2025 anchor
        b = s.cell(row=r, column=2)
        if not isinstance(c, (int, float)) or a is None: continue
        if b.value is not None: continue           # hardcoded already filled
        m = match_workbook_label(str(a), pdf_rows)
        if m is None: continue
        idx, mtype = m
        v = pdf_rows[idx]['value']
        if magnitude_violates(v, c): continue      # see L6
        pdf_rows[idx]['consumed'] = True
        b.value = v
        if mtype != 'exact':
            b.comment = Comment(f"Matched from PDF p.{pdf_rows[idx]['page']}: "
                                f"'{pdf_rows[idx]['label_id']}' ({mtype})",
                                'Append<YEAR>')
    append_unmatched_as_year_only_rows(s, pdf_rows, n_year)
```

Key knobs:

- **`NARRATIVE_ONLY_NOTES`**: Note numbers that are pure narrative
  (General, Significant Accounting Policies, Estimates, Financial
  Instruments, Risk Management, Issuance of Amendments). Skip
  auto-extraction for these; the hardcoded layer is sufficient.
- **`ROLLFORWARD_NOTES`**: Note numbers whose primary table is a
  multi-column rollforward (Fixed Assets, Intangible Assets). For these,
  `value = nums[-1]` (Saldo Akhir / Ending Balance). For everything
  else, `value = nums[0]` (first column = the target year).
- **`TOPIC_ROW_RANGES`** (per-sheet whitelist, lesson L1): only fill
  rows inside the listed ranges. Defaults to "all data rows" if not
  listed. Sheets that share content with neighbours from Run 1 MUST be
  listed here.
- **`target_note=N`** (lesson L4): inside `extract_note_rows`, detect
  per-line `^\s*(\d{1,2})\.\s+[A-Z]{3,}` headers and skip lines that
  are inside a different note's section on the same page.

#### Phase 4 row matching

```python
def match_workbook_label(sheet_label, pdf_rows):
    sl = normalize_label(sheet_label)
    if not sl or len(sl) < 4: return None
    if sl in GENERIC_LABELS: return None
    if not [w for w in sl.split() if len(w) > 2]: return None

    # Pass 1: exact normalised match against label_id or label_en
    for i, r in enumerate(pdf_rows):
        if r['consumed']: continue
        for cand in (normalize_label(r['label_id']),
                     normalize_label(r['label_en'])):
            if sl == cand: return (i, 'exact')

    # Pass 2: substring (require ≥ 6 char shared)
    for i, r in enumerate(pdf_rows):
        if r['consumed']: continue
        for cand in (normalize_label(r['label_id']),
                     normalize_label(r['label_en'])):
            if not cand or len(cand) < 6: continue
            short, long = (sl, cand) if len(sl) <= len(cand) else (cand, sl)
            if len(short) >= 6 and short in long: return (i, 'substring')

    # Pass 3: keyword overlap (asymmetric denominator)
    best, best_score = None, 0.0
    for i, r in enumerate(pdf_rows):
        if r['consumed']: continue
        for cand in (normalize_label(r['label_id']),
                     normalize_label(r['label_en'])):
            score = keyword_overlap(sl, cand)
            if score >= 0.6 and score > best_score:
                best, best_score = i, score
    return (best, 'keyword') if best is not None else None
```

#### Phase 4 appending year-only rows

After the matching loop:

1. Collect the set of all values currently in column B across this
   sheet (`existing_b_values`).
2. For each unconsumed PDF row, test:
   - Has ≥1 significant keyword in its label (after stop-word and
     length filtering).
   - Label is not in `GENERIC_LABELS`.
   - `value` is not a decimal float (rates / percentages are usually
     noise unless this is an EPS / interest-rate note).
   - `abs(value) ≥ 10,000` for ints.
   - `value` not in 1990–2100 (year token).
   - `value not in existing_b_values` (dedup, lesson L7).
3. Write an italic footnote row: `Catatan: Baris berikut hanya muncul
   di Laporan Tahunan <YEAR> (Catatan <N>) / Note: Rows added from
   <YEAR> Annual Report (Note <N>)`.
4. Emit one row per surviving candidate. Column A: `<ID label> / <EN
   label>` indented two spaces. Column B: the value. Cell comment on
   B: `From <YEAR> AR p.<N>. Line: "<raw line up to 120 chars>"`.

### Phase 5 — Key Ratios

Insert the new year's column at the correct chronological position. The
Key Ratios sheet's year-header row is NOT necessarily row 3 — detect it
dynamically by locating the row containing the existing year label
(lesson L18 in this file). Compute ratios from extracted values; use
formulas referencing the main statement sheets where possible.

For ratios needing two-year averages (ROE, ROA, Asset Turnover), now
that two years exist, switch to `(Year N + Year N-1) / 2` and add a
footnote: "Averages computed using two available years."

### Phase 6 — Integrity check, validation, save

1. **Integrity check** (run immediately after all column/row writes):
   no `#REF!` errors, merged cell ranges intact, existing year columns
   still contain their original spot-check values, headers still show
   correct years in correct order.
2. **Financial validation** for the newly added year:
   - BS: `Total Assets = Total Liab + Total Equity`.
   - BS: `Current + Non-Current = Total`.
   - IS: `Revenue − COGS = Gross Profit`.
   - CF: `Op + Inv + Fin (+ FX) = ΔCash`.
   - CF: `Beginning + ΔCash = Ending`.
   - Cross: `IS Net Profit ≈ Equity Statement Net Profit`.
   - Cross: `CF Ending Cash = BS Cash & Equivalents`.
3. **Per-row sign sanity** for income-tax and finance-cost rows: if any
   tax row's sign matches the IS column for the new year but contradicts
   the same row in another year, flag it.
4. **Save** + run recalc: `python /mnt/skills/public/xlsx/scripts/recalc.py
   [output_file]`.

If any integrity check fails, restore from backup and report. Do not
silently adjust numbers.

---

## Implementation reference

### normalize_numbers (lesson L2)

```python
def normalize_numbers(text):
    prev = None
    out = text
    while out != prev:
        prev = out
        out = re.sub(r'(\d\.\d{2})\s+(\d)(?!\d)', r'\1\2', out)
        out = re.sub(r'(\d\.\d)\s+(\d\d)(?!\d)', r'\1\2', out)
    out = re.sub(r'(?<![\w.])(\d{1,2})\s+(\d{1,2}\.\d{3}(?:\.\d{3})*)\b',
                 r'\1\2', out)
    return out
```

### parse_id_number (handles parens, dashes, em-dashes, decimals)

```python
def parse_id_number(s):
    s = s.strip()
    if not s or s in ('-', '—'): return None
    neg = s.startswith('(') and s.endswith(')')
    if neg: s = s[1:-1].strip()
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
        try: v = float(s)
        except ValueError: return None
    else:
        s2 = s.replace('.', '').replace(' ', '')
        try: v = int(s2) if ',' not in s else float(s2.replace(',', '.'))
        except ValueError: return None
    return -v if neg else v
```

### Tokeniser regex

```python
NUM_TOKEN = re.compile(
    r'\(\s*\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?\s*\)'
    r'|'
    r'\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?'
)
```

### Extract per-line label + numbers

```python
def extract_numbers_and_label(line):
    matches = list(NUM_TOKEN.finditer(line))
    nums = [parse_id_number(m.group(0)) for m in matches]
    nums = [n for n in nums if n is not None]
    if not matches: return line.strip(), '', []
    label_id = line[:matches[0].start()].strip()
    label_en = line[matches[-1].end():].strip()
    return label_id, label_en, nums
```

### Noise filter

```python
def is_noise_value(v):
    if isinstance(v, float) and 0 < v < 100: return True
    if isinstance(v, int):
        if 1990 <= v <= 2100: return True
        if 0 <= abs(v) < 100: return True
    return False
```

### Magnitude sanity check (lesson L6)

```python
def magnitude_violates(value, anchor):
    if not (isinstance(value, (int, float)) and isinstance(anchor, (int, float))):
        return False
    if abs(anchor) > 1_000_000 and abs(value) < abs(anchor) / 1000 and abs(value) < 1_000_000:
        return True
    if abs(value) > 1_000_000 and abs(anchor) < abs(value) / 1000 and abs(anchor) < 1_000_000:
        return True
    return False
```

### Generic-label skip list (lesson L3)

```python
GENERIC_LABELS = {
    'total', 'total total', 'sub total', 'sub total sub total',
    'neto net', 'net', 'neto', 'jumlah', 'jumlah total',
    'catatan note', 'note', 'catatan', 'pasal article',
    'pasal', 'article', 'perusahaan the company',
    'entitas anak subsidiaries', 'lain lain others', 'lain lain',
}
```

### Asymmetric keyword overlap (lesson L5)

```python
STOP = {'dan', 'atau', 'yang', 'untuk', 'ke', 'dari', 'dengan', 'dalam',
        'the', 'and', 'or', 'of', 'to', 'from', 'for', 'in', 'on', 'a',
        'an', 'is', 'are', 'be', 'pt', 'tbk'}

def keyword_overlap(a, b):
    wa = {w for w in a.split() if w not in STOP and len(w) > 2}
    wb = {w for w in b.split() if w not in STOP and len(w) > 2}
    if not wa or not wb: return 0.0
    n = len(wa & wb)
    if n == 0: return 0.0
    if min(len(wa), len(wb)) <= 2:
        return n / min(len(wa), len(wb))
    return n / max(len(wa), len(wb))
```

### Note section filter (lesson L4)

```python
NOTE_HEADER = re.compile(r'^\s*(\d{1,2})\.\s+[A-Z][A-Z]')

def extract_note_rows(pdf, start, end, rollforward=False, target_note=None):
    rows, year_ctx = [], None
    for pn in range(start, end + 1):
        txt = normalize_numbers(pdf.pages[pn-1].extract_text() or '')
        in_target = (target_note is None)
        for raw in txt.split('\n'):
            line = raw.rstrip()
            stripped = line.strip()
            if not stripped: continue
            m = NOTE_HEADER.match(stripped)
            if m and 'lanjutan' not in stripped.lower() and 'continued' not in stripped.lower():
                n = int(m.group(1))
                if target_note is not None: in_target = (n == target_note)
                year_ctx = None
                continue
            if not in_target: continue
            if re.fullmatch(r'2024(\s+/\s+2024)?', stripped): year_ctx = 'NEW'; continue
            if re.fullmatch(r'2023(\s+/\s+2023)?', stripped): year_ctx = 'OLD'; continue
            if re.search(r'\b<NEW>\b.*\b<OLD>\b', stripped) and 'Rp' not in stripped:
                year_ctx = 'BOTH'; continue
            if year_ctx == 'OLD': continue
            label_id, label_en, nums = extract_numbers_and_label(line)
            if not nums: continue
            real = [n for n in nums if not is_noise_value(n)]
            if not real: continue
            value = real[-1] if (rollforward and year_ctx == 'NEW') else real[0]
            rows.append({
                'label_id': label_id, 'label_en': label_en,
                'value': value, 'all_nums': real,
                'line': stripped, 'page': pn, 'consumed': False,
            })
    return rows
```

Replace `<NEW>` / `<OLD>` with the literal target / comparative year.

---

## Anti-hallucination rules

1. Never fill a cell with a calculated value without an italic footnote
   explaining the calculation.
2. Fallback calculations allowed only for `Total Liab = TA − TE` when
   direct extraction fails. Mark with footnote.
3. Comparative-year values in the new PDF: extract only the target year.
4. Consolidated only.
5. OCR-extracted values are flagged with cell comments.
6. Zero vs empty: explicit zero in the PDF → `0`; parsing failure →
   empty + comment. Never confuse the two.
7. Financial Highlights ≠ audited FS. Only extract from pages after the
   auditor's report.
8. Workbook data is never a source. Every value in the new column comes
   from the new PDF (or the hardcoded layer's manually-verified
   reading of the new PDF).

---

## Terminal output

```
=== IDX Excel — Year Append Complete ===
Company        : <TICKER> – <COMPANY>
Year Added     : <YEAR>
Source PDF     : <PATH> (<N> pages)
Output File    : <XLSX>
Backup File    : <BACKUP>
Years in File  : <list, left to right>

Source location in PDF:
  Auditor's Report : page <N>  (firm: <NAME>, license <KEP-...>)
  Audit Status     : <AUDITED | UNVERIFIED>
  Statement Type   : <CONSOLIDATED | STANDALONE>
  BS / IS / CF / EQ: <page ranges>

Column inserted: <LETTER>  (<left of | right of | between>)

Note matching:
  <N>. <TITLE> → "<sheet>" (renumbered from Note <X>)
  <N>. <TITLE> → NEW SHEET CREATED
  "<sheet>" → note NOT FOUND in <YEAR>

Phase 4 results (per sheet):
  <sheet>: matched=<M>, appended=<A>, skipped-hardcoded=<S>

Integrity:
  #REF! errors         : <NONE | count>
  Existing data spot   : <PASS | FAIL>
  Merged cells intact  : <PASS | FAIL>

Financial validation:
  BS Assets = Liab+Eq  : <PASS | FAIL>
  IS Rev − COGS = GP   : <PASS | FAIL>
  CF Op+Inv+Fin = ΔCash: <PASS | FAIL>
  CF Begin+Δ = End     : <PASS | FAIL>
  Cross IS NP ≈ Eq NP  : <PASS | FAIL>
  Cross CF End = BS Cash: <PASS | FAIL>

Notes warnings:
  - <any sign-flips applied>
  - <any topic-range guards triggered>
  - <any unmatched PDF rows by sheet>
=========================================
```

---

## Lesson cross-reference

This skill operationalises:

- **L1** TOPIC_ROW_RANGES whitelist
- **L2** normalize_numbers() pre-pass
- **L3** Two-layer extractor + GENERIC_LABELS skip-list
- **L4** target_note section filter
- **L5** Asymmetric keyword-overlap denominator
- **L6** Magnitude sanity check during matching
- **L7** Value-based dedup during append
- **L8** Note Index in English with Indonesian comments
- **L9** Sign-flip via hardcoded layer
- **L10** Documentation-drift guard in Phase 0
- **L11** cffi before pdfplumber
- **C1** Strip note-reference patterns, prefer Indonesian half
- **C2** Re-verify equity opening against the year-being-added's own AR
- **C3** Equity page mirror detection (token reversal)
- **C4** Match notes by topic, never by number
- **C5** Trust PDF sign on CF lines, not the label
- **C6** Fold absent-template equity columns into combined column
- **C7** Audit pre-existing movement rows; clear if not in the new AR
