---
name: idx-excel-anchor-ledger
description: >
  Ledger-first extraction for IDX Annual Report financial data. Builds a
  structured, self-verifying JSON ledger of every statement line and every
  note sub-line BEFORE anything is written to Excel. A value may only reach
  the workbook after it has passed arithmetic tie-out inside the ledger
  (items → sub-totals → totals → cross-statement anchors). Use this skill
  WHENEVER numbers are being extracted from a PDF for a workbook — it is the
  numbers-engine that idx-excel-run1-template-rev.md and
  idx-excel-append-year-rev.md delegate to. The ledger file is committed
  alongside the workbook as the audit trail.
---

# IDX Excel — Anchor-Ledger Extraction (extract → verify → write)

## Why this skill exists

Every costly error in this repo's history shares one shape: **a number
entered Excel before it was verified**, and was discovered later by a human
paging through sheets. Fixing errors in ledger space is minutes; finding them
in workbook space is hours. This skill inverts the pipeline:

```
OLD:  PDF ──parse──▶ Excel ──inspect──▶ find errors ──▶ re-run   (slow loop, human finds bugs)
NEW:  PDF ──parse──▶ Ledger ──tie-out──▶ FAIL? fix parse, loop   (fast loop, arithmetic finds bugs)
                        │ all green
                        ▼
                      Excel (mechanical write, no judgment)
```

**The invariant: no cell in the workbook ever holds a number that is not in a
VERIFIED or MANUAL entry of the committed ledger.**

---

## Role

Expert financial analyst and certified accountant (IDX/BEI, PSAK, IFRS)
building a reconciled trial-extraction before presentation.

## Required input

| Parameter | Example |
|---|---|
| Source PDF | `reports/WIFI_Annual_Report_2024.pdf` |
| Fiscal year to extract | `2024` |
| Ticker | `WIFI` |
| Target workbook (existing or to-be-created) | `WIFI_Financial_Statements.xlsx` |
| Ledger output path | `ledgers/WIFI_2024_ledger.json` |

---

## The ledger format

One JSON file per (company, year). Everything the workbook will contain for
that year must appear here first.

```json
{
  "meta": {
    "ticker": "WIFI",
    "year": 2024,
    "source_pdf": "reports/WIFI_Annual_Report_2024.pdf",
    "pdf_pages": 339,
    "auditor_report_page": 197,
    "audit_firm": "ANWAR & REKAN",
    "statement_pages": {"bs": [203, 205], "is": [206, 207], "eq": [208, 209], "cf": [210, 210]},
    "extraction_date": "<date>",
    "status": "VERIFIED | PARTIAL | UNVERIFIED"
  },
  "statements": {
    "bs":  {"rows": [ROW, ...], "checks": [CHECK, ...]},
    "is":  {"rows": [...], "checks": [...]},
    "cf":  {"rows": [...], "checks": [...]},
    "eq":  {"rows": [...], "checks": [...]}
  },
  "notes": {
    "8": {
      "title_id": "ASET TETAP", "title_en": "FIXED ASSETS",
      "pages": [248, 250], "layout": "rollforward",
      "status": "VERIFIED",
      "rows": [ROW, ...],
      "checks": [CHECK, ...]
    }
  }
}
```

**ROW object:**

```json
{
  "id": "bs.cash",                      // stable key, dot-namespaced
  "label_id": "Kas dan setara kas",
  "label_en": "Cash and cash equivalents",
  "value": 18495026165,
  "page": 203,
  "line": "Kas dan setara kas 2c,2d,4 18.495.026.165 40.072.539.130 Cash and cash equivalents",
  "method": "parsed | manual | ocr | computed",
  "comparative_value": 40072539130,     // the OTHER year's number on the same line — proves column choice
  "sign_note": null,                    // set when workbook sign ≠ PDF sign, e.g. "flipped to IS expense convention"
  "verified": true
}
```

Field rules:

- `method: "manual"` — a human/model read the PDF page directly and typed the
  value. Requires `page` and `line` (quote the PDF text) so it is re-checkable.
- `method: "computed"` — allowed ONLY for `Total Liabilities = TA − TE`;
  requires a `sign_note`-style explanation.
- `comparative_value` is **mandatory for every parsed row from a two-column
  table**. Recording the neighbor value proves you took the correct year's
  column — the single most common silent error in bilingual two-column notes.
- `sign_note` is mandatory whenever `value`'s sign differs from the PDF's
  printed sign (tax rows, finance costs into IS-convention columns).

**CHECK object:**

```json
{
  "kind": "sum | tie",
  "expr": "sum(note8.cost.*) == note8.cost.total",
  "lhs": 2611749297374,
  "rhs": 2611749297374,
  "diff": 0,
  "pass": true
}
```

- `sum` checks: component rows add to their sub-total; sub-totals add to total.
- `tie` checks: a note total equals its anchor in a main statement (see the
  tie-out table below). Tolerance is **0** — these are integers in full
  Rupiah; any nonzero diff means an extraction error, not rounding.

---

## Canonical tie-out table

Every quantitative note MUST declare at least one `tie` check against a main
statement. This table is the default set; extend per company as needed.

| Note topic | Ledger total | Ties to |
|---|---|---|
| Cash & equivalents | note total | BS `Kas dan setara kas` |
| Trade receivables | net (after allowance) | BS `Piutang usaha` |
| Prepaid expenses | short-term sub-total / long-term sub-total | BS current / non-current prepaid rows |
| Advances | total | BS `Uang muka` (current + non-current) |
| Fixed assets | Nilai Buku Neto (closing cost − closing acc. depr.) | BS `Aset tetap - neto` |
| Intangibles | net | BS `Aset takberwujud - neto` |
| Trade payables | total | BS `Utang usaha` |
| Other payables | current / non-current portions | BS rows 3rd-party payables |
| Accrued expenses | total | BS `Beban akrual` |
| Advances from customers | short-term / long-term split | BS current & non-current rows |
| Lease liabilities | current maturity / long-term portion | BS lease rows |
| Taxation | prepaid-tax total; taxes-payable total; income-tax expense | BS prepaid tax; BS `Utang pajak`; IS tax expense |
| Consumer financing | current / non-current | BS rows |
| Bank loans / Loans / Bonds | current portion / long-term portion | BS rows |
| Due to related parties | total | BS row |
| Employee benefits | liability total | BS row |
| Share capital | issued & paid-in | BS `Modal ditempatkan dan disetor` |
| APIC | total | BS `Tambahan modal disetor` |
| NCI | total | BS `Kepentingan nonpengendali` |
| Revenues | net revenue | IS revenue |
| Cost of revenues | total | \|IS COGS\| |
| Operating / G&A expenses | sub-total | \|IS G&A\| |
| Other income (expense) | net | IS line |
| Finance income / costs | total | IS lines |
| EPS | NP attributable to parent; per-share figure | IS attribution row; IS EPS row |
| Depreciation allocation (in Fixed Assets note) | COGS + Opex portions | sum = total depreciation additions |
| Equity statement | closing total; NP row; opening total | BS Total Equity; IS NP-to-parent; prior year's closing |

A note whose total cannot be tied to any statement (pure narrative, or truly
free-standing detail) gets `status: "NARRATIVE"` and contributes no numbers —
the workbook gets the year header plus a sheet-level comment only.

---

## Execution sequence

### Phase 0 — Setup and probe

1. `pip install cffi pdfplumber openpyxl --break-system-packages` (cffi first).
2. Run the structure probe from `idx-excel-append-year-rev.md` Phase 1:
   auditor's report located by keyword **plus firm letterhead** (name +
   `KEP-…` license). Record everything into `meta`.
3. If the auditor's report is scanned/unfindable, set
   `meta.status = "UNVERIFIED"` and **stop to ask the user** before investing
   in a full extraction.

### Phase 1 — Statement rows into the ledger

For each of BS, IS, CF, EQ:

1. Extract text from the identified pages only; run `normalize_numbers()`
   on every page (see append-year-rev Implementation — mandatory, F4).
2. Parse every line into ROW objects. For two-column comparative layouts,
   record `comparative_value`.
3. Add `sum` checks: current+non-current=total (BS both sides), Rev−COGS=GP,
   attribution sums (IS), section sub-totals and Beginning+Δ=End (CF),
   opening+movements=closing per column (EQ).
4. Add `tie` checks: TA = TL + TE; CF ending cash = BS cash; EQ closing = BS
   total equity; EQ NP row = IS NP-to-parent.
5. **Equity special case (F16):** if header text reads mirrored, reverse each
   numeric token; the EQ ties above are the acceptance test for the
   reconstruction.

### Phase 2 — Note rows into the ledger

Build the note index (headers `^\d{1,2}\. [A-Z]…`, skipping
`lanjutan`/`continued`). For each quantitative note:

1. Extract lines with the **section filter** (`target_note=N`) — page ranges
   are not boundaries (F6).
2. Declare the note's `layout`:
   - `two-column` (default): target-year value is the FIRST numeric column.
   - `rollforward` (Fixed Assets, Intangibles): target value is the LAST
     column (Saldo Akhir) within the target year's block; a `2023` block
     header on the same pages switches context off.
3. Parse rows. Apply the noise filter (lone ints <100, years, bare rates) so
   note references never become values (F8).
4. Write `sum` checks for every visible sub-total chain in the note.
5. Write the `tie` check(s) from the canonical table.
6. Where the auto-parse cannot produce a clean row (generic labels, merged
   table cells, complex reconciliations): read the PDF page yourself and add
   the row with `method: "manual"`, quoting the source line. Manual entries
   participate in checks like any other row — being manual does not exempt a
   number from tie-out.
7. Sign conventions: if the workbook column needs the opposite sign from the
   PDF (tax expense, finance costs), store the workbook-convention value and
   set `sign_note` (F10).

### Phase 3 — Verify the ledger (the gate)

Run every check. Then compute per-note and global status:

- Note `VERIFIED`: all sums pass AND at least one tie passes.
- Note `PARTIAL`: ties pass but some component rows unparsed (list them).
- Note `UNVERIFIED`: any check fails → **no numbers from this note may be
  written**. Fix the parse or downgrade rows to empty-with-comment.
- `meta.status = VERIFIED` only when: all four statements fully pass, and
  every quantitative note is VERIFIED or PARTIAL, and no check object has
  `pass: false`.

Print the check matrix (every check: expr, lhs, rhs, diff, pass). Iterate on
Phases 1–2 until green. **Do not proceed to Phase 4 with any failing check
unless the user has explicitly accepted a documented residual.**

### Phase 4 — Mechanical write to Excel

Only now touch the workbook (after backup, per CLAUDE.md §2):

1. For a new workbook, follow `idx-excel-run1-template-rev.md` structure; for
   an append, follow `idx-excel-append-year-rev.md` column-insertion and
   note-matching (topic titles, `TOPIC_ROW_RANGES`, `GENERIC_LABELS`,
   magnitude check, dedup on append).
2. Every write copies `value` from a ledger ROW. The writer performs **no
   arithmetic and no parsing** — if you find yourself computing in Phase 4,
   the ledger is incomplete; go back to Phase 2.
3. Cell comments carry provenance: `page` for manual/fuzzy entries;
   `sign_note` verbatim where set.
4. Ledger rows with no matching workbook row → appended as year-only rows
   under the italic footnote. Workbook rows with no ledger row → empty +
   `Not found in <YEAR> report` comment.

### Phase 5 — Post-write reconciliation and commit

1. Re-open the saved workbook; for every ledger row that was written, assert
   `workbook cell value == ledger value`. Zero mismatches allowed.
2. Run `skill/idx-excel-audit.md` (the independent audit). Fix and repeat
   until it passes.
3. Commit **ledger + workbook + script + doc updates in one commit**.

---

## Terminal output

```
=== Anchor-Ledger Extraction — <TICKER> <YEAR> ===
Ledger      : <path>   status: <VERIFIED | PARTIAL | UNVERIFIED>
Statements  : BS <n> rows | IS <n> | CF <n> | EQ <n>
Notes       : <n> quantitative (<v> VERIFIED, <p> PARTIAL, <u> UNVERIFIED), <m> narrative

Check matrix: <N> checks, <F> failing
  [FAIL] <expr>: lhs=<..> rhs=<..> diff=<..>     ← list every failure
  ...
Tie-outs to main statements: <T> declared, <T-pass> passing

Manual entries : <n>  (each with page + quoted line)
Sign overrides : <n>  (each with sign_note)

Excel write    : <written> cells written, <appended> rows appended,
                 <empty> cells left empty+comment
Post-write     : <mismatches> ledger↔workbook mismatches  (must be 0)
================================================
```

## Hard rules

1. No number reaches the workbook except from a VERIFIED/PARTIAL ledger note
   (and within PARTIAL, only rows with `verified: true`).
2. Tie-out tolerance is zero. A diff of 1 is a bug.
3. `comparative_value` recorded for every two-column parse — no exceptions.
4. Manual entries quote the PDF line. "I remember it said" is not a source.
5. The ledger is committed. If the workbook and ledger ever disagree, the
   workbook is wrong (the ledger passed checks; the workbook write is the
   only unchecked step, and Phase 5.1 exists to catch it).
