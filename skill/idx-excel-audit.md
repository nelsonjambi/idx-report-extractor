---
name: idx-excel-audit
description: >
  Independent post-run audit of an IDX Financial Statements workbook. Re-derives
  anchor values from the source PDF(s), runs the full arithmetic reconciliation
  matrix, measures per-note coverage, detects cloned content across note sheets,
  and sweeps for the hygiene defects this repo has historically shipped
  (N/A strings, missing comments, language mixing, note-refs-as-values,
  sign inconsistencies, misplaced headers). Run this after EVERY run that
  writes to a workbook — Run 1, appends, and manual fixes. Extraction is not
  "done" until this audit passes. The audit only reports; it never edits the
  workbook.
---

# IDX Excel — Independent Workbook Audit

## Why this skill exists

Every defect the user has had to flag by hand in this repo — an empty 2024
column on Fixed Assets, an Indonesian/English mixed Note Index, a missing
Note-15 detail block, a note reference `14` sitting in a billions row,
duplicated appended rows, identical content cloned across four note sheets —
was mechanically detectable. This skill is that mechanical detector. It is
**independent**: it re-reads the PDF itself rather than trusting the
extraction script's own summaries, so it catches category errors the
extractor cannot see in itself.

Two iron rules:

1. **The audit never modifies the workbook.** It produces a findings report.
   Fixes happen in the extraction script / ledger, then the audit re-runs.
   Never "fix" a finding by weakening the audit.
2. **Findings in columns you did not create this session are reported, not
   fixed** (CLAUDE.md §2.4). The user decides what happens to prior data.

## Required input

| Parameter | Example |
|---|---|
| Workbook | `WIFI_Financial_Statements.xlsx` |
| Source PDF per year column | `{2024: "reports/WIFI_Annual_Report_2024.pdf", 2025: "reports/WIFI_Annual_Report_2025.pdf"}` |
| Ledger files if they exist | `ledgers/WIFI_2024_ledger.json` |
| Scope | `all years` or `year just added` |

If a year column has no source PDF available, audit it structurally
(Phases 1–2, 4–6) and mark anchor re-derivation SKIPPED for that year.

---

## Phase 1 — Structural hygiene sweep

Walk every sheet, every cell. Each check is binary.

| ID | Check | Fails when |
|---|---|---|
| H1 | No `"N/A"` strings in any data cell | any found (breaks SUM) |
| H2 | No `#REF!` / `#VALUE!` / `#NAME?` in any cell | any found |
| H3 | Year headers exist on every data sheet and are chronological left→right | missing / out of order / duplicated year |
| H4 | Year headers sit on the correct row per sheet (main statements row 3; Key Ratios row 4; Note Index row 2) — detect by locating the row that contains the other year labels | a year label floats on a different row than its siblings |
| H5 | Note Index sits immediately after `Changes in Equity`; row count from row 3 equals the number of note sheets | position or count mismatch |
| H6 | Note Index language uniformity: no title cell in any year column contains Indonesian-only vocabulary (test against a wordlist: `UMUM, ASET, UTANG, BEBAN, PIUTANG, MODAL, PENDAPATAN, PENGHASILAN, PERPAJAKAN, SAHAM, IMBALAN, PINJAMAN, SEWA, KAS`) | any hit — titles must be English, Indonesian belongs in the cell comment |
| H7 | Every populated Note Index title cell carries an `Indonesian title:` comment | missing comment |
| H8 | Backup file exists for the most recent mutation | no `*_backup_before_*` matching the last run |

## Phase 2 — Arithmetic reconciliation matrix

For **every year column present**, compute with actual cell values and report
the numeric diff (not just pass/fail):

| ID | Identity | Tolerance |
|---|---|---|
| A1 | BS: `TA − (TL + TE)` | 0 |
| A2 | BS: `ΣCurrent assets rows − Total Current Assets` and non-current, and both liability sides | 0 |
| A3 | IS: `Revenue − |COGS| − Gross Profit` | 0 |
| A4 | IS: `NP − Σattribution rows` and `CI − Σattribution rows` | 0 |
| A5 | CF: `Op + Inv + Fin − ΔCash` | 0 |
| A6 | CF: `Beginning + ΔCash − Ending` | 0 |
| A7 | Cross: `CF Ending − BS Cash` | 0 |
| A8 | Cross: `IS NP-to-parent − EQ NP row` | 0 |
| A9 | EQ: `opening + Σmovements − closing` per column and per year block | 0 |
| A10 | EQ chaining: closing of year N = opening of year N+1 | 0 |

Tolerance really is zero — full-Rupiah integers. A residual of exactly one
movement row's amount usually means a stale comparative artifact (see
lesson C7); name the row in the finding.

## Phase 3 — Independent anchor re-derivation from the PDF

For each year with a source PDF:

1. Locate the audited FS (auditor's-report anchor **with firm letterhead**;
   `KONSOLIDASIAN` + `(Disajikan dalam Rupiah` on statement pages). Never
   read the Financial Highlights (F1/F2).
2. Apply `normalize_numbers()` to every page (F4).
3. Independently parse ~20 anchor values: TA, TCA, TNCA, TL, TCL, TNCL, TE,
   NCI, cash, trade receivables net, fixed assets NBV, taxes payable,
   revenue, COGS, GP, NP, NP-to-parent, Op CF, Inv CF, Fin CF, ending cash.
4. Diff each against the workbook cell. Report every nonzero diff with page
   number and both values.
5. If a ledger file exists, three-way diff (PDF ↔ ledger ↔ workbook) and
   report which leg disagrees — that localizes the bug to parse vs write.

## Phase 4 — Note coverage matrix

The user's #1 historical complaint: a year column that exists but is silently
empty. For every note sheet × every year column:

1. Count `template_rows` = rows with a numeric value in the workbook's
   reference year column (the year the template was built from), restricted
   to the sheet's `TOPIC_ROW_RANGES` if defined.
2. Count `filled` = rows with a numeric value in the audited year column
   within the same range.
3. Count `commented_empty` = empty cells in that range carrying a comment
   (legitimate absences).
4. **Finding** when `filled + commented_empty < template_rows` — those are
   unexplained gaps: list the row numbers and labels.
5. **Finding** when `filled == 0 or 1` (header/total only) for a note whose
   equivalent exists in that year's PDF index — the "only the total was
   populated" failure. Cross-check against the PDF: does the note contain a
   table? If yes, severity HIGH.
6. Also verify the renumbering footnote exists on sheets whose note number
   differs across years.

Output as a table: `sheet | year | template | filled | commented | gaps`.

## Phase 5 — Clone and duplicate detection

1. **Cross-sheet clones (F7):** for every pair of note sheets, compute the
   multiset of numeric values (|v| > 10,000) per year column and its overlap
   ratio `|A∩B| / min(|A|,|B|)`. Flag pairs over 0.30 — either Run 1 cloned
   content (report; don't fix prior columns) or this run's fill leaked across
   a topic boundary (fix).
2. **Within-sheet duplicate appends (F13):** in the appended-rows region
   (below the italic footnote), flag any value that also appears in the
   matched region above, and any label appearing twice.
3. **Bogus appended rows:** appended rows whose label has <1 significant
   keyword, or whose value is a year (1990–2100), a lone small integer, or a
   bare decimal rate.

## Phase 6 — Plausibility and sign sweep

| ID | Check | Fails when |
|---|---|---|
| P1 | Magnitude ratio across year columns in the same row ≤ 1000× (when both ≥ 1M is exempt) | e.g. `14` next to `276,941,197,068` — a note-ref-as-value (F8) |
| P2 | No lone integers < 100 or years 1990–2100 as data values | any found outside count/ratio rows |
| P3 | No decimal floats in full-Rupiah rows (EPS, ratio, and %-rows exempt) | any found |
| P4 | Sign consistency: the same row's values across years share a sign, OR the deviating cell carries a comment | uncommented sign flip (F9/F10 symptom) |
| P5 | Sub-total monotony: a "Total" row ≥ any single component in the same block (absolute values) | component exceeds its own total — usually a column-pick error |
| P6 | Comment discipline: every fuzzy/keyword-matched or hardcoded value has a provenance comment | naked value with no trace to a PDF page |

## Phase 7 — Report and fix loop

Produce the report (terminal + optionally a `audit_<date>.md` committed with
the run):

```
=== Workbook Audit — <file> ===
Scope: <years>     Sources: <pdfs>     Ledgers: <found|none>

VERDICT: <PASS | FAIL (n findings: c critical, h high, m medium, l low)>

CRITICAL   (wrong numbers a reader would trust)
  [A1|2024] BS identity diff = <n>            → <cells involved>
  [P1|Note 10!B32] value 14 vs 2025 276.9B    → note-ref parsed as value
HIGH       (missing data presented as if complete / provenance gaps)
  [C4|Note 19|2024] 11 unexplained empty rows: <rows+labels>
MEDIUM     (hygiene that will bite the next run)
  [H6|Note Index!B34] "BEBAN OPERASIONAL" is Indonesian
LOW        (cosmetic)
  ...

Coverage matrix:      <table>
Clone pairs > 0.30:   <list or none>
Anchor re-derivation: <n> anchors, <d> diffs   <per-year>
```

Severity definitions (fixed, not judgment calls):

- **CRITICAL** — any Phase 2 nonzero diff, any Phase 3 anchor diff, any P1/P5
  hit: a number is wrong or arithmetic doesn't close.
- **HIGH** — Phase 4 unexplained gaps, "total-only" notes with tabular PDF
  source, clone/duplicate findings in columns created this session, P4/P6.
- **MEDIUM** — H-checks, bogus appended rows, findings located in prior-run
  columns (report-only).
- **LOW** — formatting/width/freeze-pane deviations.

**Fix loop protocol:**

1. Fix CRITICAL and HIGH findings in the extraction script or ledger — never
   by hand-editing cells (hand edits are unreproducible) and never by
   relaxing an audit check.
2. Re-run extraction from the backup, then re-run this audit.
3. Repeat until the audit passes **twice in a row from a clean re-run**.
4. MEDIUM findings in prior-run columns: list them for the user with a
   proposed remediation; do not touch without approval.
5. Commit the audit report in the same commit as the fix.

**Exit criteria — extraction may be declared done only when:** zero CRITICAL,
zero HIGH; every MEDIUM either fixed or explicitly accepted by the user; the
final report is committed.
