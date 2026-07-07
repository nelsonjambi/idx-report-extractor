# CLAUDE.md — Operating Manual for idx-report-extractor

You are working in a repository that extracts **audited consolidated financial
statements of IDX-listed Indonesian companies** from Annual Report PDFs into
multi-year Excel workbooks. The numbers end up in front of a human who makes
decisions with them. A wrong number that *looks* right is the worst outcome
this repo can produce. Read this whole file before your first write.

---

## 1. What lives where

```
reports/                          Source Annual Report PDFs (Laporan Tahunan)
skill/idx-excel-run1-template-rev.md   How to build a NEW workbook (Run 1)
skill/idx-excel-append-year-rev.md     How to APPEND a year to an existing workbook
skill/idx-excel-audit.md               How to independently AUDIT a workbook after any run
skill/idx-excel-anchor-ledger.md       Ledger-first extraction: verify numbers BEFORE Excel
skill/idx-excel-quarterly.md           Quarterly reports (different pipeline — do not mix)
lesson-rev.md                     Curated lessons (L1–L10 new, C1–C7 carryover). Read before any run.
lesson.md                         Full historical lesson log (append new lessons here)
README-001-wifi.md                Per-company status doc for WIFI
WIFI_Financial_Statements.xlsx    The deliverable. Treat as production data.
*_backup_before_*.xlsx            Safety nets. Never delete.
extract_wifi_fs.py / append_2024_wifi.py / append_2024_notes.py   Run scripts (committed artifacts)
```

The `-rev.md` skill files supersede their non-rev counterparts. Use the rev
versions.

## 2. Non-negotiable conventions

1. **The workbook is the source of truth, not the docs.** README and lesson
   files describe; the xlsx *is*. They have drifted before (README claimed
   "Run 2 complete" while the workbook held only one year).
2. **Backup before mutation.** Copy the xlsx to
   `<name>_backup_before_<change>.xlsx` (or scratchpad) before the first write.
3. **Every number in a year column comes from that year's own Annual Report
   PDF.** Never from another workbook column, never from a later year's
   comparative column, never computed — with one exception:
   `Total Liabilities = Total Assets − Total Equity`, which requires an italic
   footnote naming the formula and why direct extraction failed.
4. **Existing rows and existing year columns are immutable.** Do not rename,
   reorder, delete, or "clean up" — even when you can see they contain a prior
   run's mistake. Constrain your own writes instead, and report the prior
   mistake to the user.
5. **New information is additive.** PDF line items with no workbook row become
   new rows at the bottom of the sheet, under an italic footnote row:
   `Catatan: Baris berikut hanya muncul di Laporan Tahunan <YEAR> / Note: Rows added from <YEAR> Annual Report`.
6. **Missing ≠ zero ≠ "N/A".** Missing data = empty cell + cell comment (e.g.
   `Not found in 2024 report`). Explicit zero or dash in the PDF = `0`, no
   comment. Never write the string "N/A" into a data cell — it breaks SUM.
7. **Bilingual labels**: Column A is `Indonesian / English`. The **Note Index
   sheet uses English titles in every year column**, with the Indonesian
   original in a cell comment (`Indonesian title: …`).
8. **Column order is chronological**: oldest year leftmost, newest rightmost.
9. **Commit the workbook, the script that produced it, and the doc updates in
   one commit**, on the designated `claude/...` branch, pushed with
   `git push -u origin <branch>`. Never create a PR unless asked (PR #9
   already tracks the current branch — pushes update it).
10. **Environment**: `pip install cffi pdfplumber openpyxl --break-system-packages`
    — `cffi` explicitly and first, or pdfplumber's import fails with a
    misleading `_cffi_backend` error.

## 3. Named failure modes — and the rule that prevents each

These are not hypothetical. Every one happened in this repo. If you feel the
pull toward any of these behaviors, that is the signal to apply the rule.

| # | Failure mode | What it looks like | Preventing rule |
|---|---|---|---|
| F1 | **Highlights trap** | Extracting BS/IS numbers from the Financial Highlights (Ikhtisar Keuangan) or MD&A tables near the front of the PDF | Only extract from pages AFTER the auditor's report. A statement page must contain `KONSOLIDASIAN` **and** `(Disajikan dalam Rupiah` on the same page. |
| F2 | **Fake auditor anchor** | Treating the Table of Contents or FS cover page as the auditor's report because the keyword matched | The real auditor's report page contains the audit firm letterhead: firm name and license number (`KEP-…/KM.1/…`). Keyword alone is insufficient. |
| F3 | **Note-number matching** | Assuming Note 8 in 2024 = Note 8 in 2025 | Note numbers shift up to ±6 positions between adjacent years. Match by **topic title** (Indonesian preferred), never by number. |
| F4 | **Naive number regex** | Silently dropping rows because the PDF renders `60.0 29.971.450` (spaces inside numbers) or `P eriklanan` (detached first letter) | Run `normalize_numbers()` (see `skill/idx-excel-append-year-rev.md`, Implementation) on every page of text before tokenizing. Drop tokens ≤2 chars when matching labels. |
| F5 | **Generic-label auto-match** | Auto-assigning a value to a row labeled `Total Total` or `Pasal Article` — dozens of PDF lines match | Rows whose normalized label is in the `GENERIC_LABELS` set get values ONLY from the hardcoded verified layer, never from the auto-matcher. |
| F6 | **Page-range bleed** | Extracting Note 27's table and getting Note 26 and 28 rows too, because one PDF page holds up to four notes | Filter lines by note **section header** (`^\d{1,2}\. [A-Z]…`, excluding `lanjutan`/`continued`), tracking which note's section each line is inside. Page ranges alone are not boundaries. |
| F7 | **Clone contamination** | Filling a value into every sheet where the label matches — when Run 1 cloned identical rows into Notes 31–36 | Respect the per-sheet `TOPIC_ROW_RANGES` whitelist. Rows outside a sheet's topic range stay empty in the new year column, even if the label matches. |
| F8 | **Note-ref as value** | Writing `14` (from `Catatan 14`) into a cell whose other year holds ~277 billion | Two guards: noise filter (drop lone integers <100, years 1990–2100, bare decimal rates) AND magnitude check (reject a match when value and same-row anchor differ by >1000× and one side is <1M). |
| F9 | **Label-sign inference** | Flipping a CF value negative because the label says "Pembayaran" (payment) | Trust the sign printed in the PDF. Then require `Op + Inv + Fin = ΔCash` to hold; if it fails, flip candidates one at a time until it balances. |
| F10 | **Note-vs-IS sign flip** | Copying `+38,286,594,180` current tax from the note's reconciliation table into an IS-convention column where it must be negative | Sign overrides live ONLY in the hardcoded layer, each with a cell comment explaining the flip. Never guess signs systematically. |
| F11 | **Doc trust** | Believing the README's "Run 2 ✅ Complete" when the workbook had only one year of data | At session start, print `wb.sheetnames`, the year headers of each main sheet, and one sample value per sheet. Reconcile against the README **before any write**. If they disagree, tell the user and stop. |
| F12 | **Comparative-column sourcing** | Taking 2024's equity opening balance from the 2025 AR's comparative column (it was 23% wrong — mirrored pages) | Each year's data comes from its own AR. On every append, re-verify the equity opening balance against the year-being-added's AR, and check `opening + movements = closing` row by row. |
| F13 | **Duplicate appends** | Appending a "2024-only" row whose value already sits in a matched row | Before appending, build the set of all values already in the year column; skip candidates whose value is in the set. |
| F14 | **Short-label miss** | Failing to match `Tanah / Land` because Jaccard-over-max scores it 0.5, under the 0.6 threshold | Asymmetric overlap: when the smaller side has ≤2 significant keywords, divide the intersection by `min(|a|,|b|)`, else by `max`. |
| F15 | **Hardcoded header row** | Writing the year header to row 3 on Key Ratios, whose headers are on row 4 | Locate the row containing the existing year label dynamically; write the new header in that row. |
| F16 | **Mirrored equity pages** | Reading `969.843.329.191` where the PDF shows `191.923.348.969` — landscape spread extracted in reverse | If header words read backwards (`YTIUQE`), reverse each numeric token's characters. Accept only if the closing balance equals BS Total Equity. |
| F17 | **Self-certified done** | Saving and committing because the script ran without error | Run the validation matrix (§4) and paste actual diffs into your summary. "Ran clean" is not "verified correct". |

## 4. Quality bar per deliverable — checkable criteria

### Workbook (after ANY run that writes to it)

Every item is a mechanical check; run them all and report the numbers.

- [ ] BS identity per year column: `Total Assets − (Total Liab + Total Equity) = 0`
- [ ] BS: `Current + Non-Current = Total` for both assets and liabilities
- [ ] IS: `Revenue − COGS − Gross Profit = 0`; `Net Profit = attribution total`
- [ ] CF: `Op + Inv + Fin − ΔCash = 0`; `Beginning + ΔCash − Ending = 0`
- [ ] Cross-statement: `CF Ending Cash = BS Cash`; `IS Net Profit ≈ Equity-statement NP row`
- [ ] Equity: `opening + movements = closing` for each year block
- [ ] Zero `#REF!` errors; zero `"N/A"` strings anywhere
- [ ] Year headers present and in chronological order on every data sheet
- [ ] Five spot-check values in pre-existing year columns identical to the backup
- [ ] Every empty cell inside a matched note's topic range carries a comment
- [ ] Every fuzzy/keyword/hardcoded/sign-flipped value carries a provenance comment (PDF page or reason)
- [ ] Note Index: every year column English; every populated title cell has an `Indonesian title:` comment; row count = note-sheet count; sheet sits immediately after `Changes in Equity`

### Extraction script

- [ ] Re-runnable: `cp backup → xlsx && python script.py` reproduces the same workbook
- [ ] Prints a per-sheet summary table: matched / appended / skipped counts
- [ ] Every entry in the hardcoded dict has a comment stating source page and, if non-obvious, the reason (e.g. a sign flip)
- [ ] No code path copies a value from one workbook column to another

### lesson.md updates

- [ ] Each new lesson has: numbered heading, "What happened" (specific, with real values), "Fix / rule to add" (generalized — usable on a different company's PDF), and a row in the run's summary table
- [ ] No lesson that is merely a one-off debugging anecdote with no forward value

### README / status docs

- [ ] "Years in File", sheet counts, and validation results match the actual workbook you just saved — verify by reading the workbook, not by memory
- [ ] Updated in the same commit as the workbook and script

### Commits

- [ ] Workbook + script + docs in one commit; message states what changed and the validation outcome
- [ ] On the designated `claude/...` branch; pushed with `-u`; no PR created unless asked

## 5. When to stop and ask the user

Stop and ask — with concrete options, not open-ended questions — when:

1. **The year you're asked to add already exists** in the workbook.
2. **The workbook contradicts the docs** (F11). Surface the discrepancy before writing anything.
3. **A validation check won't reach zero** after re-diagnosing your own extraction — the source PDF may be internally inconsistent. Show the residual and the rows involved.
4. **Two defensible mappings exist** and no rule in `lesson-rev.md` decides it — e.g., an old single line ("Saldo laba") maps to either of two template rows, or duplicated sheet content makes row placement ambiguous. Present both with consequences.
5. **You'd have to delete or overwrite anything you didn't create this session** — including "obviously wrong" prior data (F12-style artifacts). Report; don't fix silently.
6. **The source PDF looks like a standalone Laporan Keuangan (Q4/interim)**, not an Annual Report — often unaudited; the quarterly pipeline applies.
7. **The auditor's report cannot be located** and totals can't be cross-validated — everything would be `[unverified]`. Ask before proceeding at that confidence level.
8. **A presentation choice is visible to the user** and not covered by §2 conventions (language of a column, adding a new sheet, restructuring anything).
9. **Any outward-facing or hard-to-reverse action**: creating a PR, force-pushing, deleting branches or backups.

Do NOT ask when a convention in this file or a lesson in `lesson-rev.md`
already decides the question. Apply it and note that you did.

## 6. Session-start checklist

1. `git branch --show-current` — confirm you're on the designated branch.
2. Open the workbook; print sheet names, year headers, one sample value per
   main sheet. Reconcile with README (F11).
3. Read `lesson-rev.md` in full. It is short on purpose.
4. Identify which skill file governs the task: new workbook →
   `run1-template-rev`; add a year → `append-year-rev` (with
   `anchor-ledger` for the numbers); verify → `audit`.
5. Back up the workbook before the first write.
