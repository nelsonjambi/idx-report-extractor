# Lessons (Revised) — IDX FS Extraction

This file collects the lessons that fell out of the 2026-05-18 "Pass 2"
re-extraction work plus the prior lessons that proved genuinely load-bearing
across multiple runs. Lessons from `lesson.md` that were narrow one-off
debugging notes are deliberately omitted; everything below is something a
future run is **likely to hit again**.

The first block is what this task itself surfaced. The second block is the
carefully-vetted carryover from `lesson.md`.

---

## Part 1 — New lessons from this task (2026-05-18)

### L1. Run 1 can clone the same rows into adjacent note sheets

**What happened.** The starting workbook had identical content across
`Note 31` (COGS) and `Note 32` (G&A) in rows 17–35, and identical content
across `Note 33` / `Note 34` / `Note 35` / `Note 36` (Other Income /
Finance Income / Finance Costs / EPS) in rows 17–53. The bug was created by
Run 1 in 2026-05-15 — its note extractor wrote the same parsed rows into
multiple sheets because the underlying note headers all appeared on
adjacent pages. The 2025 column was already populated when Pass 2 started,
so we could not fix it without risking real data; we could only constrain
what the 2024 column got filled with.

**Rule for the future.**
- Every note sheet must carry a `TOPIC_ROW_RANGES` whitelist — a tuple of
  `(low_row, high_row)` describing which rows truly belong to the sheet's
  topic. The append run only fills cells inside the whitelisted ranges in
  the new year's column.
- The whitelist is per-workbook (it depends on Run 1's row layout). If a
  whitelist row was filled in the 2025 column but should not have been,
  flag it during the append run rather than silently overwrite.
- The cleanest place to record the whitelist is at the top of the
  append-year script, next to the `NOTE_MATCH` table. Comment each entry
  with what it represents.

### L2. PDF text inserts spaces inside numbers and at word starts

**What happened.** The 2024 AR's bilingual two-column layout produced
fragmented numeric tokens like `60.0 29.971.450 140.8 96.788.552` (two
numbers split by stray spaces, then joined to the next one's leading
digits) and label tokens like `P eriklanan`, `T otal`, `S aldo`. A naive
`\d+(?:\.\d+)+` regex silently dropped most numeric rows on these pages.

**Rule for the future.** Run a `normalize_numbers(text)` pre-pass against
every page of extracted PDF text before tokenising. The pattern is:

```python
prev = None
out = text
while out != prev:
    prev = out
    # LEFT ends with .DD (2-digit partial group), RIGHT is 1 digit → merge
    out = re.sub(r'(\d\.\d{2})\s+(\d)(?!\d)', r'\1\2', out)
    # LEFT ends with .D (1-digit partial group), RIGHT is 2 digits → merge
    out = re.sub(r'(\d\.\d)\s+(\d\d)(?!\d)', r'\1\2', out)
# Leading split: "2 58.619.980" → "258.619.980"
out = re.sub(r'(?<![\w.])(\d{1,2})\s+(\d{1,2}\.\d{3}(?:\.\d{3})*)\b', r'\1\2', out)
```

The `while` loop is required — fragments compound. Validate the
normaliser by spot-checking that every note's 2024 total now parses to a
single int and matches the BS/IS cross-reference.

For label fragments (`P eriklanan`), do not try to "un-fragment" — instead
strip tokens of length ≤ 2 when computing keyword overlap. The first
letter sometimes detaches; the rest of the word is recognisable.

### L3. Some workbook labels are too generic to auto-match

**What happened.** Pass 2's matcher couldn't reliably assign 2024 values
to rows whose Column-A label is just `Total Total`, `Sub-total Sub-total`,
`Pasal Article`, `(Catatan ) (Note )`, etc. There are many PDF lines that
satisfy the keyword overlap, so any of them might be picked, and the
wrong-row assignment cascades.

**Rule for the future.** Two-layer extraction:

1. **Hardcoded layer first.** Maintain a `NOTE_2024_DATA` (or
   `NOTE_<YEAR>_DATA`) dict keyed by `(sheet_name, row_num)` of values
   that have been manually verified against the source PDF and the
   BS/IS/CF cross-reference. Apply this layer *before* any auto-matching.
2. **Systematic layer second.** Run the keyword-overlap auto-matcher only
   against rows whose Column-A label is NOT in a `_GENERIC_LABELS` set
   (`'total'`, `'total total'`, `'sub total'`, `'neto net'`, `'pasal article'`,
   `'catatan note'`, etc.) AND whose B-column cell is still empty.

The hardcoded layer is also where sign-convention overrides live — e.g.,
the 2024 PDF presents `Pajak kini` as a positive amount in the tax
reconciliation, but the workbook follows the IS convention where it is
negative. Encode the sign override in the hardcoded dict; do not try to
infer sign systematically.

### L4. Multi-note pages need a section-header filter, not a page range

**What happened.** Several 2024 AR pages contain 2–4 notes each (e.g.,
page 288 has Notes 26 COGS + 27 G&A + 28 Other Income + 29 Finance
Income). A page-range extractor for Note 27 also pulled in rows from
Notes 26 and 28, which then matched to wrong workbook rows or were
appended as bogus 2024-only items.

**Rule for the future.** During note-content extraction, detect the note
header line via `re.match(r'^\s*(\d{1,2})\.\s+[A-Z]{3,}', line)` (and
exclude lines containing `lanjutan`/`continued`). Track which note's
section we're inside, and pass a `target_note=N` parameter so only lines
that fall inside Note N's section are emitted. With this filter, the
declared page ranges in metadata can safely overlap between adjacent
notes — the section filter does the real boundary work.

### L5. Short-label keyword overlap needs an asymmetric denominator

**What happened.** Many fixed-asset rows have 1–2 keyword labels
(`Tanah / Land`, `Bangunan / Buildings`). With a classic Jaccard
`intersection / max(|a|, |b|)`, matching the workbook label
`tanah land` against the PDF's `tanah` alone scored only 0.5, below the
0.6 threshold, so Tanah was skipped.

**Rule for the future.** Use a hybrid denominator: when the smaller side
has ≤2 significant keywords, score is `intersection / min(|a|, |b|)`;
otherwise the classic `max`. This rewards short-label coverage without
turning into a free-for-all for longer labels.

```python
def keyword_overlap(a, b):
    wa = {w for w in a.split() if w not in STOP and len(w) > 2}
    wb = {w for w in b.split() if w not in STOP and len(w) > 2}
    if not wa or not wb: return 0.0
    n = len(wa & wb)
    if n == 0: return 0.0
    return n / (min(len(wa), len(wb)) if min(len(wa), len(wb)) <= 2 else max(len(wa), len(wb)))
```

### L6. The append script needs a magnitude sanity check

**What happened.** Some labels matched against PDF lines whose only
real number was a note reference (e.g., `(Catatan 14)` extracted as
14). The 14 got written into a cell whose 2025 column had ~277B, an
obvious 10⁹× discrepancy.

**Rule for the future.** When matching cell B against PDF row `r`, if
`abs(C-column value) > 1,000,000` and `abs(r.value) < abs(C) / 1000` (or
vice versa), reject the match. This is a cheap guard against single-digit
notes references slipping through the keyword filter.

Also, filter PDF rows during extraction to drop lines whose only number
is in the noise band: single-digit integers, integers in 1990–2100
(years), decimal floats < 1 (rates). Keep these only if the line has a
second, larger number that is the real value.

### L7. Append unmatched PDF rows but dedup against existing values

**What happened.** After matching, Pass 2 appended each unconsumed PDF
row to the bottom of the sheet under a `Catatan: Baris berikut hanya
muncul di Laporan Tahunan 2024` footnote. Early iterations duplicated
rows that had matched in some earlier sheet — e.g., a customer name with
the same balance appeared twice.

**Rule for the future.** Build a `set` of all values already present in
column B (across all rows of the sheet) before the append step. When
emitting a candidate 2024-only row, skip it if its value is already in
that set. Also enforce: the candidate's normalised label has ≥1
significant keyword (after stripping stop words), the value is not in
the noise band (|v| ≥ 10,000 for ints, no decimal floats unless they're
rates that the script expects), and the label is not in `_GENERIC_LABELS`.

### L8. Note Index column language must match across years

**What happened.** Run 2's first pass put Indonesian titles in the 2024
column of the Note Index and English titles in the 2025 column. Visual
inspection of the workbook flagged the asymmetry immediately — readers
can't compare across columns when the language flips.

**Rule for the future.** The Note Index sheet uses **English titles in
every year column**, with the original Indonesian title attached as a
cell comment (`Indonesian title: <ID>`). If a renumbered topic has a
different note number in the older year, add that to the comment too
(`In 2024 this topic was Note 8.`). Drop the Indonesian-title-in-cell
pattern from the skill.

### L9. Tax-line sign convention diverges from PDF presentation

**What happened.** The 2024 AR's tax reconciliation table shows
`Pajak kini Total current tax 38.286.594.180` as a positive number. The
workbook's IS column for 2025 has the same row at −72,810,126,002 (a
negative expense). Filling the 2024 cell with +38B would create a
sign-inconsistent column.

**Rule for the future.** Tax-related rows (current tax, deferred tax,
total income tax) almost always need sign flipping when the source is the
note's reconciliation page rather than the IS. Encode these as hardcoded
overrides in `NOTE_<YEAR>_DATA` with a comment explaining the flip. Do
the same for any expense subtotal whose PDF presentation is positive but
whose IS-column convention is negative (`Finance costs`, certain CF
items per lesson L13 below).

### L10. Documentation drift is a real failure mode

**What happened.** When this task started, `README-001-wifi.md` claimed
"Run 2 ✅ Complete" but the workbook had only 2025 in column B. The
README had been written ahead of the actual data; the data state was the
ground truth.

**Rule for the future.** Treat the workbook as the source of truth for
"what is done." At the top of every append run, print `wb.sheetnames`,
the year headers, and a one-line sample value per sheet, and compare to
the README's claimed state. If they disagree, surface to the user
*before* writing anything new. Commit workbook + script + README in the
same commit so reviewers can never see a doc update without the
underlying data.

---

## Part 2 — Carryover lessons that earn their keep

The lessons below originate from earlier runs but are strong enough that
every future append should keep them in mind.

### C1. Bilingual two-column layout (from `lesson.md` #6)

Every audited FS page renders Indonesian on the left and English on the
right, both extracted as a single interleaved text stream. Labels appear
as `Indonesian text English text` joined on one line, with note
reference codes (`2h,2j,4,38,39`) sometimes between them. When parsing:
strip the note-reference patterns; use the Indonesian half for matching
(it's more stable than the English half across years).

### C2. Equity opening balance from a next-year comparative is unreliable (from `lesson.md` #10)

When Year N is added later via append, the workbook's Year (N+1)
comparative column for equity was likely extracted from the Year (N+1)
PDF — that column is mirrored/rotated in some PDFs (see C3) and
frequently wrong. On every append, **re-verify** the equity opening
balance for the year being added against the year-N AR itself, not the
year-(N+1) comparative. Validate `opening + movements = closing` for
every row before declaring success.

### C3. Equity pages can be mirrored or rotated (from `lesson.md` #3)

Landscape equity tables embedded in portrait PDFs often come out
character-reversed: `191.923.348.969` shows as `969.843.329.191`. If the
extracted header words read backwards (`YTIUQE NI SEGNAHC`), the whole
page is mirrored. Reverse each token, then cross-validate the ending
balance against Total Equity on the Balance Sheet — if they match, the
reconstruction is reliable.

### C4. Note numbering shifts up to ±6 positions across adjacent years (from `lesson.md` #11)

Between two adjacent fiscal years, new notes can be inserted mid-sequence
(Inventories, Restricted Fund, Sukuk are common 2024→2025 insertions).
This renumbers most subsequent notes. **Match by topic title, never by
note number.** Build the new year's note index from the new PDF first,
then match each new note to the workbook sheet list by Indonesian-title
similarity. A note that "disappears" likely shifted, not vanished —
search the full new index before marking it absent.

### C5. CF sign convention can contradict the line label (from `lesson.md` #13)

In the CF financing section, a line labelled `Pembayaran utang
pembiayaan konsumen / Payment of consumer financing payables` can have a
positive value when 2024 receipts exceeded repayments. Trust the PDF
sign, not the label. The definitive check is `Op + Inv + Fin = ΔCash`
— if it balances with your signs, the signs are correct.

### C6. "Uang Muka Setoran Modal" column can be absent in the template (from `lesson.md` #16)

When an equity column existed in an earlier year but was fully converted
during that year (UMSM → APIC), the template year (which lacks the
column) won't have a place for it. Fold the opening balance and the
movements involving that column into the combined "Selisih & Saldo
Laba" column in the existing equity template. Document the folding in
cell comments so the column composition can be reconciled against the
source PDF.

### C7. Pre-existing rows can contain stale comparative artifacts (from `lesson.md` #18)

When a previous append populated movement rows from the wrong year's
comparative column, those values can survive into the new run if they
land in rows the new run doesn't touch. On every append, audit each
movement row against the year-being-added's own AR; if a row is not
present in that AR, clear its value and attach a "not present in [YEAR]
report" comment. Don't assume that "the row already has data, so I'll
leave it alone."

---

## Process checklist (use before every append run)

1. Print `wb.sheetnames`, the year headers in each main sheet, and the
   first three data rows of BS/IS/CF. Compare against the README's
   "Years in File" claim. Flag any mismatch.
2. Back up the workbook to `_backup_before_<YEAR>.xlsx` BEFORE any
   writes.
3. Run the structure probe on the new PDF; print page ranges for BS, IS,
   CF, Equity, and each note's start page.
4. Verify the equity opening balance for the year being added against
   the year-being-added's own AR.
5. Build the note title index from the new PDF; match by topic to the
   workbook sheet list.
6. Apply hardcoded verified totals first; then systematic auto-fill;
   then append unmatched PDF rows as 2024-only items.
7. Run the cross-validation suite (BS identity, IS subtotals, CF
   reconciliation, CF↔BS cash agreement). Save only if all pass.
8. Commit workbook + script + docs in one commit.
