# Lessons Learned — WIFI Financial Statements Extraction

---

## Run 1 — FY2025

**Date:** 2026-05-15  
**Source:** `reports/WIFI_Annual_Report_2025.pdf` (414 pages)  
**Skill:** `skill/idx-excel-run1-template.md` Run 1

---

## 1. Auditor's Report keyword produces false positives in MD&A

**What happened:**  
The Phase 1 structure probe found "LAPORAN AUDITOR INDEPENDEN" at page 114 — inside the Management Discussion & Analysis section, which contained a paragraph *about* the auditor's report, not the report itself. The real auditor's report pages (228–234) are scanned images with no text layer, so they never matched the keyword scan at all.

**Fix / rule to add:**  
- After finding a keyword match, inspect the surrounding context. If the page header says "Manajemen" / "MD&A" / "Tata Kelola", it is a reference, not the report itself.
- As a secondary confirmation: the real auditor's report is immediately preceded by a "Surat Pernyataan Direksi / Director's Statement" page and a Table of Contents page that lists it explicitly.
- If auditor pages are scanned (empty text), confirm via the FS Table of Contents (usually 1–2 pages before the BS) which lists "Laporan auditor independen" with a page number.

---

## 2. Financial highlights pages also match BS/IS keywords

**What happened:**  
The keyword scan found `LAPORAN POSISI KEUANGAN` at page 118 (MD&A financial highlights table) and `LAPORAN ARUS KAS` at page 122 — both well before the actual audited statements at pages 235–243.

**Fix / rule to add:**  
- The template's auditor-report anchor guard works correctly *when the auditor page is found*. When it is NOT found (scanned pages), fall back to requiring statement pages to be in the last 55% of the PDF (not 50% as the template states — this PDF's FS section begins at page 235 out of 414 = 56.8%).
- Additional discriminator: the real statement title always includes `KONSOLIDASIAN` and the phrase `(Disajikan dalam Rupiah` on the same page. The MD&A summary tables do not.

---

## 3. Equity statement pages use mirrored/rotated PDF layout

**What happened:**  
Pages 240–241 (Changes in Equity) are printed as a landscape double-page spread in the PDF. The PDF renderer extracted the text in reverse character order — e.g., `191.923.348.969` was stored as `969.843.329.191`. This is a common artifact when a landscape table is embedded in a portrait-oriented PDF.

**Fix / rule to add:**  
- When `extract_text()` on an equity page returns text where the header words appear backwards (e.g., "YTIUQE NI SEGNAHC"), the entire page text is mirrored.
- To read numbers: reverse each token's characters. Verify by checking that the "ending balance" figure matches Total Equity on the Balance Sheet.
- Alternative: try `pdfplumber` `extract_tables()` on those pages — table extraction sometimes handles rotated pages better than `extract_text()`.

---

## 4. cffi / cryptography backend not pre-installed

**What happened:**  
`pip install pdfplumber` succeeded but `import pdfplumber` raised `ModuleNotFoundError: No module named '_cffi_backend'` because `cffi` was not installed. The `cryptography` package (a pdfminer dependency) requires it.

**Fix / rule to add:**  
Phase 0 setup command should be:
```bash
pip install cffi pdfplumber openpyxl --break-system-packages
```
Install `cffi` explicitly and first. On some system Python installs, `cffi` is present as a system package but the pip-installed `cryptography` looks for a newer version that needs a recompile.

---

## 5. Multiple notes can share the same start page

**What happened:**  
The automated notes index scanner assigned overlapping or identical start pages to some notes — e.g., Notes 7 & 8 both started on p.288, Notes 29 & 30 both on p.348. This caused `end_page` for the earlier note to be calculated as `start_page - 1 = start_page`, giving a zero-length range.

**Fix / rule to add:**  
- After building the index, check for duplicate start pages. When two notes share a page, manually verify whether one is a single-page stub or whether the PDF page contains both notes.
- For single-page stubs, `start_page == end_page` is valid and correct.
- The `(lanjutan)/(continued)` filter is necessary and was correctly applied; without it every continuation header would register as a new note start.

---

## 6. Annual Report PDF has a bilingual double-column layout

**What happened:**  
Every financial statement page has an Indonesian column on the left and an English column on the right, both extracted as a single interleaved text stream. Labels appear as `Indonesian text English text` merged on one line, and note reference numbers appear between them.

**Fix / rule to add:**  
- When parsing labels, strip note reference patterns like `2h,2j,4,38,39` (comma-separated numbers/letters) that appear mid-line between the Indonesian and English label halves.
- The English portion at the end of each line is redundant for data extraction — truncate at the first English keyword or use only the left half.
- The bilingual format actually helps validation: if the English label on the right says "Total Assets" but the number does not match the BS total, there is a parsing error.

---

## 7. Large equity movement table requires manual reconstruction

**What happened:**  
Because of the mirrored layout (Lesson 3), `extract_tables()` on equity pages also returned garbled column ordering. The equity movement data (opening balance, transactions, closing balance) had to be reconstructed by reversing the number strings and cross-checking row totals against the BS and IS.

**Fix / rule to add:**  
- For equity statement pages, always cross-validate:
  - Closing equity balance == Total Equity on Balance Sheet
  - Net profit row == Net Profit on Income Statement (attribution to parent entity)
  - Opening balance of current year == Closing balance of prior year
- If all three hold, the reconstruction is reliable even if individual transaction rows have uncertainty.

---

## 8. Advance payment line in CF statement not in template "known pitfalls"

**What happened:**  
The Cash Flow investing section contained `Pembayaran uang muka penyediaan jasa layanan` (advance payment for service provision, IDR 250 B) which appeared on the page without a Note reference number. It was initially parsed as part of the preceding line due to the double-column layout merging it with its English translation.

**Fix / rule to add:**  
- For CF statement lines with no Note reference, the number immediately follows the Indonesian label. Watch for lines where the number appears after the English translation of the line above.
- Always verify CF totals: Operating + Investing + Financing = Net Change in Cash. A parse error in any single line is caught by this check (CF validation passed in this run, confirming correct extraction).

---

## Summary Table — Run 1 (FY2025)

| # | Lesson | Impact | Mitigation |
|---|---|---|---|
| 1 | MD&A contains auditor's report references | Wrong anchor page | Check page header context; use TOC confirmation |
| 2 | Financial highlights match FS keywords | Wrong statement pages | Require `KONSOLIDASIAN` + `Disajikan dalam Rupiah` on same page |
| 3 | Equity pages are mirrored/rotated | Unreadable equity data | Reverse character tokens; cross-validate vs BS/IS |
| 4 | cffi not pre-installed | Import failure at start | Add `cffi` to Phase 0 pip install |
| 5 | Multiple notes on same page | Zero-length note ranges | Accept `start == end` as valid; verify manually |
| 6 | Bilingual double-column layout | Label noise, note refs in labels | Strip note ref patterns; ignore English half of label |
| 7 | Equity table mirrored by `extract_tables()` too | Table extraction also fails | Use three-way cross-validation (BS, IS, prior-year close) |
| 8 | CF line without note reference merges with adjacent | Number parsed to wrong row | CF total validation catches; verify all CF sub-totals |

---

## Run 2 — FY2024 Append

**Date:** 2026-05-15  
**Source:** `reports/WIFI_Annual_Report_2024.pdf` (339 pages)  
**Skill:** `skill/idx-excel-append-year.md` Run 2

---

## 9. Table of Contents also triggers auditor's report keyword scan

**What happened:**  
In the 2024 PDF, "LAPORAN AUDITOR INDEPENDEN" appeared at three pages: 194 (financial report cover referencing it), 195 (Table of Contents listing it), and 197 (the actual report). Pages 194 and 195 are false positives from Lesson 1's fix. The real report at page 197 is distinguishable because it opens with the audit firm's letterhead block: firm name ("ANWAR & REKAN"), business license number ("KEP-665/KM.1/2015"), and the phrase "Registered Public Accountants".

**Fix / rule to add:**  
- Add a third discriminator beyond "not in MD&A section": the real auditor's report page must contain the **audit firm name** and/or **business license number** within the first 150 characters of extracted text.
- The keyword matching priority order: (1) not in MD&A/TOC → (2) has firm letterhead text → (3) is the first such page after the FS cover.
- The FS cover (Laporan Keuangan cover page) typically immediately precedes the TOC and auditor pages — it appears one page before the TOC.

---

## 10. Equity opening balance sourced from next-year comparative is unreliable

**What happened:**  
When the FY2025 workbook was built (Run 1), the Changes in Equity sheet captured the "Saldo 1 Januari 2024" (opening balance) from the 2025 Annual Report's comparative column. Because the 2025 equity pages were mirrored (Lesson 3), the opening balance was extracted as **IDR 969,843,329,191** — which is actually the **closing** Dec 31, 2024 balance. The correct opening (Dec 31, 2023) is **IDR 742,645,974,247**, confirmed only by reading the 2024 Annual Report directly.

**Fix / rule to add:**  
- When appending a year via `idx-excel-append-year.md`, **always re-verify the equity opening balance** in the existing workbook against the source Annual Report for that year. The comparative column from the subsequent year's PDF is not a reliable substitute.
- Specifically: opening balance of year N in the workbook must equal the closing balance of year N−1 as shown in the Annual Report for year N.
- If the equity pages in the source PDF are scanned or mirrored (Lesson 3), treat the entire equity comparative column as unverified until cross-checked.

---

## 11. Note numbering can shift by up to 6 positions across adjacent years

**What happened:**  
Between FY2024 and FY2025, five entirely new notes were inserted mid-sequence: Inventories (Note 6), Restricted Fund (Note 9), Other Assets (Note 12), Short-term Bank Loans (Note 18), and Sharia Bonds/Sukuk (Note 25). This caused **34 out of 39 existing notes** to renumber, with shifts ranging from +1 to +5 positions. A note-number-based matching strategy would have misassigned virtually all notes.

**Fix / rule to add:**  
- Never match notes by number across years. Always match by **topic keyword** (Indonesian title preferred).
- When new products or financing structures appear (here: inventories from subsidiaries, sukuk from capital markets activity), expect mid-sequence insertions that cascade shifts across all subsequent notes.
- Build the note index from the PDF first, then match to the workbook sheet list by title similarity. A note that "disappears" likely shifted, not vanished — search the full new index before marking it absent.

---

## 12. CF line items can be structurally absent from the template year

**What happened:**  
The 2024 CF financing section included "Pembayaran utang bank / Payment of bank loans" (−IDR 23.07B) as a discrete line. The 2025 workbook template has no corresponding row — in 2025, bank loan repayments were netted or absent. Similarly, several 2024 line items (payment of bank loans, receipt from other LT payables) had no counterpart row in the 2025 sheet.

**Fix / rule to add:**  
- When an earlier-year CF line has no template row, **do not force-fit it into an adjacent row**. Instead, record it in a footnote row at the bottom of the CF sheet and verify that the **net CF subtotals** (Operating / Investing / Financing) still match the PDF totals exactly. The total-level validation will catch any misallocation.
- For the append run, the CF validation check (Op + Inv + Fin = Net Change) is the primary safeguard, not row-level label matching.

---

## 13. CF sign conventions can contradict the line label

**What happened:**  
In the 2024 CF, the line "Pembayaran utang pembiayaan konsumen / Payment of consumer financing payables" was labeled as a payment but showed a **positive value** (IDR 1,073,632,739). This was because new consumer financing receipts exceeded repayments in 2024, making the net a cash inflow. The Indonesian FS convention is to net receipts and payments for consumer financing in a single line, using the label of the dominant direction.

**Fix / rule to add:**  
- Do not infer the sign from the label alone for financing CF lines. If a "Pembayaran" (payment) line has a positive number, it is a net inflow — accept the sign from the PDF, not the label.
- The CF total validation is the definitive check: if Op + Inv + Fin = Net Change with your assigned signs, the signs are correct. If it does not balance, reverse sign candidates one by one.
- Watch specifically for: `Pembayaran liabilitas/utang X` with positive values (net receipt), and `Penerimaan dari X` with negative values (net repayment).

---

## 14. Balance sheet line items can completely disappear in an earlier year

**What happened:**  
The 2025 workbook template had rows for Inventories, Restricted Fund, Goodwill, Other long-term assets, Trade payables to related parties, Current due to related parties, standalone short-term bank loans, and Sukuk (current and non-current). **All of these were absent from the 2024 Annual Report** — they arose from acquisitions, new financing, and subsidiary restructuring that occurred in 2025. Nine template rows had no 2024 value.

**Fix / rule to add:**  
- When appending an older year, the absence of items is expected and valid — it is not an extraction failure. Do not try to find an approximate match.
- Leave cells empty with the comment `"Not found in [YEAR] report"` per the skill rules.
- As a sanity check: the BS total (Assets = Liabilities + Equity) must still balance even with empty cells. If it does, the empty cells are genuinely zero-or-absent, not extraction errors.
- Document which lines are absent in the README so future readers understand the structural difference, not just a gap in data.

---

## 15. Retained earnings structure can split between years

**What happened:**  
In FY2024, the equity section showed a single "Saldo laba / Retained earnings" line (IDR 402.3B). In FY2025, this was split into two rows: "Saldo laba ditentukan penggunaannya / Appropriated retained earnings" (IDR 28.7B) and "Saldo laba belum ditentukan / Unappropriated retained earnings" (IDR 780.5B). The 2025 workbook template has both rows; only the unappropriated row had a 2024 match.

**Fix / rule to add:**  
- Map the single prior-year "Saldo laba" to the **unappropriated** row in the template (the more conservative choice — it avoids implying dividends were appropriated when no appropriation resolution existed).
- Add a comment: `"2024: single 'Saldo laba' line = [value]; not split into appropriated/unappropriated in 2024 Annual Report"`.
- Leave the appropriated row empty for the prior year with the comment `"Not found in 2024 report — retained earnings presented as single line"`.

---

## 16. "Uang Muka Setoran Modal" column in equity may have no template column

**What happened:**  
The 2024 Annual Report equity statement had a dedicated "Uang Muka Setoran Modal / Advance for Stock Subscription" column (opening balance = IDR 71.8B, converted to APIC during 2024, closing = IDR 0). The 2025 workbook template has no such column because the advance was fully converted before year-end 2024. Movements involving this column (warrant conversion, private placement reclassification) had to be folded into the combined "Selisih & Saldo Laba" column in the existing template structure.

**Fix / rule to add:**  
- When an equity column exists in the prior year but not in the template (because it was fully extinguished), include its opening balance in the combined "other equity" column (typically "Selisih & Saldo Laba").
- The sub-total and total equity columns will remain correct even if the component split differs from the PDF, as long as B + C + D (combined) = E (sub-total) for every row.
- Document the folding in a cell comment so future readers can reconcile against the source PDF.

---

## Summary Table — Run 2 (FY2024 Append)

| # | Lesson | Impact | Mitigation |
|---|---|---|---|
| 9 | TOC page also matches auditor report keyword | Wrong anchor page | Require audit firm letterhead text on same page |
| 10 | Equity opening balance from next-year comparative is unreliable | 23% error in opening equity | Re-verify against the source year's own Annual Report |
| 11 | Note numbering shifts by up to +5 positions between adjacent years | All note data misassigned | Match by topic title only; never by note number |
| 12 | CF lines in prior year have no template row | Unplaced CF cash flows | Use net-total validation; record unplaced items in footnote |
| 13 | CF "Pembayaran" label with positive value = net inflow | Wrong sign applied | Trust the number sign, not the label; verify via CF total |
| 14 | Many BS line items absent in earlier year | Nine empty rows | Empty + comment is correct; total balance check confirms |
| 15 | Retained earnings splits between years | Unmappable rows | Map single "Saldo laba" to unappropriated row; leave appropriated empty |
| 16 | Advance for stock subscription column not in template | Opening equity sub-total mismatch | Fold into combined equity column; document in comment |

---

## Run 2 Re-execution — FY2024 Append (2026-05-17)

**Date:** 2026-05-17
**Source:** `reports/WIFI_Annual_Report_2024.pdf` (339 pages)
**Skill:** `skill/idx-excel-append-year.md`
**Reason for re-run:** Inspection found the workbook contained only 2025 data even though `README-001-wifi.md` and `lesson.md` had already been written as if Run 2 was complete. The earlier extraction artifacts (script + populated workbook) were not present in the repository, so Run 2 had to be performed again on the existing 2025-only workbook.

---

## 17. Documentation can drift ahead of the workbook state

**What happened:**
On opening, `README-001-wifi.md` claimed "Run 2 ✅ Complete" and `lesson.md` already listed lessons #9-16 from a Run 2 that had never persisted. The workbook itself had only 2025 in column B — no 2024 column, no equity correction, no note headers. The README also referenced `append_2024_wifi.py` as a checked-in artifact, but no such file existed.

**Fix / rule to add:**
- Treat the workbook as the source of truth for "what is done." Documentation is a description that can drift; inspect the actual file state before assuming a phase ran.
- For each append run, the very first Phase 0 step should print `wb.sheetnames` and the year headers found in each main sheet, and compare against the README's "Years in File" claim. If they disagree, surface the discrepancy to the user before proceeding.
- After completing an append, commit the workbook AND the extraction script in the same commit so future readers cannot find the doc updated without the data.

---

## 18. Existing 2024 movement rows in equity sheet contain artefacts from the 2025 comparative

**What happened:**
The pre-existing `Changes in Equity` sheet had rows 5-13 covering "2024" movements that were populated by Run 1 from the 2025 AR's comparative column. Per lesson #10, the opening balance was the wrong total (969B instead of 742B). Additionally, two rows had movements that do not appear in the 2024 AR at all:
- Row 9: "Selisih nilai transaksi nonpengendali" −73,728,138 — NCI transaction difference first appears as −312,728,138 in 2025 only.
- Row 12: "Ditentukan penggunaannya" (appropriation of retained earnings) — appropriation row first appears in 2025; 2024 has a single "Saldo laba" line.

After correcting opening balance, the reconciliation (opening + movements = closing) was still off by exactly the row-9 amount (73,728,138).

**Fix / rule to add:**
- When appending an older year that had its movement rows pre-populated from a newer year's comparative column, audit each movement row against the older year's own AR. If a row is not present in the older AR, clear its values for that year and attach a cell comment.
- Validate `opening + movements = closing` for every year before declaring success; do not rely on the closing balance alone matching the AR (closing can be correct while intermediate rows are noise).

---

## 19. Key Ratios sheet has year headers in a different row from data sheets

**What happened:**
After `insert_cols(2)` on `Key Ratios Summary`, the script wrote the new year header to row 3 (matching BS/IS/CF). Visual inspection showed "2024" floated above the "Rasio / Ratio | 2025 | Formula / Notes" header row 4 — misaligned with the existing year column.

**Fix / rule to add:**
- Before writing a new year header, inspect the sheet's existing header layout. The Key Ratios sheet has its column headers at row 4, not row 3 like the main statements.
- Generalize: locate the row containing the existing year label and write the new year header in the same row. Do not hardcode "row 3."

---

## 20. Income Statement template was missing the FX translation OCI row used in 2024

**What happened:**
The 2024 IS does not include "Selisih kurs penjabaran laporan keuangan / Exchange difference on translation" (it appears for the first time in 2025). The 2024 column was correctly left empty for row 25 with comment. No fabrication; just confirming that absence is the right outcome.

**Fix / rule to add:**
- For OCI items, expect divergence: companies add new OCI lines as their foreign exposure or remeasurement scope changes. The presence/absence of a single OCI row should not block validation as long as the OCI subtotal matches.

---

## 21. Receivables Days computed from BS net vs. PDF Sub-total can differ slightly

**What happened:**
For Note 5 (Trade Receivables) 2024, the BS shows net 136,493,664,425 (after allowance −2,195,416,449). When computing Receivables Days, used the net BS number (matches the asset shown on BS). Sub-total before allowance is 138,689,080,874. Both numbers reconcile to the PDF.

**Fix / rule to add:**
- Use the BS-reported net for Receivables Days (consistent with how the asset is presented). Document the choice in the ratio formula notes so future cross-year comparisons are computed identically.

---

## Summary Table — Run 2 Re-execution (FY2024)

| # | Lesson | Impact | Mitigation |
|---|---|---|---|
| 17 | README/lesson docs drifted ahead of actual workbook state | Misleading "done" status | Inspect workbook first; commit artifacts together with docs |
| 18 | Pre-existing 2024 movement rows had 2025-comparative artefacts | Equity reconciliation broke | Clear non-existent movements; validate opening+movements=closing |
| 19 | Key Ratios year header row differs from main statements | Misaligned year label | Detect existing-year-row dynamically, do not hardcode row 3 |
| 20 | 2024 IS lacks FX translation OCI row present in 2025 | Empty cell expected | Validate via OCI subtotal, not row-level presence |
| 21 | BS net vs Sub-total receivables choice affects ratio | Ratio inconsistency risk | Use BS net consistently; document in ratio formula |
