# Lessons Learned — WIFI FY2025 Extraction Run

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

## Summary Table

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
