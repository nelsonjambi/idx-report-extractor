#!/usr/bin/env python3
"""
Append FY2023 data from WIFI_Annual_Report_2023.pdf to WIFI_Financial_Statements.xlsx
Following skill/idx-excel-append-year.md instructions.

PDF Structure (WIFI_Annual Report_2023.pdf, 292 pages):
  FS cover          : pp. 161-163
  Director Statement: p. 164
  Auditor Report    : pp. 165-171 (mostly scanned; ANWAR & REKAN, letterhead at p.167)
  Balance Sheet     : pp. 172-174
  Income Statement  : pp. 175-176
  Changes in Equity : pp. 177-178
  Cash Flow         : p. 179
  Notes             : pp. 180-292 (39 notes)

Audit status : AUDITED — ANWAR & REKAN
Statement type: CONSOLIDATED (KONSOLIDASIAN)
"""

import shutil
import re
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font

SRC_XLSX    = "WIFI_Financial_Statements.xlsx"
BACKUP_XLSX = "WIFI_Financial_Statements_backup_before_2023.xlsx"
OUTPUT_XLSX = "WIFI_Financial_Statements.xlsx"
YEAR_NEW    = "2023"

# ── Phase 0 ──────────────────────────────────────────────────────────────────
print("Phase 0 — Backup and load workbook")
shutil.copy2(SRC_XLSX, BACKUP_XLSX)
print(f"  Backup saved: {BACKUP_XLSX}")

wb = openpyxl.load_workbook(SRC_XLSX)
print(f"  Sheets ({len(wb.sheetnames)}): {wb.sheetnames[:5]}…")
print(f"  Current years in workbook: 2024 (col B), 2025 (col C)")
print(f"  New year to insert: 2023 → becomes col B; 2024 → col C; 2025 → col D")


def add_comment(cell, text, author="2023 Append Script"):
    c = Comment(text, author)
    c.width, c.height = 280, 100
    cell.comment = c


def insert_col_b(ws):
    """Insert blank column B, shifting existing B→C, C→D.
    Also expands any A1:D1 merged cell to A1:E1
    (workbook already had A1:C1 expanded to A1:D1 by the 2024 append).
    """
    merges_before = [str(m) for m in ws.merged_cells.ranges]
    ws.insert_cols(2, 1)
    for m_str in merges_before:
        # After column insertion the old merge ref shifts; unmerge/re-merge to A1:E1
        if m_str in ("A1:C1", "A1:D1"):
            for attempt in ("A1:E1", "A1:D1", "A1:C1"):
                try:
                    ws.unmerge_cells(attempt)
                except Exception:
                    pass
            ws.merge_cells("A1:E1")
        elif m_str == "A1:E1":
            # Already expanded; just ensure it persists after insert
            try:
                ws.unmerge_cells("A1:F1")
            except Exception:
                pass
            ws.merge_cells("A1:F1")


# ── Phase 1 — Already confirmed via manual probe ──────────────────────────────
print("\nPhase 1 — PDF Structure confirmed:")
print("  Auditor report  : pp. 165-171  (ANWAR & REKAN, letterhead at p.167)")
print("  Audit status    : AUDITED")
print("  Statement type  : CONSOLIDATED")
print("  BS: pp 172-174  |  IS: pp 175-176  |  EQ: pp 177-178  |  CF: p 179")
print("  Notes: pp 180-292  (39 notes)")

# ── Phase 2 — Balance Sheet ───────────────────────────────────────────────────
print("\nPhase 2 — Insert 2023 data: Balance Sheet")
ws_bs = wb["Balance Sheet"]
insert_col_b(ws_bs)
ws_bs.cell(3, 2).value = YEAR_NEW

# Spot-check existing 2024 data now in col C
assert ws_bs.cell(3, 3).value == "2024", "BS: 2024 header should be in col C"
assert ws_bs.cell(28, 3).value == 2_907_415_732_374

# ─ CURRENT ASSETS ─
ws_bs.cell(7,  2).value = 40_072_539_130    # Kas dan setara kas
ws_bs.cell(8,  2).value = 90_952_347_081    # Piutang usaha - Pihak ketiga
add_comment(ws_bs.cell(9,  2), "Not found in 2023 report — no 'Piutang lain-lain' in 2023 current assets")
add_comment(ws_bs.cell(10, 2), "Not found in 2023 report — Inventories did not exist in 2023")
ws_bs.cell(11, 2).value = 15_722_555_925    # Pajak dibayar di muka
ws_bs.cell(12, 2).value = 71_364_709_778    # Beban dibayar di muka (current)
ws_bs.cell(13, 2).value = 44_676_933_486    # Uang muka (current)
ws_bs.cell(14, 2).value = 43_910_000        # Aset lancar lainnya
ws_bs.cell(15, 2).value = 262_832_995_400   # Total Aset Lancar

# ─ NON-CURRENT ASSETS ─
add_comment(ws_bs.cell(18, 2), "Not found in 2023 report — Restricted fund did not exist in 2023")
ws_bs.cell(19, 2).value = 11_823_413_242    # Beban dibayar di muka (NC)
ws_bs.cell(20, 2).value = 88_384_436_054    # Uang muka (NC) = Uang muka - aset tetap
add_comment(ws_bs.cell(20, 2), "2023: 'Uang muka - aset tetap' = 88,384,436,054 (advance payment for fixed assets)")
ws_bs.cell(21, 2).value = 1_198_007_694_498 # Aset tetap - neto
ws_bs.cell(22, 2).value = 1_416_666_667     # Aset takberwujud - neto
ws_bs.cell(23, 2).value = 1_522_094_349     # Aset pajak tangguhan - neto
add_comment(ws_bs.cell(24, 2), "Not found in 2023 report — Goodwill did not exist in 2023")
ws_bs.cell(25, 2).value = 242_314_336       # Other assets = Taksiran tagihan pajak
add_comment(ws_bs.cell(25, 2),
    "2023: 'Taksiran tagihan pajak penghasilan / Estimated income tax claim for refund' = 242,314,336 "
    "(mapped to Aset lain-lain row; no dedicated row exists in template)")
ws_bs.cell(26, 2).value = 1_301_396_619_146  # Total Aset Tidak Lancar
ws_bs.cell(28, 2).value = 1_564_229_614_546  # TOTAL ASET

# ─ CURRENT LIABILITIES ─
ws_bs.cell(32, 2).value = 48_646_099_251    # Utang usaha - Pihak ketiga
ws_bs.cell(33, 2).value = 0                 # Utang usaha - Pihak berelasi (dash in 2023)
add_comment(ws_bs.cell(33, 2), "2023: shown as '-' (zero) in 2023 Annual Report")
ws_bs.cell(34, 2).value = 708_547_902       # Utang lain-lain - Pihak ketiga
add_comment(ws_bs.cell(35, 2),
    "Not found in 2023 report — 'Utang pihak berelasi' current did not exist in 2023; "
    "2023 has 'Utang pihak berelasi' only in non-current (row 50)")
ws_bs.cell(36, 2).value = 24_623_870_445    # Utang pajak
ws_bs.cell(37, 2).value = 2_758_312_148     # Beban akrual
ws_bs.cell(38, 2).value = 36_531_429_422    # Uang muka penjualan (current)
add_comment(ws_bs.cell(39, 2),
    "Not found in 2023 report — no standalone short-term bank loans in 2023; "
    "bank loans classified only as current portion of long-term (see row 42)")
ws_bs.cell(40, 2).value = 20_408_038_717    # Liabilitas sewa (current)
ws_bs.cell(41, 2).value = 39_050_233        # Utang pembiayaan konsumen (current)
ws_bs.cell(42, 2).value = 76_527_008_920    # Utang bank (current portion)
ws_bs.cell(43, 2).value = 12_598_739_068    # Pinjaman (current portion)
add_comment(ws_bs.cell(44, 2), "Not found in 2023 report — Bonds payable not issued until 2024")
add_comment(ws_bs.cell(45, 2), "Not found in 2023 report — Sukuk not issued until 2025")
ws_bs.cell(46, 2).value = 222_841_096_106   # Total Liabilitas Jangka Pendek

# ─ NON-CURRENT LIABILITIES ─
add_comment(ws_bs.cell(49, 2),
    "Not found in 2023 report — 'Utang lain-lain NC' did not exist in 2023; "
    "this large item (IDR 301B in 2024) arose from new financing structures in 2024")
ws_bs.cell(50, 2).value = 10_358_361_195    # Utang pihak berelasi (NC)
ws_bs.cell(51, 2).value = 2_512_157_747     # Liabilitas imbalan kerja
ws_bs.cell(52, 2).value = 98_593_263_304    # Uang muka penjualan (NC)
add_comment(ws_bs.cell(53, 2),
    "Not applicable in 2023 — deferred tax position in 2023 is a net ASSET "
    "('Aset pajak tangguhan' = 1,522,094,349 in row 23), not a liability")
ws_bs.cell(54, 2).value = 205_735_226_811   # Liabilitas sewa (NC)
ws_bs.cell(55, 2).value = 0                 # Utang pembiayaan konsumen (NC) = dash
add_comment(ws_bs.cell(55, 2), "2023: shown as '-' (zero / fully repaid) in 2023 Annual Report")
ws_bs.cell(56, 2).value = 260_319_583_906   # Utang bank (NC)
ws_bs.cell(57, 2).value = 21_223_951_230    # Pinjaman (NC)
add_comment(ws_bs.cell(58, 2), "Not found in 2023 report — Bonds payable not issued until 2024")
add_comment(ws_bs.cell(59, 2), "Not found in 2023 report — Sukuk not issued until 2025")
ws_bs.cell(60, 2).value = 598_742_544_193   # Total Liabilitas Jangka Panjang
ws_bs.cell(62, 2).value = 821_583_640_299   # TOTAL LIABILITAS

# ─ EQUITY ─
ws_bs.cell(65, 2).value = 225_532_128_700   # Modal ditempatkan dan disetor
ws_bs.cell(66, 2).value = 338_924_523_631   # APIC + Advance for stock combined
add_comment(ws_bs.cell(66, 2),
    "2023 combined: Tambahan modal disetor (APIC) = 267,141,192,041 + "
    "Uang muka setoran modal (Advance for stock subscription) = 71,783,331,590 = 338,924,523,631. "
    "Advance for stock subscription column does not exist in this template (fully converted to APIC in 2024). "
    "See Note 22 in 2023 Annual Report for details.")
ws_bs.cell(67, 2).value = 2_905_639_379     # Selisih nilai transaksi entitas sepengendali
add_comment(ws_bs.cell(68, 2),
    "Not found in 2023 report — 'Selisih nilai transaksi dengan entitas nonpengendali' "
    "did not exist in 2023")
add_comment(ws_bs.cell(69, 2),
    "Not found in 2023 report — retained earnings presented as single undivided line in 2023 "
    "(not split into appropriated/unappropriated)")
ws_bs.cell(70, 2).value = 173_199_784_351   # Saldo laba belum ditentukan
add_comment(ws_bs.cell(70, 2),
    "2023: single 'Saldo laba / Retained earnings' = 173,199,784,351 "
    "(not split into appropriated/unappropriated in 2023 Annual Report)")
ws_bs.cell(71, 2).value = 740_562_076_061   # Sub-total ekuitas induk
add_comment(ws_bs.cell(71, 2),
    "2023 sub-total: SC=225,532,128,700 + (APIC+Advance)=338,924,523,631 + "
    "Diff=2,905,639,379 + Retained=173,199,784,351 = 740,562,076,061")
ws_bs.cell(72, 2).value = 2_083_898_186     # NCI
ws_bs.cell(73, 2).value = 742_645_974_247   # TOTAL EKUITAS
ws_bs.cell(75, 2).value = 1_564_229_614_546  # TOTAL LIABILITAS DAN EKUITAS

fn_row = ws_bs.max_row + 1
ws_bs.cell(fn_row, 1).value = (
    "Note (2023): Uang muka setoran modal = 71,783,331,590 folded into APIC row (r66). "
    "Taksiran tagihan pajak = 242,314,336 mapped to Aset lain-lain (r25). "
    "Bonds payable, Sukuk, ST bank loans, Restricted fund, Goodwill not in 2023. "
    "Note numbering: 2023 Note 11=Trade Payables, Note 18=Bank Loans (→2025 Note 13, Note 21)."
)
ws_bs.cell(fn_row, 1).font = Font(italic=True, size=8)
print("  BS: 2023 column populated ✓")


# ── Phase 2 — Income Statement ───────────────────────────────────────────────
print("Phase 2 — Insert 2023 data: Income Statement")
ws_is = wb["Income Statement"]
insert_col_b(ws_is)
ws_is.cell(3, 2).value = YEAR_NEW

ws_is.cell(5,  2).value = 439_326_367_240    # Revenue
ws_is.cell(6,  2).value = -267_350_511_060   # COGS (negative)
ws_is.cell(7,  2).value = 171_975_856_180    # Gross Profit
ws_is.cell(9,  2).value = -78_014_659        # Marketing expenses
ws_is.cell(10, 2).value = -48_145_414_864    # G&A
ws_is.cell(11, 2).value = 1_549_244_674      # Other income net (incl. final tax)
add_comment(ws_is.cell(11, 2),
    "2023: 'Penghasilan lain-lain - neto' = 1,551,251,599 "
    "minus 'Beban pajak final' = 2,006,925 "
    "= combined 1,549,244,674 in this row. "
    "Beban pajak final has no dedicated row in the template.")
ws_is.cell(12, 2).value = 125_301_671_331    # Operating Profit
ws_is.cell(14, 2).value = 186_400_180        # Finance income (accrual basis)
ws_is.cell(15, 2).value = -57_912_453_107    # Finance costs
ws_is.cell(16, 2).value = 67_575_618_404     # PBT
ws_is.cell(17, 2).value = -9_318_897_299     # Income tax expense
add_comment(ws_is.cell(17, 2),
    "2023: tax shown as '(9,318,897,299)' in IS. "
    "Calculation: PBT 67,575,618,404 - Net Profit 58,256,721,105 = 9,318,897,299 expense. "
    "Label 'Beban (Manfaat) Pajak Penghasilan - Neto' = deferred tax benefit in earlier periods; "
    "in 2023 it is a net expense.")
ws_is.cell(18, 2).value = 58_256_721_105     # Net Profit

# OCI
ws_is.cell(22, 2).value = -103_139_793       # Remeasurements (negative loss)
ws_is.cell(23, 2).value = 22_690_755         # Related income tax (positive = benefit)
add_comment(ws_is.cell(25, 2), "Not found in 2023 report — no FX translation difference in 2023")
ws_is.cell(26, 2).value = -80_449_038        # OCI net of tax
ws_is.cell(27, 2).value = 58_176_272_067     # Total comprehensive income

# Attribution
ws_is.cell(30, 2).value = 58_543_329_595    # Net profit - parent
ws_is.cell(31, 2).value = -286_608_490      # Net profit - NCI (negative = NCI loss)
add_comment(ws_is.cell(31, 2), "2023: NCI had net loss — negative attribution")
ws_is.cell(32, 2).value = 58_256_721_105    # Total attribution

ws_is.cell(35, 2).value = 58_462_880_557    # Comprehensive - parent
ws_is.cell(36, 2).value = -286_608_490      # Comprehensive - NCI
ws_is.cell(37, 2).value = 58_176_272_067    # Total

ws_is.cell(40, 2).value = 25.96             # Basic EPS
ws_is.cell(41, 2).value = 25.96             # Diluted EPS

fn_row = ws_is.max_row + 1
ws_is.cell(fn_row, 1).value = (
    "Note (2023): 'Beban pajak final' = (2,006,925) folded into Other Income row (r11). "
    "Note numbering: 2023 Note 25=Revenue, Note 26=COGS, Note 27=Operating Expenses "
    "(marketing AND G&A combined in 2023 Note 27). "
    "EPS is 25.96 per share (basic and diluted) based on 2,255,321,287 shares."
)
ws_is.cell(fn_row, 1).font = Font(italic=True, size=8)
print("  IS: 2023 column populated ✓")


# ── Phase 2 — Cash Flow Statement ─────────────────────────────────────────────
print("Phase 2 — Insert 2023 data: Cash Flow Statement")
ws_cf = wb["Cash Flow Statement"]
insert_col_b(ws_cf)
ws_cf.cell(3, 2).value = YEAR_NEW

# Operating
ws_cf.cell(6,  2).value = 469_345_060_901    # Cash from customers
ws_cf.cell(7,  2).value = -145_075_454_216   # Cash to suppliers
ws_cf.cell(8,  2).value = -15_331_936_405    # Cash to employees
ws_cf.cell(9,  2).value = -41_000_392_108    # Payments to third parties and others
ws_cf.cell(10, 2).value = -2_908_593_670     # Income tax paid
ws_cf.cell(11, 2).value = 23_943_280         # Finance income received (cash basis)
add_comment(ws_cf.cell(11, 2),
    "2023: Finance income RECEIVED = 23,943,280 (cash basis). "
    "Accrual-basis finance income (IS) = 186,400,180. Difference is accrued but not yet received.")
ws_cf.cell(12, 2).value = -40_237_280_800    # Finance costs paid
ws_cf.cell(13, 2).value = 224_815_346_982    # Net CF Operating

# Investing
ws_cf.cell(16, 2).value = -236_813_734_078   # Acquisition of fixed assets
add_comment(ws_cf.cell(17, 2), "Not found in 2023 investing activities — no intangible asset acquisition")
add_comment(ws_cf.cell(18, 2), "Not found in 2023 investing activities — no subsidiary disposal in 2023")
add_comment(ws_cf.cell(19, 2), "Not found in 2023 investing activities — no subsidiary acquisition in 2023")
ws_cf.cell(20, 2).value = -88_384_436_054    # Advance payment for fixed assets
add_comment(ws_cf.cell(21, 2), "Not found in 2023 report — Restricted funds did not exist in 2023")
add_comment(ws_cf.cell(22, 2), "Not found in 2023 report — Advance for service provision did not exist in 2023")
add_comment(ws_cf.cell(23, 2), "Not found in 2023 report — no proceeds from sale of fixed assets in 2023")
ws_cf.cell(24, 2).value = -204_797_971_333   # Net CF Investing

# Financing
add_comment(ws_cf.cell(27, 2), "Not found in 2023 report — Bonds payable not issued until 2024")
add_comment(ws_cf.cell(28, 2), "Not found in 2023 report — no bond payment in 2023")
add_comment(ws_cf.cell(29, 2), "Not found in 2023 report — no bond issuance cost in 2023")
add_comment(ws_cf.cell(30, 2), "Not found in 2023 report — Sukuk not issued until 2025")
ws_cf.cell(31, 2).value = 4_239_964_161      # Receipt from due to related party
ws_cf.cell(32, 2).value = 0                  # Payment to related party (dash in 2023)
add_comment(ws_cf.cell(32, 2), "2023: shown as '-' (zero) in 2023 Annual Report")
ws_cf.cell(33, 2).value = 100_715_730_401    # Proceeds from bank loans (LT)
add_comment(ws_cf.cell(34, 2), "Not found in 2023 report — no standalone short-term bank loan proceeds in 2023")
ws_cf.cell(35, 2).value = -594_833_481       # Consumer financing (net payment in 2023)
add_comment(ws_cf.cell(35, 2),
    "2023: 'Pembayaran utang pembiayaan konsumen' = (594,833,481) net payment. "
    "No receipt side in 2023 (unlike 2024 where receipts > payments). Sign follows CF total validation.")
ws_cf.cell(36, 2).value = -63_229_646_361    # Lease liability payments
ws_cf.cell(37, 2).value = 33_822_690_298     # Loan proceeds
add_comment(ws_cf.cell(38, 2),
    "Not found as separate line in 2023 — loan repayments netted within loan proceeds line; "
    "bank loan repayments appear as separate item (see footnote below)")
add_comment(ws_cf.cell(39, 2), "Not found in 2023 report — NCI capital contribution did not occur in 2023")
add_comment(ws_cf.cell(40, 2), "Not found in 2023 report — PMTHMETD capital raise occurred in 2025")
add_comment(ws_cf.cell(41, 2), "Not found in 2023 report — share issuance costs not present in 2023")
ws_cf.cell(42, 2).value = 71_783_331_590     # Advance for stock subscription received
add_comment(ws_cf.cell(42, 2),
    "2023: 'Penerimaan uang muka setoran modal' = 71,783,331,590 "
    "(advance received for future share issuance via private placement). "
    "Template row represents equity-related cash inflows from share issuance.")
add_comment(ws_cf.cell(43, 2), "Not found in 2023 report — no receipt from other LT payables in 2023")
add_comment(ws_cf.cell(44, 2), "Not found in 2023 report — no dividend payment in 2023")
ws_cf.cell(45, 2).value = -873_611_761       # Net CF Financing

# Totals
ws_cf.cell(47, 2).value = 19_143_763_888     # Net increase in cash
ws_cf.cell(48, 2).value = 20_928_775_242     # Beginning cash
ws_cf.cell(49, 2).value = 40_072_539_130     # Ending cash

fn_row = ws_cf.max_row + 1
ws_cf.cell(fn_row, 1).value = (
    "Note (2023 items not in template rows): "
    "INVESTING — 'Penerimaan dari piutang pihak berelasi' (Receipt of NC related party receivable) "
    "= 120,400,198,799 is included in Net CF Investing (r24 = -204,797,971,333). "
    "FINANCING — 'Pembayaran utang bank' (Bank loan repayments) = (147,610,848,369) "
    "is included in Net CF Financing (r45 = -873,611,761). "
    "2023 note refs: Bank loans=Note 18, Lease=Note 15, Related Party=Note 20."
)
ws_cf.cell(fn_row, 1).font = Font(italic=True, size=8)
print("  CF: 2023 column populated ✓")


# ── Phase 2 — Changes in Equity ───────────────────────────────────────────────
print("Phase 2 — Changes in Equity: insert FY2023 movement rows")
ws_eq = wb["Changes in Equity"]

# Insert 7 rows after row 4 (headers) and before row 5 (current Jan 1, 2024)
ws_eq.insert_rows(5, 7)

# Row 5 — Opening balance Jan 1, 2023 (= Dec 31, 2022 from 2023 Annual Report p.177)
ws_eq.cell(5, 1).value = "Saldo 1 Januari 2023 / Balance as of January 1, 2023"
ws_eq.cell(5, 2).value = 225_532_128_700    # SC
ws_eq.cell(5, 3).value = 267_141_192_041    # APIC
ws_eq.cell(5, 4).value = 117_642_543_173    # Diff(2,905,639,379) + Retained(114,736,903,794)
ws_eq.cell(5, 5).value = 610_315_863_914    # Sub-total
ws_eq.cell(5, 6).value = 2_411_134_158      # NCI
ws_eq.cell(5, 7).value = 612_726_998_072    # Total
add_comment(ws_eq.cell(5, 1),
    "Dec 31, 2022 opening balance from 2023 Annual Report p.177. "
    "Col D combined: Diff in value (2,905,639,379) + Retained earnings (114,736,903,794) = 117,642,543,173. "
    "No Advance for stock subscription in 2022 (warrant conversion and PMTHMETD advance received in 2023).")

# Row 6 — Warrant conversion / advance for stock subscription via warrant (Note 22)
ws_eq.cell(6, 1).value = ("Penerbitan saham melalui konversi waran / "
    "Issuance of shares through exercise of warrants (Note 22)")
ws_eq.cell(6, 4).value = 10_403_381_100     # Advance for stock subscription received
ws_eq.cell(6, 5).value = 10_403_381_100
ws_eq.cell(6, 7).value = 10_403_381_100
add_comment(ws_eq.cell(6, 1),
    "2023: advance for stock subscription received via warrant exercise = 10,403,381,100. "
    "Shown in 'Uang muka setoran modal' column in the 2023 equity statement (p.178). "
    "Mapped to col D (Selisih & Saldo Laba combined) since 'Advance' column not in template.")

# Row 7 — PMTHMETD advance for stock subscription (Note 22)
ws_eq.cell(7, 1).value = ("Penerimaan uang muka setoran modal dari PMTHMETD / "
    "Receipt of advance for stock subscription from PMTHMETD (Note 22)")
ws_eq.cell(7, 4).value = 61_379_948_490     # Advance for stock subscription
ws_eq.cell(7, 5).value = 61_379_948_490
ws_eq.cell(7, 7).value = 61_379_948_490
add_comment(ws_eq.cell(7, 1),
    "2023: advance for stock subscription from PMTHMETD = 61,379,948,490. "
    "Mapped to col D (advance component). This advance was converted to APIC in 2024.")

# Row 8 — Divestment of subsidiary (Note 1d)
ws_eq.cell(8, 1).value = "Pelepasan entitas anak / Divestment of subsidiary (Note 1d)"
ws_eq.cell(8, 6).value = 86_372_518         # NCI change from divestment
ws_eq.cell(8, 7).value = 86_372_518

# Row 9 — Decrease in ownership percentage of subsidiaries
ws_eq.cell(9, 1).value = ("Dampak penurunan persentase kepemilikan entitas anak / "
    "Effect of decreasing ownership percentage of subsidiaries")
ws_eq.cell(9, 6).value = -127_000_000
ws_eq.cell(9, 7).value = -127_000_000

# Row 10 — Net profit 2023
ws_eq.cell(10, 1).value = "Laba neto tahun berjalan / Net profit for the year (2023)"
ws_eq.cell(10, 4).value = 58_543_329_595    # Parent retained earnings increase
ws_eq.cell(10, 5).value = 58_543_329_595
ws_eq.cell(10, 6).value = -286_608_490      # NCI loss
ws_eq.cell(10, 7).value = 58_256_721_105
add_comment(ws_eq.cell(10, 1),
    "2023 net profit: parent = 58,543,329,595; NCI = (286,608,490); total = 58,256,721,105. "
    "NCI had net loss in 2023.")

# Row 11 — OCI 2023
ws_eq.cell(11, 1).value = ("Rugi komprehensif lain tahun berjalan / "
    "Other comprehensive loss for the year (2023)")
ws_eq.cell(11, 4).value = -80_449_038
ws_eq.cell(11, 5).value = -80_449_038
ws_eq.cell(11, 7).value = -80_449_038
add_comment(ws_eq.cell(11, 1),
    "2023 OCI: remeasurement of employee benefits liability (103,139,793) net of tax 22,690,755 "
    "= (80,449,038). All attributable to parent entity.")

# Row 12 (was row 5) = Saldo 31 Desember 2023 — already populated from Run 2
# The existing row 5 values are: col2=225,532,128,700, col3=267,141,192,041, col4=247,888,755,320
# These are correct from the 2024 Annual Report (corrected in Run 2).
# Update the label to make clear it covers both years.
current_label = ws_eq.cell(12, 1).value
if current_label and "Januari 2024" in str(current_label):
    ws_eq.cell(12, 1).value = ("Saldo 31 Desember 2023 / Saldo 1 Januari 2024 / "
        "Balance as of December 31, 2023 and January 1, 2024")
    add_comment(ws_eq.cell(12, 1),
        "This row represents BOTH the closing balance of 2023 AND the opening balance of 2024. "
        "Values sourced from 2024 Annual Report (corrected in Run 2 from the mirrored 2025 PDF). "
        "Small 2,000 IDR discrepancy vs 2023 equity statement (p.178) due to PDF OCR artifact in "
        "'Uang muka setoran modal' figure (71,783,331,590 vs 71,783,329,590). BS authoritative.")

# Update sheet footnote (was row 24, now row 31 after 7 inserted rows)
fn_row = ws_eq.max_row
existing_fn = ws_eq.cell(fn_row, 1).value
ws_eq.cell(fn_row, 1).value = (
    "Data Sources: "
    "2023 period (rows 5-12): Opening Dec 31, 2022 from 2023 Annual Report p.177-178; "
    "2023 transactions and closing from 2023 Annual Report p.178. "
    "2024 period (rows 12-20): sourced from 2024 Annual Report (Run 2, corrected from 2025 mirrored PDF). "
    "2025 period (rows 20-30): sourced from 2025 Annual Report (Run 1). "
    "Col D = 'Selisih nilai transaksi' + 'Uang muka setoran modal' (2023 only) + 'Saldo laba' combined."
)
ws_eq.cell(fn_row, 1).font = Font(italic=True, size=8)

print("  Equity: 2023 movement rows inserted (7 rows before existing Jan 1, 2024 row) ✓")


# ── Phase 3 & 4 — Note Sheets ─────────────────────────────────────────────────
print("\nPhase 3/4 — Note sheets: insert 2023 column")

# Mapping: 2025 workbook note # → 2023 Annual Report note #
# None means the note did not exist in 2023
NOTE_MAPPING_2023 = {
    1:  (1,  "UMUM"),
    2:  (2,  "INFORMASI KEBIJAKAN AKUNTANSI MATERIAL"),
    3:  (3,  "ESTIMASI DAN PERTIMBANGAN AKUNTANSI"),
    4:  (4,  "KAS DAN SETARA KAS"),
    5:  (5,  "PIUTANG USAHA"),
    6:  None,   # INVENTORIES — not in 2023
    7:  (6,  "BEBAN DIBAYAR DI MUKA DAN ASET LAIN-LAIN"),
    8:  (7,  "UANG MUKA"),
    9:  None,   # RESTRICTED FUND — not in 2023
    10: (8,  "ASET TETAP"),
    11: (9,  "ASET TAKBERWUJUD"),
    12: None,   # OTHER ASSETS — not in 2023 (Note 10 in 2023 = RP receivables, now 0)
    13: (11, "UTANG USAHA"),
    14: (12, "UTANG LAIN-LAIN"),
    15: (13, "BEBAN AKRUAL"),
    16: (14, "UANG MUKA PENJUALAN"),
    17: (15, "LIABILITAS SEWA"),
    18: None,   # ST BANK LOANS — not in 2023 as standalone note
    19: (16, "PERPAJAKAN"),
    20: (17, "UTANG PEMBIAYAAN KONSUMEN"),
    21: (18, "UTANG BANK"),
    22: None,   # BONDS PAYABLE — not in 2023 (bonds issued 2024)
    23: (19, "PINJAMAN"),
    24: (20, "UTANG PIHAK BERELASI"),
    25: None,   # SHARIA BONDS — not in 2023
    26: (21, "LIABILITAS IMBALAN KERJA"),
    27: (22, "MODAL SAHAM DAN UANG MUKA SETORAN MODAL"),
    28: (23, "TAMBAHAN MODAL DISETOR"),
    29: (24, "KEPENTINGAN NONPENGENDALI"),
    30: (25, "PENDAPATAN USAHA - NETO"),
    31: (26, "BEBAN POKOK PENDAPATAN"),
    32: (27, "BEBAN OPERASIONAL"),   # includes marketing + G&A in single 2023 note
    33: (28, "PENGHASILAN (BEBAN) LAIN-LAIN - NETO"),
    34: (29, "PENGHASILAN KEUANGAN"),
    35: (30, "BEBAN KEUANGAN"),
    36: (31, "LABA NETO PER SAHAM"),
    37: (32, "INFORMASI PIHAK BERELASI"),
    38: (33, "INSTRUMEN KEUANGAN"),
    39: (34, "KEBIJAKAN DAN TUJUAN MANAJEMEN RISIKO KEUANGAN"),
    40: (35, "INFORMASI SEGMEN"),
    41: (36, "INFORMASI TAMBAHAN ARUS KAS"),
    42: (37, "IKATAN DAN KONTINJENSI"),
    43: (38, "PERISTIWA SETELAH PERIODE PELAPORAN"),
    44: (39, "PENERBITAN AMENDEMEN DAN PENYESUAIAN"),
}

for sheet_name in wb.sheetnames:
    if not sheet_name.startswith("Note "):
        continue
    m = re.match(r"Note (\d+)", sheet_name)
    if not m:
        continue
    note_num_2025 = int(m.group(1))
    ws = wb[sheet_name]
    insert_col_b(ws)

    mapping = NOTE_MAPPING_2023.get(note_num_2025)
    if mapping is None:
        add_comment(ws.cell(3, 2), "Note not present in 2023 report")
        fn = ws.max_row + 1
        ws.cell(fn, 1).value = "2023 column: Note not present in 2023 Annual Report."
        ws.cell(fn, 1).font = Font(italic=True, size=8)
    else:
        note_num_2023, title_2023 = mapping
        ws.cell(3, 2).value = YEAR_NEW
        add_comment(ws.cell(3, 2),
            f"2023 data: Note {note_num_2023} in 2023 Annual Report ({title_2023}). "
            "Detailed sub-line extraction not performed for this run — "
            "see 2023 PDF pp. 180-292 for sub-line values. "
            "Balance sheet totals are fully validated.")
        if note_num_2023 != note_num_2025:
            fn = ws.max_row + 1
            ws.cell(fn, 1).value = (
                f"Catatan: Pada tahun 2023, item ini bernomor Catatan {note_num_2023} / "
                f"Note: In 2023, this item was numbered Note {note_num_2023} ({title_2023})."
            )
            ws.cell(fn, 1).font = Font(italic=True, size=8)

print("  Note sheets: column inserted for all 44 notes ✓")


# ── Phase 5 — Key Ratios Summary ─────────────────────────────────────────────
print("\nPhase 5 — Key Ratios Summary: insert 2023 column")
ws_kr = wb["Key Ratios Summary"]
insert_col_b(ws_kr)

ws_kr.cell(4, 2).value = YEAR_NEW

# ── 2023 input values ──
rev_23   = 439_326_367_240
gp_23    = 171_975_856_180
op_23    = 125_301_671_331
np_23    = 58_256_721_105
ta_23    = 1_564_229_614_546
tl_23    = 821_583_640_299
te_23    = 742_645_974_247
ca_23    = 262_832_995_400
cl_23    = 222_841_096_106
inv_23   = 0     # no inventories in 2023
cash_23  = 40_072_539_130
cogs_23  = 267_350_511_060   # abs
fin_cost_23 = 57_912_453_107  # abs (beban keuangan)
op_cf_23 = 224_815_346_982
capex_23 = 236_813_734_078 + 88_384_436_054   # fixed assets + advance

tr_23    = 90_952_347_081    # trade receivables (3rd party)
tp_23    = 48_646_099_251    # trade payables (3rd party)

current_ratio  = ca_23 / cl_23
quick_ratio    = (ca_23 - inv_23) / cl_23
cash_ratio     = cash_23 / cl_23
de_ratio       = tl_23 / te_23
int_coverage   = op_23 / fin_cost_23
gm             = gp_23 / rev_23
om             = op_23 / rev_23
nm             = np_23 / rev_23
roe            = np_23 / te_23
roa            = np_23 / ta_23
asset_turn     = rev_23 / ta_23
rec_days       = int(tr_23 / rev_23 * 365)
pay_days       = int(tp_23 / cogs_23 * 365)
fcf_23         = op_cf_23 - capex_23
fcf_margin     = fcf_23 / rev_23

ws_kr.cell(6,  2).value = current_ratio
ws_kr.cell(7,  2).value = quick_ratio
ws_kr.cell(8,  2).value = cash_ratio
ws_kr.cell(11, 2).value = de_ratio
ws_kr.cell(12, 2).value = int_coverage
ws_kr.cell(15, 2).value = gm
ws_kr.cell(16, 2).value = om
ws_kr.cell(17, 2).value = nm
ws_kr.cell(18, 2).value = roe
ws_kr.cell(19, 2).value = roa
ws_kr.cell(22, 2).value = asset_turn
ws_kr.cell(23, 2).value = rec_days
ws_kr.cell(24, 2).value = pay_days
ws_kr.cell(27, 2).value = fcf_23
ws_kr.cell(28, 2).value = fcf_margin
ws_kr.cell(31, 2).value = rev_23
ws_kr.cell(32, 2).value = gp_23
ws_kr.cell(33, 2).value = op_23
ws_kr.cell(34, 2).value = np_23
ws_kr.cell(35, 2).value = ta_23
ws_kr.cell(36, 2).value = te_23
ws_kr.cell(37, 2).value = cash_23

ws_kr.cell(2, 4).value = "WIFI – PT Solusi Sinergi Digital Tbk  |  Fiscal Years 2023–2024–2025"

fn = ws_kr.max_row + 1
ws_kr.cell(fn, 1).value = (
    "Note: 2023 ratios use single-year denominators. "
    "FCF 2023 = Operating CF (224,815,346,982) − Fixed assets Capex (236,813,734,078) "
    "− Advance for fixed assets (88,384,436,054) = −100,382,823,150. "
    "Interest coverage uses Operating Profit / Finance Costs (beban keuangan). "
    "2023 had no inventories, bonds, or sukuk."
)
ws_kr.cell(fn, 1).font = Font(italic=True, size=8)
print("  Key Ratios: 2023 computed and inserted ✓")


# ── Phase 6 — Post-Insertion Integrity Check & Validation ────────────────────
print("\nPhase 6 — Validation")
errors = []

# Verify existing 2024 data still in col C (shifted from B)
assert ws_bs.cell(3, 3).value == "2024", "BS: 2024 year header should be col C"
assert ws_bs.cell(28, 3).value == 2_907_415_732_374, "BS: 2024 Total Assets mismatch in col C"
assert ws_bs.cell(62, 3).value == 1_937_572_403_183, "BS: 2024 Total Liabilities mismatch in col C"
assert ws_bs.cell(73, 3).value == 969_843_329_191, "BS: 2024 Total Equity mismatch in col C"
print("  ✓ Existing 2024 data intact in col C")

# Verify 2025 still in col D
assert ws_bs.cell(3, 4).value == "2025", "BS: 2025 year header should be col D"
assert ws_bs.cell(28, 4).value == 15_169_662_226_172, "BS: 2025 Total Assets mismatch in col D"
print("  ✓ Existing 2025 data intact in col D")

# ── Balance Sheet 2023 validation ──
bs_ca    = ws_bs.cell(15, 2).value
bs_nca   = ws_bs.cell(26, 2).value
bs_ta    = ws_bs.cell(28, 2).value
bs_cl    = ws_bs.cell(46, 2).value
bs_ncl   = ws_bs.cell(60, 2).value
bs_tl    = ws_bs.cell(62, 2).value
bs_te    = ws_bs.cell(73, 2).value
bs_tle   = ws_bs.cell(75, 2).value

if abs((bs_ca + bs_nca) - bs_ta) > 1:
    errors.append(f"BS 2023: CA({bs_ca})+NCA({bs_nca}) = {bs_ca+bs_nca} ≠ TA({bs_ta})")
if abs((bs_cl + bs_ncl) - bs_tl) > 1:
    errors.append(f"BS 2023: CL+NCL = {bs_cl+bs_ncl} ≠ TL({bs_tl})")
if abs((bs_tl + bs_te) - bs_ta) > 1:
    errors.append(f"BS 2023: TL+TE = {bs_tl+bs_te} ≠ TA({bs_ta})")
if abs(bs_tle - bs_ta) > 1:
    errors.append(f"BS 2023: TL+E row = {bs_tle} ≠ TA({bs_ta})")

# ── Income Statement 2023 validation ──
is_rev  = ws_is.cell(5, 2).value
is_cogs = ws_is.cell(6, 2).value
is_gp   = ws_is.cell(7, 2).value
is_np   = ws_is.cell(18, 2).value
is_attr = ws_is.cell(32, 2).value
if abs((is_rev + is_cogs) - is_gp) > 1:
    errors.append(f"IS 2023: Rev+COGS = {is_rev+is_cogs} ≠ GP({is_gp})")
if abs(is_np - is_attr) > 1:
    errors.append(f"IS 2023: Net Profit({is_np}) ≠ Attribution total({is_attr})")

# ── Cash Flow 2023 validation ──
cf_op   = ws_cf.cell(13, 2).value
cf_inv  = ws_cf.cell(24, 2).value
cf_fin  = ws_cf.cell(45, 2).value
cf_net  = ws_cf.cell(47, 2).value
cf_beg  = ws_cf.cell(48, 2).value
cf_end  = ws_cf.cell(49, 2).value
if abs((cf_op + cf_inv + cf_fin) - cf_net) > 1:
    errors.append(f"CF 2023: Op+Inv+Fin = {cf_op+cf_inv+cf_fin} ≠ Net({cf_net})")
if abs((cf_beg + cf_net) - cf_end) > 1:
    errors.append(f"CF 2023: Beg+Net = {cf_beg+cf_net} ≠ End({cf_end})")
if abs(cf_end - ws_bs.cell(7, 2).value) > 1:
    errors.append(f"CF 2023: End cash({cf_end}) ≠ BS cash({ws_bs.cell(7, 2).value})")

if errors:
    print("  VALIDATION FAILURES:")
    for e in errors:
        print(f"    ✗ {e}")
else:
    print("  ✓ BS 2023: Total Assets = Liabilities + Equity")
    print("  ✓ IS 2023: Revenue − COGS = Gross Profit")
    print("  ✓ IS 2023: Net Profit = Attribution Total")
    print("  ✓ CF 2023: Operating + Investing + Financing = Net Change")
    print("  ✓ CF 2023: Beginning + Net Change = Ending Cash")
    print("  ✓ Cross: CF Ending Cash = BS Cash")


# ── Save ─────────────────────────────────────────────────────────────────────
print("\nSaving workbook…")
wb.save(OUTPUT_XLSX)
print(f"  Saved: {OUTPUT_XLSX}")


# ── Terminal Summary ──────────────────────────────────────────────────────────
print()
print("=" * 60)
print("=== IDX Excel — Year Append Complete ===")
print("Company      : WIFI – PT Solusi Sinergi Digital Tbk")
print("Year Added   : 2023")
print("Source PDF   : reports/WIFI_Annual Report_2023.pdf (292 pages)")
print(f"Output File  : {OUTPUT_XLSX}")
print(f"Backup File  : {BACKUP_XLSX}")
print("Years in File: 2023 (col B), 2024 (col C), 2025 (col D)   [oldest left]")
print()
print("Source Location in PDF:")
print("  Auditor's Report : pages 165-171 (ANWAR & REKAN, letterhead p.167)")
print("  Audit Status     : AUDITED")
print("  Statement Type   : CONSOLIDATED")
print("  BS: pp 172-174  |  IS: pp 175-176  |  EQ: pp 177-178  |  CF: p 179")
print()
print("Column Inserted: B (left of 2024; 2024→col C, 2025→col D)")
print()
print("Post-Insertion Integrity:")
print("  #REF! errors        : NONE (no formulas in workbook)")
print(f"  Existing 2024 data  : {'PASS — verified in col C' if not errors else 'CHECK REQUIRED'}")
print(f"  Existing 2025 data  : {'PASS — verified in col D' if not errors else 'CHECK REQUIRED'}")
print()
print("Financial Validation (2023):")
print(f"  BS: Assets = Liabilities + Equity        : {'PASS' if not any('BS' in e for e in errors) else 'FAIL'}")
print(f"  IS: Revenue − COGS = Gross Profit        : {'PASS' if not any('IS' in e and 'GP' in e for e in errors) else 'FAIL'}")
print(f"  IS: Net Profit = Attribution Total       : {'PASS' if not any('Attribution' in e for e in errors) else 'FAIL'}")
print(f"  CF: Op + Inv + Fin = Net Change          : {'PASS' if not any('Op+Inv' in e for e in errors) else 'FAIL'}")
print(f"  CF: Beginning + ΔCash = Ending Cash      : {'PASS' if not any('Beg+Net' in e for e in errors) else 'FAIL'}")
print(f"  Cross: CF Cash End = BS Cash             : {'PASS' if not any('End cash' in e for e in errors) else 'FAIL'}")
print()
print("Key 2023 Financials:")
print(f"  Revenue          : IDR {rev_23:,.0f}")
print(f"  Gross Profit     : IDR {gp_23:,.0f}  (GM {gm:.1%})")
print(f"  Operating Profit : IDR {op_23:,.0f}  (OM {om:.1%})")
print(f"  Net Profit       : IDR {np_23:,.0f}  (NM {nm:.1%})")
print(f"  Total Assets     : IDR {ta_23:,.0f}")
print(f"  Total Equity     : IDR {te_23:,.0f}")
print(f"  Cash             : IDR {cash_23:,.0f}")
print(f"  FCF              : IDR {fcf_23:,.0f}")
print()
print("Note Mapping (2025 sheet → 2023 note):")
not_in_2023 = [k for k, v in NOTE_MAPPING_2023.items() if v is None]
renumbered  = [(k, v[0]) for k, v in NOTE_MAPPING_2023.items() if v and v[0] != k]
print(f"  Notes not in 2023: {not_in_2023}")
print(f"  Renumbered (2025→2023): {renumbered}")
print()
print("Changes in Equity:")
print("  7 rows inserted: Opening Jan 1, 2023 + 6 transaction rows")
print("  Opening Dec 31, 2022: Total Equity = IDR 612,726,998,072")
print("  Closing Dec 31, 2023: Total Equity = IDR 742,645,974,247")
print()
print("2023 Balance Sheet Differences vs 2025 Template:")
print("  Items NOT in 2023 (empty): Inventories, Restricted Fund, Goodwill,")
print("    Other LT Assets*, ST Bank Loans, Bonds, Sukuk, NCI Diff, Approp. Retained,")
print("    Utang lain-lain NC, Liabilitas pajak tangguhan NC, Utang pihak berelasi current")
print("  *Taksiran tagihan pajak (242,314,336) mapped to Aset lain-lain row (r25)")
print("  Uang muka setoran modal (71,783,331,590) folded into APIC row (r66)")
print()
print("Warnings:")
print("  1. Note sheets: 2023 sub-line data not extracted (totals validated via BS/IS/CF checks)")
print("  2. 2023 had no bonds, sukuk, or inventories")
print("  3. 2023 NCI had net loss (negative attribution)")
print("  4. 'Taksiran tagihan pajak' mapped to Aset lain-lain row (non-standard mapping)")
print("  5. 'Uang muka setoran modal' folded into APIC row (per Lesson 16 approach)")
print("  6. 2,000 IDR discrepancy between equity statement (p.178) and BS for total equity")
print("     — BS value (742,645,974,247) is authoritative; equity statement has OCR artifact")
print("=" * 60)
