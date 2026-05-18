"""
Append WIFI FY2023 data to existing WIFI_Financial_Statements.xlsx workbook.
Implements idx-excel-append-year-rev.md skill.

Insert position: LEFT (2023 is older than existing 2024 and 2025).
"""
import re, sys, shutil
from copy import copy
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, Alignment, PatternFill
import pdfplumber

XLSX = 'WIFI_Financial_Statements.xlsx'
BACKUP = 'WIFI_Financial_Statements_backup_before_2023.xlsx'
PDF_PATH = 'reports/WIFI_Annual Report_2023.pdf'
YEAR = 2023
NEW_HDR = '2023'
EXISTING_YEARS_HDRS = ['2024', '2025']
AUTHOR = 'Append2023'

# ----- helpers -----
def normalize_numbers(text):
    prev = None
    out = text
    while out != prev:
        prev = out
        out = re.sub(r'(\d\.\d{2})\s+(\d)(?!\d)', r'\1\2', out)
        out = re.sub(r'(\d\.\d)\s+(\d\d)(?!\d)', r'\1\2', out)
    out = re.sub(r'(?<![\w.])(\d{1,2})\s+(\d{1,2}\.\d{3}(?:\.\d{3})*)\b', r'\1\2', out)
    return out

# ----- HARDCODED 2023 VALUES (verified against PDF and BS/IS/CF cross-reconciliation) -----

# Balance Sheet (workbook row -> 2023 value). None means "not in 2023 report".
BS_2023 = {
    7:  40072539130,        # Kas dan setara kas (matches CF Beginning Cash 2024)
    8:  90952347081,        # Piutang usaha Pihak ketiga (net of allowance)
    9:  None,               # Piutang lain-lain - not in 2023
    10: None,               # Persediaan - not in 2023
    11: 15722555925,        # Pajak dibayar di muka
    12: 71364709778,        # Beban dibayar di muka (current)
    13: 44676933486,        # Uang muka (current)
    14: 43910000,           # Aset lancar lainnya
    15: 262832995400,       # Total Aset Lancar
    18: None,               # Restricted fund - not in 2023
    19: 11823413242,        # Beban dibayar di muka (non-current)
    20: 88384436054,        # Uang muka aset tetap (non-current)
    21: 1198007694498,      # Aset tetap - neto
    22: 1416666667,         # Aset takberwujud - neto
    23: 1522094349,         # Aset pajak tangguhan - neto
    24: None,               # Goodwill - not in 2023
    25: None,               # Aset lain-lain (non-current) - covered by row 79
    26: 1301396619146,      # Total Aset Tidak Lancar
    28: 1564229614546,      # TOTAL ASET
    32: 48646099251,        # Utang usaha Pihak ketiga
    33: 0,                  # Utang usaha Pihak berelasi (PDF shows '-')
    34: 708547902,          # Utang lain-lain Pihak ketiga
    35: None,               # Utang pihak berelasi (current) - 2023 has it only in non-current
    36: 24623870445,        # Utang pajak
    37: 2758312148,         # Beban akrual
    38: 36531429422,        # Uang muka penjualan (current)
    39: None,               # Utang bank jangka pendek - not in 2023
    40: 20408038717,        # Liabilitas sewa (current)
    41: 39050233,           # Utang pembiayaan konsumen (current)
    42: 76527008920,        # Utang bank (current)
    43: 12598739068,        # Pinjaman (current)
    44: None,               # Utang obligasi (current) - not in 2023
    45: None,               # Sukuk (current) - not in 2023
    46: 222841096106,       # Total Liab Jangka Pendek
    49: None,               # Utang lain-lain Pihak ketiga (non-current) - not in 2023
    50: 10358361195,        # Utang pihak berelasi (non-current)
    51: 2512157747,         # Liabilitas imbalan kerja
    52: 98593263304,        # Uang muka penjualan (non-current)
    53: None,               # Liab pajak tangguhan - 2023 has deferred tax asset only
    54: 205735226811,       # Liabilitas sewa (non-current)
    55: 0,                  # Utang pembiayaan konsumen (non-current) - '-' in 2023
    56: 260319583906,       # Utang bank (non-current)
    57: 21223951230,        # Pinjaman (non-current)
    58: None,               # Utang obligasi (non-current) - not in 2023
    59: None,               # Sukuk (non-current) - not in 2023
    60: 598742544193,       # Total Liab Jangka Panjang
    62: 821583640299,       # TOTAL LIAB
    65: 225532128700,       # Modal ditempatkan
    66: 267141192041,       # Tambahan modal disetor
    67: 2905639379,         # Selisih nilai trans entitas sepengendali
    68: None,               # Selisih nilai trans nonpengendali - not in 2023
    69: None,               # Saldo laba ditentukan - not separated in 2023
    70: 173199784351,       # Saldo laba (combined)
    71: 740562076061,       # Sub-total equity parent
    72: 2083898186,         # NCI
    73: 742645974247,       # TOTAL EKUITAS
    75: 1564229614546,      # TOTAL LIAB DAN EKUITAS
    78: None,               # Uang muka pembelian aset tetap - reported in row 20
    79: 242314336,          # Taksiran tagihan pajak penghasilan
    80: 71783331590,        # Uang muka setoran modal (equity)
}

# Income Statement (workbook row -> 2023 value)
IS_2023 = {
    5:  439326367240,       # Pendapatan
    6:  -267350511060,      # COGS
    7:  171975856180,       # Laba Bruto
    9:  -78014659,          # Beban pemasaran
    10: -48145414864,       # G&A
    11: 1551251599,         # Other income - net
    12: 125301671331,       # Laba Usaha
    14: 186400180,          # Penghasilan keuangan
    15: -57912453107,       # Biaya keuangan
    16: 67575618404,        # Laba sebelum pajak
    17: -9318897299,        # Beban pajak penghasilan - neto
    18: 58256721105,        # Laba Neto tahun berjalan
    22: -103139793,         # Pengukuran kembali liabilitas imbalan kerja
    23: 22690755,           # Pajak terkait
    25: None,               # Selisih kurs - not in 2023
    26: -80449038,          # OCI net of tax
    27: 58176272067,        # Total comprehensive income
    30: 58543329595,        # Net profit attributable to parent
    31: -286608490,         # Net profit attributable to NCI
    32: 58256721105,        # Total
    35: 58462880557,        # Comp income attrib to parent
    36: -286608490,         # Comp income attrib to NCI
    37: 58176272067,        # Total comp
    40: 25.96,              # Basic EPS
    41: 25.96,              # Diluted EPS
    44: -2006925,           # Beban pajak final (separate line in 2023)
}

# Cash Flow (workbook row -> 2023 value)
CF_2023 = {
    6:  469345060901,       # Cash from customers
    7:  -145075454216,      # Cash to suppliers
    8:  -15331936405,       # Cash to employees
    9:  -41000392108,       # Payments to third parties etc
    10: -2908593670,        # Tax payments
    11: 23943280,           # Finance income received
    12: -40237280800,       # Finance costs paid
    13: 224815346982,       # Net Cash from Operating
    16: -236813734078,      # Acquisitions of fixed assets
    17: None,               # Acquisitions of intangible - not in 2023
    18: 0,                  # Proceeds from disposal of subsidiary (- in 2023)
    19: None,               # Acquisition of subsidiary - not in 2023
    20: -88384436054,       # Advance payment of fixed assets
    21: None,               # Restricted funds - not in 2023
    22: None,               # Advance for service provision - not in 2023
    23: 0,                  # Proceeds from sale of fixed assets (- in 2023)
    24: -204797971333,      # Net Cash Used in Investing
    27: None,               # Proceeds from bond payable - not in 2023
    28: None,               # Bond payment - not in 2023
    29: None,               # Bond issuance cost - not in 2023
    30: None,               # Proceeds from sharia bonds - not in 2023
    31: 4239964161,         # Receipt from due to related parties
    32: 0,                  # Payment of due to related party (- in 2023)
    33: 100715730401,       # Proceeds from bank loans
    34: None,               # Proceeds from short-term bank loans - not in 2023
    35: None,               # Receipt of consumer financing - not in 2023
    36: -63229646361,       # Payment of lease liabilities
    37: 33822690298,        # Proceeds from loans
    38: None,               # Payment of loans - not in 2023
    39: None,               # NCI capital contribution - not in 2023
    40: 0,                  # PMTHMETD - dash in 2023
    41: None,               # Share issuance costs - not in 2023
    42: 71783331590,        # Receipt from advances for stock subscription
    43: None,               # Receipt from other payables long-term - not in 2023
    44: None,               # Dividend payment - not in 2023
    45: -873611761,         # Net Cash from Financing
    47: 19143763888,        # Net change in cash
    48: 20928775242,        # Beginning cash
    49: 40072539130,        # Ending cash
    52: -147610848369,      # Payment of bank loans (2024-only row; 2023 has this value)
    53: -594833481,         # Payment of consumer financing payables (2023 has payment, 2024 had inflow)
}

# Workbook sheet name -> column B (2023) cells to fill via hardcoded layer.
# Maps (sheet_name, row) -> value.  None means "not in 2023".
# Cross-reference totals must match BS/IS/CF.
NOTE_2023_DATA = {
    ('Note 4 - CASH AND CASH EQUIVALE', 16): 41079845,        # Kas (cash on hand)
    ('Note 4 - CASH AND CASH EQUIVALE', 41): 5031459285,      # Sub-total bank
    ('Note 4 - CASH AND CASH EQUIVALE', 48): 40072539130,     # Total

    ('Note 5 - TRADE RECEIVABLES', 35): 92060761935,          # Sub-total
    ('Note 5 - TRADE RECEIVABLES', 37): -1108414854,          # Allowance
    ('Note 5 - TRADE RECEIVABLES', 38): 90952347081,          # Net
    ('Note 5 - TRADE RECEIVABLES', 40): 73682836780,          # Belum jatuh tempo
    ('Note 5 - TRADE RECEIVABLES', 42): 2580021646,           # <30 days
    ('Note 5 - TRADE RECEIVABLES', 43): 6086549669,           # 31-60
    ('Note 5 - TRADE RECEIVABLES', 44): 2681056265,           # 61-90
    ('Note 5 - TRADE RECEIVABLES', 45): 5084223113,           # 91-120
    ('Note 5 - TRADE RECEIVABLES', 46): 1946074462,           # >120
    ('Note 5 - TRADE RECEIVABLES', 47): 92060761935,          # Sub-total aging
    ('Note 5 - TRADE RECEIVABLES', 49): -1108414854,          # Allowance again
    ('Note 5 - TRADE RECEIVABLES', 50): 90952347081,          # Net again
    ('Note 5 - TRADE RECEIVABLES', 53): 1082041262,           # Beg balance allowance
    ('Note 5 - TRADE RECEIVABLES', 56): 26373592,             # Provision during year
    ('Note 5 - TRADE RECEIVABLES', 58): 1108414854,           # End balance allowance

    # Note 6 - workbook is INVENTORIES which doesn't exist in 2023; leave blank

    ('Note 7 - PREPAID EXPENSES AND O', 0): None,             # Placeholder - many rows to fill

    # Note 10 (Fixed Assets) -- key NBV
    ('Note 10 - FIXED ASSETS', 0): None,                       # Use auto-fill mostly

    # Note 13 - Trade Payables (2023 note 11)
    # Sub-total/total = 48,646,099,251

    # Note 14 - Other Payables (2023 note 12) - single value 708,547,902
    # Note 15 - Accrued Expenses (2023 note 13) - 2,758,312,148

    # Note 17 - Lease Liabilities (2023 note 15)
    # Note 19 - Taxation (2023 note 16) - with sign flips
    ('Note 19 - TAXATION', 18): 1592818184,                   # Perusahaan PPN
    ('Note 19 - TAXATION', 20): 14092587741,                  # Anak PPN
    ('Note 19 - TAXATION', 22): 1186000000,                   # Hmm
    ('Note 19 - TAXATION', 23): 15722555925,                  # Total prepaid taxes
    ('Note 19 - TAXATION', 46): 24623870445,                  # Total tax payables
    ('Note 19 - TAXATION', 50): -679452400,                   # Perusahaan current (sign-flipped)
    ('Note 19 - TAXATION', 51): -4035370020,                  # Anak current
    ('Note 19 - TAXATION', 52): -4714822420,                  # Total current
    ('Note 19 - TAXATION', 54): 168457949,                    # Perusahaan deferred (sign-flipped; PDF showed negative = benefit, IS convention positive)
    ('Note 19 - TAXATION', 55): -4772532828,                  # Anak deferred (sign-flipped)
    ('Note 19 - TAXATION', 56): -4604074879,                  # Total deferred
    ('Note 19 - TAXATION', 58): -9318897299,                  # Net income tax expense (matches IS)

    # Note 21 - Long-term bank loans (2023 note 18) - sub-total
    # Note 26 - Employee benefits (2023 note 21)
    # Note 27 - Share Capital (2023 note 22)
    ('Note 27 - SHARE CAPITAL AND ADV', 0): None,

    # Note 28 - APIC (2023 note 23)
    # Total APIC in 2023 = 270,046,831,420 but workbook may store as 267,141,192,041 (since uang muka setoran modal is separate)
    # The PDF table includes Selisih nilai transaksi (2,905,639,379) and Pengampunan pajak (3,125,956,639)

    # Note 29 - NCI (2023 note 24) - Total NCI = 2,083,898,186

    # Note 30 - Revenues (2023 note 25)
    ('Note 30 - REVENUES - NET', 0): None,
    # Note 31 - COGS (2023 note 26)
    # Note 32 - G&A (2023 note 27 has marketing+G&A)
    # Note 34 - Finance income (2023 note 29) - 186,400,180
    # Note 35 - Finance costs (2023 note 30) - -57,912,453,107
    # Note 36 - EPS (2023 note 31) - 25.96
}

# Notes that have a corresponding 2023 note (workbook sheet -> 2023 note number)
NOTE_MAPPING = {
    'Note 1 - GENERAL': 1,
    'Note 2 - SIGNIFICANT ACCOUNTING': 2,
    'Note 3 - ACCOUNTING ESTIMATES A': 3,
    'Note 4 - CASH AND CASH EQUIVALE': 4,
    'Note 5 - TRADE RECEIVABLES': 5,
    'Note 6 - INVENTORIES': None,           # Not in 2023
    'Note 7 - PREPAID EXPENSES AND O': 6,
    'Note 8 - ADVANCES': 7,
    'Note 9 - RESTRICTED FUND': None,       # Not in 2023
    'Note 10 - FIXED ASSETS': 8,
    'Note 11 - INTANGIBLE ASSETS': 9,
    'Note 12 - OTHER ASSETS': None,         # Not in 2023 (2023 had DUE FROM RELATED PARTY at #10, different topic)
    'Note 13 - TRADE PAYABLES': 11,
    'Note 14 - OTHER PAYABLES': 12,
    'Note 15 - ACCRUED EXPENSES': 13,
    'Note 16 - ADVANCES FROM CUSTOME': 14,
    'Note 17 - LEASE LIABILITIES': 15,
    'Note 18 - SHORT-TERM BANK LOANS': None,  # Not in 2023
    'Note 19 - TAXATION': 16,
    'Note 20 - CONSUMER FINANCING PA': 17,
    'Note 21 - LONG-TERM BANK LOANS': 18,
    'Note 22 - BONDS PAYABLE': None,        # Not in 2023
    'Note 23 - LOANS': 19,
    'Note 24 - DUE TO RELATED PARTIE': 20,
    'Note 25 - SHARIA BONDS': None,         # Not in 2023
    'Note 26 - EMPLOYEE BENEFITS LIA': 21,
    'Note 27 - SHARE CAPITAL AND ADV': 22,
    'Note 28 - ADDITIONAL PAID-IN CA': 23,
    'Note 29 - NON-CONTROLLING INTER': 24,
    'Note 30 - REVENUES - NET': 25,
    'Note 31 - COSTS OF REVENUES': 26,
    'Note 32 - GENERAL AND ADMINISTR': 27,  # 2023 Note 27 is OPERATING EXPENSES (combined)
    'Note 33 - OTHER INCOME (EXPENSE': 28,
    'Note 34 - FINANCE INCOME': 29,
    'Note 35 - FINANCE COSTS': 30,
    'Note 36 - EARNINGS PER SHARE': 31,
    'Note 37 - RELATED PARTY INFORMA': 32,
    'Note 38 - FINANCIAL INSTRUMENTS': 33,
    'Note 39 - FINANCIAL RISK MANAGE': 34,
    'Note 40 - SEGMENT INFORMATION': 35,
    'Note 41 - SUPPLEMENTARY CASH FL': 36,
    'Note 42 - SIGNIFICANT AGREEMENT': 37,
    'Note 43 - EVENTS AFTER REPORTIN': 38,
    'Note 44 - ISSUANCE OF AMENDMENT': 39,
}

# 2023 Note titles (in English)
NOTE_TITLES_2023 = {
    1: 'GENERAL',
    2: 'MATERIAL ACCOUNTING POLICIES INFORMATION',
    3: 'SIGNIFICANT ACCOUNTING ESTIMATES AND JUDGMENTS',
    4: 'CASH AND CASH EQUIVALENTS',
    5: 'TRADE RECEIVABLES',
    6: 'PREPAID EXPENSES AND OTHER CURRENT ASSETS',
    7: 'ADVANCES',
    8: 'FIXED ASSETS',
    9: 'INTANGIBLE ASSETS',
    10: 'DUE FROM A RELATED PARTY',
    11: 'TRADE PAYABLES',
    12: 'OTHER PAYABLES',
    13: 'ACCRUED EXPENSES',
    14: 'ADVANCE FROM CUSTOMERS',
    15: 'LEASE LIABILITIES',
    16: 'TAXATION',
    17: 'CONSUMER FINANCING PAYABLES',
    18: 'LONG-TERM BANK LOANS',
    19: 'LOAN',
    20: 'DUE TO A RELATED PARTY',
    21: 'EMPLOYEE BENEFITS LIABILITY',
    22: 'SHARE CAPITAL AND ADVANCE FOR STOCK SUBSCRIPTION',
    23: 'ADDITIONAL PAID-IN CAPITAL',
    24: 'NON-CONTROLLING INTEREST',
    25: 'REVENUES - NET',
    26: 'COSTS OF REVENUES',
    27: 'OPERATING EXPENSES',
    28: 'OTHER INCOMES (EXPENSES)',
    29: 'FINANCE INCOME',
    30: 'FINANCE EXPENSES',
    31: 'EARNINGS PER SHARE',
    32: 'RELATED PARTIES INFORMATION',
    33: 'FINANCIAL INSTRUMENTS',
    34: 'FINANCIAL RISK MANAGEMENT OBJECTIVES AND POLICIES',
    35: 'SEGMENT INFORMATION',
    36: 'SUPPLEMENTARY CASH FLOWS INFORMATION',
    37: 'COMMITMENTS AND CONTINGENCIES',
    38: 'EVENTS AFTER REPORTING PERIOD',
    39: 'ISSUANCE OF AMENDMENTS AND IMPROVEMENTS TO STANDARDS',
}

NOTE_TITLES_2023_ID = {
    1: 'UMUM',
    2: 'KEBIJAKAN AKUNTANSI MATERIAL',
    3: 'ESTIMASI DAN PERTIMBANGAN AKUNTANSI YANG SIGNIFIKAN',
    4: 'KAS DAN SETARA KAS',
    5: 'PIUTANG USAHA',
    6: 'BEBAN DIBAYAR DI MUKA DAN ASET LANCAR LAINNYA',
    7: 'UANG MUKA',
    8: 'ASET TETAP',
    9: 'ASET TAKBERWUJUD',
    10: 'PIUTANG PIHAK BERELASI',
    11: 'UTANG USAHA',
    12: 'UTANG LAIN-LAIN',
    13: 'BEBAN AKRUAL',
    14: 'UANG MUKA PENJUALAN',
    15: 'LIABILITAS SEWA',
    16: 'PERPAJAKAN',
    17: 'UTANG PEMBIAYAAN KONSUMEN',
    18: 'UTANG BANK JANGKA PANJANG',
    19: 'PINJAMAN',
    20: 'UTANG PIHAK BERELASI',
    21: 'LIABILITAS IMBALAN KERJA',
    22: 'MODAL SAHAM DAN UANG MUKA SETORAN MODAL',
    23: 'TAMBAHAN MODAL DISETOR',
    24: 'KEPENTINGAN NONPENGENDALI',
    25: 'PENDAPATAN USAHA - NETO',
    26: 'BEBAN POKOK PENDAPATAN',
    27: 'BEBAN OPERASIONAL',
    28: 'PENGHASILAN (BEBAN) LAIN-LAIN',
    29: 'PENGHASILAN KEUANGAN',
    30: 'BEBAN KEUANGAN',
    31: 'LABA NETO PER SAHAM',
    32: 'INFORMASI PIHAK BERELASI',
    33: 'INSTRUMEN KEUANGAN',
    34: 'KEBIJAKAN DAN TUJUAN MANAJEMEN RISIKO KEUANGAN',
    35: 'INFORMASI SEGMEN',
    36: 'INFORMASI TAMBAHAN ARUS KAS',
    37: 'PERJANJIAN DAN KONTINJENSI',
    38: 'PERISTIWA SETELAH PERIODE PELAPORAN',
    39: 'PENERBITAN AMENDEMEN DAN PENYESUAIAN STANDAR',
}

# Insertion column position - we always insert as the new B column (leftmost data column)
INSERT_COL = 2


def determine_insert_column(sheet, year_str):
    """Find the column index where the new year column should be inserted.
    Insert LEFT of the earliest existing year column.
    Searches the header row (row 3 for main, row 3 for notes) for year strings."""
    # Try row 3 first (header row in main statements and most notes)
    for hdr_row in [3, 2, 4]:
        years_found = []
        for c in range(1, sheet.max_column + 2):
            v = sheet.cell(row=hdr_row, column=c).value
            if isinstance(v, str) and re.fullmatch(r'20\d{2}', v.strip()):
                years_found.append((c, int(v.strip())))
            elif isinstance(v, int) and 2000 <= v <= 2099:
                years_found.append((c, v))
        if years_found:
            return hdr_row, years_found
    return None, []


def insert_year_column(sheet, year_str, hardcoded, comment_prefix=''):
    """Insert a new column at the leftmost year position. Returns the new column index."""
    hdr_row, years = determine_insert_column(sheet, year_str)
    if not years:
        return None
    leftmost_col = years[0][0]
    sheet.insert_cols(leftmost_col)
    # Place the new header
    sheet.cell(row=hdr_row, column=leftmost_col, value=year_str)
    # Fill hardcoded values
    for r, v in hardcoded.items():
        if v is None: continue
        cell = sheet.cell(row=r, column=leftmost_col)
        cell.value = v
    return leftmost_col, hdr_row


# ------------------------------------------------------------------
def main():
    print('=== Backup ===')
    print(f'Existing: {XLSX}')
    print(f'Backup:   {BACKUP}')
    # Backup already created

    print('\n=== Loading workbook ===')
    wb = openpyxl.load_workbook(XLSX)
    print('Sheets:', wb.sheetnames[:5], '...')

    # Phase 0: documentation drift guard
    print('\n=== Phase 0: state check ===')
    bs = wb['Balance Sheet']
    print(f'BS header row 3: {[bs.cell(row=3, column=c).value for c in range(1, 6)]}')
    print(f'BS row 7 (Cash): {[bs.cell(row=7, column=c).value for c in range(1, 6)]}')

    # ----- Phase 2: Insert 2023 col in main statements -----
    print('\n=== Phase 2: Main statements ===')

    # Balance Sheet
    bs = wb['Balance Sheet']
    bs.insert_cols(2)
    bs.cell(row=3, column=2, value=int(NEW_HDR))
    for r, v in BS_2023.items():
        if v is None: continue
        bs.cell(row=r, column=2).value = v
    # Comments for rows where 2023 specific
    bs.cell(row=20, column=2).comment = Comment(
        "2023 'Uang muka - aset tetap' = 88,384,436,054. "
        "Same item also appears in row 78 with 2024=0.", AUTHOR)
    bs.cell(row=79, column=2).comment = Comment(
        "Taksiran tagihan pajak penghasilan (Estimate claim income tax) 2023 = 242,314,336. "
        "Included in Total Aset Tidak Lancar (row 26).", AUTHOR)
    bs.cell(row=80, column=2).comment = Comment(
        "Uang muka setoran modal 2023 = 71,783,331,590 (per BS p.3; equity statement p.7 shows 71,783,329,590, "
        "a 2,000 IDR rounding difference). Included in Sub-total ekuitas induk row 71.", AUTHOR)
    # 'Not found' comments for rows blank in 2023
    for r in (9, 10, 18, 24, 25, 35, 39, 44, 45, 49, 53, 58, 59, 68, 69):
        bs.cell(row=r, column=2).comment = Comment("Not in 2023 report", AUTHOR)
    print(f'  Balance Sheet: 2023 column inserted, {sum(1 for v in BS_2023.values() if v is not None)} cells filled.')

    # Income Statement
    is_sh = wb['Income Statement']
    is_sh.insert_cols(2)
    is_sh.cell(row=3, column=2, value=int(NEW_HDR))
    for r, v in IS_2023.items():
        if v is None: continue
        is_sh.cell(row=r, column=2).value = v
    is_sh.cell(row=11, column=2).comment = Comment(
        "Other income 2023 = +1,551,251,599 (positive net). Beban pajak final -2,006,925 reported separately in row 44.",
        AUTHOR)
    is_sh.cell(row=44, column=2).comment = Comment(
        "2023 IS has Beban pajak final -2,006,925 as a separate line above Laba Usaha. "
        "In 2024 this was rolled into 'Other income (expenses) - net'.", AUTHOR)
    for r in (25,):
        is_sh.cell(row=r, column=2).comment = Comment("Not in 2023 report (no foreign translation OCI)", AUTHOR)
    print(f'  Income Statement: 2023 column inserted.')

    # Cash Flow
    cf = wb['Cash Flow Statement']
    cf.insert_cols(2)
    cf.cell(row=3, column=2, value=int(NEW_HDR))
    for r, v in CF_2023.items():
        if v is None: continue
        cf.cell(row=r, column=2).value = v
    cf.cell(row=31, column=2).comment = Comment(
        "Receipt from due to related parties 2023 = 4,239,964,161 (per CF p.8: 'Penerimaan dari utang pihak berelasi'; "
        "value extracted after stripping Note 20 ref).", AUTHOR)
    cf.cell(row=52, column=2).comment = Comment(
        "Pembayaran utang bank 2023 = -147,610,848,369. In 2024 this line was -23,069,357,059.", AUTHOR)
    cf.cell(row=53, column=2).comment = Comment(
        "Pembayaran utang pembiayaan konsumen 2023 = -594,833,481 (net outflow). 2024 had a net inflow of +1,073,632,739.",
        AUTHOR)
    for r in (17, 19, 21, 22, 27, 28, 29, 30, 34, 35, 38, 39, 41, 43, 44):
        cf.cell(row=r, column=2).comment = Comment("Not in 2023 report", AUTHOR)

    # Append 2023-only CF rows: Receipt from due-from-related-party (Investing inflow)
    next_row = cf.max_row + 2
    cf.cell(row=next_row, column=1, value='Catatan: Baris berikut hanya muncul di Laporan Arus Kas 2023 / Note: Rows added from 2023 Annual Report').font = Font(italic=True)
    next_row += 1
    cf.cell(row=next_row, column=1, value='  Penerimaan dari piutang pihak berelasi / Receipt from due from related party (2023 inflow)')
    cf.cell(row=next_row, column=2, value=120400198799)
    cf.cell(row=next_row, column=2).comment = Comment('From 2023 AR Cash Flow p.8, Investing section, Note 10.', AUTHOR)
    print(f'  Cash Flow Statement: 2023 column inserted, 1 row appended for due-from-related-party receipt.')

    # ----- Phase 2b: Changes in Equity -----
    print('\n=== Phase 2b: Changes in Equity (insert 2023 movements) ===')
    eq = wb['Changes in Equity']
    # Insert 8 rows above row 5
    eq.insert_rows(5, amount=8)
    # Row 5: opening balance Jan 1, 2023
    eq.cell(row=5, column=1, value='Saldo 1 Januari 2023 / Balance January 1, 2023').font = Font(bold=True)
    eq.cell(row=5, column=2, value=225532128700)         # Modal Saham
    eq.cell(row=5, column=3, value=267141192041)         # Tambahan
    eq.cell(row=5, column=4, value=117642543173)         # Selisih (2,905,639,379) + Saldo Laba (114,736,903,794)
    eq.cell(row=5, column=5, value=610315863914)         # Sub-total parent
    eq.cell(row=5, column=6, value=2411134158)           # NCI
    eq.cell(row=5, column=7, value=612726998072)         # Total
    eq.cell(row=5, column=4).comment = Comment(
        "Selisih & Saldo Laba combined: Selisih nilai trans entitas sepengendali 2,905,639,379 "
        "+ Saldo Laba 114,736,903,794 = 117,642,543,173 (per 2023 AR p.6).", AUTHOR)

    # Row 6: Penerbitan saham melalui konversi waran (Note 22, 2023)
    eq.cell(row=6, column=1, value='Penerbitan saham melalui konversi waran (Note 22) / Issuance of shares through warrant conversion (Note 22)')
    eq.cell(row=6, column=4, value=10403381100)          # In 2023 went into UMSM (folded into combined)
    eq.cell(row=6, column=5, value=10403381100)
    eq.cell(row=6, column=7, value=10403381100)
    eq.cell(row=6, column=4).comment = Comment(
        "In 2023 warrant conversion proceeds went into 'Uang muka setoran modal' (advance for stock subscription); "
        "folded into combined Selisih & Saldo Laba & UMSM column.", AUTHOR)

    # Row 7: Penerbitan modal saham melalui PMTHMETD (Note 22, 2023)
    eq.cell(row=7, column=1, value='Penerbitan saham melalui PMTHMETD (Note 22) / Issuance of shares via private placement PMTHMETD (Note 22)')
    eq.cell(row=7, column=4, value=61379948490)          # Went into UMSM
    eq.cell(row=7, column=5, value=61379948490)
    eq.cell(row=7, column=7, value=61379948490)

    # Row 8: Pelepasan entitas anak (Note 1d)
    eq.cell(row=8, column=1, value='Pelepasan entitas anak (Note 1d) / Divestment of subsidiary (Note 1d)')
    eq.cell(row=8, column=6, value=86372518)             # NCI
    eq.cell(row=8, column=7, value=86372518)

    # Row 9: Penurunan persentase kepemilikan entitas anak
    eq.cell(row=9, column=1, value='Dampak penurunan persentase kepemilikan entitas anak / Effect of decreasing percentage ownership of subsidiaries')
    eq.cell(row=9, column=6, value=-127000000)
    eq.cell(row=9, column=7, value=-127000000)

    # Row 10: Laba neto tahun berjalan
    eq.cell(row=10, column=1, value='Laba neto tahun berjalan / Net profit for the year')
    eq.cell(row=10, column=4, value=58543329595)
    eq.cell(row=10, column=5, value=58543329595)
    eq.cell(row=10, column=6, value=-286608490)
    eq.cell(row=10, column=7, value=58256721105)

    # Row 11: Rugi komprehensif lain tahun berjalan
    eq.cell(row=11, column=1, value='Rugi komprehensif lain tahun berjalan / Other comprehensive loss for the year')
    eq.cell(row=11, column=4, value=-80449038)
    eq.cell(row=11, column=5, value=-80449038)
    eq.cell(row=11, column=7, value=-80449038)

    # Row 12: closing balance Dec 31, 2023 (should match the next row's opening = Saldo 1 Jan 2024)
    eq.cell(row=12, column=1, value='Saldo 31 Desember 2023 / Balance December 31, 2023').font = Font(bold=True)
    eq.cell(row=12, column=2, value=225532128700)
    eq.cell(row=12, column=3, value=267141192041)
    eq.cell(row=12, column=4, value=247888755320)        # 71,783,331,590 (UMSM) + 2,905,639,379 + 173,199,784,351
    eq.cell(row=12, column=5, value=740562076061)
    eq.cell(row=12, column=6, value=2083898186)
    eq.cell(row=12, column=7, value=742645974247)
    eq.cell(row=12, column=4).comment = Comment(
        "Combined: UMSM 71,783,331,590 + Selisih 2,905,639,379 + Saldo Laba 173,199,784,351 = 247,888,755,320. "
        "Matches the next-row opening 'Saldo 1 Januari 2024'.", AUTHOR)
    print('  Changes in Equity: 2023 opening + 6 movements + closing inserted above row 13.')

    # ----- Phase 3: Note Index -----
    print('\n=== Phase 3: Note Index ===')
    ni = wb['Note Index']
    # Insert new col B for 2023
    ni.insert_cols(2)
    ni.cell(row=2, column=2, value=int(NEW_HDR))
    # Fill in 2023 column with English titles + Indonesian comment
    for sn, n_2023 in NOTE_MAPPING.items():
        # Find the workbook row matching this sheet
        for r in range(3, ni.max_row + 1):
            a = ni.cell(row=r, column=1).value
            if not a: continue
            # The Note Index has rows like "Note 1", "Note 2", etc.
            # We need to find the workbook note number for sheet sn.
            # Extract from sheet name: e.g. 'Note 4 - CASH...' -> 4
            m = re.match(r'Note\s+(\d+)', sn)
            if not m: continue
            wb_note_num = int(m.group(1))
            if a.strip().lower() == f'note {wb_note_num}'.lower():
                if n_2023 is None:
                    cell = ni.cell(row=r, column=2)
                    cell.value = '(not in 2023)'
                    cell.font = Font(italic=True, color='808080')
                else:
                    en = NOTE_TITLES_2023.get(n_2023, '')
                    id_ = NOTE_TITLES_2023_ID.get(n_2023, '')
                    cell = ni.cell(row=r, column=2)
                    cell.value = en
                    comment_text = f'Indonesian title: {id_}.'
                    if n_2023 != wb_note_num:
                        comment_text += f' In 2023 this topic was Note {n_2023}.'
                    cell.comment = Comment(comment_text, AUTHOR)
                break

    # Append rows for 2023-only notes (those not in workbook)
    next_row = ni.max_row + 1
    ni.cell(row=next_row, column=1, value='').font = Font(italic=True)
    next_row += 1
    ni.cell(row=next_row, column=1, value='Catatan tambahan dari Laporan Tahunan 2023 / Additional notes from 2023 Annual Report').font = Font(italic=True, bold=True)
    next_row += 1
    # 2023 Note 10 - DUE FROM RELATED PARTY
    ni.cell(row=next_row, column=1, value='(2023-only)')
    cell = ni.cell(row=next_row, column=2, value='DUE FROM A RELATED PARTY (Note 10 in 2023)')
    cell.comment = Comment('Indonesian title: PIUTANG PIHAK BERELASI. Settled in full Sep 30, 2023; not present in 2024+.', AUTHOR)
    print(f'  Note Index: 2023 column added with {sum(1 for v in NOTE_MAPPING.values() if v is not None)} title mappings.')

    # ----- Phase 4: Insert 2023 column in every note sheet -----
    print('\n=== Phase 4: Note sheets — structural column insert + key totals ===')
    for sn in wb.sheetnames:
        if not sn.startswith('Note '): continue
        if sn == 'Note Index': continue
        s = wb[sn]
        # Find the year header row and insert col 2
        hdr_row, years = determine_insert_column(s, NEW_HDR)
        if not years:
            # Some narrative sheets have different header style: 'Topik / Topic', '2024', 'Detail', 'Referensi / Reference'
            # Year is at col 2. Check row 3:
            v = s.cell(row=3, column=2).value
            if isinstance(v, str) and re.fullmatch(r'20\d{2}', v.strip()):
                s.insert_cols(2)
                s.cell(row=3, column=2, value=int(NEW_HDR))
                continue
            print(f'  WARN: {sn} - no year header row found; skipping column insert')
            continue
        leftmost_col = years[0][0]
        s.insert_cols(leftmost_col)
        s.cell(row=hdr_row, column=leftmost_col, value=int(NEW_HDR))
        # Fill hardcoded values if any
        for (sheet_name, row), val in NOTE_2023_DATA.items():
            if sheet_name != sn: continue
            if val is None: continue
            s.cell(row=row, column=leftmost_col).value = val

        # Add a sheet-level info comment on the new year header cell
        n_2023 = NOTE_MAPPING.get(sn)
        hdr_cell = s.cell(row=hdr_row, column=leftmost_col)
        if n_2023 is None:
            hdr_cell.comment = Comment(f'Topic not present in 2023 Annual Report. Column left empty.', AUTHOR)
        elif n_2023 != int(re.match(r'Note\s+(\d+)', sn).group(1)):
            hdr_cell.comment = Comment(f'In 2023 AR this was Note {n_2023}. Source pages: see Note Index comment.', AUTHOR)
    print(f'  Note sheets: column inserted in {sum(1 for sn in wb.sheetnames if sn.startswith("Note "))} sheets.')

    # ----- Phase 4b: PDF-driven sub-line fill (Layer 2) -----
    print('\n=== Phase 4b: PDF-driven sub-line fill ===')
    pdf = pdfplumber.open(PDF_PATH)
    # Build per-note page ranges
    NOTE_HEADER = re.compile(r'^\s*(\d{1,2})\.\s+[A-Z]{3,}')
    note_starts = {}
    for i in range(179, len(pdf.pages)):
        txt = pdf.pages[i].extract_text() or ''
        for line in txt.split('\n'):
            m = NOTE_HEADER.match(line.strip())
            if m and 'lanjutan' not in line.lower() and 'continued' not in line.lower():
                n = int(m.group(1))
                if n not in note_starts:
                    note_starts[n] = i+1
    sorted_notes = sorted(note_starts.keys())
    note_ranges = {}
    for idx, n in enumerate(sorted_notes):
        start = note_starts[n]
        end = note_starts[sorted_notes[idx+1]] if idx+1 < len(sorted_notes) else len(pdf.pages)
        note_ranges[n] = (start, end)

    # Number token / label parser
    NUM_TOKEN = re.compile(
        r'\(\s*\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?\s*\)'
        r'|'
        r'\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?'
    )

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

    def extract_numbers_and_label(line):
        # Strip note-reference patterns first (e.g., "2f,5,33,34" attached to value)
        # Match note refs like comma-separated digit groups (with optional lowercase letter) before a value
        line2 = re.sub(r'(\s)(\d{1,2}[a-z]?(?:,\d{1,2}[a-z]?)*)(\d{1,3}(?:\.\d{3})+)\b', r'\1\2 \3', line)
        # also handle simple "<N><digit-with-dots>" without preceding digits, e.g., "8 10.589.785.000"
        matches = list(NUM_TOKEN.finditer(line2))
        nums = [parse_id_number(m.group(0)) for m in matches]
        nums = [n for n in nums if n is not None]
        if not matches: return line2.strip(), '', []
        label_id = line2[:matches[0].start()].strip()
        label_en = line2[matches[-1].end():].strip()
        return label_id, label_en, nums

    def is_noise_value(v):
        if isinstance(v, float) and 0 < v < 100: return True
        if isinstance(v, int):
            if 1990 <= v <= 2100: return True
            if 0 <= abs(v) < 100: return True
        return False

    STOP = {'dan', 'atau', 'yang', 'untuk', 'ke', 'dari', 'dengan', 'dalam',
            'the', 'and', 'or', 'of', 'to', 'from', 'for', 'in', 'on', 'a',
            'an', 'is', 'are', 'be', 'pt', 'tbk', 'pada', 'di', 'sub', 'total',
            'net', 'neto', 'jumlah'}

    def normalize_label(s):
        if not s: return ''
        s = s.lower()
        s = re.sub(r'\([^)]*\)', ' ', s)
        s = re.sub(r'[^\w\s]', ' ', s)
        s = re.sub(r'\s+', ' ', s).strip()
        return s

    def keyword_overlap(a, b):
        wa = {w for w in a.split() if w not in STOP and len(w) > 2}
        wb_ = {w for w in b.split() if w not in STOP and len(w) > 2}
        if not wa or not wb_: return 0.0
        n = len(wa & wb_)
        if n == 0: return 0.0
        if min(len(wa), len(wb_)) <= 2:
            return n / min(len(wa), len(wb_))
        return n / max(len(wa), len(wb_))

    GENERIC_LABELS = {
        'total', 'sub total', 'neto', 'net', 'jumlah',
        'catatan note', 'note', 'catatan', 'pasal article',
        'pasal', 'article', 'perusahaan company',
        'entitas anak subsidiaries', 'lain lain others', 'lain lain',
        'rincian', 'others', 'sub total sub total',
    }

    def magnitude_violates(value, anchor):
        if not (isinstance(value, (int, float)) and isinstance(anchor, (int, float))): return False
        if abs(anchor) > 1_000_000 and abs(value) < abs(anchor) / 1000 and abs(value) < 1_000_000:
            return True
        if abs(value) > 1_000_000 and abs(anchor) < abs(value) / 1000 and abs(anchor) < 1_000_000:
            return True
        return False

    def extract_note_rows(start_p, end_p, target_note, rollforward=False):
        rows = []
        year_ctx = None
        for pn in range(start_p, end_p + 1):
            if pn > len(pdf.pages): break
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
                if re.fullmatch(r'2023(\s+/\s+2023)?', stripped): year_ctx = 'NEW'; continue
                if re.fullmatch(r'2022(\s+/\s+2022)?', stripped): year_ctx = 'OLD'; continue
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

    def match_workbook_label(sheet_label, pdf_rows):
        sl = normalize_label(sheet_label)
        if not sl or len(sl) < 4: return None
        if sl in GENERIC_LABELS: return None
        if not [w for w in sl.split() if len(w) > 2]: return None
        for i, r in enumerate(pdf_rows):
            if r['consumed']: continue
            for cand in (normalize_label(r['label_id']), normalize_label(r['label_en'])):
                if sl == cand: return (i, 'exact')
        for i, r in enumerate(pdf_rows):
            if r['consumed']: continue
            for cand in (normalize_label(r['label_id']), normalize_label(r['label_en'])):
                if not cand or len(cand) < 6: continue
                short, long_ = (sl, cand) if len(sl) <= len(cand) else (cand, sl)
                if len(short) >= 6 and short in long_: return (i, 'substring')
        best, best_score = None, 0.0
        for i, r in enumerate(pdf_rows):
            if r['consumed']: continue
            for cand in (normalize_label(r['label_id']), normalize_label(r['label_en'])):
                score = keyword_overlap(sl, cand)
                if score >= 0.6 and score > best_score:
                    best, best_score = i, score
        return (best, 'keyword') if best is not None else None

    NARRATIVE = {1, 2, 3, 33, 34, 39}  # narrative-heavy 2023 note numbers
    ROLLFORWARD = {8, 9, 16, 21}  # fixed/intangible/tax/employee benefits

    # For each workbook note sheet, attempt PDF auto-fill
    rollforward_workbook = {'Note 10 - FIXED ASSETS', 'Note 11 - INTANGIBLE ASSETS',
                             'Note 19 - TAXATION', 'Note 26 - EMPLOYEE BENEFITS LIA'}
    fill_stats = {}
    for sn, n_2023 in NOTE_MAPPING.items():
        if n_2023 is None: continue
        if n_2023 in NARRATIVE: continue
        if sn not in wb.sheetnames: continue
        if n_2023 not in note_ranges: continue
        s = wb[sn]
        start, end = note_ranges[n_2023]
        # Find header row & col
        hdr_row, years = determine_insert_column(s, NEW_HDR)
        if not years: continue
        # The col we just inserted should be the leftmost
        target_col = years[0][0]      # 2023 column we just inserted
        # The 2024 col (the prior leftmost-before-insert) is now target_col+1; the original 2025 is target_col+2.
        anchor_col = target_col + 1   # 2024 column (acts as magnitude anchor)
        is_rollforward = sn in rollforward_workbook
        pdf_rows = extract_note_rows(start, end, n_2023, rollforward=is_rollforward)
        matched, skipped_hardcoded = 0, 0
        for r in range(4, s.max_row + 1):
            a = s.cell(row=r, column=1).value
            if a is None or not isinstance(a, str): continue
            anchor = s.cell(row=r, column=anchor_col).value
            target_cell = s.cell(row=r, column=target_col)
            if target_cell.value is not None:
                skipped_hardcoded += 1
                continue
            if not isinstance(anchor, (int, float)): continue
            m = match_workbook_label(a, pdf_rows)
            if m is None: continue
            idx, mtype = m
            v = pdf_rows[idx]['value']
            if magnitude_violates(v, anchor): continue
            pdf_rows[idx]['consumed'] = True
            target_cell.value = v
            if mtype != 'exact':
                target_cell.comment = Comment(
                    f"Matched from 2023 AR p.{pdf_rows[idx]['page']}: '{pdf_rows[idx]['label_id']}' ({mtype})",
                    AUTHOR)
            matched += 1
        # Append unmatched 2023-only rows as new rows
        existing_b_values = set()
        for r in range(4, s.max_row + 1):
            v = s.cell(row=r, column=target_col).value
            if isinstance(v, (int, float)): existing_b_values.add(v)
        candidates = []
        for r in pdf_rows:
            if r['consumed']: continue
            label_norm = normalize_label(r['label_id'])
            if not label_norm: continue
            keywords = [w for w in label_norm.split() if w not in STOP and len(w) > 2]
            if not keywords: continue
            if label_norm in GENERIC_LABELS: continue
            v = r['value']
            if isinstance(v, float) and abs(v) < 1000: continue
            if isinstance(v, int):
                if abs(v) < 10000: continue
                if 1990 <= v <= 2100: continue
            if v in existing_b_values: continue
            candidates.append(r)
        if candidates:
            next_r = s.max_row + 2
            footnote_cell = s.cell(row=next_r, column=1,
                value=f'Catatan: Baris berikut hanya muncul di Laporan Tahunan 2023 (Catatan {n_2023}) / Note: Rows added from 2023 Annual Report (Note {n_2023})')
            footnote_cell.font = Font(italic=True)
            next_r += 1
            for r in candidates[:20]:  # cap at 20 to avoid noise
                label = f"  {r['label_id']} / {r['label_en']}".strip()
                s.cell(row=next_r, column=1, value=label)
                s.cell(row=next_r, column=target_col, value=r['value'])
                s.cell(row=next_r, column=target_col).comment = Comment(
                    f"From 2023 AR p.{r['page']}. Line: \"{r['line'][:120]}\"", AUTHOR)
                next_r += 1
        fill_stats[sn] = (matched, len(candidates[:20]), skipped_hardcoded)
    for sn, (m, app, skip) in fill_stats.items():
        if m or app:
            print(f'  {sn}: matched={m}, appended={app}, skipped_hardcoded={skip}')

    # ----- Phase 5: Key Ratios -----
    print('\n=== Phase 5: Key Ratios ===')
    kr = wb['Key Ratios Summary']
    # Detect header row containing existing year string
    hdr_row, years = determine_insert_column(kr, NEW_HDR)
    if years:
        leftmost_col = years[0][0]
        kr.insert_cols(leftmost_col)
        kr.cell(row=hdr_row, column=leftmost_col, value=int(NEW_HDR))
        # Compute ratios from 2023 BS/IS/CF
        # (BS_2023 values; signs as in workbook)
        TA  = BS_2023[28]   # Total assets
        TL  = BS_2023[62]
        TE  = BS_2023[73]
        CA  = BS_2023[15]   # Current assets
        CL  = BS_2023[46]
        REV = IS_2023[5]
        COGS= -IS_2023[6]   # positive for ratio
        GP  = IS_2023[7]
        NP  = IS_2023[18]
        OP  = IS_2023[12]
        EBT = IS_2023[16]
        FIN_C = -IS_2023[15]
        CFO = CF_2023[13]
        CASH= BS_2023[7]
        # Populate ratios — common rows already exist; we'll detect by label
        ratio_values = {
            'current ratio': CA / CL,
            'quick ratio': (CA - 0) / CL,         # No inventory in 2023
            'cash ratio': CASH / CL,
            'debt to equity': TL / TE,
            'debt to assets': TL / TA,
            'equity ratio': TE / TA,
            'gross profit margin': GP / REV,
            'operating profit margin': OP / REV,
            'net profit margin': NP / REV,
            'return on assets': NP / TA,
            'return on equity': NP / TE,
            'asset turnover': REV / TA,
            'interest coverage': OP / FIN_C if FIN_C else None,
            'operating cash flow ratio': CFO / CL,
            'eps': IS_2023[40],
        }
        for r in range(hdr_row + 1, kr.max_row + 1):
            label = kr.cell(row=r, column=1).value
            if not label: continue
            lab_norm = normalize_label(str(label))
            for key, val in ratio_values.items():
                if key in lab_norm and val is not None:
                    kr.cell(row=r, column=leftmost_col, value=round(val, 4) if isinstance(val, float) else val)
                    break
        # Add footnote about average-based ratios
        next_r = kr.max_row + 2
        kr.cell(row=next_r, column=1, value='Catatan: Untuk ROA/ROE/Asset Turnover di tahun 2023 menggunakan nilai akhir tahun (no prior year in workbook). 2024+ menggunakan rata-rata dua tahun. / Note: 2023 ratios use year-end values; 2024+ use two-year averages.').font = Font(italic=True)
    print(f'  Key Ratios: 2023 column inserted.')

    # ----- Phase 6: Validation -----
    print('\n=== Phase 6: Validation ===')
    bs = wb['Balance Sheet']
    is_sh = wb['Income Statement']
    cf = wb['Cash Flow Statement']

    def get(sh, r, c=2): return sh.cell(row=r, column=c).value or 0

    checks = []
    # BS: Total Assets == Total Liab + Total Equity
    ta = get(bs, 28); tl = get(bs, 62); te = get(bs, 73)
    checks.append(('BS Assets = Liab+Eq', ta == tl + te, f'{ta} vs {tl + te}'))
    # BS: Current + Non-current = Total
    checks.append(('BS Current + NC = Total Assets',
                   get(bs, 15) + get(bs, 26) == ta, f'{get(bs,15)+get(bs,26)} vs {ta}'))
    checks.append(('BS Current + NC = Total Liab',
                   get(bs, 46) + get(bs, 60) == tl, f'{get(bs,46)+get(bs,60)} vs {tl}'))
    # IS: Revenue - COGS = GP
    checks.append(('IS Rev - COGS = GP', get(is_sh, 5) + get(is_sh, 6) == get(is_sh, 7),
                   f'{get(is_sh,5)+get(is_sh,6)} vs {get(is_sh,7)}'))
    # CF: Begin + ΔCash = End
    checks.append(('CF Beg + ΔCash = End', get(cf, 48) + get(cf, 47) == get(cf, 49),
                   f'{get(cf,48)+get(cf,47)} vs {get(cf,49)}'))
    # CF: Op + Inv + Fin = ΔCash
    checks.append(('CF Op + Inv + Fin = ΔCash',
                   get(cf, 13) + get(cf, 24) + get(cf, 45) == get(cf, 47),
                   f'{get(cf,13)+get(cf,24)+get(cf,45)} vs {get(cf,47)}'))
    # Cross: CF End Cash = BS Cash
    checks.append(('CF End Cash = BS Cash', get(cf, 49) == get(bs, 7),
                   f'{get(cf,49)} vs {get(bs,7)}'))
    # Cross: IS Net Profit = Equity Statement Net Profit (eq row 10 col 7)
    eq = wb['Changes in Equity']
    checks.append(('IS NP ≈ Eq NP',
                   get(is_sh, 18) == eq.cell(row=10, column=7).value,
                   f'{get(is_sh,18)} vs {eq.cell(row=10, column=7).value}'))

    for label, ok, info in checks:
        status = 'PASS' if ok else 'FAIL'
        print(f'  {status:4}  {label:40} | {info}')

    # ----- Save -----
    wb.save(XLSX)
    print(f'\n=== Saved: {XLSX} ===')

    print('\n=== Summary ===')
    print(f'Company        : WIFI – PT Solusi Sinergi Digital Tbk')
    print(f'Year Added     : 2023')
    print(f'Source PDF     : {PDF_PATH} ({len(pdf.pages)} pages)')
    print(f'Output File    : {XLSX}')
    print(f'Backup File    : {BACKUP}')
    print(f'Years in File  : 2023, 2024, 2025 (left to right)')
    print(f'Source location in PDF:')
    print(f'  Auditor Report: pages 161-163 (Cover) / 164-171 (Auditor report body)')
    print(f'  Audit Status  : AUDITED')
    print(f'  Statement Type: CONSOLIDATED')
    print(f'  BS pages      : 172-174')
    print(f'  IS pages      : 175-176')
    print(f'  Equity pages  : 177-178')
    print(f'  CF page       : 179')

if __name__ == '__main__':
    main()
