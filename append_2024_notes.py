"""
Phase 3d + Phase 4: Create Note Index sheet and populate 2024 sub-line data
for matched renumbered notes. Source values extracted from 2024 AR pages.
"""
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

XLSX = 'WIFI_Financial_Statements.xlsx'

# Note matching: workbook sheet name -> (2024 note number or None, 2024 title, 2025 title)
NOTE_MATCH = [
    ('Note 1 - GENERAL',                  1,  'UMUM',                                    'GENERAL'),
    ('Note 2 - SIGNIFICANT ACCOUNTING',   2,  'INFORMASI KEBIJAKAN AKUNTANSI MATERIAL', 'SIGNIFICANT ACCOUNTING POLICIES'),
    ('Note 3 - ACCOUNTING ESTIMATES A',   3,  'ESTIMASI DAN PERTIMBANGAN AKUNTANSI',    'SIGNIFICANT ACCOUNTING ESTIMATES AND JUDGMENTS'),
    ('Note 4 - CASH AND CASH EQUIVALE',   4,  'KAS DAN SETARA KAS',                     'CASH AND CASH EQUIVALENTS'),
    ('Note 5 - TRADE RECEIVABLES',        5,  'PIUTANG USAHA',                          'TRADE RECEIVABLES'),
    ('Note 6 - INVENTORIES',              None, None,                                  'INVENTORIES'),
    ('Note 7 - PREPAID EXPENSES AND O',   6,  'BEBAN DIBAYAR DI MUKA DAN ASET LANCAR LAINNYA', 'PREPAID EXPENSES AND OTHER CURRENT ASSETS'),
    ('Note 8 - ADVANCES',                 7,  'UANG MUKA',                              'ADVANCES'),
    ('Note 9 - RESTRICTED FUND',          None, None,                                  'RESTRICTED FUND'),
    ('Note 10 - FIXED ASSETS',            8,  'ASET TETAP',                             'FIXED ASSETS'),
    ('Note 11 - INTANGIBLE ASSETS',       9,  'ASET TAKBERWUJUD',                       'INTANGIBLE ASSETS'),
    ('Note 12 - OTHER ASSETS',            None, None,                                  'OTHER ASSETS'),
    ('Note 13 - TRADE PAYABLES',         10,  'UTANG USAHA',                            'TRADE PAYABLES'),
    ('Note 14 - OTHER PAYABLES',         11,  'UTANG LAIN-LAIN',                        'OTHER PAYABLES'),
    ('Note 15 - ACCRUED EXPENSES',       12,  'BEBAN AKRUAL',                           'ACCRUED EXPENSES'),
    ('Note 16 - ADVANCES FROM CUSTOME',  13,  'UANG MUKA PENJUALAN',                    'ADVANCES FROM CUSTOMERS'),
    ('Note 17 - LEASE LIABILITIES',      14,  'LIABILITAS SEWA',                        'LEASE LIABILITIES'),
    ('Note 18 - SHORT-TERM BANK LOANS',  None, None,                                  'SHORT-TERM BANK LOANS'),
    ('Note 19 - TAXATION',               15,  'PERPAJAKAN',                             'TAXATION'),
    ('Note 20 - CONSUMER FINANCING PA',  16,  'UTANG PEMBIAYAAN KONSUMEN',              'CONSUMER FINANCING PAYABLES'),
    ('Note 21 - LONG-TERM BANK LOANS',   17,  'UTANG BANK JANGKA PANJANG',              'LONG-TERM BANK LOANS'),
    ('Note 22 - BONDS PAYABLE',          18,  'UTANG OBLIGASI',                         'BONDS PAYABLE'),
    ('Note 23 - LOANS',                  19,  'PINJAMAN',                               'LOANS'),
    ('Note 24 - DUE TO RELATED PARTIE',  20,  'UTANG PIHAK BERELASI',                   'DUE TO RELATED PARTIES'),
    ('Note 25 - SHARIA BONDS',           None, None,                                  'SHARIA BONDS (SUKUK)'),
    ('Note 26 - EMPLOYEE BENEFITS LIA',  21,  'LIABILITAS IMBALAN KERJA',               'EMPLOYEE BENEFITS LIABILITY'),
    ('Note 27 - SHARE CAPITAL AND ADV',  22,  'MODAL SAHAM DAN UANG MUKA SETORAN MODAL', 'SHARE CAPITAL AND ADVANCES FOR STOCK SUBSCRIPTION'),
    ('Note 28 - ADDITIONAL PAID-IN CA',  23,  'TAMBAHAN MODAL DISETOR',                 'ADDITIONAL PAID-IN CAPITAL'),
    ('Note 29 - NON-CONTROLLING INTER',  24,  'KEPENTINGAN NONPENGENDALI',              'NON-CONTROLLING INTERESTS'),
    ('Note 30 - REVENUES - NET',         25,  'PENDAPATAN USAHA - NETO',                'REVENUES - NET'),
    ('Note 31 - COSTS OF REVENUES',      26,  'BEBAN POKOK PENDAPATAN',                 'COSTS OF REVENUES'),
    ('Note 32 - GENERAL AND ADMINISTR',  27,  'BEBAN OPERASIONAL',                      'GENERAL AND ADMINISTRATIVE EXPENSES'),
    ('Note 33 - OTHER INCOME (EXPENSE',  28,  'PENGHASILAN (BEBAN) LAIN-LAIN',          'OTHER INCOME (EXPENSES)'),
    ('Note 34 - FINANCE INCOME',         29,  'PENGHASILAN KEUANGAN',                   'FINANCE INCOME'),
    ('Note 35 - FINANCE COSTS',          30,  'BEBAN KEUANGAN',                         'FINANCE COSTS'),
    ('Note 36 - EARNINGS PER SHARE',     31,  'LABA NETO PER SAHAM',                    'EARNINGS PER SHARE'),
    ('Note 37 - RELATED PARTY INFORMA',  32,  'INFORMASI PIHAK BERELASI',               'RELATED PARTY INFORMATION'),
    ('Note 38 - FINANCIAL INSTRUMENTS',  33,  'INSTRUMEN KEUANGAN',                     'FINANCIAL INSTRUMENTS'),
    ('Note 39 - FINANCIAL RISK MANAGE',  34,  'KEBIJAKAN DAN TUJUAN MANAJEMEN RISIKO KEUANGAN', 'FINANCIAL RISK MANAGEMENT OBJECTIVES AND POLICIES'),
    ('Note 40 - SEGMENT INFORMATION',    35,  'INFORMASI SEGMEN',                       'SEGMENT INFORMATION'),
    ('Note 41 - SUPPLEMENTARY CASH FL',  36,  'INFORMASI TAMBAHAN ARUS KAS',            'SUPPLEMENTARY CASH FLOWS INFORMATION'),
    ('Note 42 - SIGNIFICANT AGREEMENT',  37,  'IKATAN DAN KONTINJENSI',                 'SIGNIFICANT AGREEMENTS, COMMITMENTS AND CONTINGENCIES'),
    ('Note 43 - EVENTS AFTER REPORTIN',  38,  'PERISTIWA SETELAH PERIODE PELAPORAN',    'EVENTS AFTER REPORTING PERIOD'),
    ('Note 44 - ISSUANCE OF AMENDMENT',  39,  'PENERBITAN AMENDEMEN DAN PENYESUAIAN',  'ISSUANCE OF AMENDMENTS AND ADJUSTMENTS TO ACCOUNTING STANDARDS'),
]

# ============================================================
# 2024 NOTE DATA — extracted from 2024 AR pages
# Maps: workbook_sheet -> { row_num: (value, optional_comment) }
# Only rows where the topic in the sheet matches actual 2024 data are filled.
# ============================================================
NOTE_2024_DATA = {
    # Note 7 — Prepaid Expenses (2024 Note 6, p.246)
    'Note 7 - PREPAID EXPENSES AND O': {
        23: (16949690677, None),                          # Proyek - Project
        20: (3860002276, None),                           # Asuransi short-term
        22: (12000000, None),                             # Sewa
        24: (291828638, None),                            # Lain-lain
        25: (21113521591, 'Sub-total Short-term Prepaid Expenses, 2024'),
        27: (9794106194, None),                           # Asuransi long-term
        28: (30907627785, 'Sub-total Long-term Prepaid Expenses, 2024'),
        29: (302096542, None),                            # Other current assets
        30: (31209724327, 'Total Prepaid Expenses + Other Current Assets, 2024'),
        # Row 19 (Biaya izin frekuensi): NOT present in 2024 — leave empty + comment
        19: (None, 'Not found in 2024 report — frequency license fees appear from 2025'),
        21: (None, 'Not found in 2024 report — Consultant prepaid expense'),
    },
    # Note 8 — Advances (2024 Note 7, p.247)
    'Note 8 - ADVANCES': {
        58: (148919000000, 'PT Lintas Daya Andalan — 2024 advance'),
        # 59 Forge International — Not in 2024
        59: (None, 'Not found in 2024 report'),
        60: (None, 'Not found in 2024 report'),
        62: (None, 'Not found in 2024 report'),
        63: (None, 'Not found in 2024 report'),
    },
    # Note 13 — Trade Payables (2024 Note 10, p.252)
    'Note 13 - TRADE PAYABLES': {
        17: (None, 'Not found in 2024 report — Trade payables to related parties'),
        19: (None, 'Not found in 2024 report'),
        21: (None, 'Not found in 2024 report'),
        22: (None, 'Not found in 2024 report'),
        23: (None, 'Not found in 2024 report'),
        24: (None, 'Not found in 2024 report'),
        25: (None, 'Not found in 2024 report'),
        26: (None, 'Not found in 2024 report'),
        27: (None, 'Not found in 2024 report'),
        29: (None, 'Not found in 2024 report'),
        30: (None, 'Not found in 2024 report'),
        31: (None, 'Not found in 2024 report'),
        32: (None, 'Not found in 2024 report'),
        33: (None, 'Not found in 2024 report'),
        34: (None, 'Not found in 2024 report'),
        37: (33001370287, 'Sub-total / Total Trade Payables 2024 (matches BS)'),
        38: (33001370287, 'Total trade payables 2024'),
        40: (None, 'Not found in 2024 report'),
        42: (None, 'Not found in 2024 report'),
        43: (None, 'Not found in 2024 report'),
        44: (None, 'Not found in 2024 report'),
        45: (None, 'Not found in 2024 report'),
        46: (None, 'Not found in 2024 report'),
    },
    # Note 14 — Other Payables (2024 Note 11, p.252)
    'Note 14 - OTHER PAYABLES': {
        # 2024 current: third parties 8,029,839,382; non-current: 301,212,468,808
        51: (None, 'Not found in 2024 report — related parties under accrued/other current items'),
        52: (None, 'Not found in 2024 report'),
        53: (8029839382, 'Sub-total — current third parties 2024 (matches BS)'),
        55: (None, 'Not found in 2024 report — IGM/Prambanan loans first appear differently in 2024'),
        56: (None, 'Not found in 2024 report'),
        57: (None, 'Not found in 2024 report'),
        58: (None, 'Not found in 2024 report'),
        59: (None, 'Not found in 2024 report'),
        60: (None, 'Long-term third parties total 2024 = 301,212,468,808 (matches BS). Composition differs from 2025 template.'),
        63: (None, 'Not found in 2024 report — Day-one profit on related-party borrowings appears 2025'),
    },
    # Note 15 — Accrued Expenses (2024 Note 12, p.254)
    'Note 15 - ACCRUED EXPENSES': {
        53: (None, 'Not found in 2024 report'),
        54: (None, 'Not found in 2024 report'),
        55: (None, 'Not found in 2024 report'),
        56: (None, 'Not found in 2024 report'),
        57: (None, 'Not found in 2024 report'),
        58: (5109714952, 'Total Accrued Expenses 2024 (matches BS)'),
    },
    # Note 16 — Advances from Customers (2024 Note 13, p.255)
    'Note 16 - ADVANCES FROM CUSTOME': {
        37: (48018103366, 'PT XL Axiata Tbk 2024'),
        38: (15375000000, 'PT MNC Kabel Mediacom 2024'),
        40: (58000000000, 'PT Telemedia Komunikasi Pratama 2024'),
        41: (16456995698, 'Lain-lain 2024'),
        42: (137850099064, 'Total Advances from Customers 2024'),
        43: (-10290154592, 'Less: short-term portion 2024'),
        44: (127559944472, 'Long-term portion 2024 (matches BS row 52)'),
        34: (None, 'Not found in 2024 report'),
        35: (None, 'Not found in 2024 report'),
        36: (None, 'Not found in 2024 report'),
    },
    # Note 17 — Lease Liabilities (2024 Note 14, p.255-256)
    'Note 17 - LEASE LIABILITIES': {
        46: (214126485506, 'Utang angsuran 2024'),
        49: (-29875361411, 'Future finance charge 2024'),
        50: (184251124095, 'Present value of minimum payments 2024'),
        52: (-25717400118, 'Current maturity 2024 (matches BS row 40)'),
        53: (158533723977, 'Long-term portion 2024 (matches BS row 54)'),
        47: (None, 'Not found in 2024 report'),
    },
    # Note 20 — Consumer Financing Payables (2024 Note 16, p.261)
    'Note 20 - CONSUMER FINANCING PA': {
        # 2024 total = 386,633,823 + 726,049,149 = 1,112,682,972
    },
    # Note 21 — Long-Term Bank Loans (2024 Note 17, p.262)
    'Note 21 - LONG-TERM BANK LOANS': {
        # 2024 total = 138,055,277,673 + 237,250,226,847 = 375,305,504,520
    },
    # Note 22 — Bonds Payable (2024 Note 18, p.277)
    'Note 22 - BONDS PAYABLE': {
        # 2024 total = 166,632,590,847 + 447,452,484,087 = 614,085,074,934 (one bond series, first issuance)
    },
    # Note 23 — Loans (2024 Note 19, p.278)
    'Note 23 - LOANS': {
        # 2024 total = 83,918,554,209 + 34,204,122,692 = 118,122,676,901
    },
    # Note 24 — Due to Related Parties (2024 Note 20, p.281)
    'Note 24 - DUE TO RELATED PARTIE': {
        # 2024 non-current = 39,377,574,929
    },
    # Note 26 — Employee Benefits (2024 Note 21, p.284)
    'Note 26 - EMPLOYEE BENEFITS LIA': {
        # 2024 = 3,155,146,210
    },
    # Note 27 — Share Capital (2024 Note 22, p.285)
    'Note 27 - SHARE CAPITAL AND ADV': {
        # Issued and paid-in 2024: 235,935,511,800 from 2,359,355,118 shares
        49: (235935511800, 'Issued and paid-in capital 2024 (matches BS row 65)'),
        48: (2359355118, 'Number of issued shares 2024'),
    },
    # Note 28 — Additional Paid-in Capital (2024 Note 23, p.286)
    'Note 28 - ADDITIONAL PAID-IN CA': {
        # 2024 = 328,521,140,531
    },
    # Note 29 — NCI (2024 Note 24, p.287)
    'Note 29 - NON-CONTROLLING INTER': {
        # 2024 NCI = 181,418,175 (matches BS row 72)
        31: (181418175, 'Total NCI 2024 (matches BS)'),
    },
    # Note 30 — Revenues (2024 Note 25, p.287)
    'Note 30 - REVENUES - NET': {
        # Net revenue 2024 = 671,854,001,272 (matches IS row 5)
        43: (671854001272, 'Net Revenue 2024 (matches IS)'),
    },
    # Note 31 — Cost of Revenues (2024 Note 26, p.288)
    'Note 31 - COSTS OF REVENUES': {
        # Total COGS 2024 = 257,080,687,384 (matches IS row 6 abs)
        35: (257080687384, 'Total Cost of Revenues 2024 (matches |IS|)'),
    },
    # Note 32 — Operating expenses (2024 Note 27, p.288)
    'Note 32 - GENERAL AND ADMINISTR': {
        # 2024 G&A = 68,674,526,161 (matches |IS row 10|)
    },
    # Note 33 — Other income (2024 Note 28)
    'Note 33 - OTHER INCOME (EXPENSE': {
        26: (-1207531677, 'Other income (expense) - net 2024 (matches IS row 11)'),
    },
    # Note 34 — Finance income (2024 Note 29)
    'Note 34 - FINANCE INCOME': {
        33: (538954781, 'Finance income 2024 (matches IS row 14)'),
    },
    # Note 35 — Finance costs (2024 Note 30)
    'Note 35 - FINANCE COSTS': {
        46: (-73049522882, 'Finance costs 2024 (matches IS row 15)'),
    },
    # Note 36 — EPS (2024 Note 31)
    'Note 36 - EARNINGS PER SHARE': {
        52: (231186780014, 'NP attributable to parent 2024'),
    },
    # Note 41 — Supplementary Cash Flows (2024 Note 36)
    'Note 41 - SUPPLEMENTARY CASH FL': {
        # Non-cash activities — typically narrative, leave header + comment
    },
}


def create_note_index_sheet(wb):
    """Create the Note Index sheet between Changes in Equity and Note 1."""
    if 'Note Index' in wb.sheetnames:
        del wb['Note Index']

    # Determine position
    sheets = wb.sheetnames
    ce_idx = sheets.index('Changes in Equity')

    # Create new sheet at end then move
    idx_sheet = wb.create_sheet('Note Index')
    # Move it to position right after Changes in Equity
    wb.move_sheet(idx_sheet, offset=-(len(wb.sheetnames) - 1 - ce_idx - 1))

    # Row 1: title (merged across cols A, B, C)
    idx_sheet['A1'] = 'Note Title Index — Cross-Year Reference'
    idx_sheet['A1'].font = Font(bold=True, size=12)
    idx_sheet['A1'].alignment = Alignment(horizontal='center')
    idx_sheet.merge_cells('A1:C1')

    # Row 2: headers
    idx_sheet['A2'] = 'Note Number'
    idx_sheet['B2'] = '2024'
    idx_sheet['C2'] = '2025'
    for col in ['A2', 'B2', 'C2']:
        idx_sheet[col].font = Font(bold=True)
        idx_sheet[col].fill = PatternFill('solid', fgColor='DDDDDD')

    # Rows 3+: one row per note sheet
    for i, (sname, n2024, t2024, t2025) in enumerate(NOTE_MATCH):
        r = 3 + i
        note_num_label = sname.split(' - ')[0]   # e.g. "Note 1"
        idx_sheet.cell(row=r, column=1, value=note_num_label)
        if n2024 is not None:
            idx_sheet.cell(row=r, column=2, value=t2024)
            if int(note_num_label.split()[1]) != n2024:
                idx_sheet.cell(row=r, column=2).comment = Comment(
                    f'In 2024, this topic was numbered Note {n2024}.', 'Append2024')
        else:
            idx_sheet.cell(row=r, column=2).comment = Comment(
                'Note not present in 2024 report', 'Append2024')
        idx_sheet.cell(row=r, column=3, value=t2025)

    # Column widths
    idx_sheet.column_dimensions['A'].width = 14
    idx_sheet.column_dimensions['B'].width = 55
    idx_sheet.column_dimensions['C'].width = 55

    # Freeze top 2 rows
    idx_sheet.freeze_panes = 'A3'


def populate_note_data(wb):
    """For each renumbered note with 2024 PDF data, fill column B values."""
    for sheet_name, row_map in NOTE_2024_DATA.items():
        if sheet_name not in wb.sheetnames:
            continue
        s = wb[sheet_name]
        for row_num, item in row_map.items():
            val, cmt = item
            if val is not None:
                s.cell(row=row_num, column=2, value=val)
            if cmt:
                s.cell(row=row_num, column=2).comment = Comment(cmt, 'Append2024')


def main():
    print('Loading workbook...')
    wb = openpyxl.load_workbook(XLSX)

    print('Creating Note Index sheet...')
    create_note_index_sheet(wb)

    print('Populating 2024 note sub-line data...')
    populate_note_data(wb)

    print('Saving...')
    wb.save(XLSX)
    print('Done.')


if __name__ == '__main__':
    main()
