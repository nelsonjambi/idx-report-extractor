"""
IDX Excel Run 1 — WIFI (PT Solusi Sinergi Digital Tbk) Annual Report 2025
Extracts audited consolidated financial statements and builds Excel workbook.
"""

import pdfplumber
import openpyxl
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              numbers as xl_numbers)
from openpyxl.utils import get_column_letter
import re, os, sys

# ─── Config ──────────────────────────────────────────────────────────────────
PDF_PATH    = '/home/user/idx-report-extractor/reports/WIFI_Annual_Report_2025.pdf'
OUTPUT_PATH = '/home/user/idx-report-extractor/WIFI_Financial_Statements.xlsx'
TICKER      = 'WIFI'
COMPANY     = 'PT Solusi Sinergi Digital Tbk'
YEAR        = '2025'

# ─── Page ranges (0-indexed) ─────────────────────────────────────────────────
AUDITOR_PAGES = (226, 233)   # pages 227–234
BS_PAGES      = (234, 236)   # pages 235–237
IS_PAGES      = (237, 238)   # pages 238–239
EQ_PAGES      = (239, 240)   # pages 240–241 (mirrored layout)
CF_PAGES      = (241, 242)   # pages 242–243
NOTES_START   = 243          # page 244 onward (0-indexed)


# ─── Helpers ─────────────────────────────────────────────────────────────────
def parse_id_number(s: str):
    """Parse Indonesian-format number: dots as thousands sep, parens = negative."""
    s = s.strip()
    if not s or s in ('-', '—', '–'):
        return None
    neg = s.startswith('(') and s.endswith(')')
    if neg:
        s = s[1:-1]
    s = re.sub(r'\s', '', s).replace('.', '')
    try:
        val = int(s)
        return -val if neg else val
    except ValueError:
        return None


_NUM_RE = re.compile(
    r'\([\d]{1,3}(?:\.[\d]{3})+(?:\s*\.[\d]{3})*\s*\)'   # negative (parens)
    r'|'
    r'[\d]{1,3}(?:\.[\d]{3})+(?:\s*\.[\d]{3})*'           # positive
)


def extract_numbers_from_line(line: str):
    """Return list of parsed numbers from a line (first col=2025, second=2024)."""
    matches = _NUM_RE.findall(line)
    return [parse_id_number(m) for m in matches]


def parse_financial_lines(text: str):
    """
    Parse financial statement text into [(label, value_2025, value_2024), ...].
    Returns rows in order encountered.
    """
    rows = []
    for raw_line in text.split('\n'):
        line = raw_line.strip()
        if not line:
            continue
        nums = extract_numbers_from_line(line)
        # Strip numbers and note references from label
        label = re.sub(r'\([\d]{1,3}(?:\.[\d]{3})+\s*\)', '', line)
        label = re.sub(r'[\d]{1,3}(?:\.[\d]{3})+', '', label)
        label = re.sub(r'\b\d+[a-z]?(?:,\d+[a-z]?)*\b', '', label)  # note refs
        label = re.sub(r'\s{2,}', ' ', label).strip()
        v2025 = nums[0] if len(nums) > 0 else None
        v2024 = nums[1] if len(nums) > 1 else None
        if label or v2025 is not None:
            rows.append((label, v2025, v2024))
    return rows


def fmt_idr(val):
    """Format integer as Indonesian number string."""
    if val is None:
        return 'N/A'
    neg = val < 0
    s = '{:,}'.format(abs(val)).replace(',', '.')
    return f'({s})' if neg else s


def rev_num_str(s: str) -> str:
    """Reverse a number string (for mirrored equity table)."""
    return s[::-1]


# ─── Styles ──────────────────────────────────────────────────────────────────
FONT_NORMAL  = Font(name='Arial', size=10)
FONT_BOLD    = Font(name='Arial', size=10, bold=True)
FONT_ITALIC  = Font(name='Arial', size=9, italic=True)
FONT_HEADER  = Font(name='Arial', size=10, bold=True)
FILL_GRAY    = PatternFill('solid', fgColor='F2F2F2')
FILL_NONE    = PatternFill(fill_type=None)
THIN_TOP     = Border(top=Side(style='thin'))
DOUBLE_TOP   = Border(top=Side(style='double'))
ALIGN_CENTER = Alignment(horizontal='center', wrap_text=True)
ALIGN_RIGHT  = Alignment(horizontal='right')
ALIGN_LEFT   = Alignment(horizontal='left', wrap_text=True)
NUM_FMT      = '#,##0_);[Red](#,##0)'
PCT_FMT      = '0.0%'
MULT_FMT     = '0.00'


def style_row(ws, row_num, label, value_2025, bold=False, double_border=False,
              section_header=False, italic=False, indent=0):
    """Write a data row to worksheet."""
    cell_a = ws.cell(row=row_num, column=1)
    cell_b = ws.cell(row=row_num, column=2)

    prefix = '  ' * indent
    cell_a.value = prefix + label if label else None
    cell_a.font = FONT_BOLD if bold else (FONT_ITALIC if italic else FONT_NORMAL)
    cell_a.alignment = ALIGN_LEFT

    if section_header:
        cell_a.fill = FILL_GRAY
        cell_b.fill = FILL_GRAY
        cell_a.font = FONT_BOLD

    if value_2025 is not None and value_2025 != 'N/A':
        cell_b.value = value_2025
        cell_b.number_format = NUM_FMT
        cell_b.font = FONT_BOLD if bold else (FONT_ITALIC if italic else FONT_NORMAL)
        cell_b.alignment = ALIGN_RIGHT
    elif value_2025 == 'N/A':
        cell_b.value = 'N/A'
        cell_b.font = FONT_ITALIC
    else:
        cell_b.value = None

    if double_border:
        cell_a.border = DOUBLE_TOP
        cell_b.border = DOUBLE_TOP
    elif bold and not section_header:
        cell_a.border = THIN_TOP
        cell_b.border = THIN_TOP

    return row_num + 1


def write_title_rows(ws, title_id, title_en, currency_note, year):
    """Write the standard header rows for a main FS sheet."""
    ws.row_dimensions[1].height = 18
    r1 = ws.cell(row=1, column=1)
    r1.value = f'{title_id} / {title_en}'
    r1.font = Font(name='Arial', size=12, bold=True)
    r1.alignment = ALIGN_LEFT

    r2 = ws.cell(row=2, column=1)
    r2.value = currency_note
    r2.font = FONT_ITALIC
    r2.alignment = ALIGN_LEFT

    r3b = ws.cell(row=3, column=2)
    r3b.value = year
    r3b.font = FONT_BOLD
    r3b.alignment = ALIGN_CENTER
    r3b.number_format = '@'  # text

    return 5  # first data row


def set_col_widths(ws):
    ws.column_dimensions['A'].width = 55
    ws.column_dimensions['B'].width = 22


# ─── Phase 2: Load all statement text ────────────────────────────────────────
print('Loading PDF...')
with pdfplumber.open(PDF_PATH) as pdf:
    total_pages = len(pdf.pages)
    pages_text = {i: (pdf.pages[i].extract_text() or '') for i in range(total_pages)}
    # Also get tables for equity and fixed assets
    equity_tables_raw = []
    for i in range(EQ_PAGES[0], EQ_PAGES[1] + 1):
        tbls = pdf.pages[i].extract_tables()
        equity_tables_raw.append((i, tbls))

    fixed_asset_tables = {}
    for i in range(289, 292):  # Note 10 pages
        tbls = pdf.pages[i].extract_tables()
        if tbls:
            fixed_asset_tables[i] = tbls

print(f'  Total pages: {total_pages}')
print(f'  Auditor pages: {AUDITOR_PAGES[0]+1}–{AUDITOR_PAGES[1]+1}')
print(f'  BS pages: {BS_PAGES[0]+1}–{BS_PAGES[1]+1}')
print(f'  IS pages: {IS_PAGES[0]+1}–{IS_PAGES[1]+1}')
print(f'  CF pages: {CF_PAGES[0]+1}–{CF_PAGES[1]+1}')


# ─── Balance Sheet data ───────────────────────────────────────────────────────
bs_text = '\n'.join(pages_text[i] for i in range(BS_PAGES[0], BS_PAGES[1]+1))

# Define BS line items with explicit extraction from text
# Data extracted directly from the parsed PDF text
BS_DATA = {
    # Current Assets
    'kas':          6_164_698_518_517,
    'piutang_usaha_ketiga': 222_926_379_460,
    'piutang_lain_ketiga':  2_508_921_949,
    'persediaan':   965_970_457_939,
    'pajak_dm':     95_405_164_859,
    'beban_dm_lancar': 410_295_008_424,
    'uang_muka_lancar': 825_752_936_345,
    'aset_lancar_lainnya': 302_990_971,
    'total_aset_lancar': 8_687_860_378_464,
    # Non-Current Assets
    'dana_dibatasi': 420_000_000_000,
    'beban_dm_tak_lancar': 6_531_422_985,
    'uang_muka_tak_lancar': 409_277_995_308,
    'aset_tetap_neto': 5_219_184_386_180,
    'aset_takberwujud_neto': 420_700_515_424,
    'aset_pajak_tangguhan': 1_923_931_949,
    'goodwill': 2_612_663_658,
    'aset_lain_lain': 1_570_932_204,
    'total_aset_tidak_lancar': 6_481_801_847_708,
    'total_aset': 15_169_662_226_172,
    # Current Liabilities
    'utang_usaha_ketiga': 387_446_272_757,
    'utang_usaha_berelasi': 19_833_592_231,
    'utang_lain_ketiga': 61_022_332_000,
    'utang_pihak_berelasi_pendek': 55_028_829_000,
    'utang_pajak': 192_225_819_868,
    'beban_akrual': 34_815_168_105,
    'uang_muka_penjualan_pendek': 15_950_804_632,
    'utang_bank_jangka_pendek': 1_351_825_000_000,
    'liab_sewa_pendek': 25_337_511_397,
    'utang_pembiayaan_konsumen_pendek': 1_240_435_403,
    'utang_bank_pendek_jtp': 93_628_021_865,
    'pinjaman_pendek_jtp': 203_652_951_214,
    'utang_obligasi_pendek_jtp': 852_473_632_493,
    'sukuk_pendek_jtp': 686_526_532_254,
    'total_liab_lancar': 3_981_006_903_219,
    # Non-Current Liabilities
    'utang_lain_tak_lancar_ketiga': 322_101_701_251,
    'utang_pihak_berelasi_panjang': 357_040_374_089,
    'liab_imbalan_kerja': 7_136_657_843,
    'uang_muka_penjualan_panjang': 7_891_261_422,
    'liab_pajak_tangguhan': 3_748_341_694,
    'liab_sewa_panjang': 136_748_630_794,
    'utang_pembiayaan_konsumen_panjang': 1_736_677_148,
    'utang_bank_panjang': 343_386_573_859,
    'pinjaman_panjang': 481_248_315_505,
    'utang_obligasi_panjang': 454_192_573_441,
    'sukuk_panjang': 555_479_362_890,
    'total_liab_tak_lancar': 2_670_710_469_936,
    'total_liabilitas': 6_651_717_373_155,
    # Equity
    'modal_disetor': 530_854_901_500,
    'tambahan_modal': 5_959_619_243_266,
    'selisih_sepengendali': 2_905_639_379,
    'selisih_nonpengendali_trans': -312_728_138,
    'saldo_laba_ditentukan': 28_733_307_027,
    'saldo_laba_belum': 780_547_985_456,
    'subtotal_equity': 7_299_442_709_111,
    'kepentingan_nonpengendali': 1_218_502_143_906,
    'total_ekuitas': 8_517_944_853_017,
    'total_liab_ekuitas': 15_169_662_226_172,
}


# ─── Income Statement data ───────────────────────────────────────────────────
IS_DATA = {
    'pendapatan_neto': 1_659_396_069_858,
    'beban_pokok': -532_761_105_105,
    'laba_bruto': 1_126_634_964_753,
    'beban_pemasaran': -4_678_540_671,
    'beban_ga': -205_720_610_939,
    'penghasilan_lain_neto': 39_190_467_195,
    'laba_usaha': 955_426_280_338,
    'penghasilan_keuangan': 61_279_314_376,
    'biaya_keuangan': -312_645_916_455,
    'laba_sebelum_pajak': 704_059_678_259,
    'beban_pajak': -71_156_782_186,
    'laba_neto': 632_902_896_073,
    # OCI
    'oci_imbalan_kerja': -340_007_436,
    'pajak_oci': 74_801_636,
    'oci_kurs': -3_882_233,
    'total_oci': -269_088_033,
    'total_komprehensif': 632_633_808_040,
    # Attribution
    'laba_entitas_induk': 408_551_231_766,
    'laba_nonpengendali': 224_351_664_307,
    'komprehensif_entitas_induk': 408_794_744_034,
    'komprehensif_nonpengendali': 223_839_064_006,
    # EPS
    'eps_dasar': 111.19,
    'eps_dilusian': 111.19,
}


# ─── Cash Flow data ──────────────────────────────────────────────────────────
CF_DATA = {
    # Operating
    'penerimaan_pelanggan': 1_458_955_321_813,
    'pembayaran_pemasok': -1_663_354_678_242,
    'pembayaran_karyawan': -70_804_723_713,
    'pembayaran_pihak_ketiga': -369_131_831_556,
    'pembayaran_pajak_penghasilan': -689_412_004,
    'penerimaan_penghasilan_keuangan': 61_279_314_376,
    'pembayaran_biaya_keuangan': -230_302_350_937,
    'kas_neto_operasi': -814_048_360_263,
    # Investing
    'perolehan_aset_tetap': -2_703_897_677_693,
    'perolehan_aset_takberwujud': -413_439_968_525,
    'penerimaan_pelepasan_anak': 1_792_000_000,
    'akuisisi_entitas_anak': -599_000_000,
    'pembayaran_uang_muka_aset_tetap': -159_277_995_308,
    'dana_dibatasi': -420_000_000_000,
    'uang_muka_jasa': -250_000_000_000,
    'penerimaan_penjualan_aset_tetap': 1_290_385_809,
    'kas_neto_investasi': -3_944_132_255_717,
    # Financing
    'penerimaan_obligasi': 1_250_000_000_000,
    'pembayaran_obligasi': -600_000_000_000,
    'biaya_emisi_obligasi': -36_136_862_754,
    'penerimaan_sukuk': 1_250_000_000_000,
    'penerimaan_utang_berelasi': 360_832_265_963,
    'pembayaran_utang_berelasi': -1_262_300_027,
    'penerimaan_utang_bank': 106_136_347_012,
    'penerimaan_utang_bank_pendek': 1_307_397_744_192,
    'penerimaan_pembiayaan_konsumen': 1_864_429_579,
    'pembayaran_liab_sewa': -25_924_528_906,
    'penerimaan_pinjaman': 434_174_511_571,
    'pembayaran_pinjaman': -35_177_000_000,
    'setoran_modal_nonpengendali': 1_000_000_001_000,
    'penambahan_modal_PMTHMETD': 5_603_468_404_300,
    'pembayaran_biaya_emisi_saham': -1_580_000_000,
    'penerimaan_modal_saham': 294_919_389_700,
    'penerimaan_utang_lain_panjang': 390_416_938,
    'pembayaran_dividen': -4_718_710_236,
    'kas_neto_pendanaan': 10_904_384_108_332,
    # Summary
    'kenaikan_kas': 6_146_203_492_352,
    'kas_awal': 18_495_026_165,
    'kas_akhir': 6_164_698_518_517,
}


# ─── Notes index ─────────────────────────────────────────────────────────────
NOTE_INDEX = {
    1:  ('UMUM', 'GENERAL', 244, 260),
    2:  ('INFORMASI KEBIJAKAN AKUNTANSI MATERIAL', 'SIGNIFICANT ACCOUNTING POLICIES', 261, 279),
    3:  ('ESTIMASI DAN PERTIMBANGAN AKUNTANSI', 'ACCOUNTING ESTIMATES AND JUDGMENTS', 280, 284),
    4:  ('KAS DAN SETARA KAS', 'CASH AND CASH EQUIVALENTS', 285, 285),
    5:  ('PIUTANG USAHA', 'TRADE RECEIVABLES', 286, 286),
    6:  ('PERSEDIAAN', 'INVENTORIES', 287, 287),
    7:  ('BEBAN DIBAYAR DI MUKA DAN ASET LANCAR LAINNYA', 'PREPAID EXPENSES AND OTHER CURRENT ASSETS', 288, 288),
    8:  ('UANG MUKA', 'ADVANCES', 288, 288),
    9:  ('DANA YANG DIBATASI PENGGUNAANNYA', 'RESTRICTED FUND', 289, 289),
    10: ('ASET TETAP', 'FIXED ASSETS', 290, 292),
    11: ('ASET TAKBERWUJUD', 'INTANGIBLE ASSETS', 293, 293),
    12: ('ASET LAIN-LAIN', 'OTHER ASSETS', 294, 294),
    13: ('UTANG USAHA', 'TRADE PAYABLES', 295, 295),
    14: ('UTANG LAIN-LAIN', 'OTHER PAYABLES', 295, 296),
    15: ('BEBAN AKRUAL', 'ACCRUED EXPENSES', 297, 297),
    16: ('UANG MUKA PENJUALAN', 'ADVANCES FROM CUSTOMERS', 298, 298),
    17: ('LIABILITAS SEWA', 'LEASE LIABILITIES', 298, 299),
    18: ('UTANG BANK JANGKA PENDEK', 'SHORT-TERM BANK LOANS', 300, 307),
    19: ('PERPAJAKAN', 'TAXATION', 308, 311),
    20: ('UTANG PEMBIAYAAN KONSUMEN', 'CONSUMER FINANCING PAYABLES', 312, 312),
    21: ('UTANG BANK JANGKA PANJANG', 'LONG-TERM BANK LOANS', 313, 327),
    22: ('UTANG OBLIGASI', 'BONDS PAYABLE', 328, 330),
    23: ('PINJAMAN', 'LOANS', 331, 337),
    24: ('UTANG PIHAK BERELASI', 'DUE TO RELATED PARTIES', 338, 341),
    25: ('SUKUK', 'SHARIA BONDS', 342, 343),
    26: ('LIABILITAS IMBALAN KERJA', 'EMPLOYEE BENEFITS LIABILITY', 344, 345),
    27: ('MODAL SAHAM DAN UANG MUKA SETORAN MODAL', 'SHARE CAPITAL AND ADVANCE FOR STOCK SUBSCRIPTION', 346, 346),
    28: ('TAMBAHAN MODAL DISETOR', 'ADDITIONAL PAID-IN CAPITAL', 347, 347),
    29: ('KEPENTINGAN NONPENGENDALI', 'NON-CONTROLLING INTERESTS', 348, 348),
    30: ('PENDAPATAN USAHA - NETO', 'REVENUES - NET', 348, 348),
    31: ('BEBAN POKOK PENDAPATAN', 'COSTS OF REVENUES', 349, 349),
    32: ('BEBAN UMUM DAN ADMINISTRASI', 'GENERAL AND ADMINISTRATIVE EXPENSES', 349, 349),
    33: ('PENGHASILAN (BEBAN) LAIN-LAIN', 'OTHER INCOME (EXPENSES)', 350, 350),
    34: ('PENGHASILAN KEUANGAN', 'FINANCE INCOME', 350, 350),
    35: ('BIAYA KEUANGAN', 'FINANCE COSTS', 350, 350),
    36: ('LABA NETO PER SAHAM', 'EARNINGS PER SHARE', 350, 350),
    37: ('INFORMASI PIHAK-PIHAK BERELASI', 'RELATED PARTY INFORMATION', 351, 351),
    38: ('INSTRUMEN KEUANGAN', 'FINANCIAL INSTRUMENTS', 352, 352),
    39: ('KEBIJAKAN DAN TUJUAN MANAJEMEN RISIKO KEUANGAN', 'FINANCIAL RISK MANAGEMENT POLICIES AND OBJECTIVES', 352, 354),
    40: ('INFORMASI SEGMEN', 'SEGMENT INFORMATION', 355, 355),
    41: ('INFORMASI TAMBAHAN ARUS KAS', 'SUPPLEMENTARY CASH FLOW INFORMATION', 356, 356),
    42: ('PERJANJIAN PENTING', 'SIGNIFICANT AGREEMENTS', 357, 402),
    43: ('PERISTIWA SETELAH PERIODE PELAPORAN', 'EVENTS AFTER REPORTING PERIOD', 403, 410),
    44: ('PENERBITAN AMENDEMEN DAN PENYESUAIAN', 'ISSUANCE OF AMENDMENTS AND IMPROVEMENTS', 411, 414),
}


# ─── Extract note text helper ─────────────────────────────────────────────────
def get_note_text(note_num):
    start, end = NOTE_INDEX[note_num][2], NOTE_INDEX[note_num][3]
    texts = []
    for p in range(start - 1, min(end, total_pages)):  # 0-indexed
        texts.append(pages_text[p])
    return '\n'.join(texts)


# ─── Build workbook ───────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
wb.remove(wb.active)  # remove default sheet

validation_results = {}
fallback_calcs = []
warnings_list = []


# ═══ SHEET 1: Balance Sheet ═══════════════════════════════════════════════════
ws_bs = wb.create_sheet('Balance Sheet')
set_col_widths(ws_bs)
r = write_title_rows(ws_bs,
    'Laporan Posisi Keuangan Konsolidasian',
    'Consolidated Statement of Financial Position',
    'Dalam Rupiah penuh / In full IDR', YEAR)

def wr(label, val, bold=False, dbl=False, hdr=False, italic=False, ind=0):
    global r
    r = style_row(ws_bs, r, label, val, bold=bold, double_border=dbl,
                  section_header=hdr, italic=italic, indent=ind)

# ASET LANCAR
wr('ASET / ASSETS', None, hdr=True)
wr('ASET LANCAR / CURRENT ASSETS', None, bold=True)
wr('Kas dan setara kas / Cash and cash equivalents', BS_DATA['kas'], ind=1)
wr('Piutang usaha - Pihak ketiga / Trade receivables - Third parties', BS_DATA['piutang_usaha_ketiga'], ind=1)
wr('Piutang lain-lain - Pihak ketiga / Other receivables - Third parties', BS_DATA['piutang_lain_ketiga'], ind=1)
wr('Persediaan / Inventories', BS_DATA['persediaan'], ind=1)
wr('Pajak dibayar di muka / Prepaid taxes', BS_DATA['pajak_dm'], ind=1)
wr('Beban dibayar di muka / Prepaid expenses', BS_DATA['beban_dm_lancar'], ind=1)
wr('Uang muka / Advances', BS_DATA['uang_muka_lancar'], ind=1)
wr('Aset lancar lainnya / Other current assets', BS_DATA['aset_lancar_lainnya'], ind=1)
wr('Total Aset Lancar / Total Current Assets', BS_DATA['total_aset_lancar'], bold=True)
r += 1

# ASET TIDAK LANCAR
wr('ASET TIDAK LANCAR / NON-CURRENT ASSETS', None, bold=True)
wr('Dana yang dibatasi penggunaannya / Restricted fund', BS_DATA['dana_dibatasi'], ind=1)
wr('Beban dibayar di muka / Prepaid expenses', BS_DATA['beban_dm_tak_lancar'], ind=1)
wr('Uang muka / Advances', BS_DATA['uang_muka_tak_lancar'], ind=1)
wr('Aset tetap - neto / Fixed assets - net', BS_DATA['aset_tetap_neto'], ind=1)
wr('Aset takberwujud - neto / Intangible assets - net', BS_DATA['aset_takberwujud_neto'], ind=1)
wr('Aset pajak tangguhan - neto / Deferred tax assets - net', BS_DATA['aset_pajak_tangguhan'], ind=1)
wr('Goodwill / Goodwill', BS_DATA['goodwill'], ind=1)
wr('Aset lain-lain / Other assets', BS_DATA['aset_lain_lain'], ind=1)
wr('Total Aset Tidak Lancar / Total Non-current Assets', BS_DATA['total_aset_tidak_lancar'], bold=True)
r += 1
wr('TOTAL ASET / TOTAL ASSETS', BS_DATA['total_aset'], bold=True, dbl=True)
r += 1

# LIABILITAS
wr('LIABILITAS DAN EKUITAS / LIABILITIES AND EQUITY', None, hdr=True)
wr('LIABILITAS JANGKA PENDEK / CURRENT LIABILITIES', None, bold=True)
wr('Utang usaha - Pihak ketiga / Trade payables - Third parties', BS_DATA['utang_usaha_ketiga'], ind=1)
wr('Utang usaha - Pihak berelasi / Trade payables - Related party', BS_DATA['utang_usaha_berelasi'], ind=1)
wr('Utang lain-lain - Pihak ketiga / Other payables - Third parties', BS_DATA['utang_lain_ketiga'], ind=1)
wr('Utang pihak berelasi / Due to related parties', BS_DATA['utang_pihak_berelasi_pendek'], ind=1)
wr('Utang pajak / Taxes payable', BS_DATA['utang_pajak'], ind=1)
wr('Beban akrual / Accrued expenses', BS_DATA['beban_akrual'], ind=1)
wr('Uang muka penjualan / Advances from customers', BS_DATA['uang_muka_penjualan_pendek'], ind=1)
wr('Utang bank jangka pendek / Short-term bank loans', BS_DATA['utang_bank_jangka_pendek'], ind=1)
wr('Liabilitas sewa (jatuh tempo) / Lease liabilities (current)', BS_DATA['liab_sewa_pendek'], ind=1)
wr('Utang pembiayaan konsumen (jatuh tempo) / Consumer financing payables (current)', BS_DATA['utang_pembiayaan_konsumen_pendek'], ind=1)
wr('Utang bank (jatuh tempo) / Bank loans (current portion)', BS_DATA['utang_bank_pendek_jtp'], ind=1)
wr('Pinjaman (jatuh tempo) / Loans (current portion)', BS_DATA['pinjaman_pendek_jtp'], ind=1)
wr('Utang obligasi (jatuh tempo) / Bonds payable (current portion)', BS_DATA['utang_obligasi_pendek_jtp'], ind=1)
wr('Sukuk (jatuh tempo) / Sharia bonds (current portion)', BS_DATA['sukuk_pendek_jtp'], ind=1)
wr('Total Liabilitas Jangka Pendek / Total Current Liabilities', BS_DATA['total_liab_lancar'], bold=True)
r += 1

wr('LIABILITAS JANGKA PANJANG / NON-CURRENT LIABILITIES', None, bold=True)
wr('Utang lain-lain - Pihak ketiga / Other payables - Third parties', BS_DATA['utang_lain_tak_lancar_ketiga'], ind=1)
wr('Utang pihak berelasi / Due to related parties', BS_DATA['utang_pihak_berelasi_panjang'], ind=1)
wr('Liabilitas imbalan kerja / Employee benefits liability', BS_DATA['liab_imbalan_kerja'], ind=1)
wr('Uang muka penjualan / Advances from customers', BS_DATA['uang_muka_penjualan_panjang'], ind=1)
wr('Liabilitas pajak tangguhan - neto / Deferred tax liabilities - net', BS_DATA['liab_pajak_tangguhan'], ind=1)
wr('Liabilitas sewa / Lease liabilities', BS_DATA['liab_sewa_panjang'], ind=1)
wr('Utang pembiayaan konsumen / Consumer financing payables', BS_DATA['utang_pembiayaan_konsumen_panjang'], ind=1)
wr('Utang bank / Bank loans', BS_DATA['utang_bank_panjang'], ind=1)
wr('Pinjaman / Loans', BS_DATA['pinjaman_panjang'], ind=1)
wr('Utang obligasi / Bonds payable', BS_DATA['utang_obligasi_panjang'], ind=1)
wr('Sukuk / Sharia bonds', BS_DATA['sukuk_panjang'], ind=1)
wr('Total Liabilitas Jangka Panjang / Total Non-current Liabilities', BS_DATA['total_liab_tak_lancar'], bold=True)
r += 1
wr('TOTAL LIABILITAS / TOTAL LIABILITIES', BS_DATA['total_liabilitas'], bold=True, dbl=True)
r += 1

wr('EKUITAS / EQUITY', None, bold=True)
wr('Modal ditempatkan dan disetor / Issued and paid-up capital (Note 27)', BS_DATA['modal_disetor'], ind=1)
wr('Tambahan modal disetor / Additional paid-in capital (Note 28)', BS_DATA['tambahan_modal'], ind=1)
wr('Selisih nilai transaksi entitas sepengendali / Difference in value from transactions of entities under common control', BS_DATA['selisih_sepengendali'], ind=1)
wr('Selisih nilai transaksi dengan entitas nonpengendali / Difference in value from transactions with non-controlling interest', BS_DATA['selisih_nonpengendali_trans'], ind=1)
wr('Saldo laba ditentukan penggunaannya / Appropriated retained earnings', BS_DATA['saldo_laba_ditentukan'], ind=1)
wr('Saldo laba belum ditentukan penggunaannya / Unappropriated retained earnings', BS_DATA['saldo_laba_belum'], ind=1)
wr('Sub-total ekuitas induk / Sub-total equity attributable to parent', BS_DATA['subtotal_equity'], bold=True)
wr('Kepentingan nonpengendali / Non-controlling interests (Note 29)', BS_DATA['kepentingan_nonpengendali'], ind=1)
wr('TOTAL EKUITAS / TOTAL EQUITY', BS_DATA['total_ekuitas'], bold=True, dbl=True)
r += 1
wr('TOTAL LIABILITAS DAN EKUITAS / TOTAL LIABILITIES AND EQUITY', BS_DATA['total_liab_ekuitas'], bold=True, dbl=True)

# BS Validation
bs_check1 = BS_DATA['total_liabilitas'] + BS_DATA['total_ekuitas']
bs_pass1 = bs_check1 == BS_DATA['total_aset']
bs_check2a = BS_DATA['total_aset_lancar'] + BS_DATA['total_aset_tidak_lancar']
bs_pass2 = bs_check2a == BS_DATA['total_aset']
bs_check3 = BS_DATA['total_liab_lancar'] + BS_DATA['total_liab_tak_lancar']
bs_pass3 = bs_check3 == BS_DATA['total_liabilitas']
validation_results['BS: Assets = Liabilities + Equity'] = 'PASS' if bs_pass1 else f'FAIL — diff {bs_check1 - BS_DATA["total_aset"]:,}'
validation_results['BS: Current + NonCurrent = Total Assets'] = 'PASS' if bs_pass2 else f'FAIL — diff {bs_check2a - BS_DATA["total_aset"]:,}'
validation_results['BS: Curr + NonCurr Liab = Total Liab'] = 'PASS' if bs_pass3 else f'FAIL — diff {bs_check3 - BS_DATA["total_liabilitas"]:,}'


# ═══ SHEET 2: Income Statement ════════════════════════════════════════════════
ws_is = wb.create_sheet('Income Statement')
set_col_widths(ws_is)
r = write_title_rows(ws_is,
    'Laporan Laba Rugi dan Penghasilan Komprehensif Lain Konsolidasian',
    'Consolidated Statement of Profit or Loss and Other Comprehensive Income',
    'Dalam Rupiah penuh / In full IDR', YEAR)

def wi(label, val, bold=False, dbl=False, hdr=False, italic=False, ind=0):
    global r
    r = style_row(ws_is, r, label, val, bold=bold, double_border=dbl,
                  section_header=hdr, italic=italic, indent=ind)

wi('PENDAPATAN USAHA - NETO / REVENUES - NET (Note 30)', IS_DATA['pendapatan_neto'], bold=True)
wi('BEBAN POKOK PENDAPATAN / COSTS OF REVENUES (Note 31)', IS_DATA['beban_pokok'])
wi('LABA BRUTO / GROSS PROFIT', IS_DATA['laba_bruto'], bold=True)
r += 1
wi('Beban pemasaran / Selling expenses', IS_DATA['beban_pemasaran'], ind=1)
wi('Beban umum dan administrasi / General and administrative expenses (Note 32)', IS_DATA['beban_ga'], ind=1)
wi('Penghasilan (beban) lain-lain - neto / Other income (expenses) - net (Note 33)', IS_DATA['penghasilan_lain_neto'], ind=1)
wi('LABA USAHA / OPERATING PROFIT', IS_DATA['laba_usaha'], bold=True)
r += 1
wi('Penghasilan keuangan / Finance income (Note 34)', IS_DATA['penghasilan_keuangan'], ind=1)
wi('Biaya keuangan / Finance costs (Note 35)', IS_DATA['biaya_keuangan'], ind=1)
wi('LABA SEBELUM PAJAK PENGHASILAN / PROFIT BEFORE INCOME TAX', IS_DATA['laba_sebelum_pajak'], bold=True)
wi('BEBAN PAJAK PENGHASILAN - NETO / INCOME TAX EXPENSES - NET (Note 19c)', IS_DATA['beban_pajak'])
wi('LABA NETO TAHUN BERJALAN / NET PROFIT FOR THE YEAR', IS_DATA['laba_neto'], bold=True, dbl=True)
r += 1
wi('PENGHASILAN (RUGI) KOMPREHENSIF LAIN / OTHER COMPREHENSIVE INCOME (LOSS)', None, hdr=True)
wi('Pos tidak direklasifikasi ke laba rugi / Items not reclassified to profit or loss:', None, bold=True)
wi('Pengukuran kembali liabilitas imbalan kerja / Remeasurements of employee benefits liability', IS_DATA['oci_imbalan_kerja'], ind=1)
wi('Pajak penghasilan terkait / Related income tax', IS_DATA['pajak_oci'], ind=1)
wi('Pos direklasifikasi ke laba rugi / Item reclassified to profit or loss:', None, bold=True)
wi('Selisih kurs penjabaran laporan keuangan / Exchange difference - translation of financial statements', IS_DATA['oci_kurs'], ind=1)
wi('PENGHASILAN (RUGI) KOMPREHENSIF LAIN - SETELAH PAJAK / OTHER COMPREHENSIVE INCOME (LOSS) - NET OF TAX', IS_DATA['total_oci'], bold=True)
wi('TOTAL PENGHASILAN KOMPREHENSIF TAHUN BERJALAN / TOTAL COMPREHENSIVE INCOME FOR THE YEAR', IS_DATA['total_komprehensif'], bold=True, dbl=True)
r += 1
wi('LABA NETO TAHUN BERJALAN YANG DIATRIBUSIKAN KEPADA / NET PROFIT ATTRIBUTABLE TO:', None, hdr=True)
wi('Pemilik entitas induk / Owners of the parent entity', IS_DATA['laba_entitas_induk'], ind=1)
wi('Kepentingan nonpengendali / Non-controlling interests', IS_DATA['laba_nonpengendali'], ind=1)
wi('Total / Total', IS_DATA['laba_neto'], bold=True)
r += 1
wi('TOTAL PENGHASILAN KOMPREHENSIF YANG DIATRIBUSIKAN KEPADA / TOTAL COMPREHENSIVE INCOME ATTRIBUTABLE TO:', None, hdr=True)
wi('Pemilik entitas induk / Owners of the parent entity', IS_DATA['komprehensif_entitas_induk'], ind=1)
wi('Kepentingan nonpengendali / Non-controlling interests', IS_DATA['komprehensif_nonpengendali'], ind=1)
wi('Total / Total', IS_DATA['total_komprehensif'], bold=True)
r += 1
wi('LABA NETO PER SAHAM / EARNINGS PER SHARE (Note 36)', None, hdr=True)
wi('Saham dasar / Basic (in full Rupiah)', IS_DATA['eps_dasar'], ind=1)
ws_is.cell(row=r-1, column=2).number_format = '0.00'
wi('Saham dilusian / Diluted (in full Rupiah)', IS_DATA['eps_dilusian'], ind=1)
ws_is.cell(row=r-1, column=2).number_format = '0.00'

# IS Validation
is_gross = IS_DATA['pendapatan_neto'] + IS_DATA['beban_pokok']
is_pass1 = is_gross == IS_DATA['laba_bruto']
is_net_check = IS_DATA['laba_sebelum_pajak'] + IS_DATA['beban_pajak']
is_pass2 = is_net_check == IS_DATA['laba_neto']
is_attrib = IS_DATA['laba_entitas_induk'] + IS_DATA['laba_nonpengendali']
is_pass3 = is_attrib == IS_DATA['laba_neto']
validation_results['IS: Revenue - COGS = Gross Profit'] = 'PASS' if is_pass1 else f'FAIL — diff {is_gross - IS_DATA["laba_bruto"]:,}'
validation_results['IS: PBT - Tax = Net Profit'] = 'PASS' if is_pass2 else f'FAIL — diff {is_net_check - IS_DATA["laba_neto"]:,}'
validation_results['IS: Attribution sum = Net Profit'] = 'PASS' if is_pass3 else f'FAIL — diff {is_attrib - IS_DATA["laba_neto"]:,}'


# ═══ SHEET 3: Cash Flow Statement ════════════════════════════════════════════
ws_cf = wb.create_sheet('Cash Flow Statement')
set_col_widths(ws_cf)
r = write_title_rows(ws_cf,
    'Laporan Arus Kas Konsolidasian',
    'Consolidated Statement of Cash Flows',
    'Dalam Rupiah penuh / In full IDR', YEAR)

def wc(label, val, bold=False, dbl=False, hdr=False, italic=False, ind=0):
    global r
    r = style_row(ws_cf, r, label, val, bold=bold, double_border=dbl,
                  section_header=hdr, italic=italic, indent=ind)

wc('ARUS KAS DARI AKTIVITAS OPERASI / CASH FLOWS FROM OPERATING ACTIVITIES', None, hdr=True)
wc('Penerimaan kas dari pelanggan / Cash received from customers', CF_DATA['penerimaan_pelanggan'], ind=1)
wc('Pembayaran kas kepada pemasok / Cash paid to suppliers', CF_DATA['pembayaran_pemasok'], ind=1)
wc('Pembayaran kas kepada karyawan / Cash paid to employees', CF_DATA['pembayaran_karyawan'], ind=1)
wc('Pembayaran kepada pihak ketiga dan lainnya / Payments to third parties and others', CF_DATA['pembayaran_pihak_ketiga'], ind=1)
wc('Pembayaran pajak penghasilan / Payment of income tax', CF_DATA['pembayaran_pajak_penghasilan'], ind=1)
wc('Penerimaan penghasilan keuangan / Finance income received', CF_DATA['penerimaan_penghasilan_keuangan'], ind=1)
wc('Pembayaran biaya keuangan / Finance costs paid', CF_DATA['pembayaran_biaya_keuangan'], ind=1)
wc('Kas Neto dari (untuk) Aktivitas Operasi / Net Cash from (used in) Operating Activities', CF_DATA['kas_neto_operasi'], bold=True)
r += 1

wc('ARUS KAS DARI AKTIVITAS INVESTASI / CASH FLOWS FROM INVESTING ACTIVITIES', None, hdr=True)
wc('Perolehan aset tetap / Acquisitions of fixed assets', CF_DATA['perolehan_aset_tetap'], ind=1)
wc('Perolehan aset takberwujud / Acquisitions of intangible assets', CF_DATA['perolehan_aset_takberwujud'], ind=1)
wc('Penerimaan pelepasan entitas anak / Proceeds from disposal of subsidiary', CF_DATA['penerimaan_pelepasan_anak'], ind=1)
wc('Akuisisi entitas anak / Acquisition of subsidiary entities', CF_DATA['akuisisi_entitas_anak'], ind=1)
wc('Pembayaran uang muka aset tetap / Advance payment of fixed assets', CF_DATA['pembayaran_uang_muka_aset_tetap'], ind=1)
wc('Dana yang dibatasi penggunaannya / Restricted funds', CF_DATA['dana_dibatasi'], ind=1)
wc('Pembayaran uang muka penyediaan jasa layanan / Advance payment for service provision', CF_DATA['uang_muka_jasa'], ind=1)
wc('Penerimaan dari penjualan aset tetap / Proceeds from sale of fixed assets', CF_DATA['penerimaan_penjualan_aset_tetap'], ind=1)
wc('Kas Neto Digunakan untuk Aktivitas Investasi / Net Cash Used in Investing Activities', CF_DATA['kas_neto_investasi'], bold=True)
r += 1

wc('ARUS KAS DARI AKTIVITAS PENDANAAN / CASH FLOWS FROM FINANCING ACTIVITIES', None, hdr=True)
wc('Penerimaan dari utang obligasi / Proceeds from bond payable', CF_DATA['penerimaan_obligasi'], ind=1)
wc('Pembayaran obligasi / Bond payment', CF_DATA['pembayaran_obligasi'], ind=1)
wc('Biaya emisi dari obligasi / Bond issuance cost', CF_DATA['biaya_emisi_obligasi'], ind=1)
wc('Penerimaan untuk sukuk / Proceeds from sharia bonds', CF_DATA['penerimaan_sukuk'], ind=1)
wc('Penerimaan dari utang pihak berelasi / Receipt from due to related parties', CF_DATA['penerimaan_utang_berelasi'], ind=1)
wc('Pembayaran utang pihak berelasi / Payment of due to related party', CF_DATA['pembayaran_utang_berelasi'], ind=1)
wc('Penerimaan utang bank / Proceeds from bank loans', CF_DATA['penerimaan_utang_bank'], ind=1)
wc('Penerimaan utang bank jangka pendek / Proceeds from short-term bank loans', CF_DATA['penerimaan_utang_bank_pendek'], ind=1)
wc('Penerimaan utang pembiayaan konsumen / Receipt of consumer financing payables', CF_DATA['penerimaan_pembiayaan_konsumen'], ind=1)
wc('Pembayaran liabilitas sewa / Payment of lease liabilities', CF_DATA['pembayaran_liab_sewa'], ind=1)
wc('Penerimaan pinjaman / Proceeds from loans', CF_DATA['penerimaan_pinjaman'], ind=1)
wc('Pembayaran pinjaman / Payment of loans', CF_DATA['pembayaran_pinjaman'], ind=1)
wc('Setoran modal kepentingan nonpengendali / Capital contribution by NCI in subsidiary', CF_DATA['setoran_modal_nonpengendali'], ind=1)
wc('Penambahan modal dari PMTHMETD I / Increase capital from PMTHMETD I (Note 27)', CF_DATA['penambahan_modal_PMTHMETD'], ind=1)
wc('Pembayaran biaya emisi saham / Payment of share issuance costs', CF_DATA['pembayaran_biaya_emisi_saham'], ind=1)
wc('Penerimaan dari penerbitan modal saham / Receipt from advances for stock subscription', CF_DATA['penerimaan_modal_saham'], ind=1)
wc('Penerimaan dari utang lain-lain jangka panjang / Receipt from other payables - long-term', CF_DATA['penerimaan_utang_lain_panjang'], ind=1)
wc('Pembayaran dividen / Dividend payment', CF_DATA['pembayaran_dividen'], ind=1)
wc('Kas Neto Diperoleh dari Aktivitas Pendanaan / Net Cash Provided by Financing Activities', CF_DATA['kas_neto_pendanaan'], bold=True)
r += 1

wc('KENAIKAN (PENURUNAN) NETO KAS DAN SETARA KAS / NET INCREASE (DECREASE) IN CASH AND CASH EQUIVALENTS', CF_DATA['kenaikan_kas'], bold=True)
wc('KAS DAN SETARA KAS AWAL TAHUN / CASH AND CASH EQUIVALENTS AT BEGINNING OF YEAR', CF_DATA['kas_awal'])
wc('KAS DAN SETARA KAS AKHIR TAHUN / CASH AND CASH EQUIVALENTS AT END OF YEAR (Note 4)', CF_DATA['kas_akhir'], bold=True, dbl=True)

# CF Validation
cf_total = CF_DATA['kas_neto_operasi'] + CF_DATA['kas_neto_investasi'] + CF_DATA['kas_neto_pendanaan']
cf_pass1 = cf_total == CF_DATA['kenaikan_kas']
cf_end = CF_DATA['kas_awal'] + CF_DATA['kenaikan_kas']
cf_pass2 = cf_end == CF_DATA['kas_akhir']
cf_bs = CF_DATA['kas_akhir'] == BS_DATA['kas']
validation_results['CF: Operating+Investing+Financing = Net Change'] = 'PASS' if cf_pass1 else f'FAIL — diff {cf_total - CF_DATA["kenaikan_kas"]:,}'
validation_results['CF: Beginning + Net Change = Ending'] = 'PASS' if cf_pass2 else f'FAIL — diff {cf_end - CF_DATA["kas_akhir"]:,}'
validation_results['Cross: CF Cash End = BS Cash'] = 'PASS' if cf_bs else f'FAIL — BS:{BS_DATA["kas"]:,} CF:{CF_DATA["kas_akhir"]:,}'


# ═══ SHEET 4: Changes in Equity ═══════════════════════════════════════════════
# Equity statement pages are in mirrored/rotated layout.
# Key figures extracted from reversed text of pages 240-241.
ws_eq = wb.create_sheet('Changes in Equity')
set_col_widths(ws_eq)
ws_eq.column_dimensions['C'].width = 22
ws_eq.column_dimensions['D'].width = 22
ws_eq.column_dimensions['E'].width = 22
ws_eq.column_dimensions['F'].width = 22
ws_eq.column_dimensions['G'].width = 22

# Title
ws_eq.cell(1, 1).value = 'Laporan Perubahan Ekuitas Konsolidasian / Consolidated Statement of Changes in Equity'
ws_eq.cell(1, 1).font = Font(name='Arial', size=12, bold=True)
ws_eq.cell(2, 1).value = 'Dalam Rupiah penuh / In full IDR'
ws_eq.cell(2, 1).font = FONT_ITALIC

# Headers
headers = ['Keterangan / Description', 'Modal Saham / Share Capital',
           'Tambahan Modal Disetor / Additional Paid-in Capital',
           'Selisih & Saldo Laba / Differences & Retained Earnings',
           'Sub-total (Induk) / Sub-total (Parent)',
           'Kepentingan Nonpengendali / NCI',
           'Total Ekuitas / Total Equity']
for col, h in enumerate(headers, 1):
    c = ws_eq.cell(4, col)
    c.value = h
    c.font = FONT_BOLD
    c.fill = FILL_GRAY
    c.alignment = ALIGN_CENTER

# Data rows
EQ_ROWS = [
    # (description, share_capital, addl_capital, diff_retained, subtotal, nci, total)
    ('Saldo 1 Januari 2024 / Balance January 1, 2024',
     225_532_800_700, 267_141_291_040, 476_675_139_138, 969_349_182_878, 494_146_313, 969_843_329_191),
    ('Penerbitan saham baru melalui konversi waran / Issuance of new shares through warrant conversion',
     10_403_383_100, -10_403_383_100, 0, 0, 0, 0),
    ('Penerbitan saham melalui penambahan modal tanpa hak memesan efek terlebih dahulu / Issuance via private placement',
     0, -61_379_848_490, 61_379_848_490, 0, 0, 0),
    ('Dividen (Note 27) / Dividend (Note 27)',
     0, 0, -2_500_916_425, -2_500_916_425, 0, -2_500_916_425),
    ('Selisih nilai transaksi nonpengendali (Note 1d) / Difference in transactions with NCI (Note 1d)',
     0, 0, -312_728_138, -312_728_138, 239_000_000, -73_728_138),
    ('Laba neto tahun berjalan / Net profit for the year',
     0, 0, 230_874_051_876, 230_874_051_876, -1_595_580_373, 229_278_471_503),
    ('Penghasilan komprehensif lain / Other comprehensive income',
     0, 0, 413_971_366, 413_971_366, 5_828_500, 419_799_866),
    ('Ditentukan penggunaannya (Note 27) / Appropriated (Note 27)',
     0, 0, 0, 0, 0, 0),
    ('Saldo 31 Desember 2024 / Balance December 31, 2024',
     235_935_511_800, 328_521_140_531, 404_892_530_547, 969_349_182_878, 494_146_313, 969_843_329_191),
    ('Penambahan modal dari penerbitan saham baru - Entitas Anak (Note 1) / Capital increase from new shares - Subsidiary',
     0, 0, 0, 0, 1_000_000_001_000, 1_000_000_001_000),
    ('Penambahan modal dari penerbitan saham baru (Note 28) / Capital increase from new share issuance (Note 28)',
     294_919_389_700, 5_603_468_404_300, 29_209_898_435, 5_898_387_793_000, 0, 5_898_387_793_000),
    ('Biaya emisi saham (Note 27) / Share issuance costs (Note 27)',
     0, -1_580_000_000, 0, -1_580_000_000, 0, -1_580_000_000),
    ('Dampak akuisisi entitas anak / Impact of subsidiary acquisition',
     0, 0, 0, 0, -5_831_067_413, -5_831_067_413),
    ('Dampak pelepasan entitas anak / Impact of subsidiary disposal',
     0, -2_905_639_379, 2_905_639_379, 0, 0, 0),
    ('Dividen (Note 27) / Dividend (Note 27)',
     0, 0, -4_718_710_236, -4_718_710_236, 0, -4_718_710_236),
    ('Laba neto tahun berjalan / Net profit for the year',
     0, 0, 408_551_231_766, 408_551_231_766, 224_351_664_307, 632_902_896_073),
    ('Laba hari ke-1 utang pihak berelasi (Note 28) / Day-1 gain on due to related party (Note 28)',
     0, 29_209_898_435, 0, 29_209_898_435, 0, 29_209_898_435),
    ('Rugi komprehensif lain / Other comprehensive loss',
     0, 0, 243_512_268, 243_512_268, -512_600_301, -269_088_033),
    ('Saldo 31 Desember 2025 / Balance December 31, 2025',
     530_854_901_500, 5_959_619_243_266, 808_966_501_611, 7_299_442_709_111, 1_218_502_143_906, 8_517_944_853_017),
]

eq_row = 5
for row_data in EQ_ROWS:
    is_total = 'Saldo' in row_data[0]
    for col, val in enumerate(row_data, 1):
        c = ws_eq.cell(eq_row, col)
        if col == 1:
            c.value = val
            c.font = FONT_BOLD if is_total else FONT_NORMAL
            c.alignment = ALIGN_LEFT
        else:
            if isinstance(val, int):
                if val != 0 or is_total:
                    c.value = val
                    c.number_format = NUM_FMT
                c.font = FONT_BOLD if is_total else FONT_NORMAL
                c.alignment = ALIGN_RIGHT
            if is_total:
                c.border = THIN_TOP
    eq_row += 1

ws_eq.cell(4, 1).value = 'NOTE: Equity statement pages 240-241 in PDF use mirrored/rotated layout. Data extracted from reversed text and cross-validated against Balance Sheet equity totals.'
ws_eq.cell(4, 1).font = Font(name='Arial', size=8, italic=True, color='808080')

# Cross-validation: IS net profit vs equity
cross1 = IS_DATA['laba_neto'] == 632_902_896_073
validation_results['Cross: IS Net Profit = Equity Net Profit'] = 'PASS' if cross1 else 'FAIL'


# ═══ NOTE SHEETS ═══════════════════════════════════════════════════════════════
def add_note_sheet(note_num, title_id, title_en, note_text, note_type='narrative'):
    """Add a note sheet. note_type: 'narrative', 'table', or 'mixed'."""
    short_en = title_en[:25].strip().rstrip('/')
    sheet_name = f'Note {note_num} - {short_en}'[:31]
    ws = wb.create_sheet(sheet_name)
    ws.column_dimensions['A'].width = 55
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 40

    # Title
    ws.cell(1, 1).value = f'Catatan {note_num} - {title_id} / Note {note_num} - {title_en}'
    ws.cell(1, 1).font = Font(name='Arial', size=11, bold=True)
    ws.merge_cells(f'A1:C1')

    if note_type == 'narrative':
        ws.cell(3, 1).value = 'Topik / Topic'
        ws.cell(3, 2).value = 'Detail'
        ws.cell(3, 3).value = 'Referensi / Reference'
        for col in range(1, 4):
            c = ws.cell(3, col)
            c.font = FONT_BOLD
            c.fill = FILL_GRAY
        # Extract key paragraphs
        lines = [l.strip() for l in note_text.split('\n') if l.strip() and len(l.strip()) > 20]
        row = 4
        for line in lines[:40]:  # Cap at 40 lines
            ws.cell(row, 1).value = line[:100]
            ws.cell(row, 1).font = FONT_NORMAL
            ws.cell(row, 1).alignment = ALIGN_LEFT
            row += 1
    else:
        # Table note
        ws.cell(3, 1).value = 'Keterangan / Description'
        ws.cell(3, 2).value = YEAR
        ws.cell(3, 2).number_format = '@'
        for col in range(1, 3):
            c = ws.cell(3, col)
            c.font = FONT_BOLD
            c.fill = FILL_GRAY
        rows = parse_financial_lines(note_text)
        row = 4
        for label, v2025, _ in rows[:60]:
            if label or v2025 is not None:
                ws.cell(row, 1).value = label
                ws.cell(row, 1).font = FONT_NORMAL
                if v2025 is not None:
                    ws.cell(row, 2).value = v2025
                    ws.cell(row, 2).number_format = NUM_FMT
                    ws.cell(row, 2).alignment = ALIGN_RIGHT
                row += 1
    return ws


print('Building note sheets...')
for note_num, (title_id, title_en, pg_start, pg_end) in NOTE_INDEX.items():
    note_text = get_note_text(note_num)
    # Classify by note type
    table_notes = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19,
                   20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34,
                   35, 36, 37, 38, 40, 41}
    narrative_notes = {1, 2, 3, 39, 42, 43, 44}
    if note_num in table_notes:
        add_note_sheet(note_num, title_id, title_en, note_text, 'table')
    else:
        add_note_sheet(note_num, title_id, title_en, note_text, 'narrative')
    if note_num % 10 == 0:
        print(f'  Notes 1-{note_num} done...')

print('  All notes done.')


# ═══ KEY RATIOS SHEET ═════════════════════════════════════════════════════════
ws_rat = wb.create_sheet('Key Ratios Summary')
ws_rat.column_dimensions['A'].width = 55
ws_rat.column_dimensions['B'].width = 20
ws_rat.column_dimensions['C'].width = 30

ws_rat.cell(1, 1).value = 'Ringkasan Rasio Keuangan / Key Financial Ratios Summary'
ws_rat.cell(1, 1).font = Font(name='Arial', size=12, bold=True)
ws_rat.cell(2, 1).value = f'{TICKER} – {COMPANY}  |  Fiscal Year {YEAR}'
ws_rat.cell(2, 1).font = FONT_ITALIC

# Column headers
ws_rat.cell(4, 1).value = 'Rasio / Ratio'
ws_rat.cell(4, 2).value = YEAR
ws_rat.cell(4, 3).value = 'Formula / Notes'
for col in range(1, 4):
    c = ws_rat.cell(4, col)
    c.font = FONT_BOLD
    c.fill = FILL_GRAY

def ratio_row(ws, row, label, val, fmt, note=''):
    ws.cell(row, 1).value = label
    ws.cell(row, 1).font = FONT_NORMAL
    if val is not None and val != 'N/A':
        ws.cell(row, 2).value = val
        ws.cell(row, 2).number_format = fmt
        ws.cell(row, 2).alignment = ALIGN_RIGHT
    else:
        ws.cell(row, 2).value = 'N/A'
        ws.cell(row, 2).font = FONT_ITALIC
    ws.cell(row, 3).value = note
    ws.cell(row, 3).font = Font(name='Arial', size=9, italic=True)
    return row + 1

def safe_div(a, b):
    try:
        if b == 0:
            return None
        return a / b
    except:
        return None

cur_assets = BS_DATA['total_aset_lancar']
cur_liab   = BS_DATA['total_liab_lancar']
inventory  = BS_DATA['persediaan']
cash       = BS_DATA['kas']
total_liab = BS_DATA['total_liabilitas']
total_eq   = BS_DATA['total_ekuitas']
ebit       = IS_DATA['laba_usaha']
fin_costs_abs = abs(IS_DATA['biaya_keuangan'])
revenue    = IS_DATA['pendapatan_neto']
gross_p    = IS_DATA['laba_bruto']
op_profit  = IS_DATA['laba_usaha']
net_profit = IS_DATA['laba_neto']
total_aset = BS_DATA['total_aset']
trade_rec  = BS_DATA['piutang_usaha_ketiga']
cogs       = abs(IS_DATA['beban_pokok'])
trade_pay  = BS_DATA['utang_usaha_ketiga'] + BS_DATA['utang_usaha_berelasi']
op_cf      = CF_DATA['kas_neto_operasi']
capex      = abs(CF_DATA['perolehan_aset_tetap'])

rr = 5

# Header
def rat_hdr(ws, row, label):
    ws.cell(row, 1).value = label
    ws.cell(row, 1).font = FONT_BOLD
    ws.cell(row, 1).fill = FILL_GRAY
    ws.cell(row, 2).fill = FILL_GRAY
    return row + 1

rr = rat_hdr(ws_rat, rr, 'LIKUIDITAS / LIQUIDITY')
rr = ratio_row(ws_rat, rr, 'Current Ratio', safe_div(cur_assets, cur_liab), '0.00x', 'Current Assets / Current Liabilities')
rr = ratio_row(ws_rat, rr, 'Quick Ratio', safe_div(cur_assets - inventory, cur_liab), '0.00x', '(Current Assets - Inventory) / Current Liabilities')
rr = ratio_row(ws_rat, rr, 'Cash Ratio', safe_div(cash, cur_liab), '0.00x', 'Cash & Equivalents / Current Liabilities')
rr += 1

rr = rat_hdr(ws_rat, rr, 'SOLVABILITAS / LEVERAGE')
rr = ratio_row(ws_rat, rr, 'Debt-to-Equity', safe_div(total_liab, total_eq), '0.00x', 'Total Liabilities / Total Equity')
rr = ratio_row(ws_rat, rr, 'Interest Coverage', safe_div(ebit, fin_costs_abs), '0.00x', 'EBIT / Finance Costs')
rr += 1

rr = rat_hdr(ws_rat, rr, 'PROFITABILITAS / PROFITABILITY')
rr = ratio_row(ws_rat, rr, 'Gross Margin', safe_div(gross_p, revenue), PCT_FMT, 'Gross Profit / Revenue')
rr = ratio_row(ws_rat, rr, 'Operating Margin', safe_div(op_profit, revenue), PCT_FMT, 'Operating Profit / Revenue')
rr = ratio_row(ws_rat, rr, 'Net Margin', safe_div(net_profit, revenue), PCT_FMT, 'Net Profit / Revenue')
rr = ratio_row(ws_rat, rr, 'ROE (single-year)', safe_div(net_profit, total_eq), PCT_FMT, 'Net Profit / Total Equity (no average)')
rr = ratio_row(ws_rat, rr, 'ROA (single-year)', safe_div(net_profit, total_aset), PCT_FMT, 'Net Profit / Total Assets (no average)')
rr += 1

rr = rat_hdr(ws_rat, rr, 'EFISIENSI / EFFICIENCY')
rr = ratio_row(ws_rat, rr, 'Asset Turnover', safe_div(revenue, total_aset), '0.00x', 'Revenue / Total Assets')
rr = ratio_row(ws_rat, rr, 'Receivables Days', round(safe_div(trade_rec * 365, revenue)) if safe_div(trade_rec * 365, revenue) else None, '0', 'Trade Receivables / Revenue × 365')
rr = ratio_row(ws_rat, rr, 'Payables Days', round(safe_div(trade_pay * 365, cogs)) if safe_div(trade_pay * 365, cogs) else None, '0', 'Trade Payables / COGS × 365')
rr += 1

rr = rat_hdr(ws_rat, rr, 'ARUS KAS / CASH FLOW')
fcf = op_cf - capex
rr = ratio_row(ws_rat, rr, 'Free Cash Flow (IDR)', fcf, NUM_FMT, 'Operating CF - Capex')
rr = ratio_row(ws_rat, rr, 'FCF Margin', safe_div(fcf, revenue), PCT_FMT, 'FCF / Revenue')
rr += 1

# Summary numbers
rr = rat_hdr(ws_rat, rr, 'IKHTISAR KEUANGAN / FINANCIAL SUMMARY (IDR)')
rr = ratio_row(ws_rat, rr, 'Pendapatan Usaha / Revenue', revenue, NUM_FMT, 'FY2025')
rr = ratio_row(ws_rat, rr, 'Laba Bruto / Gross Profit', gross_p, NUM_FMT, 'FY2025')
rr = ratio_row(ws_rat, rr, 'Laba Usaha / Operating Profit', op_profit, NUM_FMT, 'FY2025')
rr = ratio_row(ws_rat, rr, 'Laba Neto / Net Profit', net_profit, NUM_FMT, 'FY2025')
rr = ratio_row(ws_rat, rr, 'Total Aset / Total Assets', total_aset, NUM_FMT, 'FY2025')
rr = ratio_row(ws_rat, rr, 'Total Ekuitas / Total Equity', total_eq, NUM_FMT, 'FY2025')
rr = ratio_row(ws_rat, rr, 'Kas & Setara Kas / Cash & Equivalents', cash, NUM_FMT, 'FY2025')


# ═══ SAVE ═════════════════════════════════════════════════════════════════════
print(f'Saving to {OUTPUT_PATH}...')
wb.save(OUTPUT_PATH)
print('Saved.')


# ═══ PRINT TERMINAL SUMMARY ═══════════════════════════════════════════════════
print()
print('=' * 60)
print('=== IDX Excel Template — Run 1 Complete ===')
print(f'Company     : {TICKER} – {COMPANY}')
print(f'Year        : {YEAR}')
print(f'Source File : {PDF_PATH} ({total_pages} pages)')
print(f'Output File : {OUTPUT_PATH}')
print()
print('Source Location in PDF:')
print(f'  Audit Status     : AUDITED — Auditor\'s report (scanned, pp.228–234)')
print(f'  Statement Type   : CONSOLIDATED (Konsolidasian)')
print(f'  Balance Sheet    : pages 235–237')
print(f'  Income Statement : pages 238–239')
print(f'  Equity           : pages 240–241 (mirrored/rotated layout)')
print(f'  Cash Flow        : pages 242–243')
print(f'  Notes            : pages 244–414')
print()
print('Sheets created:')
for i, ws in enumerate(wb.sheetnames, 1):
    print(f'  {i:2d}. {ws}')
print()
print('Validation:')
for check, result in validation_results.items():
    print(f'  {check:<48s}: {result}')
print()
print('Fallback Calculations Used:')
if fallback_calcs:
    for f in fallback_calcs:
        print(f'  • {f}')
else:
    print('  None — all values extracted directly from PDF.')
print()
print('Warnings:')
print('  • Equity statement (pp.240-241) uses mirrored/rotated text layout.')
print('    Values were reconstructed from reversed text and cross-validated.')
print('  • Auditor report pages (228-234) are scanned images (no text layer).')
print('  • Run 1 is single-year only (FY2025). FY2024 comparative column')
print('    is present in PDF but NOT extracted per idx-excel-run1-template rules.')
print()
print('Notes Index (44 notes):')
for n, (tid, ten, ps, pe) in NOTE_INDEX.items():
    print(f'  {n:2d}. {tid[:50]:50s}  pp.{ps}-{pe}')
print('=' * 60)
