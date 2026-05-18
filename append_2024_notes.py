"""
Phase 3d + Phase 4 (re-run): Note Index in English + Systematic 2024 sub-line extraction.

Pass 1 — Update Note Index column B to English titles (Indonesian kept in cell comment).
Pass 2 — Apply hardcoded NOTE_2024_DATA (verified totals from manual reading).
Pass 3 — Systematic PDF-driven fill for rows still empty in each note sheet's topic range.
Pass 4 — Append 2024-only PDF rows that have no matching workbook row.

Source: reports/WIFI_Annual_Report_2024.pdf
"""
import re
import openpyxl
import pdfplumber
from openpyxl.comments import Comment
from openpyxl.styles import Font

XLSX = 'WIFI_Financial_Statements.xlsx'
PDF = 'reports/WIFI_Annual_Report_2024.pdf'

# ============================================================
# Note matching: (sheet_name, 2024_note_num, id_title_2024, en_title_2024)
# When n2024 is None: note not in 2024 report.
# ============================================================
NOTE_MATCH = [
    ('Note 1 - GENERAL',                  1,  'UMUM',                                              'GENERAL'),
    ('Note 2 - SIGNIFICANT ACCOUNTING',   2,  'INFORMASI KEBIJAKAN AKUNTANSI MATERIAL',            'MATERIAL ACCOUNTING POLICY INFORMATION'),
    ('Note 3 - ACCOUNTING ESTIMATES A',   3,  'ESTIMASI DAN PERTIMBANGAN AKUNTANSI YANG SIGNIFIKAN','SIGNIFICANT ACCOUNTING ESTIMATES AND JUDGMENTS'),
    ('Note 4 - CASH AND CASH EQUIVALE',   4,  'KAS DAN SETARA KAS',                                'CASH AND CASH EQUIVALENTS'),
    ('Note 5 - TRADE RECEIVABLES',        5,  'PIUTANG USAHA',                                     'TRADE RECEIVABLES'),
    ('Note 6 - INVENTORIES',              None, None,                                              'INVENTORIES'),
    ('Note 7 - PREPAID EXPENSES AND O',   6,  'BEBAN DIBAYAR DI MUKA DAN ASET LANCAR LAINNYA',     'PREPAID EXPENSES AND OTHER CURRENT ASSETS'),
    ('Note 8 - ADVANCES',                 7,  'UANG MUKA',                                         'ADVANCES'),
    ('Note 9 - RESTRICTED FUND',          None, None,                                              'RESTRICTED FUND'),
    ('Note 10 - FIXED ASSETS',            8,  'ASET TETAP',                                        'FIXED ASSETS'),
    ('Note 11 - INTANGIBLE ASSETS',       9,  'ASET TAKBERWUJUD',                                  'INTANGIBLE ASSETS'),
    ('Note 12 - OTHER ASSETS',            None, None,                                              'OTHER ASSETS'),
    ('Note 13 - TRADE PAYABLES',         10,  'UTANG USAHA',                                       'TRADE PAYABLES'),
    ('Note 14 - OTHER PAYABLES',         11,  'UTANG LAIN-LAIN',                                   'OTHER PAYABLES'),
    ('Note 15 - ACCRUED EXPENSES',       12,  'BEBAN AKRUAL',                                      'ACCRUED EXPENSES'),
    ('Note 16 - ADVANCES FROM CUSTOME',  13,  'UANG MUKA PENJUALAN',                               'ADVANCES FROM CUSTOMERS'),
    ('Note 17 - LEASE LIABILITIES',      14,  'LIABILITAS SEWA',                                   'LEASE LIABILITIES'),
    ('Note 18 - SHORT-TERM BANK LOANS',  None, None,                                               'SHORT-TERM BANK LOANS'),
    ('Note 19 - TAXATION',               15,  'PERPAJAKAN',                                        'TAXATION'),
    ('Note 20 - CONSUMER FINANCING PA',  16,  'UTANG PEMBIAYAAN KONSUMEN',                         'CONSUMER FINANCING PAYABLES'),
    ('Note 21 - LONG-TERM BANK LOANS',   17,  'UTANG BANK JANGKA PANJANG',                         'LONG-TERM BANK LOANS'),
    ('Note 22 - BONDS PAYABLE',          18,  'UTANG OBLIGASI',                                    'BONDS PAYABLE'),
    ('Note 23 - LOANS',                  19,  'PINJAMAN',                                          'LOANS'),
    ('Note 24 - DUE TO RELATED PARTIE',  20,  'UTANG PIHAK BERELASI',                              'DUE TO A RELATED PARTY'),
    ('Note 25 - SHARIA BONDS',           None, None,                                               'SHARIA BONDS (SUKUK)'),
    ('Note 26 - EMPLOYEE BENEFITS LIA',  21,  'LIABILITAS IMBALAN KERJA',                          'EMPLOYEE BENEFITS LIABILITY'),
    ('Note 27 - SHARE CAPITAL AND ADV',  22,  'MODAL SAHAM',                                       'SHARE CAPITAL'),
    ('Note 28 - ADDITIONAL PAID-IN CA',  23,  'TAMBAHAN MODAL DISETOR',                            'ADDITIONAL PAID-IN CAPITAL'),
    ('Note 29 - NON-CONTROLLING INTER',  24,  'KEPENTINGAN NONPENGENDALI',                         'NON-CONTROLLING INTEREST'),
    ('Note 30 - REVENUES - NET',         25,  'PENDAPATAN USAHA - NETO',                           'REVENUES - NET'),
    ('Note 31 - COSTS OF REVENUES',      26,  'BEBAN POKOK PENDAPATAN',                            'COSTS OF REVENUES'),
    ('Note 32 - GENERAL AND ADMINISTR',  27,  'BEBAN OPERASIONAL',                                 'OPERATING EXPENSES'),
    ('Note 33 - OTHER INCOME (EXPENSE',  28,  'PENGHASILAN (BEBAN) LAIN-LAIN',                     'OTHER INCOMES (EXPENSES)'),
    ('Note 34 - FINANCE INCOME',         29,  'PENGHASILAN KEUANGAN',                              'FINANCE INCOME'),
    ('Note 35 - FINANCE COSTS',          30,  'BEBAN KEUANGAN',                                    'FINANCE EXPENSES'),
    ('Note 36 - EARNINGS PER SHARE',     31,  'LABA NETO PER SAHAM',                               'EARNINGS PER SHARE'),
    ('Note 37 - RELATED PARTY INFORMA',  32,  'INFORMASI PIHAK BERELASI',                          'RELATED PARTIES INFORMATION'),
    ('Note 38 - FINANCIAL INSTRUMENTS',  33,  'INSTRUMEN KEUANGAN',                                'FINANCIAL INSTRUMENTS'),
    ('Note 39 - FINANCIAL RISK MANAGE',  34,  'KEBIJAKAN DAN TUJUAN MANAJEMEN RISIKO KEUANGAN',    'FINANCIAL RISK MANAGEMENT OBJECTIVES AND POLICIES'),
    ('Note 40 - SEGMENT INFORMATION',    35,  'INFORMASI SEGMEN',                                  'SEGMENT INFORMATION'),
    ('Note 41 - SUPPLEMENTARY CASH FL',  36,  'INFORMASI TAMBAHAN ARUS KAS',                       'SUPPLEMENTARY CASH FLOWS INFORMATION'),
    ('Note 42 - SIGNIFICANT AGREEMENT',  37,  'IKATAN DAN KONTINJENSI',                            'SIGNIFICANT AGREEMENTS, COMMITMENTS AND CONTINGENCIES'),
    ('Note 43 - EVENTS AFTER REPORTIN',  38,  'PERISTIWA SETELAH PERIODE PELAPORAN',               'EVENTS AFTER REPORTING PERIOD'),
    ('Note 44 - ISSUANCE OF AMENDMENT',  39,  'PENERBITAN AMENDEMEN DAN PENYESUAIAN',              'ISSUANCE OF AMENDMENTS AND ADJUSTMENTS TO ACCOUNTING STANDARDS'),
]

# ============================================================
# 2024 note page ranges (from prior reading of the PDF)
# ============================================================
NOTE_PAGE_RANGES_2024 = {
    1:  (211, 220),  2:  (221, 238),  3:  (239, 243),  4:  (244, 244),
    5:  (245, 246),  6:  (246, 246),  7:  (247, 247),  8:  (248, 250),
    9:  (251, 251), 10:  (252, 253), 11:  (252, 253), 12:  (254, 254),
    13: (255, 256), 14:  (255, 256), 15:  (257, 260), 16:  (261, 261),
    17: (262, 276), 18:  (277, 277), 19:  (278, 280), 20:  (281, 283),
    21: (284, 284), 22:  (285, 285), 23:  (286, 286), 24:  (287, 287),
    25: (287, 287), 26:  (288, 288), 27:  (288, 288), 28:  (288, 288),
    29: (288, 289), 30:  (289, 289), 31:  (289, 289), 32:  (289, 292),
    33: (290, 290), 34:  (291, 292), 35:  (293, 293), 36:  (294, 294),
    37: (295, 324), 38:  (325, 335), 39:  (336, 339),
}

# Notes that contain rollforward tables — column extraction takes LAST number.
# Default for other notes is "first number = 2024 column" (two-column layout).
ROLLFORWARD_NOTES = {8, 9}  # Fixed Assets, Intangibles

# ============================================================
# Topic row ranges (1-indexed inclusive). Run 1 stuffed identical row
# contents into several adjacent note sheets. Restrict each sheet's
# auto-fill to the rows that genuinely belong to its topic.
# A sheet not listed below is auto-filled across its entire data range.
# ============================================================
TOPIC_ROW_RANGES = {
    'Note 31 - COSTS OF REVENUES':       [(17, 35)],   # COGS items + sub-totals + total
    'Note 32 - GENERAL AND ADMINISTR':   [(45, 65)],   # G&A items only (rows 17-44 are stale COGS clones)
    'Note 33 - OTHER INCOME (EXPENSE':   [(17, 26)],   # Other income items + net
    'Note 34 - FINANCE INCOME':          [(27, 33)],
    'Note 35 - FINANCE COSTS':           [(36, 46)],
    'Note 36 - EARNINGS PER SHARE':      [(48, 53)],
}

# ============================================================
# Hardcoded 2024 sub-line data (verified from manual reading of 2024 AR).
# Applied first; protects against keyword-match errors on totals.
# ============================================================
NOTE_2024_DATA = {
    'Note 7 - PREPAID EXPENSES AND O': {
        23: (16949690677, None),
        20: (3860002276, None),
        22: (12000000, None),
        24: (291828638, None),
        25: (21113521591, 'Sub-total Short-term Prepaid Expenses, 2024'),
        27: (9794106194, None),
        28: (30907627785, 'Sub-total Long-term Prepaid Expenses, 2024'),
        29: (302096542, None),
        30: (31209724327, 'Total Prepaid Expenses + Other Current Assets, 2024'),
        19: (None, 'Not found in 2024 report — frequency license fees appear from 2025'),
        21: (None, 'Not found in 2024 report — Consultant prepaid expense'),
    },
    'Note 8 - ADVANCES': {
        58: (148919000000, 'PT Lintas Daya Andalan — 2024 advance'),
        59: (None, 'Not found in 2024 report'),
        60: (None, 'Not found in 2024 report'),
        62: (None, 'Not found in 2024 report'),
        63: (None, 'Not found in 2024 report'),
    },
    'Note 13 - TRADE PAYABLES': {
        37: (33001370287, 'Sub-total / Total Trade Payables 2024 (matches BS)'),
        38: (33001370287, 'Total trade payables 2024'),
    },
    'Note 14 - OTHER PAYABLES': {
        53: (8029839382, 'Sub-total — current third parties 2024 (matches BS)'),
        60: (None, 'Long-term third parties total 2024 = 301,212,468,808 (matches BS). Composition differs from 2025 template.'),
    },
    'Note 15 - ACCRUED EXPENSES': {
        58: (5109714952, 'Total Accrued Expenses 2024 (matches BS)'),
    },
    'Note 16 - ADVANCES FROM CUSTOME': {
        37: (48018103366, 'PT XL Axiata Tbk 2024'),
        38: (15375000000, 'PT MNC Kabel Mediacom 2024'),
        40: (58000000000, 'PT Telemedia Komunikasi Pratama 2024'),
        41: (16456995698, 'Lain-lain 2024'),
        42: (137850099064, 'Total Advances from Customers 2024'),
        43: (-10290154592, 'Less: short-term portion 2024'),
        44: (127559944472, 'Long-term portion 2024 (matches BS row 52)'),
    },
    'Note 17 - LEASE LIABILITIES': {
        46: (214126485506, 'Utang angsuran 2024'),
        49: (-29875361411, 'Future finance charge 2024'),
        50: (184251124095, 'Present value of minimum payments 2024'),
        52: (-25717400118, 'Current maturity 2024 (matches BS row 40)'),
        53: (158533723977, 'Long-term portion 2024 (matches BS row 54)'),
    },
    'Note 19 - TAXATION': {
        23: (16256512269, 'Total Pajak Dibayar di Muka 2024 (matches BS row 11)'),
        34: (40185553486, 'Sub-total Perusahaan Utang Pajak 2024'),
        45: (72184515853, 'Sub-total Entitas Anak Utang Pajak 2024'),
        46: (112370069339, 'Total Utang Pajak 2024 (matches BS row 36)'),
        50: (-13049014880, 'Pajak kini - Perusahaan 2024 (sign flipped per IS expense convention)'),
        51: (-25237579300, 'Pajak kini - Entitas Anak 2024 (sign flipped per IS expense convention)'),
        52: (-38286594180, 'Total Pajak Kini 2024 (sign flipped per IS expense convention)'),
        54: (-327771270, 'Pajak tangguhan - Perusahaan 2024'),
        55: (5143393536, 'Pajak tangguhan - Entitas Anak 2024 (net benefit)'),
        56: (4815622266, 'Total Pajak Tangguhan 2024 (net benefit)'),
        58: (-43102216446, 'Beban pajak penghasilan - net 2024 (matches IS row 17)'),
    },
    'Note 20 - CONSUMER FINANCING PA': {
        17: (386633823, 'Current portion 2024 (matches BS row 41)'),
        20: (726049149, 'Long-term portion 2024 (matches BS row 55)'),
    },
    'Note 21 - LONG-TERM BANK LOANS': {
        46: (-138055277673, 'Current portion 2024 (matches BS row 42)'),
        47: (375305504520, 'Sub-total LT bank loans 2024'),
        50: (237250226847, 'Long-term portion 2024 (matches BS row 56)'),
    },
    'Note 22 - BONDS PAYABLE': {
        25: (614085074934, 'Sub-total Bonds 2024'),
        27: (-166632590847, 'Current portion 2024 (matches BS row 44)'),
        28: (447452484087, 'Long-term portion 2024 (matches BS row 58)'),
    },
    'Note 23 - LOANS': {
        45: (118122676901, 'Total Loans 2024'),
        47: (-83918554209, 'Current portion 2024 (matches BS row 43)'),
        48: (34204122692, 'Long-term portion 2024 (matches BS row 57)'),
    },
    'Note 24 - DUE TO RELATED PARTIE': {
        59: (39377574929, 'Total Due to Related Parties 2024 (matches BS row 50)'),
    },
    'Note 26 - EMPLOYEE BENEFITS LIA': {
        # Liabilities total 2024 (matches BS row 51)
        # Need to find correct row in this sheet; populate via systematic where possible
    },
    'Note 28 - ADDITIONAL PAID-IN CA': {
        62: (328521140531, 'Total APIC 2024 (matches BS row 66)'),
    },
    'Note 27 - SHARE CAPITAL AND ADV': {
        49: (235935511800, 'Issued and paid-in capital 2024 (matches BS row 65)'),
        48: (2359355118, 'Number of issued shares 2024'),
    },
    'Note 29 - NON-CONTROLLING INTER': {
        31: (181418175, 'Total NCI 2024 (matches BS)'),
    },
    'Note 30 - REVENUES - NET': {
        43: (671854001272, 'Net Revenue 2024 (matches IS)'),
    },
    'Note 31 - COSTS OF REVENUES': {
        35: (257080687384, 'Total Cost of Revenues 2024 (matches |IS|)'),
    },
    'Note 33 - OTHER INCOME (EXPENSE': {
        26: (-1207531677, 'Other income (expense) - net 2024 (matches IS row 11)'),
    },
    'Note 34 - FINANCE INCOME': {
        33: (538954781, 'Finance income 2024 (matches IS row 14)'),
    },
    'Note 35 - FINANCE COSTS': {
        46: (-73049522882, 'Finance costs 2024 (matches IS row 15)'),
    },
    'Note 36 - EARNINGS PER SHARE': {
        52: (231186780014, 'NP attributable to parent 2024'),
    },
}

# ============================================================
# Parsing helpers
# ============================================================
_STOP = {'dan', 'atau', 'yang', 'untuk', 'ke', 'dari', 'dengan', 'dalam',
         'the', 'and', 'or', 'of', 'to', 'from', 'for', 'in', 'on', 'a',
         'an', 'is', 'are', 'be', 'pt', 'tbk'}


def normalize_numbers(text):
    """Iteratively merge PDF-fragmented Indonesian thousand-separated numbers."""
    prev = None
    out = text
    while out != prev:
        prev = out
        out = re.sub(r'(\d\.\d{2})\s+(\d)(?!\d)', r'\1\2', out)
        out = re.sub(r'(\d\.\d)\s+(\d\d)(?!\d)', r'\1\2', out)
    out = re.sub(r'(?<![\w.])(\d{1,2})\s+(\d{1,2}\.\d{3}(?:\.\d{3})*)\b', r'\1\2', out)
    return out


def normalize_label(s):
    if s is None:
        return ''
    s = str(s).lower().strip()
    s = re.sub(r'[^\w\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def keyword_overlap(a, b):
    wa = {w for w in a.split() if w not in _STOP and len(w) > 2}
    wb = {w for w in b.split() if w not in _STOP and len(w) > 2}
    if not wa or not wb:
        return 0.0
    intersect = len(wa & wb)
    if intersect == 0:
        return 0.0
    # Favor short-label matches: when one side has ≤2 significant words,
    # measure coverage against the smaller side. Otherwise use Jaccard-style.
    if min(len(wa), len(wb)) <= 2:
        return intersect / min(len(wa), len(wb))
    return intersect / max(len(wa), len(wb))


def parse_id_number(s):
    s = s.strip()
    if not s or s in ('-', '—'):
        return None
    neg = s.startswith('(') and s.endswith(')')
    if neg:
        s = s[1:-1].strip()
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
        try:
            v = float(s)
        except ValueError:
            return None
    else:
        s = s.replace('.', '').replace(' ', '')
        try:
            v = int(s) if ',' not in s else float(s.replace(',', '.'))
        except ValueError:
            return None
    return -v if neg else v


_NUM_TOKEN_RE = re.compile(
    r'\(\s*\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?\s*\)'
    r'|'
    r'\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?'
)


def extract_numbers_and_label(line):
    """From a normalized line, extract (label_id, label_en, [numbers]).
    Label_id = text before first numeric token. Label_en = text after last numeric token.
    """
    matches = list(_NUM_TOKEN_RE.finditer(line))
    nums = []
    for m in matches:
        v = parse_id_number(m.group(0))
        if v is not None:
            nums.append(v)
    if not matches:
        return line.strip(), '', []
    label_id = line[:matches[0].start()].strip()
    label_en = line[matches[-1].end():].strip()
    return label_id, label_en, nums


def _is_noise_value(v):
    """A single isolated number that's probably a note-ref, year, or article number."""
    if isinstance(v, float) and 0 < v < 100:
        return True
    if isinstance(v, int):
        if 1990 <= v <= 2100:
            return True
        if 0 <= abs(v) < 100:
            return True
    return False


_NOTE_HEADER_RE = re.compile(r'^\s*(\d{1,2})\.\s+[A-Z][A-Z]')


def extract_2024_note_rows(pdf, start_page, end_page, rollforward=False, target_note=None):
    """Return list of dicts with keys: label_id, label_en, value, line.
    For non-rollforward layout, value = first numeric column (2024).
    For rollforward, value = last numeric column (Saldo Akhir 2024).
    Skips lines inside a 2023-only section.
    If target_note is given, only rows inside that note's section on each page are kept.
    """
    rows = []
    current_year_context = None  # None | 2024 | 2023 | 'BOTH'
    for pn in range(start_page, end_page + 1):
        txt = pdf.pages[pn - 1].extract_text() or ''
        txt = normalize_numbers(txt)
        # Track which note we're currently inside on this page (None until first note header)
        in_target_note = target_note is None  # if no target filter, always True
        for raw in txt.split('\n'):
            line = raw.rstrip()
            stripped = line.strip()
            if not stripped:
                continue
            # Detect note header lines: "26. BEBAN..." etc.
            m = _NOTE_HEADER_RE.match(stripped)
            if m and 'lanjutan' not in stripped.lower() and 'continued' not in stripped.lower():
                note_num = int(m.group(1))
                if target_note is not None:
                    in_target_note = (note_num == target_note)
                current_year_context = None
                continue
            if target_note is not None and not in_target_note:
                continue
            # Year-context detection
            if re.fullmatch(r'2024(\s+/\s+2024)?', stripped):
                current_year_context = 2024
                continue
            if re.fullmatch(r'2023(\s+/\s+2023)?', stripped):
                current_year_context = 2023
                continue
            if re.search(r'\b2024\b.*\b2023\b', stripped) and 'Rp' not in stripped:
                current_year_context = 'BOTH'
                continue
            if current_year_context == 2023:
                continue
            label_id, label_en, nums = extract_numbers_and_label(line)
            if not nums:
                continue
            # Filter out lines whose only number is a note-reference / year / single-digit
            real_nums = [n for n in nums if not _is_noise_value(n)]
            if not real_nums:
                continue
            if rollforward and current_year_context == 2024:
                value = real_nums[-1]
            elif current_year_context == 'BOTH':
                value = real_nums[0]
            elif current_year_context == 2024:
                value = real_nums[-1] if rollforward else real_nums[0]
            else:
                value = real_nums[0]
            rows.append({
                'label_id': label_id,
                'label_en': label_en,
                'value': value,
                'all_nums': real_nums,
                'line': stripped,
                'page': pn,
                'consumed': False,
            })
    return rows


# Generic labels — too vague to match by keyword alone; skip auto-matching.
_GENERIC_LABELS = {
    'total', 'total total', 'sub total', 'sub total sub total',
    'neto net', 'net', 'neto', 'jumlah', 'jumlah total',
    'catatan note', 'note', 'catatan', 'pasal article',
    'pasal', 'article', 'perusahaan the company',
    'entitas anak subsidiaries', 'lain lain others', 'lain lain',
}


def match_workbook_label(sheet_label, pdf_rows):
    """Find best PDF row index matching the workbook label.
    Returns (index, match_type) or None.
    """
    sl = normalize_label(sheet_label)
    if not sl or len(sl) < 4:
        return None
    if sl in _GENERIC_LABELS:
        return None
    # Skip labels that are mostly punctuation/dashes (after normalize they're tiny)
    if len([w for w in sl.split() if len(w) > 2]) == 0:
        return None

    # Pass 1: exact match against ID or EN label (normalized)
    for i, r in enumerate(pdf_rows):
        if r['consumed']:
            continue
        n_id = normalize_label(r['label_id'])
        n_en = normalize_label(r['label_en'])
        if sl == n_id or sl == n_en:
            return (i, 'exact')

    # Pass 2: substring match (one direction)
    for i, r in enumerate(pdf_rows):
        if r['consumed']:
            continue
        n_id = normalize_label(r['label_id'])
        n_en = normalize_label(r['label_en'])
        for cand in (n_id, n_en):
            if not cand or len(cand) < 6:
                continue
            short, long = (sl, cand) if len(sl) <= len(cand) else (cand, sl)
            if len(short) >= 6 and short in long:
                return (i, 'substring')

    # Pass 3: keyword overlap ≥ 0.6
    best = None
    best_score = 0.0
    for i, r in enumerate(pdf_rows):
        if r['consumed']:
            continue
        for cand in (normalize_label(r['label_id']), normalize_label(r['label_en'])):
            if not cand:
                continue
            score = keyword_overlap(sl, cand)
            if score >= 0.6 and score > best_score:
                best = i
                best_score = score
    if best is not None:
        return (best, 'keyword')
    return None


# ============================================================
# Note Index update — switch col B to English titles
# ============================================================
def update_note_index_english(wb):
    if 'Note Index' not in wb.sheetnames:
        return
    ni = wb['Note Index']
    for i, (sname, n2024, title_id, title_en) in enumerate(NOTE_MATCH):
        r = 3 + i
        a = ni.cell(row=r, column=1).value
        if not a or 'Note' not in str(a):
            continue
        cell_b = ni.cell(row=r, column=2)
        if n2024 is None:
            cell_b.value = None
            cell_b.comment = Comment('Note not present in 2024 report', 'Append2024')
        else:
            cell_b.value = title_en
            note_in_workbook = int(str(a).split()[1])
            if note_in_workbook != n2024:
                cmt = (f"In 2024 this topic was Note {n2024}. "
                       f"Indonesian title: {title_id}")
            else:
                cmt = f"Indonesian title: {title_id}"
            cell_b.comment = Comment(cmt, 'Append2024')


# ============================================================
# Pass 2 — Apply hardcoded NOTE_2024_DATA
# ============================================================
def apply_hardcoded(wb):
    for sheet_name, row_map in NOTE_2024_DATA.items():
        if sheet_name not in wb.sheetnames:
            continue
        s = wb[sheet_name]
        for row_num, (val, cmt) in row_map.items():
            if val is not None:
                s.cell(row=row_num, column=2, value=val)
            if cmt:
                s.cell(row=row_num, column=2).comment = Comment(cmt, 'Append2024')


# ============================================================
# Pass 3 — Systematic PDF-driven fill
# ============================================================
def topic_rows_for(sheet_name, max_row):
    ranges = TOPIC_ROW_RANGES.get(sheet_name)
    if not ranges:
        return list(range(4, max_row + 1))
    out = []
    for lo, hi in ranges:
        out.extend(range(lo, hi + 1))
    return out


def systematic_fill(wb, pdf):
    summary = []
    for sheet_name, n2024, _ti, _te in NOTE_MATCH:
        if n2024 is None or sheet_name not in wb.sheetnames:
            continue
        if n2024 not in NOTE_PAGE_RANGES_2024:
            continue
        start, end = NOTE_PAGE_RANGES_2024[n2024]
        # Skip narrative-only or extremely complex notes
        if n2024 in {1, 2, 3, 33, 34, 39}:
            continue
        rollforward = n2024 in ROLLFORWARD_NOTES
        pdf_rows = extract_2024_note_rows(pdf, start, end, rollforward=rollforward, target_note=n2024)
        if not pdf_rows:
            continue

        s = wb[sheet_name]
        rows_to_check = topic_rows_for(sheet_name, s.max_row)
        matched, skipped_existing = 0, 0
        for r in rows_to_check:
            a = s.cell(row=r, column=1).value
            c = s.cell(row=r, column=3).value
            b_cell = s.cell(row=r, column=2)
            if not isinstance(c, (int, float)) or a is None:
                continue
            # Don't overwrite hardcoded values from Pass 2
            if b_cell.value is not None:
                skipped_existing += 1
                continue
            m = match_workbook_label(str(a), pdf_rows)
            if m is None:
                continue
            idx, mtype = m
            value = pdf_rows[idx]['value']
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            # Magnitude sanity: if column C is large and value is tiny (or vice versa)
            # by more than 1000x, treat as a spurious match.
            if isinstance(c, (int, float)) and isinstance(value, (int, float)):
                if abs(c) > 1_000_000 and abs(value) < abs(c) / 1000 and abs(value) < 1_000_000:
                    continue
                if abs(value) > 1_000_000 and abs(c) < abs(value) / 1000 and abs(c) < 1_000_000:
                    continue
            pdf_rows[idx]['consumed'] = True
            b_cell.value = value
            if mtype != 'exact':
                pdf_label = pdf_rows[idx]['label_id'] or pdf_rows[idx]['label_en']
                b_cell.comment = Comment(
                    f"Matched from PDF p.{pdf_rows[idx]['page']}: '{pdf_label}' ({mtype})",
                    'Append2024')
            matched += 1

        # Build set of values already in column B of this sheet (across all rows)
        # to dedup against the append step.
        existing_b_values = set()
        for r2 in range(4, s.max_row + 1):
            v = s.cell(row=r2, column=2).value
            if isinstance(v, (int, float)):
                existing_b_values.add(int(v) if isinstance(v, float) and v.is_integer() else v)

        # Pass 4: append unconsumed PDF rows as 2024-only items
        unmatched = [r for r in pdf_rows
                     if not r['consumed']
                     and r['label_id']
                     and _is_appendable(r, existing_b_values)]
        appended = append_2024_only(s, unmatched, n2024)
        summary.append((sheet_name, n2024, matched, appended, skipped_existing))
    return summary


def _is_appendable(row, existing_b_values):
    label = normalize_label(row['label_id'])
    significant = [w for w in label.split() if len(w) > 2 and w not in _STOP]
    if len(significant) < 1:
        return False
    if label in _GENERIC_LABELS:
        return False
    v = row['value']
    if isinstance(v, float) and not v.is_integer():
        return False  # decimal floats are usually rates / percentages / noise
    if isinstance(v, int) and abs(v) < 10_000:
        return False
    if 1990 <= abs(v) <= 2100 and isinstance(v, int):
        return False
    key = int(v) if isinstance(v, float) and v.is_integer() else v
    if key in existing_b_values:
        return False
    return True


def append_2024_only(sheet, unmatched_rows, note_num):
    if not unmatched_rows:
        return 0
    next_row = sheet.max_row + 2
    cell = sheet.cell(row=next_row, column=1,
                      value=f'Catatan: Baris berikut hanya muncul di Laporan Tahunan 2024 (Catatan {note_num}) / '
                            f'Note: Rows added from 2024 Annual Report (Note {note_num})')
    cell.font = Font(italic=True)
    next_row += 1
    appended = 0
    seen = set()
    for r in unmatched_rows:
        label = f"  {r['label_id']} / {r['label_en']}".strip(' /')
        key = normalize_label(label)
        if key in seen:
            continue
        seen.add(key)
        value = r['value']
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        sheet.cell(row=next_row, column=1, value=label)
        b = sheet.cell(row=next_row, column=2, value=value)
        b.comment = Comment(
            f"From 2024 AR p.{r['page']}. Line: \"{r['line'][:120]}\"",
            'Append2024')
        next_row += 1
        appended += 1
    return appended


# ============================================================
# Main
# ============================================================
def main():
    print(f'Loading workbook {XLSX}')
    wb = openpyxl.load_workbook(XLSX)

    print('Pass 1 — Updating Note Index column B to English titles')
    update_note_index_english(wb)

    print('Pass 2 — Applying hardcoded verified totals')
    apply_hardcoded(wb)

    print(f'Pass 3+4 — Systematic PDF extraction from {PDF}')
    pdf = pdfplumber.open(PDF)
    try:
        summary = systematic_fill(wb, pdf)
    finally:
        pdf.close()

    print()
    print(f'{"Sheet":36s} {"2024N":5s} {"Matched":>7s} {"Appended":>8s} {"Skipped":>7s}')
    for sname, n, m, a, sk in summary:
        print(f'{sname:36s} {n:5d} {m:7d} {a:8d} {sk:7d}')
    print()
    print(f'Saving {XLSX}')
    wb.save(XLSX)
    print('Done.')


if __name__ == '__main__':
    main()
