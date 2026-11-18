#!/usr/bin/env python3
"""Tạo sổ kiểm thực (bước 2, bước 3, mẫu 24 giờ) cho 1 ngày từ file THỰC ĐƠN VIỆT.

Cách chạy:
    python3 tao_so.py <thực_đơn_việt.xlsx> <việt_mẫu.xlsx> <ngày, vd 23.10> [thư_mục_ra]
"""
import ast, collections, copy, datetime, difflib, io, json, operator, os, re, sys, unicodedata, zipfile
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.pagebreak import Break

CA = ['SÁNG', 'TRƯA', 'TỐI', 'ĐÊM']
SHEET2 = {'SÁNG': ' bước 2(SÁNG)', 'TRƯA': ' bước 2 (TRƯA', 'TỐI': ' bước 2 (TỐI)', 'ĐÊM': ' bước 2 (ĐÊM)'}
SHEET3, SHEET4 = ' bước 3) (2)', 'MẪU 24 giờ'
VANG = PatternFill('solid', fgColor='FFF2CC')
CAM = PatternFill('solid', fgColor='F4B183')
VANG_DAM = PatternFill('solid', fgColor='FFFF00')
SO_THAT = 'General'    # mẫu để 0.0 hoặc 0.00 làm Excel hiện 2.25 thành 2.3; General hiện đúng số đã tính

# Quy tắc người dùng sửa được trong app; quy_tac.json chỉ chứa phần đã đổi so với mặc định
MAC_DINH = {
    'nam': datetime.date.today().year,
    'cong_thuc': 'dinh_luong * suat / 1000',
    'cong_thuc_rieng': [],                      # [{bat_dau, cong_thuc}] cho nguyên liệu tính khác
    'cho_phep_lech': 0.011,
    'khi_lech': 'thuc_don',                     # thực đơn là căn cứ; công thức chỉ để phát hiện lệch
    'dl_dem': 1000, 'don_vi_can': 'KG', 'don_vi_dem': 'CÁI',
    'don_vi_theo_ten': [{'bat_dau': 'TRỨNG', 'don_vi': 'QUẢ'}, {'bat_dau': 'TVL', 'don_vi': 'QUẢ'},
                        {'bat_dau': 'QUẢ', 'don_vi': 'QUẢ'}, {'bat_dau': 'MÌ TÔM', 'don_vi': 'GÓI'},
                        {'bat_dau': 'MỲ TÔM', 'don_vi': 'GÓI'}, {'bat_dau': 'LÁ NEM', 'don_vi': 'LÁ'}],
    'dl_com': 170, 'dl_gao_toi_thieu': 50, 'dl_thit_canh': 3,
    'tu_khac_phuc': 'co',                       # 'co': tự xếp món theo vị trí dòng
    'tu_dien_suat': 'khong',                    # sổ hoàn thiện không ghi bộ món để trống suất
    'sai_so_kg': 0.5, 'sai_so_pt': 5,           # lệch trong các mức này là sai số nhỏ, chấp nhận được
    'mon_thieu_nl': [{'mon': 'TRỨNG ỐP', 'nl': 'TRỨNG GÀ CN >60G/EA', 'moi_suat': 1, 'don_vi': 'QUẢ'}],
    'doi_nl': [{'bat_dau': 'GẠO THƠM', 'ten': 'GẠO DẺO TA'}, {'bat_dau': 'GẠO NẾP CÁI HOA VÀNG', 'ten': 'GẠO NẾP SÁP NGỖNG'},
               {'bat_dau': 'THỊT GÀ GÓC TƯ', 'ten': 'THỊT GÀ GÓC TƯ'}, {'bat_dau': 'BÍ XANH LUỘC', 'ten': 'BÍ XANH'}],
    'dl_canh_bao': 15,                          # dòng từ mức này chưa khớp món nào thì tô vàng
    'gao_khong_phai_com': ['GẠO KHANG DÂN', 'GẠO NẾP', 'GẠO NỨT'],
    'tuan_bo_qua': ['TRỨNG GÀ, KIM CHI', 'TRỨNG GÀ,SỮA'],
    'kiem_dinh_luong': 'co',                    # cảnh báo định lượng ngoài khoảng thường dùng
    'trai_cay': [{'ten': 'DƯA HẤU ', 'don_vi': 'KG'}, {'ten': 'THANH LONG', 'don_vi': 'KG'},
                 {'ten': 'CHUỐI', 'don_vi': 'QUẢ'}, {'ten': 'ỔI', 'don_vi': 'KG'}],
    'hang_cuoi_ca': [{'ten': 'THẠCH', 'suat': 100, 'khoi_luong': '1 HỘP', 'dung_cu': 'CÔNG INOX'},
                     {'ten': 'SỮA CHUA', 'suat': None, 'khoi_luong': '1 HỘP', 'dung_cu': 'CÔNG INOX'},
                     {'ten': 'SỮA IZZI', 'suat': None, 'khoi_luong': '1 HỘP', 'dung_cu': 'CÔNG INOX'},
                     {'ten': 'SỮA FAMI', 'suat': None, 'khoi_luong': '1 HỘP', 'dung_cu': 'CÔNG INOX'}],
    # Tên sữa trong thực đơn khác tên trong sổ, dùng để gợi ý số suất
    'goi_y_sua': [],
    'mau_khoi_luong': '150G', 'mau_canh': '150ML', 'mau_dung_cu': 'HỘP INOX',
    'luu_cong': [{'bat_dau': 'BÁNH MỲ', 'khoi_luong': '1 HỘP', 'dung_cu': 'CÔNG INOX'},
                 {'bat_dau': 'BÁNH MÌ', 'khoi_luong': '1 HỘP', 'dung_cu': 'CÔNG INOX'}],
    'tach_ba_dong': [{'bat_dau': 'BÚN'}, {'bat_dau': 'PHỞ'}, {'bat_dau': 'BÁNH ĐA'}, {'bat_dau': 'MÌ GÀ TIỀM'}],
    'khong_tach': ['XÀO', 'CUỐN'],
    'bo_nguyen_lieu': ['SỮA', 'SUA', 'TM', 'THẠCH', 'TRÁNG MIỆNG', 'TỔNG KCAL', 'KCAL', 'YAOURT'],
    'bo_bo_mon': ['BÀ BẦU', 'TRÁNG MIỆNG'],
    # Viết tắt / cách gọi trong thực đơn -> tên chuẩn trong sổ
    'doi_ten': [{'tim': r'\bSU SU\s*,\s*CÀ RỐT LUỘC', 'thay': 'SU SU CÀ RỐT LUỘC'},
                {'tim': r'\bSET NGÔ KHOAI LUỘC\s*\+\s*TRỨNG', 'thay': 'NGÔ LUỘC, KHOAI LANG, TRỨNG LUỘC'},
                {'tim': r'\bCƠM ĐÔI\s+', 'thay': ''}, {'tim': r'\bTVL\b', 'thay': 'TRỨNG VỊT LỘN'},
                {'tim': r'\bXÔI\b(?=.*\bSC\b)', 'thay': 'XÔI'}, {'tim': r'\bCẢI CẢI\b', 'thay': 'CẢI'},
                {'tim': r'\bRAU CẢI\b', 'thay': 'CẢI'}, {'tim': r'\bBÍ NGÔ\b', 'thay': 'BÍ ĐỎ'},
                {'tim': r'BÚN RIÊU CUA\s*,\s*TÓP MỠ', 'thay': 'BÚN RIÊU CUA TÓP MỠ'},
                {'tim': r'\bCƠM (TẤM|NIÊU)\s+', 'thay': ''}],
}
Q = dict(MAC_DINH)
BO_NGUYEN_LIEU = None


def dung_quy_tac(q):
    """Đặt bộ quy tắc đang dùng, phần thiếu lấy theo mặc định."""
    global Q, BO_NGUYEN_LIEU
    Q = {**MAC_DINH, **q}
    # Từ ngắn phải khớp nguyên từ để "TM" không bắt nhầm tên khác
    mau = [re.escape(x.upper()) + (r'\b' if len(x) <= 3 else '') for x in Q['bo_nguyen_lieu'] if x.strip()]
    BO_NGUYEN_LIEU = re.compile('^(' + '|'.join(mau) + ')') if mau else re.compile(r'(?!)')


# Nơi lưu những gì người dùng sửa (quy tắc, sổ tay món). Trên máy chủ đám mây là ổ đĩa riêng SKT_DU_LIEU.
THU_MUC_GHI = os.environ.get('SKT_DU_LIEU') or os.path.dirname(os.path.abspath(__file__))


def file_ghi(ten):
    """File người dùng sửa được: đọc ở THU_MUC_GHI, chưa có thì dùng bản đi kèm app."""
    p = os.path.join(THU_MUC_GHI, ten)
    return p if os.path.exists(p) else os.path.join(os.path.dirname(os.path.abspath(__file__)), ten)


def tai_quy_tac():
    p = os.path.join(THU_MUC_GHI, 'quy_tac.json')
    return json.load(open(p, encoding='utf-8')) if os.path.exists(p) else {}


def la_gao(nl):
    k = khoa_nl(nl)
    return k.startswith('GAO') or 'JASMINE' in k

def la_com(nl, dl):
    """Gạo nấu cơm trắng: gạo cháo định lượng rất nhỏ, gạo cơm rang, nếp, gạo nứt thuộc món riêng."""
    return la_gao(nl) and dl >= Q['dl_gao_toi_thieu'] \
        and not any(khoa_nl(nl).startswith(khoa_nl(x)) for x in Q['gao_khong_phai_com'])


_PHEP = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul, ast.Div: operator.truediv}

def tinh(bieu_thuc, dinh_luong, suat):
    """Tính công thức người dùng nhập; chỉ cho phép số, 2 biến và + - * / ()."""
    bien = {'dinh_luong': dinh_luong, 'suat': suat}
    def g(n):
        if isinstance(n, ast.Expression): return g(n.body)
        if isinstance(n, ast.BinOp) and type(n.op) in _PHEP: return _PHEP[type(n.op)](g(n.left), g(n.right))
        if isinstance(n, ast.UnaryOp) and isinstance(n.op, ast.USub): return -g(n.operand)
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)) and not isinstance(n.value, bool): return n.value
        if isinstance(n, ast.Name) and n.id in bien: return bien[n.id]
        raise ValueError(f'Công thức không hợp lệ: {bieu_thuc}')
    try:
        return g(ast.parse(str(bieu_thuc).replace(',', '.'), mode='eval'))
    except (SyntaxError, ZeroDivisionError) as e:
        raise ValueError(f'Công thức không hợp lệ: {bieu_thuc}') from e


def cong_thuc_cua(nl):
    k = khoa_nl(nl)
    for x in Q['cong_thuc_rieng']:
        if k.startswith(khoa_nl(x['bat_dau'])): return x['cong_thuc']
    return Q['cong_thuc']

def dien_giai(ct, dl, suat):
    return ct.replace('dinh_luong', str(dl)).replace('suat', str(suat)).replace('*', 'x')


BO_PHAN_MON = re.compile(r'^(SC|SỮA|SỮA CHUA|SỮA FAMI|SỮA IZZI|TM|TRÁNG MIỆNG|BÀ BẦU.*)$')
# Từ không dùng để so nguyên liệu với tên món
TU_BO = set('XAO LUOC CHIEN NAU RANG HAP SOT XOT OM KHO TRAN RAN VA MON CO'.split())
# Nguyên liệu không xuất hiện trong tên món -> từ khoá món
GOI_Y_MANH = {'RAU RAM': 'TRUNG VIT LON', 'TVL': 'TRUNG VIT LON', 'BI DO': 'BI NGO'}
GOI_Y = {'NGO NGOT': 'NGO', 'BOT MI': 'BANH', 'MA GA': 'GA', 'DUI GA': 'GA',
         'NAC VAI': 'THIT', 'BA CHI': 'THIT', 'KHOANH GIO': 'THIT', 'CAI CANH': 'CANH'}
RAU_SONG = ('HUNG QUE', 'XA LACH', 'TIA TO', 'KINH GIOI')
_mon_biet = {'ds': None, 'dl': None}
TEN_CA = {'SANG': 'SÁNG', 'TRUA': 'TRƯA', 'TOI': 'TỐI', 'DEM': 'ĐÊM'}


def khong_dau(s):
    s = unicodedata.normalize('NFD', str(s)).upper().replace('Đ', 'D').replace('Ð', 'D')
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')

def tu(s):
    return {t for t in re.findall(r'[A-Z0-9]+', khong_dau(s)) if t not in TU_BO}

def la_canh(ten):
    return unicodedata.normalize('NFC', ten).strip().upper().startswith('CANH ')

def khoa_mon(ten):
    s = khong_dau(ten)
    s = re.sub(r'\b(NAU|XANH)\b', ' ', s).replace('SU SU', 'SUSU').replace('XOT', 'SOT')
    if 'BASA' in s.replace(' ', ''): s = s.replace('BA SA', 'BASA')
    return ' '.join(s.split())

def khoa_nl(ten):
    s = khong_dau(ten)
    s = re.sub(r'\(.*?\)', ' ', s)
    return ' '.join(s.replace(' DL', ' ').split()).replace('THIT NAC VAI', 'NAC VAI').replace('XUONG ONG LON', 'XUONG ONG').replace('XA LACH THUONG', 'XA LACH')


def mo_wb(path, **kw):
    """Một số file có custom property lỗi khiến openpyxl không đọc được; bỏ phần đó khi đọc."""
    b = io.BytesIO()
    with zipfile.ZipFile(path) as z, zipfile.ZipFile(b, 'w', zipfile.ZIP_DEFLATED) as w:      # Windows khoá file còn mở
        for i in z.infolist():
            if i.filename == 'docProps/custom.xml':
                w.writestr(i, b'<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"/>')
            else:
                w.writestr(i, z.read(i.filename))
    b.seek(0)
    return openpyxl.load_workbook(b, **kw)


def ngay_cua(ws):
    """Ngày của sheet thực đơn, 'dd.mm'. Đầu sheet ghi 'ngày 23.10', 'ca sáng 26/10', 'SÁNG 29.10', 'BỮA SÁNG 31.10'."""
    if ws.sheet_state != 'visible': return None
    t = ' '.join(str(v) for row in ws.iter_rows(min_row=1, max_row=3, max_col=3, values_only=True)
                 for v in row if v is not None).upper()
    m = re.search(r'SÁNG\s+(?:NGÀY\s*)?(\d{1,2})\s*[./]\s*(\d{1,2})(?![\d./])', t)
    if not m and 'SÁNG' in t: m = re.search(r'NGÀY\s*(\d{1,2})\s*[./]\s*(\d{1,2})(?![\d./])', t)
    if not m or not (1 <= int(m.group(1)) <= 31 and 1 <= int(m.group(2)) <= 12): return None
    return f'{int(m.group(1)):02d}.{int(m.group(2)):02d}'


def tim_sheet(wb, ngay):
    """Tìm sheet theo ngày ở đầu sheet, không dựa vào tên sheet vì tên không thống nhất."""
    if not re.fullmatch(r'\d{1,2}\.\d{1,2}', str(ngay)): raise SystemExit('Hãy chọn ngày')
    d, m = ngay.split('.')
    for ws in wb.worksheets:
        if ngay_cua(ws) == f'{int(d):02d}.{int(m):02d}': return ws
    raise SystemExit(f'Không tìm thấy sheet của ngày {ngay}')


def doc_ngay(ws):
    """Đọc sheet ngày -> {ca: [dòng]}; mỗi dòng: bộ món gốc + nguyên liệu + định lượng + suất + tổng."""
    gop = {}
    for m in ws.merged_cells.ranges:
        if m.min_col == 2:
            for r in range(m.min_row, m.max_row + 1): gop[r] = m.min_row
    out, ca, ten, moi = {c: [] for c in CA}, 'SÁNG', None, True
    out['_mo_coi'], khoi = [], 0
    co_tieu_de = set()
    for r in range(1, ws.max_row + 1):
        dau = ' '.join(str(x) for x in (ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value) if isinstance(x, str)).upper()
        m = re.search(r'\bCA\s*(SÁNG|TRƯA|TỐI|ĐÊM)|^\s*(?:BỮA\s*)?(SÁNG|TRƯA|TỐI|ĐÊM)\s*\d', dau)
        if m and not isinstance(ws.cell(r, 4).value, (int, float)): co_tieu_de.add(m.group(1) or m.group(2))
    for r in range(1, ws.max_row + 1):
        b, c, d, e, f = [ws.cell(r, i).value for i in (2, 3, 4, 5, 6)]
        dau = ' '.join(str(x) for x in (b, c) if isinstance(x, str)).upper()
        m = re.search(r'\bCA\s*(SÁNG|TRƯA|TỐI|ĐÊM)|^\s*(?:BỮA\s*)?(SÁNG|TRƯA|TỐI|ĐÊM)\s*\d', dau)
        if m and not isinstance(d, (int, float)):
            ca, ten, moi = m.group(1) or m.group(2), None, True
            continue
        # Nhiều sheet thiếu tiêu đề ca thì dựa vào cột ca ở cuối dòng. Có tiêu đề thì tin tiêu đề vì cột này hay ghi sót
        g = TEN_CA.get(khong_dau(ws.cell(r, 7).value or '').strip())
        if g and g not in co_tieu_de and CA.index(g) > CA.index(ca) and isinstance(d, (int, float)): ca, ten, moi = g, None, True
        if str(ws.cell(r, 1).value or '').strip().upper() == 'CỬA':
            if not moi and CA.index(ca) < 3: ca, ten = CA[CA.index(ca) + 1], None
            moi = False
            continue
        if r in gop: b = ws.cell(gop[r], 2).value
        # Dòng gạo đôi khi nằm ngay trên ô tên bộ món, cùng số suất với bộ món đó
        sau = ws.cell(r + 1, 2).value
        if not (isinstance(b, str) and b.strip()) and r not in gop and isinstance(c, str) and la_gao(c) \
                and isinstance(sau, str) and sau.strip() and e == ws.cell(r + 1, 5).value:
            b, khoi = sau, r + 1
        elif isinstance(b, str) and b.strip():
            khoi = gop.get(r, r)
        if isinstance(b, str) and b.strip(): ten = b.strip()
        if not isinstance(c, str) or not c.strip() or c.strip().upper() == 'NGUYÊN LIỆU': continue
        if not isinstance(d, (int, float)): continue
        if ten is None:
            out['_mo_coi'].append((ca, r, '', c.strip(), 'Dòng không thuộc bộ món nào nên không đưa vào sổ'))
            continue
        out[ca].append(dict(row=r, set=ten, nl=c.strip(), nl_goc=c, dl=d, khoi=khoi, suat=e if isinstance(e, (int, float)) and e > 0 else None,
                            tong=f if isinstance(f, (int, float)) else None))
    return out


def tach_mon(ten):
    s = ten.upper()
    for x in Q['doi_ten']: s = re.sub(x['tim'], x['thay'], s)
    s = re.sub(r'(?<=[^\s,+/;.])\s+(?=CANH\s)', ', ', s)      # thiếu dấu phẩy trước món canh
    co_rs = 'RAU SỐNG' in s
    s = re.sub(r'\bRAU SỐNG\b', ',', s)
    mon = []
    for p in re.split(r'[,+/;.]', s):
        p = re.sub(r'\b(SC|SỮA CHUA|SỮA)\s*$', '', p.strip()).strip()
        if p and not BO_PHAN_MON.match(p):
            for q in tach_dinh(p):
                if q not in mon: mon.append(q)
    if co_rs: mon.append('RAU SỐNG')
    return mon


def tach_dinh(p):
    """Thực đơn hay viết liền nhiều món không có dấu phẩy, sổ hoàn thiện tách riêng từng món."""
    sau = []
    for hau in ('TRỨNG VỊT LỘN', 'TRỨNG ỐP'):
        m = re.match(rf'(.*\S)\s+{hau}$', p)
        if m: p, sau = m.group(1), [hau] + sau
    m = re.match(r'(CƠM RANG|XÔI LẠC|BÁNH CUỐN)\s+(.+)$', p)
    if m and m.group(2) not in ('XÌ DẦU', 'THẬP CẨM', 'DƯA BÒ'): return [m.group(1)] + tach_tu_dien(m.group(2)) + sau
    return tach_tu_dien(p) + sau


def tach_tu_dien(p):
    """Tên ghép từ hai món đã biết (học từ sổ hoàn thiện) thì tách ra; tên đã biết thì giữ nguyên."""
    if khoa_mon(p) in mon_biet(): return [p]
    tu_p = p.split()
    for i in range(1, len(tu_p)):
        a, b = ' '.join(tu_p[:i]), ' '.join(tu_p[i:])
        if khoa_mon(a) in mon_biet() and tach_tu_dien(b)[0] != b or khoa_mon(a) in mon_biet() and khoa_mon(b) in mon_biet():
            return [a] + tach_tu_dien(b)
    return [p]


def dinh_luong_thuong():
    """Khoảng định lượng thường dùng theo nguyên liệu, học từ sổ hoàn thiện."""
    if _mon_biet['dl'] is None:
        p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dinh_luong.json')
        _mon_biet['dl'] = json.load(open(p, encoding='utf-8')) if os.path.exists(p) else {}
    return _mon_biet['dl']


def mon_biet():
    if _mon_biet['ds'] is None:
        p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'mon_biet.json')
        _mon_biet['ds'] = set(json.load(open(p, encoding='utf-8'))) if os.path.exists(p) else set()
    return _mon_biet['ds']


def diem(nl, mon, so_tay):
    k, n = khoa_nl(nl), khoa_mon(mon)
    if any(n.startswith(mon_st) and any(k.startswith(x) for x in ds) for mon_st, ds in so_tay.items()): return 5
    if n == 'RAU SONG': return 4 if k.startswith(RAU_SONG) else -1
    d = len(tu(nl) & tu(mon))
    for a, b in GOI_Y.items():
        if k.startswith(a) and b in n: d = max(d, 1)
    for a, b in GOI_Y_MANH.items():
        if k.startswith(a) and b in n: d = max(d, 2)
    return d


def gan_mon(ten_set, dong, so_tay):
    """Gán từng nguyên liệu vào 1 món. Thứ tự nguyên liệu đi theo thứ tự món trong tên, nên
    dùng quy hoạch động để tìm cách gán đơn điệu có tổng điểm khớp cao nhất."""
    mon = tach_mon(ten_set)
    if mon and khong_dau(mon[0]).startswith(('MI TOM', 'MY TOM')): mon = mon[:1]   # trứng, rau đi kèm thuộc suất mì
    ds, gao = [], []
    com_rieng = next((m for m in mon if re.match(r'COM GAO NUT\b|XOI\b', khong_dau(m))), None)
    ket = []
    for x in dong:
        if not la_com(x['nl'], x['dl']): ds.append(x)
        elif com_rieng: ket.append((x, com_rieng, True))      # gạo trong món xôi, cơm gạo nứt
        else: gao.append(x)
    if any(khoa_nl(x['nl']).startswith('CAI CANH') for x in ds) and not any(la_canh(m) for m in mon):
        mon.append('CANH CẢI THỊT')              # canh đi kèm không ghi trong tên bộ món
    if 'RAU SỐNG' not in mon and not khong_dau(mon[0] if mon else '').startswith(('MI TOM', 'MY TOM')) \
            and sum(khoa_nl(x['nl']).startswith(RAU_SONG) for x in ds) >= 3:
        mon.append('RAU SỐNG')                   # rau ăn kèm không ghi trong tên bộ món
    if not any(tu('TRỨNG') & tu(m) for m in mon) and not khong_dau(mon[0] if mon else '').startswith(('MI TOM', 'MY TOM')) \
            and any(khoa_nl(x['nl']).startswith('TRUNG GA') and x['dl'] >= Q['dl_dem'] for x in ds):
        mon.append('TRỨNG ỐP')                   # trứng đi kèm không ghi trong tên bộ món
    ket += [(x, 'CƠM TRẮNG', True) for x in gao]
    if not ds: return ket
    if not mon: mon = [ten_set.upper()]
    n, m = len(ds), len(mon)
    tot = [[-1e9] * m for _ in range(n)]; truoc = [[0] * m for _ in range(n)]
    for j in range(m): tot[0][j] = diem(ds[0]['nl'], mon[j], so_tay) - 1.5 * j
    for i in range(1, n):
        for j in range(m):
            for p in range(j + 1):
                v = tot[i - 1][p] + diem(ds[i]['nl'], mon[j], so_tay) - 1.5 * max(0, j - p - 1) - (0.01 if j != p else 0)
                if v > tot[i][j]: tot[i][j], truoc[i][j] = v, p
    j = max(range(m), key=lambda q: tot[-1][q] - 1.5 * (m - 1 - q))
    gan = [0] * n
    for i in range(n - 1, -1, -1):
        gan[i] = j
        j = truoc[i][j]
    # Nguyên liệu khớp rõ với món khác (tên món và nguyên liệu không cùng thứ tự) thì chuyển sang món đó
    d = [[diem(x['nl'], mm, so_tay) for mm in mon] for x in ds]
    for i in range(n):
        tot_nhat = max(range(m), key=lambda q: d[i][q])
        if d[i][tot_nhat] >= 2 and d[i][tot_nhat] > d[i][gan[i]]: gan[i] = tot_nhat
        # Chưa khớp món nào nhưng cả tên nguyên liệu nằm trong đúng một tên món thì chuyển sang món đó
        if khoa_nl(ds[i]['nl']).startswith('TRUNG GA') and ds[i]['dl'] >= Q['dl_dem'] and 'TRỨNG ỐP' in mon:
            gan[i] = mon.index('TRỨNG ỐP'); d[i][gan[i]] = 2
        trung = [q for q in range(m) if tu(ds[i]['nl']) and tu(ds[i]['nl']) <= tu(mon[q])]
        if d[i][gan[i]] <= 0 and len(trung) == 1: gan[i] = trung[0]
        # Nạc vai định lượng rất nhỏ là thịt nấu canh
        canh = [q for q in range(m) if la_canh(mon[q])]
        if canh and khoa_nl(ds[i]['nl']).startswith('NAC VAI') and ds[i]['dl'] <= Q['dl_thit_canh']:
            gan[i] = canh[-1]; d[i][gan[i]] = 2
    chac_vi_tri = xep_theo_vi_tri(gan, d, mon) if Q['tu_khac_phuc'] == 'co' else set()
    for j, mm in enumerate(mon):
        if j not in set(gan) and mm not in [t for x, t, _ in ket if x]: ket.append((None, mm, False))
    # Chỉ cảnh báo dòng chưa khớp món nào và đủ lớn; gia vị nhỏ không đáng dừng lại soát
    for i, x in enumerate(ds):
        j = gan[i]
        ket.append((x, mon[j], i in chac_vi_tri or d[i][j] > 0 or (d[i][j] == 0 and (m == 1 or x['dl'] < Q['dl_canh_bao']))))
    return ket


def xep_theo_vi_tri(gan, d, mon):
    """Nguyên liệu đi theo thứ tự món, nên dòng chưa khớp được suy ra từ các dòng đã khớp hai bên:
    nằm giữa 2 dòng cùng món thì thuộc món đó; nằm ở khoảng có đúng 1 món chưa có dòng khớp thì thuộc món đó."""
    n, m, chac = len(gan), len(mon), set()
    khop = [i for i in range(n) if d[i][gan[i]] > 0]
    co_dong = {gan[i] for i in khop}
    for i in range(n):
        if d[i][gan[i]] > 0: continue
        truoc = next((q for q in reversed(khop) if q < i), None)
        sau = next((q for q in khop if q > i), None)
        a = gan[truoc] if truoc is not None else -1
        b = gan[sau] if sau is not None else m
        if a == b:
            gan[i] = a; chac.add(i)
            continue
        lo, hi = min(a, b), max(a, b)
        trong = [j for j in range(lo + 1, hi) if j not in co_dong and khoa_mon(mon[j]) != 'RAU SONG']
        if len(trong) == 1: gan[i] = trong[0]; chac.add(i)
    return chac


def don_vi(nl, dl):
    k = khoa_nl(nl)
    if k.startswith('TRUNG GA'): return 'QUẢ'      # cột TỔNG của trứng luôn là số quả
    if dl >= Q['dl_dem']:
        for x in Q['don_vi_theo_ten']:
            if k.startswith(khoa_nl(x['bat_dau'])): return x['don_vi']
        return Q['don_vi_dem']
    return Q['don_vi_can']


def lech_lon(a, b):
    """Sai số nhỏ khi lệch không quá sai_so_kg hoặc không quá sai_so_pt phần trăm."""
    d = abs(a - b)
    return d > Q['sai_so_kg'] and d > abs(a) * Q['sai_so_pt'] / 100


TONG_LECH = r'Cột TỔNG ghi ([\d.]+) nhưng .* = ([\d.]+);'
TUAN_LECH = r'Thực đơn tuần ghi ([\d.]+) suất, sheet ngày ghi ([\d.]+)'


def muc_sai_so(nd):
    """'kiem': cần kiểm tra, được tô trong sổ. 'tin': khác Thực đơn tuần, sổ hoàn thiện luôn theo sheet ngày nên chỉ liệt kê."""
    if nd.startswith('Cảnh báo'): return 'kiem'
    m = re.match(TONG_LECH, nd)
    if m: return 'kiem' if lech_lon(float(m.group(1)), float(m.group(2))) else None
    m = re.match(TUAN_LECH, nd)
    if m: return 'tin' if lech_lon(float(m.group(2)), float(m.group(1))) else None
    return None


def la_sai_so(nd):
    return muc_sai_so(nd) is not None


def vi_tri_bang(sach):
    """Dòng của mỗi nguyên liệu trên sheet bước 2, tính giống ghi_buoc2: {ca: {dong: {dòng thực đơn: hàng}, y: {hàng: nguyên liệu}, com: hàng Cơm trắng}}."""
    kq = {}
    for ca in CA:
        r, dong, com, theo_hang = 6, {}, None, {}
        for m in sach[ca]:
            for y in m['nl']:
                for d in y.get('dong', ()): dong[d] = r
                theo_hang[r] = y
                if com is None and khoa_mon(m['ten']) == 'COM TRANG': com = r
                r += 1
        kq[ca] = dict(dong=dong, y=theo_hang, com=com)
    return kq


def ghi_chu_sai_so(sach, nhat_ky):
    """Ghi chú kèm ô trong sổ. gia_tri: số đang hiện trong ô; so: phần của bộ món này; doi: số đối chiếu; gop: các ô gộp có bộ món này."""
    vt, ds = vi_tri_bang(sach), []
    for ca, dong, bo, nl, nd in nhat_ky:
        muc = muc_sai_so(nd)
        if not muc: continue
        cot, so, doi, gop = 'E', None, None, []
        hang, dong_y = vt[ca]['dong'].get(dong), vt[ca]['y']
        khoi = next((y['dong'][dong] for y in dong_y.values() if dong in y.get('dong', ())), None)
        m, t, g = re.match(TONG_LECH, nd), re.match(TUAN_LECH, nd), re.match(r'Cảnh báo: định lượng ([\d.]+)', nd)
        if m:
            so, doi = float(m.group(1)), float(m.group(2))
            if Q['khi_lech'] != 'thuc_don': so, doi = doi, so
        elif t:
            cot, so, doi = 'G', float(t.group(2)), float(t.group(1))
            # Số suất là của cả bộ: trỏ tới dòng riêng của bộ, dòng gộp nhiều bộ (như Cơm trắng) chỉ ghi kèm
            cua_bo = [r for r, y in dong_y.items() if any((k == khoi if khoi is not None else p['bo'] == bo) and p['suat'] == so for k, p in y['phan'].items())]
            rieng = [r for r in cua_bo if len(dong_y[r]['phan']) == 1 and dong_y[r]['suat'] == so]
            hang = (rieng or cua_bo or [hang])[0]
            for r in cua_bo:
                y = dong_y[r]
                if r != hang and len(y['phan']) > 1:
                    gop.append(dict(o=f'G{r}', ten=y['ten'], tong=y['suat'], phan=[(p['bo'], p['suat']) for p in y['phan'].values()]))
        elif g:
            y = dong_y.get(hang)
            so = y['phan'][khoi]['sl'] if y and khoi in y['phan'] else None
        else:
            hang = vt[ca]['com']          # bộ cơm thiếu gạo: kiểm ở dòng Cơm trắng của ca
        y = dong_y.get(hang)
        gia_tri = (y['suat'] if cot == 'G' else y['sl']) if y else None
        nd = nd.replace('Cảnh báo: ', '')
        if y and so is not None and abs(gia_tri - so) > 1e-6:
            nd += f"; ô {cot}{hang} là số gộp {gia_tri:g}, phần của bộ này là {so:g}"
        if gop:
            nd += '; bộ này cũng được cộng vào ' + ', '.join(f"ô {x['o']} ({x['ten']}, {x['tong']:g} suất)" for x in gop)
        ds.append(dict(ca=ca, muc=muc, sheet=SHEET2[ca], hang=hang, cot=cot, o=f'{cot}{hang}' if hang else '', bo=bo, nl=y['ten'] if y else nl,
                       gia_tri=gia_tri, so=so, doi=doi, lech=round(so - doi, 2) if so is not None and doi is not None else None,
                       pt=round(abs(so - doi) / doi * 100) if so is not None and doi else None, gop=gop, nd=nd))
    return sorted(ds, key=lambda x: (x['muc'] != 'kiem', CA.index(x['ca']), x['hang'] or 0))


def ten_chuan(nl):
    k = khoa_nl(nl)
    return next((x['ten'] for x in Q['doi_nl'] if k.startswith(khoa_nl(x['bat_dau']))), nl)


def xu_ly_ca(dong, ca, nhat_ky, so_tay):
    """Trả về [món: {ten, nl: [{ten, sl, dv, suat}]}] sau khi lọc, gán món và cộng trùng."""
    theo_set, thu_tu, khong_suat = {}, [], set()
    for x in dong:
        if BO_NGUYEN_LIEU.match(x['nl'].upper().strip()): continue
        if any(t.upper() in x['set'].upper() for t in Q['bo_bo_mon']): continue
        if x['suat'] is None:
            if x.get('khong_phuc_vu'): continue
            if x['set'] not in khong_suat:
                khong_suat.add(x['set'])
                toan_bo = all(y['suat'] is None for y in dong if y['set'] == x['set'] and not BO_NGUYEN_LIEU.match(y['nl'].upper().strip()))
                nhat_ky.append((ca, x['row'], x['set'], '', 'Bộ món không có số suất (cột E trống) nên không đưa vào sổ' if toan_bo
                                else 'Bộ món có nguyên liệu không có số suất nên các dòng đó không đưa vào sổ'))
            continue
        ct = cong_thuc_cua(x['nl'])
        dung = round(tinh(ct, x['dl'], x['suat']), 4)
        if x['tong'] is None and x.get('tu_suat'):
            x['tong'] = dung
        elif x['tong'] is None:
            x['tong'] = dung
            nhat_ky.append((ca, x['row'], x['set'], x['nl'], f"Thiếu công thức cột TỔNG; đã tự tính {dien_giai(ct, x['dl'], x['suat'])} = {dung}"))
        elif abs(x['tong'] - dung) > Q['cho_phep_lech']:
            giu = Q['khi_lech'] == 'thuc_don'
            nhat_ky.append((ca, x['row'], x['set'], x['nl'], f"Cột TỔNG ghi {x['tong']} nhưng {dien_giai(ct, x['dl'], x['suat'])} = {dung}; sổ dùng {x['tong'] if giu else dung}"))
            if not giu: x['tong'] = dung
        if not x['tong']: continue
        khoang = dinh_luong_thuong().get(khoa_nl(x['nl'])) if Q['kiem_dinh_luong'] == 'co' else None
        if khoang and not khoang[0] / 2 <= x['dl'] <= khoang[1] * 2:
            nhat_ky.append((ca, x['row'], x['set'], x['nl'], f"Cảnh báo: định lượng {x['dl']} lạ, thường dùng {khoang[0]} đến {khoang[1]}"))
        # Theo từng khối dòng: cùng tên bộ món ở 2 cửa vẫn là 2 bộ có số suất riêng
        if x['khoi'] not in theo_set: theo_set[x['khoi']] = []; thu_tu.append(x['khoi'])
        theo_set[x['khoi']].append(x)
    mon, chi_so = [], {}

    def mon_cua(ten):
        k = khoa_mon(ten)
        if k not in chi_so:
            chi_so[k] = dict(ten=ten, nl=[], idx={}, theo_bo={}); mon.append(chi_so[k])
        return chi_so[k]

    for khoi in thu_tu:
        ds = theo_set[khoi]
        s = ds[0]['set']
        mon_bo = tach_mon(s)
        if len(mon_bo) >= 4 and any(la_canh(t) for t in mon_bo) and not any(khoa_nl(x['nl']).startswith(('GAO', 'NEP')) for x in ds):
            nhat_ky.append((ca, ds[0]['row'], s, 'GẠO', f"Cảnh báo: bộ cơm {ds[0]['suat']} suất không có dòng gạo trong thực đơn, sổ giữ theo thực đơn"))
        gan = gan_mon(s, ds, so_tay)
        for ten in {t for x, t, _ in gan if x}:
            su = sorted({x['suat'] for x, t, _ in gan if x and t == ten})
            if len(su) > 1: nhat_ky.append((ca, ds[0]['row'], s, ten, f'Cùng một món nhưng có nhiều số suất khác nhau: {su}; sổ lấy số lớn nhất ở bước 3'))
        for x, ten, chac in gan:
            if x is None:
                su = collections.Counter(y['suat'] for y in ds).most_common(1)[0][0]
                nhat_ky.append((ca, ds[0]['row'], s, ten, f'Món có trong tên bộ món nhưng máy chưa tìm được nguyên liệu; bước 3 và mẫu lưu vẫn ghi {su} suất'))
                m = mon_cua(ten)
                m['theo_bo'][khoi] = max(m['theo_bo'].get(khoi, 0), su)
                them = next((r for r in Q['mon_thieu_nl'] if khoa_mon(r['mon']) == khoa_mon(ten)), None)
                if them:
                    nhat_ky.pop()
                    k = khoa_nl(them['nl'])
                    y = m['idx'].setdefault(k, None) or dict(ten=them['nl'], sl=0, dv=them['don_vi'], suat=0, chac=True, bo=set(), dong={}, phan={})
                    if not y['bo']: m['idx'][k] = y; m['nl'].append(y)
                    y['sl'] = round(y['sl'] + su * them['moi_suat'], 4); y['suat'] += su; y['bo'].add(khoi)
                    y['phan'][khoi] = dict(bo=s, suat=su, sl=su * them['moi_suat'])
                    nhat_ky.append((ca, ds[0]['row'], s, them['nl'], f"Đã tự sửa: thực đơn không có nguyên liệu của món {ten}, sổ thêm {su * them['moi_suat']} {them['don_vi']}"))
                continue
            x = dict(x, nl=ten_chuan(x['nl']))
            kg = khoa_nl(x['nl'])
            if ten == 'CƠM TRẮNG' and x['dl'] != Q['dl_com']:
                nhat_ky.append((ca, x['row'], s, x['nl'], f"Định lượng gạo là {x['dl']}, khác mức thường dùng {Q['dl_com']}"))
            if not chac: nhat_ky.append((ca, x['row'], s, x['nl'], f'Máy tự xếp vào món "{ten}", cần xem lại'))
            m = mon_cua(ten)
            m['theo_bo'][khoi] = max(m['theo_bo'].get(khoi, 0), x['suat'])
            if kg in m['idx']:
                y = m['idx'][kg]
                y['sl'] = round(y['sl'] + x['tong'], 4); y['dong'][x['row']] = khoi
                if khoi not in y['bo']:
                    y['bo'].add(khoi); y['suat'] += x['suat']
            else:
                y = dict(ten='TRỨNG VỊT LỘN' if kg == 'TVL' else x['nl'], sl=x['tong'], dv=don_vi(x['nl'], x['dl']), suat=x['suat'], chac=chac, bo={khoi}, dong={x['row']: khoi}, phan={})
                m['idx'][kg] = y; m['nl'].append(y)
            p = y['phan'].setdefault(khoi, dict(bo=s, suat=x['suat'], sl=0))
            p['sl'] = round(p['sl'] + x['tong'], 4)
    # Cơm trắng và rau sống dùng chung cho nhiều bộ: mọi dòng của món mang cùng tổng suất
    for m in mon:
        m['suat'] = sum(m['theo_bo'].values())
        if khoa_mon(m['ten']) == 'RAU SONG':
            for y in m['nl']: y['suat'] = m['suat']
    return mon


# ---------------- Ghi ra mẫu ----------------
def chup_hang(ws, r, cot):
    return [(ws.cell(r, c)._style, ws.cell(r, c).value) for c in range(1, cot + 1)], ws.row_dimensions[r].height

def dat_hang(ws, r, chup, gt=True):
    for c, (st, v) in enumerate(chup[0], 1):
        o = ws.cell(r, c); o._style = copy.copy(st)
        o.value = v if gt else None
    ws.row_dimensions[r].height = chup[1]
    ws.row_dimensions[r].hidden = False

def xoa_than(ws, tu_hang):
    for m in list(ws.merged_cells.ranges):
        if m.max_row >= tu_hang: ws.unmerge_cells(str(m))
    for r in range(tu_hang, ws.max_row + 3):
        for c in range(1, ws.max_column + 1):
            o = ws.cell(r, c); o.value = None; o.style = 'Normal'
        ws.row_dimensions[r].hidden = False
        ws.row_dimensions[r].height = None
    ws.row_breaks.brk = []
    ws.auto_filter.ref = None

def ghi_buoc2(ws, mon, ca_lb, ngay_txt, nhap_ca=None, to_mau=True):
    ws['E1'] = ngay_txt
    mau = chup_hang(ws, 6, 15)
    xoa_than(ws, 6)
    ws.conditional_formatting = type(ws.conditional_formatting)()
    r = 6; tt = 0; dau_b = 6
    for m in mon:
        if not m['nl']: continue                 # món chưa có nguyên liệu chỉ ghi ở bước 3 và mẫu lưu
        d0 = r
        for y in m['nl']:
            dat_hang(ws, r, mau, False); tt += 1
            ws.cell(r, 1, tt); ws.cell(r, 4, y['ten']); ws.cell(r, 5, round(y['sl'], 4)); ws.cell(r, 6, y['dv']); ws.cell(r, 7, y['suat'])
            ws.cell(r, 5).number_format = SO_THAT
            if to_mau and not y.get('chac', True): ws.cell(r, 4).fill = VANG
            r += 1
        ws.cell(d0, 3, m['ten'])
        if r - 1 > d0:
            for c in (3, 8, 9): ws.merge_cells(start_row=d0, end_row=r - 1, start_column=c, end_column=c)
    for x in Q['trai_cay']:
        ten, dv = x['ten'], x['don_vi']
        dat_hang(ws, r, mau, False); tt += 1
        ws.cell(r, 1, tt); ws.cell(r, 3, ten); ws.cell(r, 4, ten); ws.cell(r, 6, dv)
        so = (nhap_ca or {}).get(ten.strip(), {})
        ws.cell(r, 5, so.get('sl')); ws.cell(r, 7, so.get('suat'))
        ws.cell(r, 5).number_format = SO_THAT
        r += 1
    ws.cell(6, 2, ca_lb); ws.merge_cells(start_row=6, end_row=r - 1, start_column=2, end_column=2)
    ws.print_area = f'A1:O{r - 1}'

def khoi_header(ws):
    """Tách các khối ca trong sheet bước 3 / mẫu 24 giờ: (hàng đầu, hàng tiêu đề cuối, nhãn ca)."""
    dau = [r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or '').startswith(('Tên cơ sở', '=A', "=' bước 3"))
           and r < ws.max_row]
    return dau

def ghi_khoi(ws, cot, dong_fn, ngay_txt, cot_ngay, wv):
    pos = [r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == 'TT']
    tieu_de = []
    for tt_r in pos:
        r0 = tt_r - 3
        hang = list(range(r0, tt_r + 2))
        cs = []
        for r in hang:
            ch = chup_hang(ws, r, cot)
            vals = [(st, wv.cell(r, i + 1).value) for i, (st, _) in enumerate(ch[0])]
            cs.append((vals, ch[1]))
        gop = [(m.min_row - r0, m.max_row - r0, m.min_col, m.max_col) for m in ws.merged_cells.ranges
               if r0 <= m.min_row and m.max_row <= tt_r + 1]
        tieu_de.append((cs, gop))
    than = chup_hang(ws, pos[0] + 2, cot)
    than_b = [(st, None) for st, _ in than[0]]
    xoa_than(ws, 1)
    ws.conditional_formatting = type(ws.conditional_formatting)()
    r = 1; ngat = []
    for i, ca in enumerate(CA):
        cs, gop = tieu_de[i]
        if i: ngat.append(r - 1)
        r0 = r
        for ch in cs:
            dat_hang(ws, r, ch); r += 1
        for a, b, c, d in gop: ws.merge_cells(start_row=r0 + a, end_row=r0 + b, start_column=c, end_column=d)
        # Mẫu đặt ô ngày không đồng nhất giữa các ca nên tìm ô ngày cũ để thay
        cu = [ws.cell(h, c) for h in range(r0, r0 + len(cs)) for c in range(1, cot + 1)
              if re.search(r'ngày\s*\d+\s*tháng', str(ws.cell(h, c).value or ''), re.I)]
        for o in cu or [ws.cell(r0, cot_ngay)]: o.value = ngay_txt
        d0 = r
        for tt, ghi in enumerate(dong_fn(ca), 1):
            dat_hang(ws, r, (than_b, than[1]), False); ws.cell(r, 1, tt); ghi(ws, r); r += 1
        yield ca, d0, r - 1
        r += 1
    for b in ngat: ws.row_breaks.append(Break(id=b))

def mau_luu(mon, nhap_ca=None):
    """Danh sách mẫu lưu của 1 ca: bún/phở tách 3 dòng; canh, nước dùng tính ml; hàng đóng gói lưu trong công."""
    kq, kl, ml, hop = [], Q['mau_khoi_luong'], Q['mau_canh'], Q['mau_dung_cu']
    suat_nhap = lambda ten: (nhap_ca or {}).get(ten.strip(), {}).get('suat')
    for m in mon:
        t = m['ten'].strip()
        k = khoa_mon(t)
        tach = next((x for x in Q['tach_ba_dong'] if k == khoa_mon(x['bat_dau']) or k.startswith(khoa_mon(x['bat_dau']) + ' ')), None)
        cong = next((x for x in Q['luu_cong'] if k.startswith(khoa_mon(x['bat_dau']))), None)
        if tach and not any(khong_dau(x) in k for x in Q['khong_tach']):
            kq += [(t, m['suat'], kl, hop), (f'NHÂN {t}', m['suat'], kl, hop), ('NƯỚC DÙNG', m['suat'], ml, hop)]
        elif la_canh(t):
            kq.append((t, m['suat'], ml, hop))
        elif cong:
            kq.append((t, m['suat'], cong['khoi_luong'], cong['dung_cu']))
        else:
            kq.append((t, m['suat'], kl, hop))
    kq += [(x['ten'].strip() + ' ' if x['ten'].startswith('DƯA') else x['ten'], suat_nhap(x['ten']), kl, hop) for x in Q['trai_cay']]
    kq += [(x['ten'], suat_nhap(x['ten']) if suat_nhap(x['ten']) is not None else x['suat'], x['khoi_luong'], x['dung_cu']) for x in Q['hang_cuoi_ca']]
    return kq

def ghi_bia(wb, ngay, ds_ngay):
    """Cập nhật dòng tháng trên bìa theo các ngày của tháng đó trong file thực đơn."""
    th = ngay.split('.')[1]
    ngay_thang = sorted(int(x.split('.')[0]) for x in (ds_ngay or [ngay]) if x.split('.')[1] == th) or [int(ngay.split('.')[0])]
    for row in wb['BÌA'].iter_rows():
        for o in row:
            if isinstance(o.value, str) and re.fullmatch(r'\s*THÁNG\s*\d+\s*\([\d\s-]+\)\s*', o.value):
                o.value = f'THÁNG {int(th):02d} ({ngay_thang[0]:02d}-{ngay_thang[-1]:02d})'


def danh_dau(wb, ghi_chu):
    """Bản kiểm tra: cam là cần kiểm tra, vàng là khác Thực đơn tuần; mỗi ô có chú thích; sheet ghi chú đứng đầu file, có liên kết tới ô."""
    o_to = {}

    def them(x, o, muc, dong, dau=None):
        v = o_to.setdefault((x['sheet'], o), dict(kiem=False, dau=dau, dong=[]))
        v['kiem'] = v['kiem'] or muc == 'kiem'
        if dong not in v['dong']: v['dong'].append(dong)

    for x in ghi_chu:
        if not x['hang']: continue
        them(x, x['o'], x['muc'], x['nd'])
        for g in x['gop']:
            dau = f"Dòng gộp nhiều bộ món, tổng {g['tong']:g} suất:\n" + '\n'.join(f'- {b}: {s:g}' for b, s in g['phan'])
            them(x, g['o'], x['muc'], f"Lệch: {x['bo']}, Thực đơn tuần ghi {x['doi']:g}", dau)
    for (sheet, o), v in o_to.items():
        o = wb[sheet][o]
        o.fill = CAM if v['kiem'] else VANG_DAM
        chu = '\n'.join(([v['dau']] if v['dau'] else []) + v['dong'])
        c = Comment(chu, 'Sổ kiểm thực'); c.width, c.height = 420, min(90 + 18 * chu.count('\n') + 18 * (len(chu) // 60), 420)
        o.comment = c
    ws = wb.create_sheet('GHI CHÚ SAI SỐ', 0)
    ws['A1'] = 'Cần kiểm tra'; ws['A1'].fill = CAM
    ws['B1'] = 'Ô tô cam: cột TỔNG lệch, bộ cơm thiếu gạo, định lượng lạ.'
    ws['A2'] = 'Thông tin'; ws['A2'].fill = VANG_DAM
    ws['B2'] = 'Ô tô vàng: Thực đơn tuần khác sheet ngày. Dòng gộp nhiều bộ cũng tô vàng, rê chuột vào ô để xem từng bộ.'
    ws['A3'] = 'Cách đọc'
    ws['B3'] = 'Số trong ô là số thấy khi bấm Vị trí. Số của bộ món là phần của riêng bộ đó, hai số khác nhau khi ô là dòng gộp.'
    ws.append([])
    dau = ['Vị trí', 'Mức', 'Ca', 'Bộ món', 'Nguyên liệu trong ô', 'Số trong ô', 'Số của bộ món', 'Số đối chiếu', 'Lệch', '%', 'Ô gộp liên quan', 'Nội dung']
    ws.append(dau)
    tieu_de = ws.max_row
    for x in ghi_chu:
        ws.append([f"{x['sheet'].strip()}!{x['o']}" if x['hang'] else '', 'Cần kiểm tra' if x['muc'] == 'kiem' else 'Thông tin', x['ca'], x['bo'], x['nl'],
                   x['gia_tri'], x['so'], x['doi'], x['lech'], x['pt'], ', '.join(g['o'] for g in x['gop']), x['nd']])
        if x['hang']:
            o = ws.cell(ws.max_row, 1)
            o.hyperlink = Hyperlink(ref=o.coordinate, location=f"'{x['sheet']}'!{x['o']}")
            o.font = Font(color='0563C1', underline='single')
        ws.cell(ws.max_row, 2).fill = CAM if x['muc'] == 'kiem' else VANG_DAM
        for c in ws[ws.max_row]: c.alignment = Alignment(wrap_text=True, vertical='top')
    for c in ws[tieu_de]:
        c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical='center')
    for c, w in zip('ABCDEFGHIJKL', (20, 13, 8, 42, 26, 9, 9, 10, 8, 6, 12, 60)): ws.column_dimensions[c].width = w
    ws.freeze_panes = ws.cell(tieu_de + 1, 1)
    ws.auto_filter.ref = f'A{tieu_de}:L{max(ws.max_row, tieu_de)}'


def ghi_ra(mau_path, sach, ngay, nhat_ky, ra, nhap=None, ds_ngay=None, ban='kiem'):
    """ban='chot': sạch, đúng mẫu để in. ban='kiem': thêm đánh dấu và các sheet kiểm tra."""
    nhap = nhap or {}
    wb, wv = mo_wb(mau_path), mo_wb(mau_path, data_only=True)
    d, th = ngay.split('.')
    txt = f'Thời gian kiểm tra:ngày {int(d):02d} tháng {th} năm {Q["nam"]}'
    nhan = {}
    for ca in CA:
        ws = wb[SHEET2[ca]]
        nhan[ca] = wv[SHEET2[ca]]['B6'].value
        ghi_buoc2(ws, sach[ca], nhan[ca], txt, nhap.get(ca), ban == 'kiem')

    def dong3(ca):
        for m in sach[ca]:
            yield (lambda m=m: (lambda ws, r: (ws.cell(r, 3, m['ten']), ws.cell(r, 4, m['suat']))))()
        for x in Q['trai_cay']:
            yield (lambda t=x['ten'], s=nhap.get(ca, {}).get(x['ten'].strip(), {}).get('suat'): (lambda ws, r: (ws.cell(r, 3, t), ws.cell(r, 4, s))))()

    ws, v = wb[SHEET3], wv[SHEET3]
    for ca, a, b in list(ghi_khoi(ws, 10, dong3, txt, 6, v)):
        ws.cell(a, 2, nhan[ca]); ws.merge_cells(start_row=a, end_row=b, start_column=2, end_column=2)
    ws.print_area = f'A1:J{b}'

    def dong4(ca):
        for t, s, kl, dc in mau_luu(sach[ca], nhap.get(ca)):
            def g(ws, r, t=t, s=s, kl=kl, dc=dc):
                ws.cell(r, 2, t); ws.cell(r, 4, s); ws.cell(r, 5, kl); ws.cell(r, 6, dc)
            yield g

    ws, v = wb[SHEET4], wv[SHEET4]
    for ca, a, b in list(ghi_khoi(ws, 12, dong4, txt, 6, v)):
        ws.cell(a, 3, nhan[ca]); ws.merge_cells(start_row=a, end_row=b, start_column=3, end_column=3)
    ws.print_area = f'A1:L{b}'
    # Mẫu 72 giờ giống mẫu 24 giờ, chỉ khác nhiệt độ bảo quản
    w72 = wb.copy_worksheet(ws)
    w72.title = 'MẪU 72 giờ'
    w72.print_area = ws.print_area
    for g in ws.row_breaks.brk: w72.row_breaks.append(Break(id=g.id))
    for hang in w72.iter_rows():
        for o in hang:
            if isinstance(o.value, str) and '2-8' in o.value: o.value = o.value.replace('2-8', '≤-18')
    wb.move_sheet(w72, wb.sheetnames.index(ws.title) + 1 - wb.sheetnames.index(w72.title))

    if 'BÌA' in wb.sheetnames: ghi_bia(wb, ngay, ds_ngay)
    if ban == 'kiem':
        ws = wb.create_sheet('CẦN SOÁT')
        ws.append(['Ca', 'Dòng ở thực đơn', 'Bộ món', 'Nguyên liệu', 'Nội dung cần xem'])
        for x in nhat_ky: ws.append(list(x))
        for c, w in zip('ABCDE', (8, 10, 55, 28, 70)): ws.column_dimensions[c].width = w
        danh_dau(wb, ghi_chu_sai_so(sach, nhat_ky))
    else:
        for ten in [n for n in wb.sheetnames if n.strip() in ('Sheet1', 'MẪU 24 giờ (3)')]: del wb[ten]
    dau = 'GHI CHÚ SAI SỐ' if ban == 'kiem' else SHEET2['SÁNG']
    wb.active = wb.sheetnames.index(dau)
    for s in wb.worksheets: s.sheet_view.tabSelected = s.title == dau
    wb.save(ra)


def tai_so_tay():
    p = file_ghi('so_tay_mon.json')
    d = json.load(open(p, encoding='utf-8')) if os.path.exists(p) else {}
    kq = {}
    for k, ds in d.items():                      # khoá phải cùng dạng với khoa_mon mới khớp được
        kq.setdefault(khoa_mon(k), []).extend(x for x in ds if x not in kq.get(khoa_mon(k), []))
    return kq


_td = {}

# ---------------- So sổ hoàn thiện với sổ của app ----------------
def _nhom_nl(s):
    """Cùng một hàng nhưng người ghi và thực đơn gọi khác nhau."""
    k = re.sub(r'^THIT |^RAU |\b(VN|XOAN|DAI|TUOI)\b|\b2-3KG/CON\b', ' ', khoa_nl(s))
    k = ' '.join(k.split())
    for dau, ten in (('BA CHI', 'BA CHI'), ('HUNG QUE', 'HUNG QUE'), ('HUNG CHO', 'HUNG QUE'), ('GA XIEN', 'GA XIEN'), ('GAO DEO', 'GAO DEO'),
                     ('GAO THOM', 'GAO DEO'), ('GAO BAC THOM', 'GAO DEO'), ('GAO NEP', 'GAO NEP'), ('NEP', 'GAO NEP'), ('TRUNG GA', 'TRUNG GA'),
                     ('QUA BI XANH', 'BI XANH'), ('THIT GA GOC TU', 'GA GOC TU'), ('GA GOC TU', 'GA GOC TU')):
        if k.startswith(dau): return ten
    return 'NEM NGON' if k.startswith('NEM') and 'NGON' in k else k


def doc_so(f):
    """Đọc sổ đã tạo hoặc sổ hoàn thiện: số lượng bước 2 theo ca, món bước 3, mẫu lưu."""
    w = mo_wb(f, data_only=True)
    kd = {n: khong_dau(n) for n in w.sheetnames}
    bo = {khoa_nl(x['ten']) for x in Q['trai_cay'] + Q['hang_cuoi_ca']} | {khoa_nl(x['thuc_don']) for x in Q['goi_y_sua']}
    buoc2 = [n for n in w.sheetnames if 'BUOC 2' in kd[n].replace('  ', ' ')]
    if len(buoc2) < 4: raise ValueError('File này không có đủ 4 sheet bước 2')
    the = {c: next((n for n in buoc2 if khong_dau(c) in kd[n]), buoc2[i]) for i, c in enumerate(CA)}
    so, dong = {}, {}
    for c in CA:
        ws, g, mon = w[the[c]], collections.Counter(), None
        dong[c] = []
        for r in range(1, ws.max_row + 1):
            t = ws.cell(r, 3).value
            if t and str(t).strip() and 'TÊN' not in str(t).upper(): mon = str(t).strip()
            d, e = ws.cell(r, 4).value, ws.cell(r, 5).value
            if d and isinstance(e, (int, float)) and khoa_nl(str(d)) not in bo and not khoa_nl(str(d)).startswith('SUA'):
                g[_nhom_nl(str(d))] += e
                dong[c].append((str(d).strip(), mon, e, ws.cell(r, 7).value))
        so[c] = g

    def theo_ca(ten, cot_ten, cot_ca, cot_suat):
        ws, i, ra, truoc = w[ten], -1, [[], [], [], []], 0
        for r in range(1, ws.max_row + 1):
            so_tt = isinstance(ws.cell(r, 1).value, (int, float))
            if so_tt and (str(ws.cell(r, cot_ca).value or '').strip().lower().startswith('ca ') or (ws.cell(r, 1).value == 1 and truoc >= 5)):
                i = min(i + 1, 3)
            if so_tt: truoc = ws.cell(r, 1).value
            t = ws.cell(r, cot_ten).value
            if i >= 0 and t and so_tt: ra[i].append((str(t).strip(), ws.cell(r, cot_suat).value))
        return ra
    b3 = next(n for n in w.sheetnames if 'BUOC 3' in kd[n])
    mau = next(n for n in w.sheetnames if '24' in n and not n.strip().endswith('(3)'))
    tieu_de = str(w[b3]['F1'].value or '')
    m = re.search(r'ngày\s*(\d+)\s*tháng\s*(\d+)', tieu_de, re.I)
    return dict(so=so, dong=dong, b3=theo_ca(b3, 3, 2, 4), mau=theo_ca(mau, 2, 3, 4), co_72=any('72' in n for n in w.sheetnames),
                ngay=f'{int(m.group(1)):02d}.{int(m.group(2)):02d}' if m else None)


def _chinh_ta(s):
    k = khoa_mon(s)
    for x, y in (('Y', 'I'), ('X', 'S'), ('TR', 'CH'), ('GI', 'D'), (' ', '')): k = k.replace(x, y)
    return k


def giong_ten(a, b):
    """Cùng một tên, chỉ khác chính tả hoặc lỗi gõ: mì/mỳ, sốt/xốt, sả/xả, thiếu một vài chữ cái."""
    a, b = _chinh_ta(a), _chinh_ta(b)
    return a == b or (min(len(a), len(b)) >= 5 and difflib.SequenceMatcher(None, a, b).ratio() >= .88)


def so_sanh(h, a):
    """So sổ hoàn thiện h với sổ app a (theo thực đơn): khác tên, sai số nhỏ, sai số lớn."""
    kq = dict(tong=0, khop=0, khac_ten=[], nho=[], lon=[], mon_khop=0, mon_tong=0, mau_khop=0, mau_tong=0, dong_dung=0, dong_tong=0, chinh_ta=[])
    for i, c in enumerate(CA):
        H, A = h['so'][c], a['so'][c]
        khoa = set(H) | set(A)
        lech = {k: (H[k], A[k]) for k in khoa if abs(H[k] - A[k]) > .02}
        kq['tong'] += len(khoa); kq['khop'] += len(khoa) - len(lech)
        chi_so = [k for k, v in lech.items() if v[1] == 0]
        for k in [k for k, v in lech.items() if v[0] == 0]:
            t = next((z for z in chi_so if abs(H[z] - A[k]) <= .02), None)
            if t:
                chi_so.remove(t); del lech[t]; del lech[k]
                kq['khac_ten'].append(dict(ca=c, so=t, thuc_don=k, sl=round(A[k], 2)))
        for k in sorted(lech):
            x, y = lech[k]
            kq['lon' if lech_lon(y, x) else 'nho'].append(dict(ca=c, nl=k, so=round(x, 2), thuc_don=round(y, 2), lech=round(x - y, 2)))
        # Nguyên liệu nằm đúng món ở bước 2: tên món trùng, hoặc tên này chứa tên kia
        con = collections.defaultdict(list)
        for nl, mon, *_ in a['dong'][c]: con[_nhom_nl(nl)].append(khoa_mon(mon or ''))
        for nl, mon, *_ in h['dong'][c]:
            ds, k = con.get(_nhom_nl(nl), []), khoa_mon(mon or '')
            hit = next((z for z in ds if z == k or z.startswith(k) or k.startswith(z) or giong_ten(z, k)), None)
            kq['dong_tong'] += 1
            if hit is not None: ds.remove(hit); kq['dong_dung'] += 1
        for nhan, g in (('mon', 'b3'), ('mau', 'mau')):
            ds = [z[0] for z in a[g][i]]
            for z in h[g][i]:
                kq[nhan + '_tong'] += 1
                hit = next((y for y in ds if khoa_mon(y) == khoa_mon(z[0])), None) or next((y for y in ds if giong_ten(y, z[0])), None)
                if hit is not None:
                    ds.remove(hit); kq[nhan + '_khop'] += 1
                    if khoa_mon(hit) != khoa_mon(z[0]): kq['chinh_ta'].append((z[0], hit))
    return kq


def doc_td(td_path):
    """Mở file thực đơn, giữ lại để xem nhiều ngày liên tiếp không phải đọc lại."""
    khoa = (td_path, os.path.getmtime(td_path))
    if khoa not in _td:
        _td.clear(); _td[khoa] = mo_wb(td_path, data_only=True)
    return _td[khoa]


def cac_ngay(td_path):
    """Các ngày có trong file thực đơn, theo thứ tự sheet."""
    kq = []
    for ws in doc_td(td_path).worksheets:
        n = ngay_cua(ws)
        if n and n not in kq: kq.append(n)
    return kq


def _giong(a, b):
    """Hai tên bộ món viết khác nhau chút ít: trùng ít nhất 60% số từ."""
    ta, tb = [set(re.findall(r'[A-Z0-9]+', khong_dau(x))) - {'SC', 'SUA'} for x in (a, b)]
    return len(ta & tb) / max(len(ta | tb), 1) >= 0.6


def tinh_suat_thieu(wb, ws, raw):
    """Suy ra số suất còn trống. Cả bộ trống: lấy theo Thực đơn tuần. Trống một phần: phần còn lại
    của số ở Thực đơn tuần, không có thì lấy số suất phổ biến trong bộ. Trả {ca: {bộ món: (số suất, căn cứ)}}."""
    tuan, kq = thuc_don_tuan(wb, ws), {}
    for ca in CA:
        dong = [x for x in raw[ca] if not BO_NGUYEN_LIEU.match(x['nl'].upper().strip())
                and not any(t.upper() in x['set'].upper() for t in Q['bo_bo_mon'])]
        cac_bo = list(dict.fromkeys(x['set'] for x in dong))
        for s in dict.fromkeys(x['set'] for x in dong if x['suat'] is None):
            ten_tuan = [(t, sa) for c, t, sa in tuan if c == ca]
            w = next((sa for t, sa in ten_tuan if _gon(t) == _gon(s)), None)
            if w is None:                        # tên hơi khác: chỉ nhận khi bên tuần chưa khớp bộ món nào khác
                gan_dung = [sa for t, sa in ten_tuan if _giong(t, s) and not any(_gon(t) == _gon(b) for b in cac_bo)]
                w = gan_dung[0] if len(gan_dung) == 1 else None
            co = [x['suat'] for x in dong if x['set'] == s and x['suat'] is not None]
            if not co:
                if w: kq.setdefault(ca, {})[s] = (w, f'Thực đơn tuần ghi {w} suất')
            elif w and w > max(co):
                kq.setdefault(ca, {})[s] = (w - max(co), f'Thực đơn tuần ghi {w} suất, trừ {max(co)} suất đã có')
            else:
                n = collections.Counter(co).most_common(1)[0][0]
                kq.setdefault(ca, {})[s] = (n, 'số suất của các dòng khác trong bộ món')
    return kq


def lap_so(td_path, ngay, suat_bo=None):
    """suat_bo: {ca: {bộ món: số suất}} người dùng điền cho bộ món thiếu suất; 0 là không phục vụ."""
    wb = doc_td(td_path)
    ws = tim_sheet(wb, ngay)
    so_tay, nhat_ky = tai_so_tay(), []
    raw = doc_ngay(ws)
    nhat_ky += raw['_mo_coi']
    tu_dong = tinh_suat_thieu(wb, ws, raw) if Q['tu_dien_suat'] == 'co' else {}
    for ca in CA:
        da_bao = set()
        for x in raw[ca]:
            n = (suat_bo or {}).get(ca, {}).get(x['set'])
            if x['suat'] is not None: continue
            if n is not None:
                if n > 0: x['suat'], x['tong'] = n, None
                else: x['khong_phuc_vu'] = True
            elif x['set'] in tu_dong.get(ca, {}):
                x['suat'], x['tong'], x['tu_suat'] = tu_dong[ca][x['set']][0], None, True
                if x['set'] not in da_bao:
                    da_bao.add(x['set'])
                    nhat_ky.append((ca, x['row'], x['set'], '', f"Đã tự sửa: thực đơn để trống số suất, sổ dùng {x['suat']} suất theo {tu_dong[ca][x['set']][1]}"))
    doi_chieu(wb, ws, raw, nhat_ky)
    return {ca: xu_ly_ca(raw[ca], ca, nhat_ky, so_tay) for ca in CA}, nhat_ky, ws.title


def goi_y_nhap(td_path, ngay):
    """Số trái cây, sữa thực đơn đã có sẵn theo ca, để gợi ý khi nhập. Tên sữa chưa rõ loại xếp vào 'khac'."""
    raw = doc_ngay(tim_sheet(doc_td(td_path), ngay))
    ten_so = {khong_dau(x['ten']): x['ten'] for x in Q['hang_cuoi_ca']}
    ten_so.update({khong_dau(x['thuc_don']): x['so'] for x in Q['goi_y_sua']})
    kq = {}
    for ca in CA:
        g = kq[ca] = dict(hang={}, khac={})
        for x in raw[ca]:
            if x['suat'] is None: continue
            k, ten = khoa_nl(x['nl']), None
            if 'TRÁNG MIỆNG' in x['set'].upper():
                ten = next((t['ten'].strip() for t in Q['trai_cay'] if k.startswith(khoa_nl(t['ten']))), None)
            elif k.startswith('SUA'):
                ten = ten_so.get(k)
            elif k == 'TM':
                g['khac']['TM'] = g['khac'].get('TM', 0) + x['suat']
                continue
            else:
                continue
            if ten is None:
                g['khac'][x['nl'].upper()] = g['khac'].get(x['nl'].upper(), 0) + x['suat']
            else:
                y = g['hang'].setdefault(ten, dict(suat=0, sl=0))
                y['suat'] += x['suat']
                y['sl'] = round(y['sl'] + (x['tong'] or 0), 4)
    return kq


def _gon(x):
    return re.sub(r'[^A-Z0-9]', '', khong_dau(x))


def thuc_don_tuan(wb, ws):
    """Các bộ món của ngày này ở sheet thực đơn cả tuần: [(ca, tên, suất)]."""
    hang = {}
    for s in wb.worksheets:        # sheet tuần tên không cố định: nhận qua dòng tiêu đề có 'Thứ 2'
        if s.sheet_state != 'visible' or s.title == ws.title: continue
        r = next((r for r in range(1, 7) for c in range(1, min(s.max_column, 12) + 1) if str(s.cell(r, c).value or '').strip().upper() == 'THỨ 2'), None)
        if r: hang[s] = r; break
    if not hang: return []
    t, tieu_de = next(iter(hang.items()))
    d, m = ngay_cua(ws).split('.')
    thu = ('THỨ 2', 'THỨ 3', 'THỨ 4', 'THỨ 5', 'THỨ 6', 'THỨ 7', 'CHỦ NHẬT')[datetime.date(Q['nam'], int(m), int(d)).weekday()]
    cot = next((c for c in range(1, t.max_column + 1) if str(t.cell(tieu_de, c).value or '').strip().upper() == thu), None)
    if not cot: return []
    kq, ca = [], None
    for r in range(tieu_de + 1, t.max_row + 1):
        m = re.search(r'SÁNG|TRƯA|TỐI|ĐÊM', str(t.cell(r, 1).value or '').upper())
        if m: ca = m.group()
        ten, sa = t.cell(r, cot).value, t.cell(r, cot + 2).value
        if ca and isinstance(ten, str) and isinstance(sa, (int, float)): kq.append((ca, ten.strip(), sa))
    return kq


def doi_chieu(wb, ws, raw, nhat_ky):
    """So hai chiều sheet ngày với sheet 'Thực đơn tuần': số suất và bộ món có ở bên này mà thiếu bên kia."""
    for ca, ten, sa in thuc_don_tuan(wb, ws):
        if any(t.upper() in ten.upper() for t in Q['bo_bo_mon']): continue
        tim = [x for x in raw[ca] if _gon(x['set']) == _gon(ten) or _gon(ten) in _gon(x['set'])
               or _gon(x['set']) in _gon(ten) or _giong(x['set'], ten)]
        if not tim:
            nhat_ky.append((ca, '', ten, '', f'Thực đơn tuần có bộ món này ({sa} suất) nhưng sheet ngày không có'))
            continue
        if '/' in ten: continue
        for x in tim:
            if x['suat'] is not None and not BO_NGUYEN_LIEU.match(x['nl'].upper()):
                if x['suat'] != sa:
                    nhat_ky.append((ca, x['row'], x['set'], x['nl'], f"Thực đơn tuần ghi {sa} suất, sheet ngày ghi {x['suat']}"))
                break


def thieu_suat(td_path, ngay, suat_bo=None):
    """Các bộ món còn dòng không có số suất, kèm số ở Thực đơn tuần nếu có."""
    wb = doc_td(td_path)
    ws = tim_sheet(wb, ngay)
    raw, tuan, kq = doc_ngay(ws), thuc_don_tuan(wb, ws), []
    tu_dong = tinh_suat_thieu(wb, ws, raw) if Q['tu_dien_suat'] == 'co' else {}
    for ca in CA:
        dong = [x for x in raw[ca] if not BO_NGUYEN_LIEU.match(x['nl'].upper().strip())
                and not any(t.upper() in x['set'].upper() for t in Q['bo_bo_mon'])]
        for s in dict.fromkeys(x['set'] for x in dong if x['suat'] is None):
            cua_bo = [x for x in dong if x['set'] == s]
            nl = [x['nl'] for x in cua_bo if x['suat'] is None]
            n = (suat_bo or {}).get(ca, {}).get(s)
            tu = tu_dong.get(ca, {}).get(s) if n is None else None
            kq.append(dict(ca=ca, bo=s, nl=nl, toan_bo=len(nl) == len(cua_bo), suat=tu[0] if tu else n, xong=n is not None or tu is not None,
                           can_cu=tu[1] if tu else None,
                           goi_y=next((sa for c, t, sa in tuan if c == ca and _gon(t) == _gon(s)), None)))
    return kq



try:
    dung_quy_tac(tai_quy_tac())
except (ValueError, TypeError, AttributeError, re.error) as e:
    # quy_tac.json hỏng thì vẫn mở được app để khôi phục từ Lịch sử
    print('Không đọc được quy_tac.json, tạm dùng quy tắc mặc định:', e)
    dung_quy_tac({})


if __name__ == '__main__':
    if len(sys.argv) < 4: raise SystemExit(__doc__)
    td, mau, ngay = sys.argv[1:4]
    ra_dir = sys.argv[4] if len(sys.argv) > 4 else 'ket_qua'
    os.makedirs(ra_dir, exist_ok=True)
    sach, nhat_ky, ten = lap_so(td, ngay)
    ra = os.path.join(ra_dir, f'So_{ngay}.xlsx')
    ghi_ra(mau, sach, ngay, nhat_ky, ra, ds_ngay=cac_ngay(td))
    print(f'{ngay} (sheet {ten}): ' + ', '.join(f"{c} {sum(len(m['nl']) for m in sach[c])} dòng/{len(sach[c])} món" for c in CA) + f'; cần soát {len(nhat_ky)}; -> {ra}')
