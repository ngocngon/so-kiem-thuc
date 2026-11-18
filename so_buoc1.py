"""Sổ bước 1: sổ kiểm tra hàng khô và sổ kiểm tra hàng tươi, điền từ thực đơn Việt và thực đơn Hàn.

Lượng lấy từ thực đơn: cộng cột TỔNG theo tên nguyên liệu và theo ca. Những gì thực đơn không có
(gia vị, ngày nhập, giờ nhập, mã kiểm dịch, đơn vị) lấy theo các ngày đã điền trước đó trong chính sổ.
Sheet ngày đã có số thì giữ nguyên; kết quả ghi ra file mới, không sửa file gốc.
"""
import collections, copy, datetime, math, re
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.styles import Font, Alignment
import tao_so as T

CA4 = ['SÁNG', 'TRƯA', 'TỐI', 'ĐÊM']
TEN_CA = {'SANG': 'SÁNG', 'TRUA': 'TRƯA', 'TOI': 'TỐI', 'DEM': 'ĐÊM'}
TRAI_CAY = ('DUA HAU', 'THANH LONG', 'CHUOI', 'OI')
# Nguyên liệu khô trong thực đơn -> tên mặt hàng trong sổ khô (khớp đầu tên, theo thứ tự)
KHO = [('GAO KHANG DAN', 'GẠO KHANG DÂN MB, 25KG/PAC'), ('GAO NEP', 'GẠO NẾP VỎ TẺ, 25KG/PAC'), ('NEP', 'GẠO NẾP VỎ TẺ, 25KG/PAC'),
       ('GAO LUT', 'GẠO LỨT ĐEN'), ('GAO', 'GẠO THƠM TA'),
       ('BANH DA CUA', 'BÁNH ĐA KHÔ (BÁNH ĐA CUA HẢI PHÒNG )'), ('DO XANH', 'HẠT ĐỖ XANH SÁT VỎ'), ('DO DEN', 'HẠT ĐỖ ĐEN LÒNG XANH (VN)'),
       ('DO TRANG', 'HẠT ĐỖ TRẮNG'), ('LAC NHAN', 'HẠT LẠC NHÂN'), ('HAT LAC NHAN', 'HẠT LẠC NHÂN'), ('MI TOM HA VIET', 'MỲ TÔM HÀ VIỆT'),
       ('MI TOM', 'MỲ TÔM KOKOMI 65G,MASAN,65G*100EA*BOX'), ('MY TOM', 'MỲ TÔM KOKOMI 65G,MASAN,65G*100EA*BOX'),
       ('MOC NHI', 'MỘC NHĨ SẠCH CẮT CHÂN 1KG/PAC'), ('NAM HUONG', 'NẤM HƯƠNG KHÔ,SEN VÀNG,500G*PAC'),
       ('PHONG TOM', 'BÁNH PHỒNG TÔM GÓI 200G,CƯỜNG GIANG,180G*200 GÓI*THÙNG'), ('BANH MI COM', 'BÁNH MỲ CỐM SỮA STAFF 40G,HỮU NGHỊ,40G*60EA*BOX'),
       ('MIEN DONG', 'MIẾN DONG VĨNH PHÚC 5KG/PAC'), ('MANG KHO', 'MĂNG KHÔ SỢI'), ('MANG SOI', 'MĂNG KHÔ SỢI'),
       ('CU CAI KHO', 'CỦ CẢI KHÔ'), ('RONG BIEN', 'RONG BIỂN KHÔ'),
       ('LA NEM', 'LÁ NEM GÓI'), ('ME CHAY', 'MÈ CHAY ÂU LẠC 160G*EA'), ('TUONG CA', 'TƯƠNG CÀ CHUA CHOLIMEX 2.1KG,CHOLIMEX,2,1KG*6CAN*BOX'),
       ('VUNG TRANG', 'VỪNG TRẮNG 1KG/GÓI'), ('BOT MI', 'BỘT MỲ MEIZAN 1KG,MEIZAN CLV,1KG*10EA/BOX'), ('BOT XU', 'BỘT CHIÊN XÙ PANKO VÀNG SOFUTO GÓI 1KG'),
       ('BOT TRA XANH', 'BỘT TRÀ XANH'), ('BOT KA KAO', 'BỘT CA CAO'), ('GIA VI LAU THAI', 'GIA VỊ LẨU THÁI'),
       ('BANH MI RUOC', 'BÁNH MỲ RUỐC'), ('BANH MY RUOC', 'BÁNH MỲ RUỐC')]
# Hàng khô trong thực đơn Hàn (mì gói, lá kim, tương, bột, gạo): ghi vào sổ khô, đúng ca
HAN_KHO = re.compile(r'^(MY|MI|MYJIN|LA KIM|GAO|TUONG|BOT)\b|^MYJIN')
BO_QUA = re.compile(r'^(SC|TM|SUA CHUA|TRANG MIENG|THACH.*|BA BAU.*)$')

# ---- Giờ, ngày nhập hàng tươi (anh chị xác nhận ở sổ 04.11) ----
# Rau củ: sáng 15H hôm trước; trưa, tối 7H; đêm 15H. Trái cây: sáng 15H hôm trước; trưa, tối 8H; đêm 15H.
GIO_RAU = ((15, -1), (7, 0), (7, 0), (15, 0))
GIO_TRAI = ((15, -1), (8, 0), (8, 0), (15, 0))
# Bún, phở, đậu phụ, bánh cuốn nhập theo khung giờ cùng ngày
GIO_KHUNG = ('1H', '8H', '14H', '20H')
KHUNG = re.compile(r'^(TC, )?(BUN|PHO|DAU PHU|BANH CUON)\b')
# Đồ muối, đồ ngâm: 7H ngày hôm trước
MUOI = re.compile(r'\b(MUOI|NGAM)\b|^DUA CHUA|^DUA GOP')
# Giò chả nhập 7H30 sáng hàng ngày (giò cái, giò lụa, chả lụa, giò bò) và hải sản sống:
# ca sáng ăn hàng 7H30 hôm trước; ca trưa, tối, đêm ăn hàng 7H30 cùng ngày
GIO_7H30 = re.compile(r'^(NGAO|CON CHAI|GIO BO|GIO CAI|GIO LUA|GIO CHA|CHA LUA|HEN|OC)\b')
# Hàng nhập theo lịch: thứ trong tuần (0 là thứ 2), giờ. Lấy lần giao gần nhất trước ngày làm sổ.
TRUNG = re.compile(r'^(TRUNG|TVL)\b')
LICH_TRUNG = ((0, 2, 4), '8H30')               # thứ 2, 4, 6
LICH_KIM_CHI = ((1,), '10H')                   # thứ 3
LICH_DL = ((1, 3, 5), '8H30')                  # hàng đông lạnh, thịt: thứ 3, 5, 7
# Hàng đông lạnh và thịt: nhận theo đầu tên, có dấu để không lẫn CÁ với CÀ
HANG_LO = ('THỊT', 'BA CHỈ', 'CÁ ', 'CHẢ', 'CHẠO', 'GÀ', 'TRÂU', 'BÒ', 'BẮP BÒ', 'SƯỜN', 'XƯƠNG', 'NEM', 'TAI LỢN', 'BÁNH BAO', 'GIÒ CÁI', 'GIÒ LỤA',
           'MÓNG', 'LƯỜN', 'ĐÙI', 'VỊT', 'TÔM', 'MỰC', 'ẾCH', 'XÚC XÍCH', 'HEM', 'DẺ SƯỜN', 'CUA', 'NẠC', 'MÁ ', 'KHOANH', 'CHIM', 'MAN DU', 'LÒNG', 'ĐUÔI')
# Tên trong sổ -> dòng trong sheet MÃ SỐ KIỂM DỊCH khi không đoán được bằng mã hay bằng tên
# (lấy từ mã anh chị đã ghi ở sổ 01-02.11 và ghi chú sổ 04.11)
BI_DANH_KD = {'TOM': 'TOM THE DL', 'SUON': 'SUON HEO', 'THIT VIT': 'VIT SONG HONG', 'VIT': 'VIT SONG HONG', 'THIT GA 1/4': 'GA GOC TU', 'GA 1/4': 'GA GOC TU',
              'LUON GA': 'GA GOC TU', 'BAP BO': 'THIT BO BAP SHTT', 'TRAU VAI': 'TRAU DUI'}
# Giấy kiểm dịch anh chị xác nhận nhưng sheet MÃ SỐ KIỂM DỊCH chưa ghi: (dòng trong sheet, ngày, tháng, mã).
# App coi như ô đó đã có mã; giấy mới hơn ghi trong sheet vẫn được ưu tiên, quá CU_NHAT ngày thì hết lô.
BO_SUNG_KD = [('CÁ BASA FILE', 31, 10, '208/HP-CBCLSP/2026')]
CU_NHAT = 14          # giấy kiểm dịch cũ hơn số ngày này thì coi như hết lô, ghi theo lịch giao 3-5-7


def K(s):
    return re.sub(r'\s+', ' ', T.khong_dau(str(s or '')).upper()).strip()


def ngay_sheet(ten, nam):
    m = re.fullmatch(r'\s*(\d{1,2})\.(\d{1,2})\s*', ten)
    if not m: return None
    try: return datetime.date(nam, int(m.group(2)), int(m.group(1)))
    except ValueError: return None


def lam_tron(x, buoc=0.5):
    return max(buoc, round(x / buoc) * buoc)


# ---------------- Đọc nguồn ----------------
def khoa_pivot(ten):
    """Gộp như bảng Pivot của Excel: không phân biệt hoa thường nhưng giữ nguyên dấu cách,
    nên 'CÀ RỐT' và ' CÀ RỐT' hay 'NẠC VAI ' là các dòng riêng (đúng như sổ anh chị điền)."""
    return str(ten).upper()


def tong_viet(td_path, ngay):
    """{ca: {khoá Pivot: [tên hiển thị, tổng, nl gốc, dl]}} và {ca: [(bộ món, suất, có dòng sữa)]}.
    Dòng có trong thực đơn nhưng tổng bằng 0 (chưa có suất) vẫn giữ, như Pivot."""
    wb = T.doc_td(td_path)
    raw = T.doc_ngay(T.tim_sheet(wb, ngay))
    kq, bo = {ca: collections.OrderedDict() for ca in CA4}, {ca: collections.OrderedDict() for ca in CA4}
    for ca in CA4:
        for x in raw[ca]:
            b = bo[ca].setdefault(x['khoi'], dict(ten=x['set'], suat=x['suat'], sua=False))
            if x['suat'] and not b['suat']: b['suat'] = x['suat']
            tong = x['tong']
            if tong is None and x['suat'] and x['dl'] is not None:
                tong = T.tinh(T.cong_thuc_cua(x['nl']), x['dl'], x['suat'])
            if tong is None: tong = 0
            k = K(x['nl'])
            if k.startswith('SUA') and tong: b['sua'] = True
            v = kq[ca].setdefault(khoa_pivot(x.get('nl_goc', x['nl'])), [str(x['nl']).strip().upper(), 0, x['nl'], x['dl']])
            v[1] = round(v[1] + tong, 4)
    return kq, {ca: list(v.values()) for ca, v in bo.items()}


def tong_han(han_path, ngay):
    """Thực đơn Hàn: sheet có 'ngày dd.mm' ở các dòng đầu, cột G ghi ca. {ca: {khoá: [tên, tổng]}}."""
    kq = {ca: collections.OrderedDict() for ca in CA4}
    if not han_path: return kq, None
    d, m = ngay.split('.')
    mau = re.compile(rf'NGAY\s*0?{int(d)}\.0?{int(m)}\b')
    wb, wv = T.mo_wb(han_path), T.mo_wb(han_path, data_only=True)
    ws = next((s for s in wb.worksheets if s.sheet_state == 'visible'
               and any(mau.search(K(s.cell(r, c).value)) for r in range(1, 4) for c in range(1, 8))), None)
    if ws is None: return kq, None
    v = wv[ws.title]
    for r in range(1, ws.max_row + 1):
        ca = TEN_CA.get(K(ws.cell(r, 7).value))
        ten, dl, suat = ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value
        if not ca or not isinstance(ten, str) or not ten.strip() or K(ten) == 'NGUYEN LIEU': continue
        tong = v.cell(r, 6).value
        if not isinstance(tong, (int, float)):
            if isinstance(dl, (int, float)) and isinstance(suat, (int, float)): tong = dl * suat / 1000
            elif ws.cell(r, 6).value is not None or isinstance(suat, (int, float)): tong = 0
            else: continue
        # Gộp như Pivot (giữ dấu cách); tên gõ thừa chữ như DĐẬU PHỤ vẫn là dòng riêng, chỉ sửa chữ hiển thị
        hien = re.sub(r'^D(?=Đ)', '', ten.strip().upper())
        x = kq[ca].setdefault(khoa_pivot(ten), [hien, 0])
        x[1] = round(x[1] + tong, 4)
    return kq, ws.title


# ---------------- Đọc sổ ----------------
def tach_khoi(ws, cot_so, mau_ca):
    """Các khối ca trong sheet sổ: [(nhãn, ca, hq, hàng tiêu đề, [hàng mặt hàng])]."""
    kq = []
    for r in range(7, ws.max_row + 1):
        b = ws.cell(r, 2).value
        if not isinstance(b, str) or not b.strip(): continue
        m = mau_ca.fullmatch(K(b))
        if m:
            kq.append((b, TEN_CA[m.group(1)], bool(m.group(2)), r, []))
        elif kq:
            kq[-1][4].append(r)
    return kq


MAU_KHO = re.compile(r'CA (SANG|TRUA|TOI|DEM)()')
MAU_TUOI = re.compile(r'(SANG|TRUA|TOI|DEM)( HQ)?')


def co_so(ws, cot, mau):
    return any(isinstance(ws.cell(r, cot).value, (int, float)) for k in tach_khoi(ws, cot, mau) for r in k[4])


def sheet_ngay(wb, nam):
    return sorted(((ngay_sheet(s.title, nam), s) for s in wb.worksheets if ngay_sheet(s.title, nam)), key=lambda x: x[0])


# ---------------- Ghi sheet ----------------
def chup(ws, r):
    hang = {c: (copy.copy(ws.cell(r, c)._style), ws.cell(r, c).value) for c in range(1, ws.max_column + 1)}
    gop = [(m.min_col, m.max_col) for m in ws.merged_cells.ranges if m.min_row == m.max_row == r]
    return dict(r=r, hang=hang, gop=gop, cao=ws.row_dimensions[r].height)


def dat(ws, r, mau, du_lieu, so_dau=None):
    """Ghi một hàng theo hàng mẫu: kiểu, công thức dời hàng, ô gộp; du_lieu {cột: giá trị} ghi đè."""
    for c, (st, v) in mau['hang'].items():
        o = ws.cell(r, c)
        o._style = copy.copy(st)
        if isinstance(v, str) and v.startswith('='):
            v = Translator(v, origin=f'A{mau["r"]}').translate_formula(f'A{r}')
        o.value = v if c not in du_lieu else du_lieu[c]
    if so_dau is not None: ws.cell(r, 1).value = so_dau
    for a, b in mau['gop']: ws.merge_cells(start_row=r, end_row=r, start_column=a, end_column=b)
    ws.row_dimensions[r].height = mau['cao']


def lam_sheet(wb, truoc, ten, ngay, nam, cot_so, mau_ca):
    """Sheet mới chép từ sheet ngày trước, giữ phần đầu, xoá thân. Trả (ws, mẫu tiêu đề ca, mẫu hàng)."""
    khoi = tach_khoi(truoc, cot_so, mau_ca)
    if not khoi or not any(k[4] for k in khoi): raise ValueError(f'Sheet {truoc.title} không có khối ca để làm mẫu')
    ws = wb.copy_worksheet(truoc)
    ws.title = ten
    wb.move_sheet(ws, wb.sheetnames.index(truoc.title) + 1 - wb.sheetnames.index(ws.title))
    k2 = next((k for k in khoi[1:] if k[4]), khoi[0])
    hang_mau = next((r for k in khoi for r in k[4] if r > k[3] + 1), khoi[0][4][0])
    m_ca, m_hang = chup(truoc, k2[3]), chup(truoc, hang_mau)
    ws.print_title_rows = truoc.print_title_rows
    for row in ws.iter_rows(min_row=1, max_row=3):
        for o in row:
            if isinstance(o.value, str) and 'kiểm tra' in o.value and re.search(r'Ngày\s*\d', o.value):
                o.value = re.sub(r'Ngày\s*\d+\s*([Tt]háng)\s*\d+\s*năm\s*\d+', lambda m: f'Ngày {ngay.day:02d} {m.group(1)} {ngay.month:02d} năm {nam}', o.value)
    T.xoa_than(ws, 7)
    return ws, m_ca, m_hang


# ---------------- Lịch sử trong sổ ----------------
class LichSu:
    """Thông tin mặt hàng ở các ngày trước: lần gần nhất theo tên (ưu tiên cùng ca)."""
    def __init__(self, ds, cot):
        self.lan = collections.defaultdict(list)      # khoá tên -> [(ngày, ca, hq, {cột: giá trị})], mới nhất trước
        self.so = collections.defaultdict(list)       # (ca, khoá) -> [số] các ngày trước, mới nhất trước
        self.thu_tu = collections.defaultdict(list)   # (ca, hq) -> tên theo thứ tự của ngày gần nhất
        self.ngay = []
        for d, ws, mau in reversed(ds):
            self.ngay.append(d)
            for nhan, ca, hq, _, hang in tach_khoi(ws, cot, mau):
                for r in hang:
                    ten = ws.cell(r, 2).value
                    k = K(ten)
                    v = {c: ws.cell(r, c).value for c in range(2, 12)}
                    self.lan[k].append((d, ca, hq, v))
                    if isinstance(ws.cell(r, cot).value, (int, float)): self.so[(ca, k)].append(ws.cell(r, cot).value)
                    if len(self.ngay) == 1: self.thu_tu[(ca, hq)].append(str(ten).strip())

    def gan(self, k, ca=None, hq=False):
        ds = self.lan.get(k, [])
        return next((x for x in ds if x[1] == ca and x[2] == hq), None) or next(iter(ds), None)

    def don_vi(self, k, cot):
        """Đơn vị hay gặp nhất, ưu tiên ca sáng và trưa vì sổ hay lệch cột đơn vị ở ca tối, đêm."""
        c = collections.Counter(x[3].get(cot) for x in self.lan.get(k, []) if x[1] in ('SÁNG', 'TRƯA') and x[3].get(cot))
        c = c or collections.Counter(x[3].get(cot) for x in self.lan.get(k, []) if x[3].get(cot))
        return c.most_common(1)[0][0] if c else None


# ---------------- Sổ hàng khô ----------------
def nhom_kho(k):
    for a, ten in KHO:
        if k == a or k.startswith(a + ' ') or (a not in ('GAO', 'NEP') and a in k): return ten
    return None


def ten_ds(wb, ten):
    """Tên đúng như sheet 'ds mặt hàng' để công thức VLOOKUP tra được; kèm cờ có tìm thấy không."""
    ds = next((s for s in wb.worksheets if K(s.title) == 'DS MAT HANG'), None)
    if ds is None: return ten, True
    k = K(ten)
    for (v,) in ds.iter_rows(min_row=2, max_col=1, values_only=True):
        if v and K(v) == k: return v, True
    for (v,) in ds.iter_rows(min_row=2, max_col=1, values_only=True):
        if v and K(v).startswith(k): return v, True
    chu = set(re.findall(r'[A-Z0-9]+', k))
    for (v,) in ds.iter_rows(min_row=2, max_col=1, values_only=True):
        if v and chu <= set(re.findall(r'[A-Z0-9]+', K(v))): return v, True
    return ten, False


def dien_kho(wb, ngay, nam, viet, bo, han, ls, truoc, ghi_chu):
    ten = f'{ngay.day:02d}.{ngay.month:02d}'
    ws, m_ca, m_hang = lam_sheet(wb, truoc, ten, ngay, nam, 7, MAU_KHO)
    # Gia vị: mặt hàng có ở cả 4 ca trong phần lớn các ngày trước
    du_ca = lambda k: sum(1 for d in ls.ngay if len({x[1] for x in ls.lan[k] if x[0] == d}) == 4)
    gia_vi = [t for t in ls.thu_tu[('SÁNG', False)] if du_ca(K(t)) * 2 > len(ls.ngay) and not re.match(r'(GAO|SUA|THACH)', K(t))]
    sua_ten = [t for t in ls.thu_tu[('SÁNG', False)] if K(t).startswith('SUA')][:3] or ['SỮA CHUA HANOI MILK 100G,SỮA HÀ NỘI,100G*48EA*BOX']
    thach = next((t for t in ls.thu_tu[('SÁNG', False)] if K(t).startswith('THACH')), None)
    r, so = 7, None
    for ca in CA4:
        dong = collections.OrderedDict()            # tên sổ -> [số, ghi chú]
        them = collections.defaultdict(float)
        for k, (hien, tong, *_) in viet[ca].items():
            k = K(k)
            if k.startswith('SUA') or not tong: continue
            t = nhom_kho(k)
            if not t: continue
            t, co = ten_ds(wb, t)
            if any(K(t) == K(g) for g in gia_vi): them[K(t)] += tong
            else: dong.setdefault(t, [0, None if co else 'Chưa có trong sheet ds mặt hàng nên các cột tra cứu sẽ báo #N/A; hãy thêm vào danh sách hoặc sửa tên'])[0] += tong
        for k, (hien, tong) in han[ca].items():
            k = K(hien)
            if not HAN_KHO.search(k) or not tong: continue
            k2 = re.sub(r'^MYJIN\b', 'MY JIN', k)               # mì gói: tìm theo tên trong ds mặt hàng, không gộp với rong biển
            t, co = ten_ds(wb, k2 if re.match(r'(MY|MI)\b', k2) else (nhom_kho(k) or hien))
            if not co: t = hien
            dong.setdefault(t, [0, 'Hàng khô bếp Hàn' + ('' if co else ', chưa có trong sheet ds mặt hàng nên các cột tra cứu sẽ báo #N/A')])[0] += tong
        gao = sorted((t for t in dong if K(t).startswith('GAO')), key=lambda t: next((i for i, a in enumerate(('GAO THOM TA', 'GAO KHANG', 'GAO NEP')) if K(t).startswith(a)), 9))
        sua = sum(v[1] for k, v in viet[ca].items() if K(k).startswith('SUA'))
        bo_sua = [b for b in bo[ca] if not b['sua'] and b['suat'] and re.search(r'\bSUA\b', K(b['ten']))]
        sua += sum(b['suat'] for b in bo_sua)
        hang = [(t, dong[t][0], dong[t][1]) for t in gao]
        if sua:
            n = int(round(sua))
            for i, t in enumerate(sua_ten):
                hang.append((t, math.ceil(n / len(sua_ten)) if i == 0 else n // len(sua_ten),
                             'Có bộ món ghi sữa nhưng thực đơn không có dòng sữa: ' + ', '.join(f"{b['ten']} ({b['suat']})" for b in bo_sua) if bo_sua and i == 0 else None))
        if thach:
            v = ls.so.get((ca, K(thach)))
            if v: hang.append((thach, collections.Counter(v[:7]).most_common(1)[0][0], None))
        hang += [(t, v[0], v[1]) for t, v in dong.items() if t not in gao]
        for t in gia_vi:
            v = ls.so.get((ca, K(t)), [])[:7]
            if not v: continue
            sl = lam_tron(sorted(v)[len(v) // 2] if len(v) % 2 else sum(sorted(v)[len(v) // 2 - 1:len(v) // 2 + 1]) / 2)
            nd = f'Gia vị ước theo số giữa của {len(v)} ngày trước'
            if them.get(K(t)):
                sl = round(sl + them[K(t)], 2); nd += f', cộng {them[K(t)]:g} theo thực đơn'
            hang.append((t, sl, nd))
        gv = []
        dat(ws, r, m_ca, {2: f'CA {ca}'})
        if r == 7: ws.cell(r, 1).value = None
        r += 1
        for t, sl, nd in hang:
            k = K(t)
            cu = ls.gan(k, ca)
            ngay_nhap = cu[3].get(5) if cu else None
            if not cu:
                ngay_nhap = datetime.datetime.combine(ngay - datetime.timedelta(days=1), datetime.time())
                ghi_chu.append(dict(sheet=ten, o=f'E{r}', ca=ca, ten=t, nd='Mặt hàng mới: ngày nhập để ngày hôm trước, đơn vị cần xem lại'))
            dv = ls.don_vi(k, 6) or ('GÓI' if re.match(r'(MY|MI)\b|MYJIN|LA KIM', k) else 'KG')
            dat(ws, r, m_hang, {2: t, 5: ngay_nhap, 6: dv, 7: round(sl, 4)}, 1 if so is None else None)
            ws.cell(r, 7).number_format = T.SO_THAT
            so = r
            if nd and nd.startswith('Gia vị'):
                gv.append(r)
                if 'cộng' in nd: ghi_chu.append(dict(sheet=ten, o=f'G{r}', ca=ca, ten=t, nd=nd))
            elif nd and nd != 'Hàng khô bếp Hàn': ghi_chu.append(dict(sheet=ten, o=f'G{r}', ca=ca, ten=t, nd=nd))
            r += 1
        if gv:
            ghi_chu.append(dict(sheet=ten, o=f'G{gv[0]}', ca=ca, ten=f'Gia vị, G{gv[0]}:G{gv[-1]}',
                                nd=f'Thực đơn không có gia vị: số ước theo số giữa (trung vị) của {min(len(ls.ngay), 7)} ngày đã điền trước, làm tròn 0,5. Hãy sửa theo thực tế'))
    ws.print_area = f'A1:{_cot_cuoi(truoc)}{r - 1}'
    return ws


def _cot_cuoi(ws):
    m = re.search(r'\$?([A-Z]+)\$?\d+$', str(ws.print_area or ''))
    return m.group(1) if m else 'Q'


# ---------------- Sổ hàng tươi ----------------
def chu_gio(gio, d):
    return f'{gio} {d.day:02d}.{d.month:02d}.{d.year}'


def lan_giao(ngay, thu, cung_ngay=False):
    """Lần giao gần nhất trước ngày làm sổ; cung_ngay thì tính cả lần giao sáng chính ngày đó."""
    d = ngay if cung_ngay else ngay - datetime.timedelta(days=1)
    while d.weekday() not in thu: d -= datetime.timedelta(days=1)
    return d


def la_lo(ten):
    return str(ten).strip().upper().startswith(HANG_LO)


class KiemDich:
    """Sheet MÃ SỐ KIỂM DỊCH: mỗi dòng một mặt hàng, mỗi cột một lần giao (ngày ở dòng 1), ô là mã kiểm dịch."""
    def __init__(self, wb, nam):
        ws = next((s for s in wb.worksheets if K(s.title) == 'MA SO KIEM DICH'), None)
        self.cot, self.hang, self.o, self.ten = [], {}, {}, {}
        if ws is None: return
        # Excel ghi 1.10 thành số 1.1 và 20.10 thành 20.1: tháng 1 hay tháng 10 thì chọn theo cột ngày rõ ràng gần nhất
        tho = []
        for c in range(2, ws.max_column + 1):
            m = re.match(r'\s*(\d{1,2})\.(\d{1,2})', str(ws.cell(1, c).value or ''))
            if not m: continue
            d, th = int(m.group(1)), int(m.group(2))
            ung = [x for x in ([1, 10] if th == 1 else [th]) if 1 <= x <= 12 and _hop_le(d, x)]
            if ung: tho.append((c, d, ung))
        ro = [(i, datetime.date(nam, u[0], d)) for i, (c, d, u) in enumerate(tho) if len(u) == 1]
        for i, (c, d, ung) in enumerate(tho):
            if len(ung) == 1: ngay = datetime.date(nam, ung[0], d)
            else:
                gan = min(ro, key=lambda x: (abs(x[0] - i), x[0] < i))[1] if ro else datetime.date(nam, 10, 1)
                ngay = min((datetime.date(nam, x, d) for x in ung), key=lambda x: abs((x - gan).days))
            if self.cot and (self.cot[-1][1] - ngay).days > 150: ngay = ngay.replace(year=ngay.year + 1)   # sang năm mới
            self.cot.append((c, ngay))
        for r in range(2, ws.max_row + 1):
            ten = ws.cell(r, 1).value
            if not ten: continue
            self.hang[r] = K(ten)
            self.ten[r] = re.sub(r'\s+', ' ', str(ten)).strip()
            for c, _ in self.cot:
                v = ws.cell(r, c).value
                ma = [x.strip().lstrip("'-").strip() for x in str(v or '').split('\n') if x.strip().lstrip("'-").strip()]
                if ma: self.o[(r, c)] = ma
        for ten, d, th, ma in BO_SUNG_KD:
            r = next((r for r, h in self.hang.items() if h == K(ten)), None)
            if r is None or not _hop_le(d, th): continue
            ngay = datetime.date(nam, th, d)
            c = next((c for c, x in self.cot if x == ngay), None)
            if c is None:                       # sheet chưa có cột ngày này: thêm cột ảo đúng thứ tự ngày
                c = ('bo_sung', ngay)
                self.cot.append((c, ngay)); self.cot.sort(key=lambda x: x[1])
            self.o.setdefault((r, c), [ma])

    def _ung_vien(self, k, ma_cu):
        """Các dòng có thể là của mặt hàng, kèm độ ưu tiên: 3 dòng có mã đã dùng, 2 bí danh, 1 trùng tên, 0 giống tên."""
        chu = set(re.findall(r'[A-Z0-9/-]+', k)) - {'KO', 'KHONG', 'CO', 'DA', 'THIT', 'DL'}
        giong = lambda r: len(chu & set(re.findall(r'[A-Z0-9/-]+', self.hang[r])))
        gon = lambda x: re.sub(r'[^A-Z0-9]', '', re.sub(r'\b(THIT|DL)\b', '', x))
        uv = {}
        if ma_cu:
            for r in self.hang:
                if any(ma_cu in self.o.get((r, c), ()) for c, _ in self.cot): uv[r] = 3
        for a, b in BI_DANH_KD.items():
            if k == a or k.startswith(a + ' '):
                for r, h in self.hang.items():
                    if h == b: uv[r] = max(uv.get(r, 0), 2)
        gk = gon(k)
        for r, h in self.hang.items():
            gh = gon(h)
            if len(gh) >= 3 and (gh in gk or (len(gk) >= 3 and gh.startswith(gk))): uv[r] = max(uv.get(r, 0), 1)
        if not uv:
            for r in self.hang:
                if giong(r) >= 2: uv[r] = 0
        return uv, giong

    def _lan_moi(self, r, ngay):
        return next(((c, d) for c, d in reversed(self.cot) if d < ngay and (r, c) in self.o), None)

    def dong_cua(self, k, ma_cu, ngay=None):
        """Dòng của mặt hàng. Nhiều dòng hợp (cá basa có dòng FILE, NGUYỄN KHANH, HÙNG PHÁT...) thì lấy dòng
        có giấy mới nhất trước ngày làm sổ; bằng ngày thì ưu tiên dòng có mã đã dùng, rồi bí danh, rồi giống tên."""
        uv, giong = self._ung_vien(k, ma_cu)
        if not uv: return None
        ngay = ngay or datetime.date.max
        def diem(r):
            x = self._lan_moi(r, ngay)
            return (x[1] if x else datetime.date.min, uv[r], giong(r))
        return max(uv, key=diem)

    def tra(self, k, ma_cu, ngay, gan_day=()):
        """(ngày giao, mã) của lần giao gần nhất trước ngày làm sổ, không quá CU_NHAT ngày.
        Ô có nhiều mã thì giữ mã đã dùng các ngày trước."""
        r = self.dong_cua(k, ma_cu, ngay)
        if r is None: return None
        x = self._lan_moi(r, ngay)
        if not x or (ngay - x[1]).days > CU_NHAT: return None
        ma = self.o[(r, x[0])]
        return x[1], next((m for m in [ma_cu, *gan_day] if m in ma), ma[0])


def _hop_le(d, th):
    try: datetime.date(2000, th, d); return True
    except ValueError: return False


def gio_tuoi(t, k, i, hq, ngay, kd, ma_cu, trai=False, gan_day=()):
    """Giờ nhập (cột C), mã kiểm dịch (cột K) và cách tính, theo quy tắc anh chị đã xác nhận."""
    hom_truoc = ngay - datetime.timedelta(days=1)
    if trai:
        g, lui = GIO_TRAI[i]
        return chu_gio(f'{g}H', ngay + datetime.timedelta(days=lui)), None, 'trái cây'
    if MUOI.search(k): return chu_gio('7H', hom_truoc), None, 'đồ muối, ngâm'
    if KHUNG.match(k): return chu_gio(GIO_KHUNG[i], ngay), None, 'khung giờ'
    # Trứng giao 8H30 thứ 2-4-6: ca sáng ăn lần giao trước, ca trưa, tối, đêm ăn lần giao 8H30 sáng cùng ngày
    if TRUNG.match(k): return chu_gio(LICH_TRUNG[1], lan_giao(ngay, LICH_TRUNG[0], cung_ngay=i > 0)), None, 'trứng 2-4-6'
    if k.startswith('KIM CHI'): return chu_gio(LICH_KIM_CHI[1], lan_giao(ngay, LICH_KIM_CHI[0])), None, 'kim chi thứ 3'
    if GIO_7H30.match(k): return chu_gio('7H30', ngay if i > 0 else hom_truoc), None, '7H30'
    x = kd.tra(k, ma_cu, ngay, gan_day) if la_lo(t) or ma_cu else None
    if x: return chu_gio('8H30', x[0]), x[1], 'kiểm dịch'
    if la_lo(t): return chu_gio(LICH_DL[1], lan_giao(ngay, LICH_DL[0])), None, 'đông lạnh 3-5-7'
    g, lui = GIO_RAU[i]
    return chu_gio(f'{g}H', ngay + datetime.timedelta(days=lui)), None, 'rau củ'


def dien_tuoi(wb, ngay, nam, viet, han, ls, truoc, nhap, ghi_chu, kd):
    ten = f'{ngay.day:02d}.{ngay.month:02d}'
    ws, m_ca, m_hang = lam_sheet(wb, truoc, ten, ngay, nam, 4, MAU_TUOI)
    nhan = {}
    for nh, ca, hq, *_ in tach_khoi(truoc, 4, MAU_TUOI): nhan[(ca, hq)] = nh
    r, dau = 7, True
    # Mã kiểm dịch đã dùng các ngày trước, mới nhất trước: ô có nhiều mã thì ưu tiên mã đang dùng
    gan_day = [str(x[3].get(11)).strip() for d in ls.ngay for ds in ls.lan.values() for x in ds if x[0] == d and x[3].get(11)]
    for hq in (False, True):
        for i, ca in enumerate(CA4):
            nguon = han[ca] if hq else viet[ca]
            hang = []
            for k, v in nguon.items():
                k = K(v[0])
                if hq and HAN_KHO.search(k): continue
                if k.startswith('SUA') or BO_QUA.match(k) or nhom_kho(k): continue
                hang.append((v[0], v[1], v[2] if not hq else v[0], v[3] if not hq else None))
            if not hang: continue
            hang.sort(key=lambda x: T.khong_dau(x[0]).upper())
            trai = [] if hq else [t for t in ls.thu_tu[(ca, False)] if K(t) in TRAI_CAY] or \
                ['DƯA HẤU', 'THANH LONG', 'CHUỐI'] + ([] if ca == 'SÁNG' else ['ỔI'])
            dat(ws, r, m_ca, {2: nhan.get((ca, hq), f'{ca} HQ' if hq else ca)}, 1 if dau else None)
            if dau: ws.cell(r, 1).value = None
            r += 1
            for t in trai:
                cu = ls.gan(K(t), ca)
                gio = gio_tuoi(t, K(t), i, hq, ngay, kd, None, trai=True)[0]
                so = (nhap.get(ca) or {}).get(t.strip(), {}).get('sl')
                dat(ws, r, m_hang, {2: t, 3: gio, 4: so, 5: (cu[3].get(5) if cu else None) or 'KG', 11: None, 17: None, 18: ca},
                    1 if dau else None)
                dau = False; r += 1
            for t, sl, nl, dl in hang:
                k = K(t)
                cu = ls.gan(k, ca, hq)
                ma_cu = next((str(x[3].get(11)).strip() for x in ls.lan.get(k, []) if x[3].get(11) and len(str(x[3].get(11)).strip()) > 5), None)
                gio, ma, cach = gio_tuoi(t, k, i, hq, ngay, kd, ma_cu, gan_day=gan_day)
                dv = ls.don_vi(k, 5) or T.don_vi(nl, dl if dl is not None else 0)
                dat(ws, r, m_hang, {2: t, 3: gio, 4: round(sl, 4), 5: dv, 11: ma, 17: ca if ma else None, 18: ca}, 1 if dau else None)
                ws.cell(r, 4).number_format = T.SO_THAT
                if not sl:
                    ghi_chu.append(dict(sheet=ten, o=f'D{r}', ca=ca + (' HQ' if hq else ''), ten=t,
                                        nd='Thực đơn có mặt hàng này nhưng tổng bằng 0 (chưa có suất hoặc định lượng); vẫn ghi dòng như Pivot, hãy điền số thực tế'))
                if la_lo(t) and cach == 'đông lạnh 3-5-7' and kd.dong_cua(k, ma_cu, ngay) is not None:
                    ghi_chu.append(dict(sheet=ten, o=f'K{r}', ca=ca + (' HQ' if hq else ''), ten=t,
                                        nd='Có trong sheet MÃ SỐ KIỂM DỊCH nhưng chưa có mã trước ngày này: ghi theo lịch giao 3-5-7, chưa có mã'))
                elif cu is None and cach in ('kiểm dịch', 'đông lạnh 3-5-7'):
                    ghi_chu.append(dict(sheet=ten, o=f'C{r}', ca=ca + (' HQ' if hq else ''), ten=t,
                                        nd=('Mặt hàng mới, ngày nhập và mã lấy theo dòng ' + kd.ten[kd.dong_cua(k, ma_cu, ngay)] + ' trong sheet MÃ SỐ KIỂM DỊCH')
                                        if cach == 'kiểm dịch' else 'Mặt hàng mới, coi là hàng đông lạnh giao thứ 3-5-7; nếu có giấy kiểm dịch hãy thêm vào sheet MÃ SỐ KIỂM DỊCH'))
                dau = False; r += 1
    ws.print_area = f'A1:{_cot_cuoi(truoc)}{r - 1}'
    return ws


def ghi_chu_sheet(wb, ghi_chu, tieu_de):
    if 'GHI CHÚ TỰ ĐIỀN' in wb.sheetnames: del wb['GHI CHÚ TỰ ĐIỀN']
    ws = wb.create_sheet('GHI CHÚ TỰ ĐIỀN')
    ws.append([tieu_de]); ws['A1'].font = Font(bold=True)
    ws.append(['Vị trí', 'Ca', 'Mặt hàng', 'Ghi chú'])
    for c in ws[2]: c.font = Font(bold=True)
    for x in ghi_chu:
        ws.append([f"{x['sheet']}!{x['o']}" if x['o'] else x['sheet'], x['ca'], x['ten'], x['nd']])
        if x['o']:
            o = ws.cell(ws.max_row, 1)
            o.hyperlink = Hyperlink(ref=o.coordinate, location=f"'{x['sheet']}'!{x['o']}")
            o.font = Font(color='0563C1', underline='single')
        for c in ws[ws.max_row]: c.alignment = Alignment(wrap_text=True, vertical='top')
    for c, w in zip('ABCD', (14, 10, 40, 70)): ws.column_dimensions[c].width = w
    ws.freeze_panes = 'A3'


# ---------------- Chạy ----------------
def tao(kho_path, tuoi_path, td_path, han_path, ds_ngay, nam, nhap_cua, ra_kho, ra_tuoi):
    """Điền các ngày theo thứ tự; ngày sau học từ ngày vừa điền. Trả kết quả từng ngày và ghi chú."""
    kq = dict(ngay=[], ghi_chu_kho=[], ghi_chu_tuoi=[])
    wk, wt = T.mo_wb(kho_path), T.mo_wb(tuoi_path)
    kd = KiemDich(wt, nam)
    for ngay_txt in sorted(ds_ngay, key=lambda n: ngay_sheet(n, nam)):
        ngay = ngay_sheet(ngay_txt, nam)
        ten = f'{ngay.day:02d}.{ngay.month:02d}'
        viet, bo = tong_viet(td_path, ngay_txt)
        han, sh_han = tong_han(han_path, ngay_txt)
        dong = dict(ngay=ten, kho=None, tuoi=None, han=sh_han)
        for loai, wb, cot, mau in (('kho', wk, 7, MAU_KHO), ('tuoi', wt, 4, MAU_TUOI)):
            gc = kq['ghi_chu_' + loai]
            cu = next((s for s in wb.worksheets if ngay_sheet(s.title, nam) == ngay), None)
            if cu is not None and co_so(cu, cot, mau):
                dong[loai] = 'Đã có số, giữ nguyên'; continue
            truoc_ds = [(d, s, mau) for d, s in sheet_ngay(wb, nam) if d < ngay and co_so(s, cot, mau)][-7:]
            if not truoc_ds:
                dong[loai] = 'Chưa có ngày nào trước đó đã điền để làm mẫu'; continue
            if cu is not None: del wb[cu.title]
            ls = LichSu(truoc_ds, cot)
            truoc = truoc_ds[-1][1]
            if loai == 'kho': dien_kho(wb, ngay, nam, viet, bo, han, ls, truoc, gc)
            else: dien_tuoi(wb, ngay, nam, viet, han, ls, truoc, nhap_cua(ngay_txt), gc, kd)
            if loai == 'tuoi' and han_path and not sh_han:
                gc.append(dict(sheet=ten, o='', ca='HQ', ten='', nd='Thực đơn Hàn không có ngày này nên phần HQ để trống'))
            dong[loai] = 'Đã điền'
        kq['ngay'].append(dong)
    ghi_chu_sheet(wk, kq['ghi_chu_kho'], 'Ô cần xem lại trong sổ hàng khô do app tự điền')
    ghi_chu_sheet(wt, kq['ghi_chu_tuoi'], 'Ô cần xem lại trong sổ hàng tươi do app tự điền')
    for wb in (wk, wt): wb.calculation.fullCalcOnLoad = True     # công thức VLOOKUP chưa có giá trị, Excel tính khi mở
    wk.save(ra_kho); wt.save(ra_tuoi)
    return kq
