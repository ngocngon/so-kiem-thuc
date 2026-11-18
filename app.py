#!/usr/bin/env python3
"""App web tạo sổ kiểm thực, chạy trên Mac hoặc Windows. Máy cùng mạng vào được bằng mã truy cập; file không gửi ra internet.

Cách chạy:
    python3 app.py [--khong-mo]
"""
import datetime, hashlib, hmac, json, os, re, secrets, shutil, signal, socket, subprocess, sys, tempfile, threading, time, urllib.parse, urllib.request, webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
import tao_so as T
import so_buoc1 as B1

GOC = os.path.dirname(os.path.abspath(__file__))
# Chế độ máy chủ đám mây (Render...): bật bằng SKT_DAM_MAY=1. Không bật thì app chạy như cũ trên máy này.
DAM_MAY = os.environ.get('SKT_DAM_MAY', '').strip() not in ('', '0')
WEB = os.environ.get('SKT_WEB', '').strip() not in ('', '0')   # bản chạy trong trình duyệt (GitHub Pages)
GHI = os.environ.get('SKT_DU_LIEU') or GOC          # nơi lưu dữ liệu; trên máy chủ là ổ đĩa riêng
DU_LIEU = os.path.join(GHI, 'du_lieu')
KET_QUA = os.path.join(GHI, 'ket_qua')
LICH_SU = os.path.join(GHI, 'lich_su')
P_QT = os.path.join(GHI, 'quy_tac.json')
P_ST = os.path.join(GHI, 'so_tay_mon.json')
P_TEN = os.path.join(DU_LIEU, 'ten_file.json')
P_NHAP = os.path.join(DU_LIEU, 'nhap_tay.json')
P_MA = os.path.join(DU_LIEU, 'ma_truy_cap.txt')
FILE = {'thuc_don': os.path.join(DU_LIEU, 'thuc_don.xlsx'), 'mau': os.path.join(DU_LIEU, 'mau.xlsx'),
        'so_kho': os.path.join(DU_LIEU, 'so_kho.xlsx'), 'so_tuoi': os.path.join(DU_LIEU, 'so_tuoi.xlsx'), 'thuc_don_han': os.path.join(DU_LIEU, 'thuc_don_han.xlsx')}
CONG = int(os.environ.get('PORT') or os.environ.get('CONG_APP') or 8765)
PHIEN = set()      # phiên đăng nhập của máy khác, mất khi app tắt
# Trên máy chủ: mã nhân viên và mã quản trị đặt trong biến môi trường; phiên ký bằng SKT_BI_MAT nên
# máy chủ khởi động lại vẫn không phải đăng nhập lại
MA_DUNG = os.environ.get('SKT_MA_TRUY_CAP', '').strip().upper()
MA_QT = os.environ.get('SKT_MA_QUAN_TRI', '').strip().upper()
BI_MAT = (os.environ.get('SKT_BI_MAT') or secrets.token_hex(32)).encode()
HAN_PHIEN = 30 * 86400
CAN_QT = 'Việc này cần đăng nhập bằng mã quản trị' if DAM_MAY else 'Việc này chỉ làm được trên máy chạy app'
SAI = {}           # ip -> (số lần nhập sai, khoá đến lúc)
# Máy khác chỉ làm sổ; sửa cấu hình và điều khiển app chỉ từ chính máy chạy app
CHI_MAY_NAY = {'/api/quy_tac', '/api/thu_quy_tac', '/api/so_tay', '/api/khoi_phuc', '/api/mo_thu_muc', '/api/doi_ma', '/api/tat'}
PHIEN_BAN = hashlib.sha1(b''.join(open(os.path.join(GOC, f), 'rb').read() for f in ('app.py', 'tao_so.py', 'so_buoc1.py'))).hexdigest()


def doc_json(p, mac_dinh):
    return json.load(open(p, encoding='utf-8')) if os.path.exists(p) else mac_dinh

def ghi_json(p, d):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    tam = p + '.tam'
    with open(tam, 'w', encoding='utf-8') as f: json.dump(d, f, ensure_ascii=False, indent=1)
    os.replace(tam, p)

def sao_luu(p, loai):
    """Giữ bản hiện tại trước khi ghi đè để khôi phục được."""
    if not os.path.exists(p): return
    os.makedirs(LICH_SU, exist_ok=True)
    gio = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    shutil.copy2(p, os.path.join(LICH_SU, f'{loai}_{gio}.json'))


def kiem_quy_tac(q, goc=None):
    """Trả lỗi dễ hiểu trước khi lưu; quy tắc sai không được ghi xuống."""
    d = {**(T.Q if goc is None else goc), **{k: v for k, v in q.items() if k in T.MAC_DINH}}
    for k in ('nam', 'cho_phep_lech', 'dl_dem', 'dl_com', 'dl_gao_toi_thieu', 'dl_thit_canh', 'dl_canh_bao', 'sai_so_kg', 'sai_so_pt'):
        if not isinstance(d[k], (int, float)) or isinstance(d[k], bool) or d[k] < 0:
            raise ValueError('Các ô năm, định lượng và mức lệch phải là số không âm')
    for x in d['mon_thieu_nl']:
        if not isinstance(x.get('moi_suat'), (int, float)) or x['moi_suat'] <= 0 or not str(x.get('mon', '')).strip() or not str(x.get('nl', '')).strip():
            raise ValueError('Món thiếu nguyên liệu phải có tên món, tên nguyên liệu và số lượng mỗi suất')
    for k in ('don_vi_can', 'don_vi_dem', 'mau_khoi_luong', 'mau_canh', 'mau_dung_cu'):
        if not isinstance(d[k], str) or not d[k].strip(): raise ValueError('Ô đơn vị hoặc khối lượng không được để trống')
    for k in ('khong_tach', 'bo_nguyen_lieu', 'bo_bo_mon', 'gao_khong_phai_com', 'tuan_bo_qua'):
        if not isinstance(d[k], list) or any(not isinstance(x, str) for x in d[k]): raise ValueError('Danh sách chỉ được chứa chữ')
        d[k] = [x.strip() for x in d[k] if x.strip()]      # dòng trống sẽ khớp mọi tên
    if d['kiem_dinh_luong'] not in ('co', 'khong'): raise ValueError('Lựa chọn kiểm định lượng không hợp lệ')
    if d['khi_lech'] not in ('cong_thuc', 'thuc_don'): raise ValueError('Cách xử lý khi lệch không hợp lệ')
    if d['tu_khac_phuc'] not in ('co', 'khong') or d['tu_dien_suat'] not in ('co', 'khong'): raise ValueError('Lựa chọn tự khắc phục không hợp lệ')
    for ct in [d['cong_thuc']] + [x.get('cong_thuc', '') for x in d['cong_thuc_rieng']]:
        T.tinh(ct, 170, 400)
    for x in d['cong_thuc_rieng']:
        if not str(x.get('bat_dau', '')).strip(): raise ValueError('Công thức riêng phải ghi tên nguyên liệu')
    for x in d['doi_ten']:
        x['thay'] = x.get('thay') or ''
        try: re.sub(x['tim'], x['thay'], 'THU')
        except (re.error, KeyError, IndexError): raise ValueError(f'Mẫu đổi tên không hợp lệ: {x.get("tim")}')
        if not x['tim']: raise ValueError('Dòng đổi tên phải có ô Tìm')
    for k, cot in (('goi_y_sua', ('thuc_don', 'so')), ('don_vi_theo_ten', ('bat_dau', 'don_vi')), ('luu_cong', ('bat_dau', 'khoi_luong', 'dung_cu')),
                   ('tach_ba_dong', ('bat_dau',)), ('trai_cay', ('ten', 'don_vi')),
                   ('hang_cuoi_ca', ('ten', 'khoi_luong', 'dung_cu'))):
        for x in d[k]:
            if any(not str(x.get(c) or '').strip() for c in cot): raise ValueError('Có dòng trong bảng chưa điền đủ ô')
    for x in d['hang_cuoi_ca']:
        x['suat'] = x.get('suat')
        if x['suat'] is not None and (not isinstance(x['suat'], (int, float)) or x['suat'] < 0): raise ValueError('Số suất phải là số không âm')
    # Chạy thử trên dữ liệu mẫu để quy tắc lỗi không lọt xuống file
    hien = T.Q
    try:
        T.dung_quy_tac(d)
        T.tach_mon('CƠM RANG, CANH CẢI NẤU THỊT'); T.don_vi('TRỨNG GÀ', 1000); T.mau_luu([dict(ten='BÚN BÒ', suat=1)])
    except (ValueError, TypeError, KeyError, AttributeError, re.error) as e:
        raise ValueError(f'Quy tắc chưa hợp lệ: {e}')
    finally:
        T.dung_quy_tac(hien)
    return d


def chua_doi(q):
    """Chỉ lưu phần khác mặc định để lần nâng cấp sau vẫn nhận được quy tắc mới."""
    return {k: v for k, v in q.items() if k in T.MAC_DINH and v != T.MAC_DINH[k]}


def can_file():
    if not os.path.exists(FILE['thuc_don']): raise ValueError('Chưa tải file thực đơn')
    return FILE['thuc_don']


def ban_nhap(ngay):
    """Số nhập tay của 1 ngày: {ca, suat_bo, td, da_biet, da_xem_ca}; td là mã file thực đơn lúc nhập."""
    v = doc_json(P_NHAP, {}).get(ngay, {})
    ban = dict(ca=v.get('ca', {}), suat_bo=v.get('suat_bo', {}), td=v.get('td')) if 'ca' in v or 'suat_bo' in v \
        else dict(ca=v, suat_bo={}, td=None)
    return dict(ban, da_biet=v.get('da_biet', {}), da_xem_ca=v.get('da_xem_ca', []))


def ghi_ban_nhap(ngay, ban):
    d = doc_json(P_NHAP, {})
    d[ngay] = {**ban, 'td': doc_json(P_TEN, {}).get('ma_thuc_don')}
    ghi_json(P_NHAP, d)


def lap(ngay):
    ban = ban_nhap(ngay)
    sach, nhat_ky, ten = T.lap_so(can_file(), ngay, ban['suat_bo'])
    for ca in ban['da_xem_ca']:          # đã xem cả ca: chỉ có hiệu lực cho ngày này, không ghi sổ tay
        for m in sach[ca]:
            for y in m['nl']: y['chac'] = True
    nhat_ky = [x for x in nhat_ky if not (x[0] in ban['da_xem_ca'] and x[4].startswith('Máy tự xếp'))]
    return sach, nhat_ky, ten


def nhap_cua(ngay):
    return ban_nhap(ngay)['ca']


def xem_nhap(ngay):
    hang = [dict(ten=x['ten'].strip(), dv=x['don_vi'], loai='trai_cay') for x in T.Q['trai_cay']] + \
           [dict(ten=x['ten'].strip(), mac_dinh=x['suat'], loai='cuoi_ca') for x in T.Q['hang_cuoi_ca']]
    ban, canh_bao = ban_nhap(ngay), []
    if any(ban['ca'].values()) and ban['td'] and ban['td'] != doc_json(P_TEN, {}).get('ma_thuc_don'):
        canh_bao.append('Số ở đây được nhập khi dùng file thực đơn khác. Hãy kiểm tra lại.')
    mo_coi = {t for ds in ban['ca'].values() for t in ds} - {x['ten'] for x in hang}
    if mo_coi: canh_bao.append('Có số đã nhập cho hàng không còn trong danh sách: ' + ', '.join(sorted(mo_coi)))
    return dict(ngay=ngay, hang=hang, ca=ban['ca'], goi_y=T.goi_y_nhap(can_file(), ngay), canh_bao=canh_bao)


def luu_nhap(ngay, ca):
    """Số nhập tay theo ngày; ô trống thì bỏ để sổ để trống như mẫu."""
    if not re.fullmatch(r'\d{1,2}\.\d{1,2}', str(ngay)): raise ValueError('Ngày không hợp lệ')
    sach = {}
    for c, ds in ca.items():
        if c not in T.CA: raise ValueError('Ca không hợp lệ')
        for ten, so in ds.items():
            for k in ('sl', 'suat'):
                v = so.get(k)
                if v is None: continue
                if not isinstance(v, (int, float)) or isinstance(v, bool) or v < 0:
                    raise ValueError(f'{ten}: số nhập không hợp lệ')
                sach.setdefault(c, {}).setdefault(ten, {})[k] = v
    ghi_ban_nhap(ngay, {**ban_nhap(ngay), 'ca': sach})
    return dict(ok=True)


def luu_suat_bo(ngay, ca, bo, suat):
    """Số suất cho bộ món thực đơn để trống; 0 là không phục vụ."""
    if ca not in T.CA or not isinstance(suat, (int, float)) or isinstance(suat, bool) or suat < 0:
        raise ValueError('Số suất không hợp lệ')
    ban = ban_nhap(ngay)
    ban['suat_bo'].setdefault(ca, {})[bo] = suat
    ghi_ban_nhap(ngay, ban)
    return xem(ngay)


def gon(sach, nhap=None):
    """Chuyển kết quả sang dạng hiển thị."""
    kq, nhap = {}, nhap or {}
    for ca in T.CA:
        mon = [dict(ten=m['ten'], suat=m['suat'], nl=[dict(ten=y['ten'], sl=y['sl'], dv=y['dv'], suat=y['suat'],
                                                           chac=y.get('chac', True)) for y in m['nl']]) for m in sach[ca]]
        kq[ca] = dict(mon=mon,
                      mau=[dict(ten=t, suat=s, kl=kl, dc=dc) for t, s, kl, dc in T.mau_luu(sach[ca], nhap.get(ca))])
    return kq


def loai_soat(nd, bo=''):
    if nd.startswith('Đã tự sửa'): return 'tu_sua'
    if T.la_sai_so(nd): return 'sai_so'
    if 'sheet ngày không có' in nd:
        return 'tin' if any(T.khong_dau(t) in T.khong_dau(bo) for t in T.Q['tuan_bo_qua']) else 'tuan'
    return 'viec' if nd.startswith(('Bộ món', 'Máy tự xếp', 'Món có trong tên')) else 'tin'


def xem(ngay):
    sach, nhat_ky, ten = lap(ngay)
    ban = ban_nhap(ngay)
    return dict(ngay=ngay, sheet=ten, ca=gon(sach, ban['ca']),
                thieu_suat=T.thieu_suat(can_file(), ngay, ban['suat_bo']), ghi_chu=T.ghi_chu_sai_so(sach, nhat_ky),
                can_soat=[dict(ca=a, dong=b, bo=c, nl=d, nd=e, loai=loai_soat(e, c), da_biet=c in ban['da_biet'].get(a, []))
                          for a, b, c, d, e in nhat_ky])


def tom_tat(ds_ngay):
    """Số việc của từng ngày: con_viec cần người quyết định, nen_xem không làm đổi số kg."""
    kq = {}
    for ngay in ds_ngay:
        sach, nhat_ky, _ = lap(ngay)
        ban = ban_nhap(ngay)
        nen_xem = sum(1 for c in T.CA for m in sach[c] for y in m['nl'] if not y.get('chac', True)) \
            + sum(1 for c in T.CA for m in sach[c] if not m['nl'])
        nen_xem += sum(1 for x in T.thieu_suat(can_file(), ngay, ban['suat_bo']) if not x['xong'])
        tuan = sum(1 for a, _, c, _, e in nhat_ky if loai_soat(e, c) == 'tuan' and c not in ban['da_biet'].get(a, []))
        kq[ngay] = dict(con_viec=tuan, nen_xem=nen_xem, sai_so=sum(1 for x in nhat_ky if T.muc_sai_so(x[4]) == 'kiem'), ca_da_nhap=sum(1 for c in T.CA if ban['ca'].get(c)))
    return kq


def phang(sach):
    d = {}
    for ca in T.CA:
        for m in sach[ca]:
            for y in m['nl']:
                d[('Bước 2', ca, m['ten'], y['ten'])] = f"{y['sl']} {y['dv']}, {y['suat']} suất"
        dem = {}
        for t, s, kl, dc in T.mau_luu(sach[ca]):
            dem[t] = dem.get(t, 0) + 1
            d[('Mẫu 24 giờ', ca, f'{t} #{dem[t]}' if dem[t] > 1 else t, '')] = f"{s if s is not None else '-'} suất, {kl}, {dc}"
    return d


def thu_quy_tac(moi, ngay):
    """So kết quả của quy tắc đang lưu với quy tắc đang sửa, chưa ghi gì xuống."""
    moi = kiem_quy_tac(moi)
    cu, _, _ = lap(ngay)
    a = phang(cu)
    hien = T.Q
    try:
        T.dung_quy_tac(moi)
        moi_sach, _, _ = lap(ngay)
        b = phang(moi_sach)
    finally:
        T.dung_quy_tac(hien)
    doi = [dict(so=k[0], ca=k[1], mon=k[2], nl=k[3], cu=a.get(k, ''), moi=b.get(k, ''))
           for k in list(a) + [k for k in b if k not in a] if a.get(k) != b.get(k)]
    return dict(tong=len(doi), doi=doi[:200])


def tao(ds_ngay):
    if not os.path.exists(FILE['mau']): raise ValueError('Chưa tải file mẫu')
    os.makedirs(KET_QUA, exist_ok=True)
    kq = []
    for ngay in ds_ngay:
        if not re.fullmatch(r'\d{1,2}\.\d{1,2}', str(ngay)): raise ValueError('Ngày không hợp lệ')
        try:
            sach, nhat_ky, _ = lap(ngay)
            ten = f'So_{ngay}.xlsx'
            ra = os.path.join(KET_QUA, ten)
            tam = os.path.join(KET_QUA, f'.dang_tao_{ngay}.xlsx')
            nhap = nhap_cua(ngay)
            ds_ngay = T.cac_ngay(can_file())
            T.ghi_ra(FILE['mau'], sach, ngay, nhat_ky, tam, nhap, ds_ngay, 'chot')
            ten_kt = f'So_{ngay}_kiem_tra.xlsx'
            T.ghi_ra(FILE['mau'], sach, ngay, nhat_ky, os.path.join(KET_QUA, f'.dang_tao_kt_{ngay}.xlsx'), nhap, ds_ngay, 'kiem')
            # Bản cũ có thể đã điền tay số trái cây, sữa nên không ghi đè
            if os.path.exists(ra):
                os.replace(ra, os.path.join(KET_QUA, f"So_{ngay}_cu_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx"))
            os.replace(tam, ra)
            os.replace(os.path.join(KET_QUA, f'.dang_tao_kt_{ngay}.xlsx'), os.path.join(KET_QUA, ten_kt))
            kq.append(dict(ngay=ngay, file=ten, file_kt=ten_kt, can_soat=len(nhat_ky), ca_da_nhap=sum(1 for c in T.CA if nhap.get(c))))
        except PermissionError:       # Windows không cho đổi tên file đang mở trong Excel
            for p in (tam, os.path.join(KET_QUA, f'.dang_tao_kt_{ngay}.xlsx')):
                if os.path.exists(p): os.remove(p)
            kq.append(dict(ngay=ngay, loi=f'Hãy đóng file {ten} trong Excel rồi tạo lại'))
        except (Exception, SystemExit) as e:
            kq.append(dict(ngay=ngay, loi=str(e)))
    return kq


def tao_buoc1(ds_ngay):
    """Sổ bước 1 (hàng khô, hàng tươi) ra file mới; file anh chị tải lên không bị sửa."""
    for k, t in (('so_kho', 'sổ hàng khô'), ('so_tuoi', 'sổ hàng tươi')):
        if not os.path.exists(FILE[k]): raise ValueError(f'Chưa tải {t}')
    td = can_file()
    ds_ngay = [n for n in ds_ngay if re.fullmatch(r'\d{1,2}\.\d{1,2}', str(n))]
    if not ds_ngay: raise ValueError('Hãy chọn ít nhất một ngày')
    thang = ds_ngay[0].split('.')[1].zfill(2)
    os.makedirs(KET_QUA, exist_ok=True)
    ten = {'kho': f'So_buoc1_hang_kho_{thang}.xlsx', 'tuoi': f'So_buoc1_hang_tuoi_{thang}.xlsx'}
    tam = {k: os.path.join(KET_QUA, f'.dang_tao_{v}') for k, v in ten.items()}
    try:
        kq = B1.tao(FILE['so_kho'], FILE['so_tuoi'], td, FILE['thuc_don_han'] if os.path.exists(FILE['thuc_don_han']) else None,
                    ds_ngay, T.Q['nam'], nhap_cua, tam['kho'], tam['tuoi'])
        luc = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        for k, v in ten.items():
            ra = os.path.join(KET_QUA, v)
            if os.path.exists(ra): os.replace(ra, ra.replace('.xlsx', f'_cu_{luc}.xlsx'))
            os.replace(tam[k], ra)
    except PermissionError:
        raise ValueError('Hãy đóng file sổ bước 1 đang mở trong Excel rồi tạo lại')
    finally:
        for p in tam.values():
            if os.path.exists(p): os.remove(p)
    return dict(kq, file_kho=ten['kho'], file_tuoi=ten['tuoi'], han=os.path.exists(FILE['thuc_don_han']))


def doi_chieu_so(ten, du_lieu):
    """So một sổ hoàn thiện với thực đơn đang chọn, cùng ngày. Thực đơn là căn cứ đúng."""
    with tempfile.TemporaryDirectory() as tam:
        sach, ra = os.path.join(tam, 'so.xlsx'), os.path.join(tam, 'app.xlsx')
        with open(sach, 'wb') as f: f.write(du_lieu)
        try:
            h = T.doc_so(sach)
        except Exception:
            raise ValueError('Không đọc được file này, hãy chọn sổ .xlsx có các sheet bước 2, bước 3')
        ngay = h['ngay']
        if ngay not in T.cac_ngay(can_file()): raise ValueError(f'Thực đơn đang chọn không có ngày {ngay}. Hãy chọn đúng thực đơn tuần ở bước 1.')
        s, nk, _ = lap(ngay)
        T.ghi_ra(FILE['mau'], s, ngay, nk, ra, nhap_cua(ngay), T.cac_ngay(can_file()), 'chot')
        return dict(T.so_sanh(h, T.doc_so(ra)), ngay=ngay, ten=ten, kg=T.Q['sai_so_kg'], pt=T.Q['sai_so_pt'])


def tai_len(loai, ten, du_lieu):
    if loai not in FILE: raise ValueError('Loại file không hợp lệ')
    os.makedirs(DU_LIEU, exist_ok=True)
    fd, tam = tempfile.mkstemp(suffix='.xlsx', dir=DU_LIEU)
    with os.fdopen(fd, 'wb') as f: f.write(du_lieu)
    try:
        try:
            wb = T.mo_wb(tam)
        except Exception:
            raise ValueError('Không đọc được file này, hãy chọn file .xlsx')
        if loai == 'mau':
            thieu = [s for s in list(T.SHEET2.values()) + [T.SHEET3, T.SHEET4] if s not in wb.sheetnames]
            if thieu: raise ValueError('Đây không phải file mẫu: thiếu sheet ' + ', '.join(s.strip() for s in thieu))
        elif loai in ('so_kho', 'so_tuoi'):
            cot, mau = (7, B1.MAU_KHO) if loai == 'so_kho' else (4, B1.MAU_TUOI)
            if not any(B1.co_so(s, cot, mau) for _, s in B1.sheet_ngay(wb, T.Q['nam'])):
                raise ValueError('Đây không phải sổ ' + ('hàng khô' if loai == 'so_kho' else 'hàng tươi') +
                                 ': cần ít nhất một sheet tên ngày (như 02.11) đã điền số theo ca')
        elif loai == 'thuc_don_han':
            if not any(re.search(r'NGAY\s*\d{1,2}\.\d{1,2}', B1.K(s.cell(r, c).value)) for s in wb.worksheets for r in range(1, 4) for c in range(1, 8)):
                raise ValueError('Không thấy sheet nào ghi "ngày 02.11" ở các dòng đầu, đây không phải thực đơn Hàn')
        elif not T.cac_ngay(tam):
            raise ValueError('Không thấy sheet ngày nào có dòng đầu kiểu "CA SÁNG 26/10" trong file này')
        os.replace(tam, FILE[loai])
    finally:
        if os.path.exists(tam): os.remove(tam)
    ten_file = doc_json(P_TEN, {})
    ten_file[loai] = ten
    if loai == 'thuc_don': ten_file['ma_thuc_don'] = hashlib.sha1(du_lieu).hexdigest()
    ghi_json(P_TEN, ten_file)


def trang_thai():
    co = os.path.exists(FILE['thuc_don'])
    ds = []
    if os.path.isdir(KET_QUA):
        ds = sorted((f for f in os.listdir(KET_QUA) if f.endswith('.xlsx') and not f.startswith('.')),
                    key=lambda f: os.path.getmtime(os.path.join(KET_QUA, f)), reverse=True)[:30]
    ten = doc_json(P_TEN, {})
    return dict(thuc_don=ten.get('thuc_don') if co else None, mau=ten.get('mau') if os.path.exists(FILE['mau']) else None,
                ngay=T.cac_ngay(FILE['thuc_don']) if co else [], ket_qua=ds,
                **{k: ten.get(k) if os.path.exists(FILE[k]) else None for k in ('so_kho', 'so_tuoi', 'thuc_don_han')})


def lich_su():
    if not os.path.isdir(LICH_SU): return []
    kq = []
    for f in sorted(os.listdir(LICH_SU), reverse=True)[:50]:
        m = re.fullmatch(r'(quy_tac|so_tay)_(\d{8})_(\d{6})_\d+\.json', f)
        if m:
            d, g = m.group(2), m.group(3)
            kq.append(dict(file=f, loai=m.group(1), luc=f'{d[6:]}/{d[4:6]}/{d[:4]} {g[:2]}:{g[2:4]}:{g[4:]}'))
    return kq


def khoi_phuc(f):
    if not re.fullmatch(r'(quy_tac|so_tay)_\d{8}_\d{6}_\d+\.json', f or ''): raise ValueError('Bản lưu không hợp lệ')
    p = os.path.join(LICH_SU, f)
    if not os.path.exists(p): raise ValueError('Không tìm thấy bản lưu')
    d = json.load(open(p, encoding='utf-8'))
    if f.startswith('quy_tac'):
        luu_quy_tac(d, T.MAC_DINH)               # bản lưu chỉ chứa phần khác mặc định
    else:
        luu_so_tay(d)


def luu_quy_tac(q, goc=None):
    d = kiem_quy_tac(q, goc)
    sao_luu(P_QT, 'quy_tac')
    ghi_json(P_QT, chua_doi(d))
    T.dung_quy_tac(d)


def luu_so_tay(st):
    sach = {}
    for k, ds in st.items():
        ten = T.khoa_mon(k)
        ds = [T.khoa_nl(x) for x in ds if str(x).strip()]
        if ten and ds: sach[ten] = list(dict.fromkeys(sach.get(ten, []) + ds))
    sao_luu(P_ST, 'so_tay')
    ghi_json(P_ST, sach)


def ghi_nho(ngay, ca, nl, mon):
    """Dạy sổ tay: nguyên liệu này thuộc món này; xem lại ngay để biết đã có tác dụng chưa."""
    st = T.tai_so_tay()
    k = T.khoa_mon(mon)
    x = T.khoa_nl(nl)
    if x not in st.get(k, []): st.setdefault(k, []).append(x)
    luu_so_tay(st)
    kq = xem(ngay)
    kq['da_chuyen'] = any(T.khoa_mon(m['ten']) == k and any(T.khoa_nl(y['ten']) == x for y in m['nl'])
                          for m in kq['ca'][ca]['mon'])
    return kq


def ghi_nho_ca(ngay, ca):
    """Đã xem cả ca: áp dụng cho ngày này, không ghi sổ tay vì sổ tay có hiệu lực mọi tuần."""
    if ca not in T.CA: raise ValueError('Ca không hợp lệ')
    ban = ban_nhap(ngay)
    if ca not in ban['da_xem_ca']: ban['da_xem_ca'].append(ca)
    ghi_ban_nhap(ngay, ban)
    return xem(ngay)


def luu_da_biet(ngay, ca, bo):
    """Đánh dấu đã biết bộ món chỉ có ở Thực đơn tuần."""
    if ca not in T.CA: raise ValueError('Ca không hợp lệ')
    ban = ban_nhap(ngay)
    ds = ban['da_biet'].setdefault(ca, [])
    if bo not in ds: ds.append(bo)
    ghi_ban_nhap(ngay, ban)
    return xem(ngay)


def ma_truy_cap():
    if DAM_MAY: return MA_DUNG
    if not os.path.exists(P_MA):
        os.makedirs(DU_LIEU, exist_ok=True)
        ma = ''.join(secrets.choice('ABCDEFGHJKMNPQRSTUVWXYZ23456789') for _ in range(8))
        with os.fdopen(os.open(P_MA, os.O_WRONLY | os.O_CREAT, 0o600), 'w') as f: f.write(ma)
    return open(P_MA).read().strip()

def doi_ma():
    """Đổi mã thì các máy đang đăng nhập phải nhập lại mã mới."""
    if DAM_MAY: raise ValueError('Trên máy chủ, mã đặt ở mục Environment của Render (SKT_MA_TRUY_CAP, SKT_MA_QUAN_TRI)')
    if os.path.exists(P_MA): os.remove(P_MA)
    PHIEN.clear()
    return dict(ma=ma_truy_cap())

def mo_thu_muc():
    if DAM_MAY: raise ValueError('App đang chạy trên máy chủ: hãy bấm tên file để tải về')
    os.makedirs(KET_QUA, exist_ok=True)
    if os.name == 'nt': os.startfile(KET_QUA)
    else: subprocess.run(['open', KET_QUA])
    return dict(ok=True)

def dia_chi():
    ds = []
    if sys.platform == 'darwin':
        ten = subprocess.run(['scutil', '--get', 'LocalHostName'], capture_output=True, text=True).stdout.strip()
        if ten: ds.append(f'http://{ten}.local:{CONG}')
    else:
        ds.append(f'http://{socket.gethostname()}:{CONG}')
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('192.0.2.1', 9))          # chỉ để hỏi địa chỉ mạng, không gửi gì đi
        ds.append(f'http://{s.getsockname()[0]}:{CONG}')
        s.close()
    except OSError:
        pass
    return ds


def ky_phien(vai):
    het = int(time.time()) + HAN_PHIEN
    nd = f'{vai}.{het}'
    return nd + '.' + hmac.new(BI_MAT, nd.encode(), 'sha256').hexdigest()

def doc_phien(tk):
    """Vai trò ('dung' hoặc 'quan_tri') của phiên còn hạn, không hợp lệ thì None."""
    try:
        vai, het, ky = tk.split('.')
        dung = hmac.new(BI_MAT, f'{vai}.{het}'.encode(), 'sha256').hexdigest()
        if hmac.compare_digest(ky, dung) and int(het) > time.time() and vai in ('dung', 'quan_tri'): return vai
    except ValueError:
        pass
    return None


class XuLy(BaseHTTPRequestHandler):
    timeout = 30       # app xử lý từng yêu cầu một, máy treo không được giữ cả app

    def log_message(self, *a): pass

    def tra(self, d, ma=200, them=()):
        b = json.dumps(d, ensure_ascii=False).encode('utf-8')
        self.send_response(ma)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(b)))
        for k, v in them: self.send_header(k, v)
        self.end_headers()
        self.wfile.write(b)

    def gui_trang(self, ten):
        b = open(os.path.join(GOC, ten), 'rb').read()
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(b)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(b)

    def la_may_nay(self):
        # Trên máy chủ mọi yêu cầu đi qua proxy nên không coi ai là "máy chạy app"
        return not DAM_MAY and self.client_address[0] in ('127.0.0.1', '::1')

    def vai_tro(self):
        m = re.search(r'(?:^|;\s*)phien=([\w.-]+)', self.headers.get('Cookie', ''))
        if not m: return None
        if DAM_MAY: return doc_phien(m.group(1))
        return 'dung' if m.group(1) in PHIEN else None

    def co_phien(self):
        return self.vai_tro() is not None

    def la_quan_tri(self):
        """Vào được Quản trị: máy chạy app, hoặc trên máy chủ là người đăng nhập bằng mã quản trị."""
        return self.la_may_nay() or self.vai_tro() == 'quan_tri'

    def ip(self):
        if DAM_MAY:              # proxy của máy chủ thêm IP thật vào cuối X-Forwarded-For
            x = [v.strip() for v in self.headers.get('X-Forwarded-For', '').split(',') if v.strip()]
            if x: return x[-1]
        return self.client_address[0]

    def cung_goc(self, goc, may):
        return not goc or goc in (f'http://{may}', f'https://{may}' if DAM_MAY else None)

    def dung_nguon(self):
        # Chặn trang web khác gọi vào app; máy khác phải đăng nhập
        goc = self.headers.get('Origin')
        may = self.headers.get('Host', '')
        if self.la_may_nay():
            return may in (f'127.0.0.1:{CONG}', f'localhost:{CONG}') and (not goc or goc == f'http://{may}')
        return self.co_phien() and self.cung_goc(goc, may)

    def tu_choi(self):
        self.tra(dict(loi='Không được phép'), 403 if self.la_may_nay() or self.co_phien() else 401)

    def dang_nhap(self):
        than = self.rfile.read(min(int(self.headers.get('Content-Length', 0) or 0), 1024))
        goc, may = self.headers.get('Origin'), self.headers.get('Host', '')
        if not self.cung_goc(goc, may): return self.tra(dict(loi='Không được phép'), 403)
        ip = self.ip()
        so, khoa = SAI.get(ip, (0, 0))
        if time.time() < khoa: return self.tra(dict(loi='Sai quá nhiều lần. Hãy thử lại sau 1 phút.'), 429)
        try:
            ma = str(json.loads(than).get('ma', ''))
        except (ValueError, AttributeError):
            ma = ''
        ma = ma.strip().upper().encode()
        vai = None
        if DAM_MAY and MA_QT and hmac.compare_digest(ma, MA_QT.encode()): vai = 'quan_tri'
        elif ma_truy_cap() and hmac.compare_digest(ma, ma_truy_cap().encode()): vai = 'dung'
        if vai is None:
            SAI[ip] = (0, time.time() + 60) if so >= 4 else (so + 1, 0)
            return self.tra(dict(loi='Mã truy cập không đúng'), 401)
        SAI.pop(ip, None)
        if DAM_MAY:
            phien = ky_phien(vai)
        else:
            phien = secrets.token_urlsafe(32)
            PHIEN.add(phien)
        an_toan = '; Secure' if self.headers.get('X-Forwarded-Proto') == 'https' else ''
        self.tra(dict(ok=True), them=[('Set-Cookie', f'phien={phien}; Path=/; HttpOnly; SameSite=Strict; Max-Age={HAN_PHIEN}{an_toan}')])

    def chay(self, ham):
        if not self.dung_nguon(): return self.tu_choi()
        try:
            self.tra(ham())
        except (Exception, SystemExit) as e:
            self.tra(dict(loi=str(e) or type(e).__name__), 400)

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        h = dict(urllib.parse.parse_qsl(u.query))
        if u.path == '/':
            return self.gui_trang('giao_dien.html' if self.la_may_nay() or self.co_phien() else 'dang_nhap.html')
        if u.path == '/healthz': return self.tra(dict(ok=True))
        if u.path == '/api/ping' and self.la_may_nay():
            return self.tra(dict(app='so_kiem_thuc', phien_ban=PHIEN_BAN, pid=os.getpid()))
        if u.path == '/api/truy_cap':
            if not self.la_quan_tri(): return self.tra(dict(loi=CAN_QT), 403)
            if DAM_MAY: return self.chay(lambda: dict(dia_chi=['https://' + self.headers.get('Host', '')], ma=MA_DUNG))
            return self.chay(lambda: dict(dia_chi=dia_chi(), ma=ma_truy_cap()))
        if u.path == '/api/trang_thai': return self.chay(lambda: dict(trang_thai(), may_chu=self.la_quan_tri(), dam_may=DAM_MAY, web=WEB))
        if u.path == '/api/xem': return self.chay(lambda: xem(h.get('ngay', '')))
        if u.path == '/api/quy_tac': return self.chay(lambda: dict(q=T.Q, mac_dinh=T.MAC_DINH))
        if u.path == '/api/nhap': return self.chay(lambda: xem_nhap(h.get('ngay', '')))
        if u.path == '/api/tom_tat': return self.chay(lambda: tom_tat([x for x in h.get('ngay', '').split(',') if x]))
        if u.path == '/api/so_tay': return self.chay(T.tai_so_tay)
        if u.path == '/api/lich_su': return self.chay(lich_su)
        if u.path == '/api/tai_ve':
            ten = h.get('ten', '')
            p = os.path.join(KET_QUA, ten)
            if not self.dung_nguon() or not re.fullmatch(r'So_[\w.]+\.xlsx', ten) or not os.path.exists(p):
                return self.tra(dict(loi='Không tìm thấy file'), 404)
            b = open(p, 'rb').read()
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{ten}"')
            self.send_header('Content-Length', str(len(b)))
            self.end_headers()
            return self.wfile.write(b)
        self.tra(dict(loi='Không có trang này'), 404)

    def do_POST(self):
        u = urllib.parse.urlparse(self.path)
        h = dict(urllib.parse.parse_qsl(u.query))
        if u.path == '/api/dang_nhap': return self.dang_nhap()
        if not self.dung_nguon(): return self.tu_choi()
        if u.path in CHI_MAY_NAY and not self.la_quan_tri():
            return self.tra(dict(loi=CAN_QT), 403)
        if WEB and u.path in ('/api/tat', '/api/mo_thu_muc', '/api/doi_ma'):
            return self.tra(dict(loi='Bản web không có việc này: file tải về bằng cách bấm tên file'), 400)
        if DAM_MAY and u.path == '/api/tat':
            return self.tra(dict(loi='App trên máy chủ không tắt từ đây; tạm dừng ở trang quản lý của Render'), 400)
        than = self.rfile.read(int(self.headers.get('Content-Length', 0) or 0))
        if u.path == '/api/doi_chieu_so': return self.chay(lambda: doi_chieu_so(h.get('ten', ''), than))
        if u.path == '/api/tai_len':
            return self.chay(lambda: (tai_len(h.get('loai'), h.get('ten', ''), than), trang_thai())[1])
        try:
            d = json.loads(than or b'{}')
        except ValueError:
            return self.tra(dict(loi='Dữ liệu không hợp lệ'), 400)
        ham = {'/api/tao': lambda: dict(kq=tao(d.get('ngay', [])), **trang_thai()),
               '/api/buoc1': lambda: dict(kq=tao_buoc1(d.get('ngay', [])), **trang_thai()),
               '/api/quy_tac': lambda: (luu_quy_tac(d), dict(q=T.Q))[1],
               '/api/nhap': lambda: luu_nhap(d.get('ngay'), d.get('ca', {})),
               '/api/suat_bo': lambda: luu_suat_bo(d.get('ngay'), d.get('ca'), d.get('bo'), d.get('suat')),
               '/api/ghi_nho_ca': lambda: ghi_nho_ca(d.get('ngay'), d.get('ca')),
               '/api/da_biet': lambda: luu_da_biet(d.get('ngay'), d.get('ca'), d.get('bo')),
               '/api/thu_quy_tac': lambda: thu_quy_tac(d.get('q', {}), d.get('ngay', '')),
               '/api/tinh': lambda: dict(kq=round(T.tinh(d.get('cong_thuc'), d.get('dinh_luong', 0), d.get('suat', 0)), 4)),
               '/api/so_tay': lambda: (luu_so_tay(d), T.tai_so_tay())[1],
               '/api/ghi_nho': lambda: ghi_nho(d['ngay'], d['ca'], d['nl'], d['mon']),
               '/api/khoi_phuc': lambda: (khoi_phuc(d.get('file')), dict(ok=True))[1],
               '/api/mo_thu_muc': mo_thu_muc,
               '/api/doi_ma': doi_ma,
               '/api/tat': lambda: (threading.Timer(0.3, os._exit, [0]).start(), dict(ok=True))[1],
               }.get(u.path)
        if ham: return self.chay(ham)
        self.tra(dict(loi='Không có trang này'), 404)


def mo_cong(url):
    """Mở cổng cho app. Cùng bản đang chạy thì trả None; bản cũ hơn thì thay bằng bản này."""
    # Hỏi trước khi mở cổng: macOS cho 0.0.0.0 mở chồng lên bản cũ đang giữ 127.0.0.1
    try:
        d = json.load(urllib.request.urlopen(url + 'api/ping', timeout=2))
    except (OSError, ValueError):
        d = {}
    if d.get('app') == 'so_kiem_thuc':
        if d.get('phien_ban') == PHIEN_BAN: return None
        print('Đang thay bản app cũ...')
        os.kill(d['pid'], signal.SIGTERM)
        time.sleep(1)
    try:
        return HTTPServer(('0.0.0.0', CONG), XuLy)
    except OSError:
        raise SystemExit(f'Cổng {CONG} đang bị chương trình khác dùng')


def chay_may_chu():
    if not MA_DUNG or not MA_QT:
        raise SystemExit('Cần đặt SKT_MA_TRUY_CAP và SKT_MA_QUAN_TRI trước khi chạy trên máy chủ')
    if MA_DUNG == MA_QT: raise SystemExit('Mã quản trị phải khác mã nhân viên')
    if len(MA_DUNG) < 8 or len(MA_QT) < 8: raise SystemExit('Mã cần ít nhất 8 ký tự')
    if not os.environ.get('SKT_BI_MAT'): print('Cảnh báo: chưa đặt SKT_BI_MAT nên mỗi lần khởi động lại phải đăng nhập lại')
    for d in (DU_LIEU, KET_QUA, LICH_SU): os.makedirs(d, exist_ok=True)
    print('App chạy chế độ máy chủ, cổng', CONG, '- dữ liệu ở', GHI, flush=True)
    HTTPServer(('0.0.0.0', CONG), XuLy).serve_forever()


def main():
    if DAM_MAY: return chay_may_chu()
    url = f'http://127.0.0.1:{CONG}/'
    may = mo_cong(url)
    if may is None:
        if '--khong-mo' not in sys.argv: webbrowser.open(url)
        print('App đã chạy sẵn:', url)
        return
    print('App đang chạy tại', url)
    if os.name == 'nt': print('Giữ cửa sổ này mở trong lúc dùng app, đóng cửa sổ để tắt app.')
    if '--khong-mo' not in sys.argv: webbrowser.open(url)
    try:
        may.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == '__main__':
    main()
