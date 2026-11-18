"""Cầu nối cho bản chạy trong trình duyệt (GitHub Pages + Pyodide).

Giao diện vẫn gọi /api/... như với app trên máy. Ở bản web không có máy chủ, nên mỗi lời gọi được đưa
thẳng vào bộ xử lý XuLy của app.py (cùng một code, cùng kiểm tra), đọc câu trả lời HTTP rồi trả về trình duyệt.
"""
import io
import app


class _Goi(app.XuLy):
    """XuLy không có kết nối mạng: đọc yêu cầu từ bộ nhớ, ghi câu trả lời vào bộ nhớ."""
    def __init__(self, phuong, duong, than):
        self.command, self.path, self.request_version = phuong, duong, 'HTTP/1.1'
        self.requestline = f'{phuong} {duong} HTTP/1.1'
        self.client_address = ('127.0.0.1', 0)                 # bản web: người dùng là chủ máy
        self.headers = {'Host': f'127.0.0.1:{app.CONG}', 'Content-Length': str(len(than))}
        self.rfile, self.wfile = io.BytesIO(than), io.BytesIO()


def goi(phuong, duong, than=b''):
    """Trả (mã HTTP, kiểu nội dung, tên file tải về hoặc '', nội dung bytes)."""
    if hasattr(than, 'to_bytes'): than = than.to_bytes()        # Uint8Array từ trình duyệt
    h = _Goi(phuong, duong, bytes(than or b''))
    (h.do_POST if phuong == 'POST' else h.do_GET)()
    tho = h.wfile.getvalue()
    dau, _, noi = tho.partition(b'\r\n\r\n')
    dong = dau.decode('latin-1').split('\r\n')
    ma = int(dong[0].split()[1])
    tt = {}
    for x in dong[1:]:
        k, _, v = x.partition(':')
        tt[k.strip().lower()] = v.strip()
    ten = ''
    if 'filename=' in tt.get('content-disposition', ''):
        ten = tt['content-disposition'].split('filename=')[1].strip('"')
    return ma, tt.get('content-type', ''), ten, noi
