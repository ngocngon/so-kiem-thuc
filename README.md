# Sổ kiểm thực

App lập sổ kiểm thực (bước 1, bước 2, bước 3, mẫu 24 giờ) cho nhà ăn từ file thực đơn tuần.
Viết bằng Python, chỉ cần thư viện `openpyxl` (và `pillow`).

## Chạy trên máy

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python app.py
```

App mở tại http://127.0.0.1:8765. Máy khác cùng mạng vào bằng mã truy cập trong mục **Máy khác và tắt app**.
Thực đơn, sổ và kết quả nằm ở `du_lieu/` và `ket_qua/`, không đưa lên GitHub.

## Chạy trên internet (Render)

1. Trên https://render.com, chọn **New** › **Blueprint** và chọn kho GitHub này. Render đọc `render.yaml`.
2. Nhập 2 mã (mỗi mã ít nhất 8 ký tự, khác nhau):
   - `SKT_MA_TRUY_CAP`: mã cho nhân viên dùng app.
   - `SKT_MA_QUAN_TRI`: mã vào Quản trị (quy tắc, sổ tay, lịch sử).
3. Bấm **Apply**. Vài phút sau app chạy ở địa chỉ dạng `https://so-kiem-thuc.onrender.com`.

Gói Starter có ổ đĩa giữ dữ liệu. Gói Free không có ổ đĩa: mỗi lần máy chủ ngủ hoặc khởi động lại sẽ mất file đã tải lên và quy tắc đã sửa.

Biến môi trường của chế độ máy chủ:

| Biến | Ý nghĩa |
|---|---|
| `SKT_DAM_MAY=1` | Bật chế độ máy chủ: bắt buộc đăng nhập, không có "máy chạy app" |
| `SKT_DU_LIEU` | Thư mục lưu dữ liệu (ổ đĩa của máy chủ) |
| `SKT_MA_TRUY_CAP`, `SKT_MA_QUAN_TRI` | Mã nhân viên, mã quản trị |
| `SKT_BI_MAT` | Khoá ký phiên đăng nhập, để khởi động lại không phải đăng nhập lại |
| `PORT` | Cổng, máy chủ tự đặt |
