#!/bin/bash
# Bấm đúp để tạo sổ. Chạy lệnh thì truyền: Tao_so.command <thực đơn> <mẫu> <ngày> [ngày...]
cd "$(dirname "$0")" || exit 1

dung() { echo; echo "Lỗi: $1"; read -r -p "Nhấn Enter để đóng..." _; exit 1; }

# Lần đầu chạy (hoặc sau khi chuyển thư mục) thì cài thư viện vào .venv
if ! .venv/bin/python -c 'import openpyxl, PIL' 2>/dev/null; then
    echo "Đang chuẩn bị thư viện (chỉ lần đầu, cần mạng)..."
    rm -rf .venv
    python3 -m venv .venv && .venv/bin/python -m pip install -q --disable-pip-version-check openpyxl pillow \
        || dung "không cài được thư viện. Kiểm tra mạng rồi chạy lại."
fi

chon_file() {
    osascript -e "POSIX path of (choose file with prompt \"$1\" of type {\"org.openxmlformats.spreadsheetml.sheet\"} default location (path to downloads folder))" 2>/dev/null
}

if [ $# -ge 3 ]; then
    td=$1; mau=$2; shift 2; ngay=("$@")
else
    td=$(chon_file "Chọn file THỰC ĐƠN VIỆT của tuần") || exit 0
    mau=$(chon_file "Chọn file VIỆT MẪU") || exit 0
    nhap=$(osascript -e 'text returned of (display dialog "Nhập ngày cần làm, nhiều ngày cách nhau bằng dấu cách (vd: 23.10 24.10)" default answer "")' 2>/dev/null) || exit 0
    read -r -a ngay <<< "$nhap"
fi
[ ${#ngay[@]} -gt 0 ] || dung "chưa nhập ngày."

mkdir -p ket_qua
loi=0
for d in "${ngay[@]}"; do
    # Giữ lại bản cũ vì có thể đã điền tay số trái cây, sữa
    cu="ket_qua/So_$d.xlsx"
    [ -e "$cu" ] && mv "$cu" "ket_qua/So_${d}_cu_$(date +%Y%m%d_%H%M%S).xlsx"
    .venv/bin/python tao_so.py "$td" "$mau" "$d" ket_qua || { echo "Không tạo được sổ ngày $d"; loi=1; }
done

[ -n "$KHONG_MO" ] || open ket_qua
[ $loi -eq 0 ] || dung "có ngày chưa tạo được, xem thông báo ở trên."
echo "Xong. Sheet CẦN SOÁT trong từng file ghi những chỗ cần xem lại."
