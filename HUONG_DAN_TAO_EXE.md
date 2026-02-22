# Hướng dẫn tạo file .exe (Chương trình chạy)

Do máy tính của bạn chưa cài Python, tôi không thể tự tạo trực tiếp file `.exe` cho bạn được. Tuy nhiên, bạn có thể tự làm việc này rất dễ dàng chỉ với vài bước đơn giản.

## Bước 1: Cài đặt Python
1. Tải Python tại: [python.org/downloads](https://www.python.org/downloads/)
2. Chạy file cài đặt `python-installer.exe`.
3. **QUAN TRỌNG:** Ở màn hình đầu tiên, hãy tích vào ô **"Add Python to PATH"** trước khi bấm **Install Now**.

## Bước 2: Cài đặt thư viện cần thiết
1. Mở thư mục chứa code này: `d:\code\chia-file`
2. Chuột phải vào khoảng trống, chọn **"Open in Terminal"** (hoặc gõ `cmd` vào thanh địa chỉ folder rồi Enter).
3. Copy và dán dòng lệnh sau vào bảng đen (Terminal) để cài thư viện:
   ```bash
   pip install xlwings pyinstaller
   ```
   *(Đợi nó chạy xong...)*

## Bước 3: Tạo file .exe
Copy và dán dòng lệnh sau vào Terminal:
   ```bash
   pyinstaller --onefile --noconsole --name "LocExcel_GiuDinhDang" gui_app.py
   ```

## Kết quả
Sau khi chạy xong, bạn sẽ thấy thư mục `dist` trong folder hiện tại.
Bên trong đó có file **`LocExcel_GiuDinhDang.exe`**.
Bạn có thể copy file này ra màn hình Desktop và sử dụng vĩnh viễn (không cần cài gì thêm nữa).

---
**Lưu ý:** Chương trình này sử dụng Excel của bạn để xử lý nên sẽ giữ nguyên 100% định dạng file gốc.
