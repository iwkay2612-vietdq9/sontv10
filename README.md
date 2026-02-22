# Dự án Chia File Son

Ứng dụng hỗ trợ xử lý file Excel và chỉnh sửa ảnh trực tuyến.

## Cách chạy ứng dụng

### 1. Xem trực tiếp qua GitHub Pages (Khuyên dùng)
Bạn có thể xem ứng dụng chạy trực tiếp mà không cần cài đặt gì qua link:
`https://iwkay2612-vietdq9.github.io/sontv10/`

**Cách bật GitHub Pages:**
- Truy cập vào repository của bạn trên GitHub.
- Chọn **Settings** (Cài đặt).
- Chọn **Pages** ở menu bên trái.
- Tại mục **Build and deployment** -> **Branch**, chọn `main` và thư mục `/ (root)`.
- Nhấn **Save**. Sau vài phút, link website của bạn sẽ hiện ra ở phía trên.

### 2. Chạy trên máy tính cá nhân (Local)
- Tải toàn bộ code về hoặc dùng `git clone`.
- Mở file `index.html` bằng trình duyệt (Chrome, Edge, Firefox).
- Hoặc sử dụng extension **Live Server** trên VS Code để chạy tốt hơn.

## Cấu trúc thư mục
- `index.html`: Trang chủ chính.
- `app.js`, `script.js`: Xử lý logic.
- `style.css`: Định dạng giao diện.
- `sua-anh/`: Thư mục chứa ứng dụng chỉnh sửa ảnh riêng biệt.
- `gui_app.py`: Phiên bản ứng dụng chạy trên máy tính (Python).
