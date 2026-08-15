# Office to Markdown Converter PRO (High Speed Client-Side)

Ứng dụng chuyển đổi tài liệu Office sang Markdown hoạt động **100% phía trình duyệt (Client-side)** với tốc độ tức thì, không phụ thuộc vào máy chủ.

## Điểm nổi bật

- **Khởi động tức thì (Zero Latency)**: Không cần nạp máy ảo nặng nề, ứng dụng ở trạng thái sẵn sàng ngay khi vừa mở trang.
- **Tải máy chủ = 0%**: Toàn bộ quá trình đọc, phân tích và chuyển đổi dữ liệu diễn ra 100% trên RAM và CPU của máy khách.
- **Bảo mật tuyệt đối**: Dữ liệu tài liệu của người dùng không bao giờ gửi ra môi trường internet.
- **Hỗ trợ định dạng chuyên sâu**:
  - **Excel (.xlsx, .xlsm, .xltx, .xltm)**: Trích xuất công thức gốc `(=FORMULA)`, tái tạo màu nền bảng tính (ARGB & Tinting), ô gộp và trích xuất toàn bộ ảnh nhúng trong sheet.
  - **PowerPoint (.pptx, .pptm, .potx)**: Trích xuất nội dung văn bản slide, bảng biểu và toàn bộ hình ảnh nhúng.
  - **Word (.docx, .docm)**: Chuyển đổi định dạng văn bản, danh sách, bảng biểu và hình ảnh qua `Mammoth.js`.
  - **PDF (.pdf)**: Trích xuất văn bản phân trang qua `Mozilla PDF.js`.
  - **VBA Macros**: Tự động phát hiện và trích xuất module VBA từ `vbaProject.bin`.
  - **Đóng gói ZIP tức thì**: Tự động tạo file `.zip` (gồm file `.md` và toàn bộ ảnh) trực tiếp trên trình duyệt bằng `JSZip`.

## Cấu trúc thư mục

- `index.html`: Giao diện Web Tĩnh chính (Full Width, Native Instant Engine).
- `pro/`: Gói Web Tĩnh PRO độc lập (có kèm script chạy `run_server.py`).
- `app.py`: Máy chủ tĩnh FastAPI siêu nhẹ cho Docker / VPS.
- `CODEMAP.md`: Bản đồ chi tiết toàn bộ mã nguồn.
- `structure/index_structure.md`: Tài liệu thiết kế kiến trúc, UI Tree và luồng dữ liệu phục vụ bảo trì.

## Hướng dẫn Khởi chạy

### Cách 1: Chạy trực tiếp từ thư mục `pro/`
```bash
cd pro
python run_server.py
```
Trình duyệt sẽ tự động mở `http://localhost:8080`.

### Cách 2: Triển khai Web Tĩnh lên Hosting
Bạn có thể upload trực tiếp toàn bộ thư mục `pro/` lên bất kỳ nền tảng tĩnh nào (GitHub Pages, Cloudflare Pages, Vercel, Netlify, Nginx) mà không cần cài đặt Python hay máy chủ backend.
