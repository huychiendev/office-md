# CODEMAP - Bản đồ Kiến trúc Mã nguồn Office to Markdown

Tài liệu này ghi nhận toàn bộ cấu trúc mã nguồn, vai trò của từng tệp tin và hướng dẫn luồng dữ liệu để phục vụ bảo trì, nâng cấp nhanh chóng.

## 1. Cấu trúc Thư mục

```
d:\Source-Code\Tool\office-md/
├── index.html                   # Giao diện Web Tĩnh chính (Root) - Native Client Engine
├── app.py                       # Máy chủ tĩnh siêu nhẹ (FastAPI Static Server)
├── Dockerfile                   # Cấu hình container Docker tối giản
├── docker-compose.yml           # Khởi chạy dịch vụ Docker
├── requirements.txt             # Chỉ gồm FastAPI và Uvicorn cho server tĩnh
├── README.md                    # Hướng dẫn tổng quan dự án
├── CODEMAP.md                   # Bản đồ cấu trúc mã nguồn
│
├── pro/                         # THƯ MỤC WEB TĨNH PRO ĐỘC LẬP (100% STATIC - ZERO PYTHON)
│   ├── index.html               # Ứng dụng SPA Full-Width Native Client-Side Engine
│   ├── libs/                    # Thư viện JS Vendor độc lập (Hoạt động cả Offline)
│   │   ├── exceljs.min.js       # Core phân tích Excel (công thức, màu sắc, merged cells)
│   │   ├── mammoth.browser.min.js # Core chuyển đổi Word (.docx) sang Markdown & trích xuất ảnh
│   │   ├── jszip.min.js         # Core phân tích PPTX, VBA Macros và tạo file ZIP
│   │   ├── pdf.min.js           # Core trích xuất văn bản PDF
│   │   ├── pdf.worker.min.js    # Web Worker hỗ trợ giải mã PDF
│   │   ├── marked.min.js        # Parser Markdown sang HTML
│   │   └── lucide.min.js        # Icon vector giao diện
│   └── README.md                # Hướng dẫn triển khai gói Pro
│
├── static/                      # Thư mục phân phối tĩnh cho FastAPI
│   └── index.html               # Đồng bộ với giao diện chính
│
└── structure/                   # Tài liệu thiết kế kiến trúc chi tiết
    └── index_structure.md       # Phân tích UI Tree, States, Logic Flow & Hướng dẫn bảo trì
```

## 2. Bản đồ Trách nhiệm Mô-đun (Module Responsibilities)

| Mô-đun / Tệp tin | Trách nhiệm chính | Thư viện sử dụng |
| :--- | :--- | :--- |
| **`pro/` (Gói Web Tĩnh Độc Lập)** | - Ném trực tiếp lên máy chủ tĩnh bất kỳ (Nginx, S3, Cloudflare Pages, GitHub Pages...).<br>- **100% tĩnh**: Không chứa bất kỳ file `.py` hay backend nào.<br>- Tích hợp sẵn toàn bộ thư viện trong `pro/libs/` giúp chạy offline hoàn toàn. | `ExcelJS`, `Mammoth.js`, `JSZip`, `PDF.js`, `Marked.js`, `Lucide`, `Tailwind` |
| **`app.py`** | Phục vụ các file tĩnh khi người dùng muốn chạy dạng Docker / VPS. | `FastAPI`, `uvicorn` |

## 3. Các Hàm Chuyển Đổi Cốt Lõi (Core Conversion Functions)

Các hàm này nằm trực tiếp trong thẻ `<script>` của `pro/index.html` và `index.html`:

- `convertExcel(buffer, filename, excludeHidden)`: Phân tích file `.xlsx`, `.xlsm`, đọc công thức `(=FORMULA)`, màu nền ARGB, ô gộp, trích xuất ảnh nhúng.
- `convertPPTX(buffer, filename)`: Phân tích file `.pptx`, trích xuất văn bản slide, bảng biểu và toàn bộ ảnh nhúng trong `ppt/media/`.
- `convertDOCX(buffer)`: Chuyển đổi `.docx` sang Markdown qua `Mammoth.js`, xử lý ảnh nhúng thành Data URI.
- `convertPDF(buffer)`: Trích xuất văn bản phân trang qua `Mozilla PDF.js`.
- `extractVBA(buffer)`: Phát hiện và trích xuất module mã nguồn VBA từ `vbaProject.bin`.
- `convertDocumentNative(buffer, filename, excludeHidden)`: Router điều phối chính tiếp nhận file và gọi hàm tương ứng.
- `processFiles(files)`: Đọc mảng `ArrayBuffer` và kích hoạt chu trình chuyển đổi hàng loạt.

## 4. Hướng dẫn Nâng cấp & Thêm Định Dạng Mới

Khi muốn thêm hỗ trợ một định dạng mới (Ví dụ: `.rtf`, `.epub`, `.odt`):
1. Thêm extension vào thuộc tính `accept` của thẻ `<input type="file" id="fileInput">`.
2. Tạo hàm phân tích mới: `async function convertNewFormat(buffer, filename)`.
3. Khai báo nhánh điều kiện trong hàm `convertDocumentNative()`.
4. Cập nhật lại tài liệu trong `CODEMAP.md` và `structure/index_structure.md`.
