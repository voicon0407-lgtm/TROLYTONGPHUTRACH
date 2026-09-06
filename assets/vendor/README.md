# Thư viện cục bộ dùng khi ngoại tuyến

Các gói dưới đây được phân phối trong ứng dụng để Trung tâm nhập dữ liệu không phụ thuộc CDN:

| Thành phần | Phiên bản | Chức năng | Giấy phép |
|---|---:|---|---|
| SheetJS Community Edition | 0.18.5 | Đọc/ghi XLS, XLSX | Apache-2.0 |
| Mammoth.js | 1.12.1 | Trích nội dung DOCX | BSD-2-Clause |
| PDF.js | 5.6.205 | Đọc văn bản và dựng trang PDF | Apache-2.0 |
| Tesseract.js | 7.x | OCR ảnh/PDF scan | Apache-2.0 |
| `vie.traineddata` | tessdata_fast | Nhận dạng tiếng Việt | Apache-2.0 |

Tệp giấy phép gốc nằm trong thư mục tương ứng. Ứng dụng chỉ tải các thư viện này khi người dùng chọn đúng loại tệp; Service Worker lưu chúng trong bộ nhớ đệm ứng dụng để dùng ngoại tuyến. Không có tệp nào gọi CDN hoặc gửi tài liệu nhập lên máy chủ bên thứ ba.
