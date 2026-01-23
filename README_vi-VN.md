[English](README.md) | [简体中文](README_zh-CN.md) | [繁體中文](README_zh-TW.md) | [Deutsch](README_de-DE.md) | [Français](README_fr-FR.md) | [Русский](README_ru-RU.md) | [Português](README_pt-BR.md) | [日本語](README_ja-JP.md) | [한국어](README_ko-KR.md) | [Español](README_es-ES.md) | [Tiếng Việt](README_vi-VN.md)

# DocWen

Công cụ chuyển đổi định dạng tài liệu và bảng biểu: hỗ trợ chuyển đổi hai chiều Word/Markdown/Excel. Chạy trên máy (offline), đảm bảo an toàn dữ liệu.

## 📖 Bối cảnh dự án

Phần mềm được tạo ra để giải quyết các vấn đề thường gặp trong công việc văn phòng:
- Tài liệu từ nhiều nguồn có định dạng không thống nhất, cần chuẩn hoá.
- Nhiều loại file với yêu cầu định dạng khác nhau.
- Cần chạy offline trong môi trường intranet/thiết bị cũ.

**Triết lý thiết kế**: công cụ nhẹ, dễ dùng, chi phí học thấp. Không nhằm thay thế các công cụ chuyên nghiệp như LaTeX/Pandoc.

## ✨ Tính năng chính

- **📄 Chuyển đổi tài liệu** - Word ↔ Markdown, hỗ trợ công thức và ánh xạ dấu phân cách (---/***/___) với ngắt trang/ngắt mục/dòng kẻ. DOCX/DOC/WPS/RTF/ODT.
- **📊 Chuyển đổi bảng tính** - Excel ↔ Markdown. XLSX/XLS/ET/ODS/CSV. Có công cụ tóm tắt bảng.
- **📑 PDF & file bố cục** - PDF/XPS/OFD → Markdown hoặc DOCX. Hỗ trợ gộp/tách PDF.
- **🖼️ Ảnh** - Chuyển đổi và nén JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC.
- **🔍 Nhận dạng văn bản OCR** - Tích hợp RapidOCR để trích xuất văn bản từ ảnh và PDF.
- **✏️ Kiểm tra lỗi** - Quy tắc tuỳ chỉnh cho ký hiệu, lỗi chính tả, từ nhạy cảm.
- **📝 Mẫu (Template)** - Cơ chế template linh hoạt cho tài liệu/báo cáo.
- **💻 GUI + CLI** - Giao diện đồ hoạ và dòng lệnh.
- **🔒 Chạy hoàn toàn cục bộ** - Chạy offline và có cơ chế cách ly mạng tích hợp để đảm bảo an toàn dữ liệu.
- **🔗 Chạy đơn phiên bản** - Tự động quản lý instance của chương trình và hỗ trợ tích hợp với plugin Obsidian đi kèm.

## Nhật ký thay đổi

### v0.6.0 (2025-01-20)

- Hỗ trợ quốc tế hóa đầy đủ (GUI và CLI hỗ trợ 11 ngôn ngữ).
- Thay thế PaddleOCR bằng RapidOCR để tương thích tốt hơn.
- Thêm các mẫu Word/Excel đa ngôn ngữ.
- Tự động phát hiện và chèn kiểu mẫu.
- Các tối ưu hóa và sửa lỗi khác.

### v0.5.1 (2025-01-01)

- Thêm chuyển đổi công thức toán học hai chiều (Word OMML ↔ Markdown LaTeX).
- Thêm chuyển đổi chú thích cuối trang/cuối văn bản hai chiều.
- Thêm kiểu ký tự và đoạn văn cho mã, trích dẫn, v.v.
- Cải thiện xử lý danh sách (lồng nhiều cấp, đánh số tự động).
- Cải thiện chức năng bảng (phát hiện/chèn kiểu, bảng ba dòng, v.v.).
- Tối ưu hóa việc dọn dẹp và thêm số tiêu đề phụ.
- Cải thiện tương tác giao diện và liên kết cài đặt.

### v0.4.1 (2025-12-05)

- Tái cấu trúc CLI để cải thiện trải nghiệm người dùng.
- Thêm hỗ trợ cho nhiều loại tài liệu hơn.
- Triển khai nhiều tùy chọn có thể cấu hình hơn.

## 🚀 Bắt đầu nhanh

### Khởi chạy chương trình

Nhấp đúp `DocWen.exe` để mở giao diện đồ hoạ.

### Hướng dẫn nhanh

1.  **Chuẩn bị file Markdown**:

    ```markdown
    ---
    title: Test Document
    ---
    
    ## Test Title
    
    This is the test body content.
    ```

2.  **Chuyển đổi bằng kéo thả**:
    - Khởi chạy chương trình.
    - Kéo file `.md` vào cửa sổ.
    - Chọn template.
    - Nhấn "Convert to DOCX".

3.  **Nhận kết quả**:
    - Tài liệu Word chuẩn hoá sẽ được tạo trong cùng thư mục.

**Mẹo**: Có thể dùng các file mẫu trong thư mục `samples/` để trải nghiệm nhanh.

## 📝 Quy ước Markdown

### Ánh xạ cấp tiêu đề

Để dễ ghi nhớ, tiêu đề Markdown trong phần mềm tương ứng **1:1** với tiêu đề Word:
- Tiêu đề và phụ đề của tài liệu đặt trong YAML metadata.
- Markdown `# Heading 1` tương ứng Word "Heading 1".
- Markdown `## Heading 2` tương ứng Word "Heading 2".
- Và tiếp tục như vậy, hỗ trợ tối đa 9 cấp.

**Mẹo**: Nếu bạn muốn dùng `#` làm tiêu đề tài liệu và dùng `##` trở đi cho tiêu đề nội dung, hãy chỉnh style "Heading 1" trong template Word để trông giống tiêu đề (ví dụ: căn giữa, in đậm, cỡ chữ lớn) và chọn scheme đánh số bỏ qua cấp 1 trong phần cài đặt.

### Xuống dòng và đoạn văn

**Quy tắc cơ bản**: Mỗi dòng không rỗng được xem là một đoạn riêng theo mặc định.

**Đoạn trộn**: Khi một tiêu đề phụ cần trộn với nội dung trong cùng một đoạn, phải thỏa các điều kiện:
1.  Tiêu đề phụ kết thúc bằng dấu kết câu (hỗ trợ dấu kết câu đa ngôn ngữ).
2.  Nội dung nằm ở **dòng ngay bên dưới** tiêu đề phụ.
3.  Dòng nội dung không được là phần tử Markdown đặc biệt (tiêu đề, code block, bảng, danh sách, trích dẫn, khối công thức, dấu phân cách, ...).

**Ví dụ**:
```markdown
## I. Work Requirements.
This meeting requires all units to earnestly implement...
```
Hai dòng trên sẽ được gộp thành một đoạn: "I. Work Requirements." giữ định dạng tiêu đề phụ và "This meeting..." giữ định dạng nội dung.

**Lưu ý**:
- Không được có dòng trống giữa tiêu đề phụ và nội dung; nếu có sẽ bị nhận diện thành hai đoạn riêng.
- Nếu tiêu đề phụ không kết thúc bằng dấu kết câu và nội dung không có dòng trống, nội dung sẽ bị gộp vào dòng tiêu đề và định dạng sẽ được điều chỉnh.

### Chuyển đổi dấu phân cách hai chiều

Hỗ trợ chuyển đổi hai chiều giữa dấu phân cách Markdown và ngắt trang/ngắt mục/dòng kẻ trong Word:

-   **DOCX → MD**: Tự động chuyển ngắt trang, ngắt mục và dòng kẻ của Word thành dấu phân cách Markdown.
-   **MD → DOCX**: Tự động chuyển `---`, `***`, `___` thành phần tử Word tương ứng.
-   **Có thể cấu hình**: Quan hệ ánh xạ có thể tuỳ chỉnh trong phần cài đặt.

## 📖 Hướng dẫn sử dụng chi tiết

### Word sang Markdown

1.  Kéo file `.docx` vào cửa sổ chương trình.
2.  Chương trình tự phân tích cấu trúc tài liệu.
3.  Tạo file `.md` có chứa YAML metadata.

**Định dạng hỗ trợ**:
-   `.docx` - Tài liệu Word chuẩn.
-   `.doc` - Tự chuyển sang DOCX để xử lý.
-   `.wps` - Tự chuyển tài liệu WPS để xử lý.

**Tuỳ chọn xuất**:

| Tuỳ chọn | Mô tả |
| :--- | :--- |
| **Trích xuất ảnh** | Nếu bật, ảnh trong tài liệu sẽ được xuất ra thư mục và chèn link ảnh vào file MD. |
| **OCR ảnh** | Nếu bật, chạy OCR trên ảnh và tạo file ảnh `.md` (chứa văn bản nhận dạng). |
| **Dọn số tiêu đề phụ** | Nếu bật, xoá số trước tiêu đề phụ (ví dụ: "一、", "（一）", "1.", ...). |
| **Thêm số tiêu đề phụ** | Nếu bật, tự thêm số theo cấp tiêu đề (có thể cấu hình). |

### Markdown sang Word

1.  Chuẩn bị file `.md` có YAML header.
2.  Kéo vào cửa sổ và chọn template Word tương ứng.
3.  Chương trình tự điền template và tạo tài liệu.

**Tuỳ chọn chuyển đổi**:

| Tuỳ chọn | Mô tả |
| :--- | :--- |
| **Dọn số tiêu đề phụ** | Nếu bật, xoá số trước tiêu đề phụ. |
| **Thêm số tiêu đề phụ** | Nếu bật, tự thêm số theo cấp tiêu đề. |

**Lưu ý**: Nếu có đoạn trộn giữa tiêu đề phụ và nội dung, cần giữ xuống dòng nghiêm ngặt trong file MD (xem "Xuống dòng và đoạn văn").

### Tự động xử lý style của template

Trong quá trình Markdown → DOCX, chương trình tự phát hiện và xử lý style của template:

#### Phân loại style

**Style đoạn (Paragraph Style)**: Áp dụng cho toàn bộ đoạn.

| Style | Hành vi phát hiện | Chèn khi thiếu | Nguồn |
| :--- | :--- | :--- | :--- |
| Heading (1~9) | Phát hiện style đoạn | Style heading trong template | Word built-in |
| Code Block | Phát hiện style đoạn | Font Consolas + nền xám | Định nghĩa bởi phần mềm |
| Quote (1~9) | Phát hiện style đoạn | Nền xám + viền trái | Định nghĩa bởi phần mềm |
| Formula Block | Phát hiện style đoạn | Style công thức | Định nghĩa bởi phần mềm |
| Separator (1~3) | Phát hiện style đoạn | Style đoạn có viền dưới | Định nghĩa bởi phần mềm |

**Style ký tự (Character Style)**: Áp dụng cho vùng chữ được chọn.

| Style | Hành vi phát hiện | Chèn khi thiếu | Nguồn |
| :--- | :--- | :--- | :--- |
| Inline Code | Phát hiện style ký tự | Font Consolas + shading xám | Định nghĩa bởi phần mềm |
| Inline Formula | Phát hiện style ký tự | Style công thức | Định nghĩa bởi phần mềm |

**Style bảng (Table Style)**: Áp dụng cho toàn bộ bảng.

| Style | Hành vi phát hiện | Chèn khi thiếu | Nguồn |
| :--- | :--- | :--- | :--- |
| Three-Line Table | Ưu tiên cấu hình người dùng | Định nghĩa style bảng 3 dòng | Định nghĩa bởi phần mềm |
| Grid Table | Ưu tiên cấu hình người dùng | Định nghĩa style bảng lưới | Định nghĩa bởi phần mềm |

**Định nghĩa đánh số (Numbering Definition)**: Dùng cho định dạng danh sách.

| Loại | Hành vi phát hiện | Xử lý khi thiếu |
| :--- | :--- | :--- |
| List Numbering | Quét các định nghĩa danh sách trong template | Dùng preset decimal/bullet |

#### Quốc tế hoá tên style

-   **Style Word built-in** (heading 1~9):
    -   Tên style dùng tên chuẩn tiếng Anh (ví dụ: `heading 1`).
    -   Word sẽ hiển thị tên đã được bản địa hoá theo ngôn ngữ hệ thống.
-   **Style do phần mềm định nghĩa** (Code Block, Quote, Formula, Separator, Table, ...):
    -   Chèn tên style theo ngôn ngữ giao diện của phần mềm.

**Gợi ý**: Sau khi bạn tuỳ chỉnh style trong template, chương trình sẽ ưu tiên dùng style đó; nếu template không có thì dùng preset mặc định.

### Xử lý file bảng tính

1.  **Excel/CSV sang Markdown**: Kéo file `.xlsx` hoặc `.csv` để tự chuyển sang bảng Markdown.
2.  **Markdown sang Excel**: Chuẩn bị file MD và chọn template Excel để chuyển đổi.

**Định dạng hỗ trợ**:
-   `.xlsx` - Excel chuẩn.
-   `.xls` - Tự chuyển sang XLSX để xử lý.
-   `.et` - Tự chuyển bảng tính WPS để xử lý.
-   `.csv` - Bảng văn bản CSV.

### Chức năng kiểm tra lỗi văn bản

Chương trình cung cấp 4 quy tắc kiểm tra có thể tuỳ chỉnh:

1.  **Kiểm tra cặp dấu câu** - Kiểm tra ngoặc kép, ngoặc đơn, ... có khớp cặp không.
2.  **Kiểm tra ký hiệu** - Phát hiện dùng lẫn dấu câu tiếng Trung/tiếng Anh.
3.  **Kiểm tra lỗi chính tả** - Dựa trên từ điển tuỳ chỉnh.
4.  **Phát hiện từ nhạy cảm** - Dựa trên từ điển tuỳ chỉnh.

**Từ điển tuỳ chỉnh**: Chỉnh sửa trực quan từ điển lỗi chính tả và từ nhạy cảm trong "Cài đặt".

**Cách dùng**:
1.  Kéo tài liệu Word cần kiểm tra vào chương trình.
2.  Chọn các quy tắc cần dùng.
3.  Nhấn nút "Text Proofreading".
4.  Kết quả hiển thị dưới dạng comment trong tài liệu.

## 🛠️ Hệ thống template

### Dùng template có sẵn

Chương trình có sẵn nhiều template (bao gồm đa ngôn ngữ). File template nằm trong thư mục `templates/`.

### Template tuỳ chỉnh

1.  Tạo file template bằng Word hoặc WPS.
2.  Tham khảo template có sẵn và chèn placeholder như `{{Title}}`, `{{DocumentNumber}}`, ... vào vị trí cần điền.
3.  Trong template, các style Heading 1 ~ Heading 5 built-in cần chỉnh sửa thủ công.
4.  Lưu template vào thư mục `templates/`.
5.  Khởi động lại chương trình, template mới sẽ tự được tải.

Bạn cũng có thể copy một template có sẵn, chỉnh sửa và đổi tên.

### Cách dùng placeholder

#### Placeholder trong template Word

**Placeholder theo trường YAML**: Dùng dạng `{{Field Name}}` trong template, sẽ được thay thế bằng giá trị tương ứng trong YAML header của file Markdown.

| Placeholder | Mô tả |
| :--- | :--- |
| `{{Title}}` | Tiêu đề tài liệu (xem thứ tự ưu tiên bên dưới) |
| `{{Body}}` | Vị trí chèn nội dung Markdown |
| Khác | Hỗ trợ mọi trường tuỳ chỉnh |

**Thứ tự ưu tiên lấy tiêu đề**:

| Ưu tiên | Nguồn | Mô tả |
| :--- | :--- | :--- |
| 1 | YAML `Title` | Cao nhất |
| 2 | YAML `aliases` | Lấy phần tử đầu tiên của danh sách hoặc chuỗi |
| 3 | Tên file | Tên file không gồm đuôi `.md` |

**Hỗ trợ đa ngôn ngữ**: Placeholder tiêu đề và nội dung hỗ trợ nhiều ngôn ngữ, ví dụ tiêu đề `{{title}}`, `{{标题}}`, `{{Titel}}`, ...; nội dung `{{body}}`, `{{正文}}`, `{{Inhalt}}`, ...

#### Placeholder trong template Excel

Template Excel hỗ trợ 3 loại placeholder:

**1. Placeholder theo trường YAML** `{{Field Name}}`

Điền một giá trị từ YAML header:

```markdown
---
ReportName: 2024 Annual Sales Statistics
Unit: Sales Dept
---
```

`{{ReportName}}`, `{{Unit}}` sẽ được thay bằng giá trị tương ứng. Trường Title cũng theo cùng thứ tự ưu tiên.

**2. Placeholder điền theo cột** `{{↓Field Name}}`

Trích dữ liệu từ bảng Markdown và điền **xuống dưới** theo từng dòng từ vị trí placeholder:

```markdown
| ProductName | Quantity |
|:--- |:--- |
| Product A | 100 |
| Product B | 200 |
```

`{{↓ProductName}}` sẽ được thay bằng "Product A", dòng tiếp theo điền "Product B".

**3. Placeholder điền theo hàng** `{{→Field Name}}`

Trích dữ liệu từ bảng Markdown và điền **sang phải** theo từng cột từ vị trí placeholder:

```markdown
| Month |
|:--- |
| Jan |
| Feb |
| Mar |
```

`{{→Month}}` sẽ được điền "Jan", "Feb", "Mar" sang phải.

**Xử lý ô gộp**: Chương trình tự bỏ qua các ô không phải ô đầu của vùng gộp để đảm bảo điền đúng.

**Gộp dữ liệu nhiều bảng**: Nếu Markdown có nhiều bảng dùng cùng tiêu đề cột, dữ liệu sẽ được gộp theo thứ tự và điền liên tục.

## 🖥️ Sử dụng giao diện đồ hoạ

Hầu hết người dùng sử dụng phần mềm qua giao diện đồ hoạ. Dưới đây là hướng dẫn chi tiết.

### Tổng quan giao diện

Chương trình dùng **bố cục 3 cột thích ứng**:

| Khu vực | Mô tả | Khi hiển thị |
| :--- | :--- | :--- |
| **Cột giữa (khu vực chính)** | Khu kéo thả file, panel thao tác, thanh trạng thái | Luôn hiển thị |
| **Cột phải** | Bộ chọn template / panel chuyển đổi | Tự mở sau khi chọn file |
| **Cột trái** | Danh sách file theo lô (nhóm theo loại) | Hiển thị khi bật chế độ theo lô |

### Quy trình thao tác cơ bản

1.  **Khởi chạy**: Nhấp đúp `DocWen.exe`.
2.  **Nhập file**:
    -   Cách 1: Kéo thả file vào cửa sổ.
    -   Cách 2: Nhấn nút "Add" trong vùng kéo thả để chọn file.
3.  **Chọn template** (nếu cần): Panel template bên phải tự mở, chọn template phù hợp.
4.  **Chọn tuỳ chọn**: Tick các tuỳ chọn xuất/chuyển đổi trong panel thao tác.
5.  **Thực thi**: Nhấn nút chức năng tương ứng (ví dụ: "Export MD", "Convert to DOCX", ...).
6.  **Xem kết quả**: Thanh trạng thái hiển thị tiến độ; nhấn biểu tượng 📍 để mở vị trí file output.

### Chế độ 1 file vs chế độ theo lô

Chương trình hỗ trợ hai chế độ xử lý, chuyển đổi bằng nút trong vùng kéo thả:

**Chế độ 1 file** (mặc định):
-   Xử lý từng file một.
-   Giao diện gọn, phù hợp sử dụng hằng ngày.

**Chế độ theo lô**:
-   Nhập nhiều file cùng lúc.
-   Cột trái hiển thị danh sách file theo nhóm.
-   Hỗ trợ thêm/xoá/sắp xếp theo lô.
-   Nhấp vào file trong danh sách để đổi mục tiêu thao tác.

### Chức năng panel thao tác

Panel thao tác tự điều chỉnh chức năng theo loại file:

| Loại file | Thao tác khả dụng |
| :--- | :--- |
| Tài liệu Word | Export MD, Convert PDF, Text Proofreading, OCR |
| Markdown | Convert DOCX, Convert PDF |
| Bảng tính Excel | Export MD, Convert PDF, Table Summary |
| PDF | Export MD, Merge, Split, OCR |
| Ảnh | Chuyển đổi định dạng, Nén, OCR |

### Màn hình cài đặt

Nhấn nút ⚙️ để mở cài đặt:

-   **Chung**: Chủ đề, ngôn ngữ, độ trong suốt cửa sổ
-   **Chuyển đổi**: Giá trị mặc định cho các tuỳ chọn chuyển đổi
-   **Đầu ra**: Thư mục output mặc định, quy tắc đặt tên
-   **Kiểm tra lỗi**: Sửa từ điển lỗi chính tả/từ nhạy cảm
-   **Style**: Cấu hình style cho code block, trích dẫn, bảng

### Phím tắt

-   **Kéo file ngoài**: Kéo trực tiếp vào cửa sổ để nhập
-   **Nhấp đúp kết quả trên thanh trạng thái**: Mở nhanh thư mục output
-   **Chuột phải vào template**: Mở vị trí file template

---

## 🔧 Dùng dòng lệnh

Ngoài GUI, chương trình cung cấp CLI phù hợp tự động hoá và xử lý theo lô.

### Chế độ chạy

-   **Chế độ tương tác**: Hiển thị menu hướng dẫn sau khi truyền file, tương tự GUI.
-   **Chế độ headless**: Chạy trực tiếp bằng tham số `--action`, phù hợp cho script.

### Ví dụ thường dùng

```bash
# Chế độ tương tác
DocWen.exe document.docx

# Xuất Word sang Markdown (Trích ảnh + OCR)
DocWen.exe report.docx --action export_md --extract-img --ocr

# Markdown sang Word (Chỉ định template)
DocWen.exe document.md --action convert --target docx --template "Template Name"

# Chuyển đổi theo lô (Bỏ xác nhận, tiếp tục khi lỗi)
DocWen.exe *.docx --action export_md --batch --yes --continue-on-error

# Kiểm tra lỗi tài liệu
DocWen.exe document.docx --action validate --check-typo --check-punct

# Gộp/Tách PDF
DocWen.exe *.pdf --action merge_pdfs
DocWen.exe report.pdf --action split_pdf --pages "1-3,5,7-10"
```

### Tham số chính

| Tham số | Mô tả |
| :--- | :--- |
| `--action` | Loại thao tác: `export_md`, `convert`, `validate`, `merge_pdfs`, `split_pdf` |
| `--target` | Định dạng đích: `pdf`, `docx`, `xlsx`, `md` |
| `--template` | Tên template (ví dụ: `Template Name`) |
| `--extract-img` | Trích xuất ảnh |
| `--ocr` | Bật OCR |
| `--batch` | Chế độ theo lô |
| `--yes` / `-y` | Bỏ qua hỏi xác nhận |
| `--continue-on-error` | Lỗi vẫn tiếp tục |
| `--json` | Xuất kết quả dạng JSON |
| `--quiet` / `-q` | Chế độ yên lặng |

## 🔌 Plugin Obsidian

Dự án có plugin Obsidian đi kèm để dùng cùng với bộ chuyển đổi:

### Tính năng cốt lõi

-   **🚀 Khởi chạy 1 lần nhấn** - Icon ở sidebar để mở nhanh bộ chuyển đổi.
-   **📂 Bàn giao tự động** - Tự truyền đường dẫn file đang mở.
-   **🔄 Quản lý đơn phiên bản** - Nếu chương trình đang chạy, chỉ gửi file, không cần khởi chạy lại.
-   **💪 Khôi phục khi crash** - Tự phát hiện trạng thái process và dọn file dư.

### Nguyên lý hoạt động

Plugin tương tác với bộ chuyển đổi thông qua IPC dựa trên hệ thống file:

1.  **Nhấn lần đầu** → Khởi chạy bộ chuyển đổi và truyền file hiện tại.
2.  **Nhấn lại (có file)** → Thay file mới (chế độ 1 file).
3.  **Nhấn lại (không có file)** → Kích hoạt cửa sổ bộ chuyển đổi.

### Cài đặt

Plugin được phát hành ở repo riêng. Xem [docwen-obsidian](https://github.com/ZHYX91/docwen-obsidian) để biết cách cài và phiên bản mới nhất.

## ❓ FAQ

### Nếu chuyển đổi thất bại thì sao?

-   Kiểm tra file có đang bị ứng dụng khác mở/chiếm dụng không.
-   Xác nhận định dạng file đúng.
-   Xem log trong thư mục `logs/`.

### Template không hiển thị?

-   Xác nhận file template nằm trong `templates/`.
-   Kiểm tra template có bị hỏng không.
-   Khởi động lại chương trình để tải lại template.

### Chức năng kiểm tra lỗi không hoạt động?

-   Xác nhận tài liệu là `.docx`.
-   Kiểm tra tài liệu có chứa text có thể chỉnh sửa không.
-   Xác nhận các quy tắc kiểm tra đã bật trong cài đặt.

### Định dạng output không như mong đợi?

-   Chương trình tạo tài liệu dựa trên style của template. Nếu muốn điều chỉnh output, hãy sửa trực tiếp style trong file template.
-   Template nằm trong `templates/`.
-   Sau khi sửa style, mọi tài liệu chuyển đổi với template đó sẽ áp dụng style mới.

### Ô công thức bị trống sau khi chuyển Excel sang Markdown?

Đây là hành vi dự kiến. Chương trình đọc **giá trị cache** của ô thay vì công thức.

**Lý do kỹ thuật**:
-   Ô công thức trong Excel lưu cả công thức và kết quả tính gần nhất (cache).
-   Chương trình dùng `data_only=True` nên chỉ đọc cache.
-   Nếu file chưa từng mở trong Excel hoặc chưa lưu lại sau khi tính, cache có thể trống.

**Giải pháp**:
1.  Mở file trong Excel.
2.  Đợi tính toán hoàn tất.
3.  Lưu file.
4.  Chuyển đổi lại.

## 🔒 Tính năng bảo mật

-   **Chạy hoàn toàn cục bộ**: Mọi xử lý diễn ra trên máy, không phụ thuộc mạng.
-   **Cách ly mạng**: Cơ chế cách ly mạng tích hợp để tránh rò rỉ dữ liệu.
-   **Không upload dữ liệu**: File của người dùng không được upload lên bất kỳ server nào.

## 📜 Giấy phép

Dự án dùng giấy phép **GNU Affero General Public License v3.0 (AGPL-3.0)**.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   Dự án sử dụng PyMuPDF (AGPL-3.0), nên toàn bộ dự án cũng theo AGPL-3.0.
-   Bạn có thể tự do sử dụng, sửa đổi và phân phối phần mềm.
-   Nếu bạn sửa phần mềm và cung cấp dịch vụ qua mạng, bạn phải cung cấp mã nguồn đã sửa cho người dùng.
-   Xem thêm trong [LICENSE](LICENSE).

### Liên hệ

-   **GitHub**: https://github.com/ZHYX91/docwen
-   **Email**: zhengyx91@hotmail.com

---

**Author**: ZhengYX
