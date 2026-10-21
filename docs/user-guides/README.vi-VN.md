# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md) · [العربية](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ar.md)

DocWen là công cụ chuyển đổi tài liệu và bảng biểu cục bộ cho Word, Markdown, Excel và các định dạng văn phòng liên quan. Công cụ ưu tiên quy trình offline ổn định và xử lý tài liệu ngay trên thiết bị của người dùng.

Ngôn ngữ giao diện chỉ điều khiển văn bản và bố cục ứng dụng. Ngôn ngữ đầu ra tác vụ trong cài đặt chung điều khiển riêng nhãn được tạo, kiểu tài liệu và chú thích hiệu đính; không dịch nội dung. Đổi giao diện vẫn giữ ngôn ngữ đầu ra và mẫu đã lưu. OCR Tự động theo ngôn ngữ đầu ra. Tùy chọn tạo một đoạn Word cho mỗi dòng văn bản Markdown thông thường mặc định tắt.

## 📖 Bối cảnh dự án

Phần mềm được tạo ra để giải quyết các vấn đề thường gặp trong công việc văn phòng:
- Tài liệu từ nhiều nguồn có định dạng không thống nhất, cần chuẩn hoá.
- Nhiều loại file với yêu cầu định dạng khác nhau.
- Cần chạy offline trong môi trường intranet/thiết bị cũ.

**Triết lý thiết kế**: DocWen ưu tiên các quy trình chuyển đổi và sắp xếp tài liệu hằng ngày dễ học và có thể dùng ngay. Công cụ không nhằm thay thế các hệ thống dàn trang hoặc chuyển đổi chuyên biệt như LaTeX hay Pandoc; mục tiêu là làm cho các tác vụ tài liệu phổ biến trở nên trực tiếp, rõ ràng và vẫn giữ ranh giới kỹ thuật minh bạch.

## ✨ Tính năng chính

- **📄 Chuyển đổi tài liệu** - Word ↔ Markdown, hỗ trợ công thức, ánh xạ dấu phân cách (---/***/___) với ngắt trang/ngắt mục/dòng kẻ và khôi phục dấu gộp `<` / `^` trong bảng Markdown thành gộp ô hình chữ nhật trong Word. DOCX/DOC/WPS/RTF/ODT.
- **📊 Chuyển đổi bảng tính** – Excel ↔ Markdown với XLSX/XLS/ET/ODS/CSV/TSV. Khi xuất ô gộp, có thể lặp lại nội dung, để trống các ô bị gộp hoặc giữ cấu trúc bằng ký hiệu. Mẫu XLSX hỗ trợ trường YAML, hàng và cột của bảng, hình ảnh và ô gộp. Có thể tổng hợp nhiều bảng với nhau.
- **📑 PDF & file bố cục** - PDF/XPS/OFD → Markdown hoặc DOCX. Hỗ trợ gộp/tách PDF.
- **🖼️ Ảnh** - Chuyển đổi và nén JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC.
- **📥 Nhập định dạng khác** - Hỗ trợ chuyển đổi một chiều HTML/MHTML/ENEX/EPUB/PPTX/PPT sang Markdown.
- **🔍 Nhận dạng văn bản OCR** - Tích hợp RapidOCR để trích xuất văn bản từ ảnh và PDF.
- **✏️ Kiểm tra lỗi** - Quy tắc tuỳ chỉnh cho dấu câu, ký hiệu, lỗi chính tả và từ nhạy cảm trong file Word (.docx) và Markdown (.md). Có thể chỉnh sửa quy tắc trong giao diện cài đặt.
- **📝 Mẫu (Template)** - Cơ chế template linh hoạt cho tài liệu/báo cáo.
- **💻 GUI + CLI** - Giao diện đồ hoạ và dòng lệnh.
- **🔒 Xử lý cục bộ với bảo vệ kết nối đi của thư viện** - Việc chuyển đổi không phụ thuộc dịch vụ trực tuyến. Khi DocWen chạy, tiến trình Python chặn DNS và IPv4/IPv6 của thư viện trong tiến trình; các ứng dụng Office bên ngoài vẫn theo chính sách mạng của hệ thống.
- **🔗 Chạy đơn phiên bản** - Tự động quản lý instance của chương trình và hỗ trợ tích hợp với plugin Obsidian đi kèm.

## 📸 Ảnh chụp màn hình

Chọn mẫu và các tùy chọn tạo tài liệu từ Markdown.

![Tạo tài liệu từ Markdown](../assets/screenshots/markdown-generation-en.png)


Nhật ký thay đổi: xem [CHANGELOG.md](../CHANGELOG.md)

## 🚀 Bắt đầu nhanh

### Cài đặt từ mã nguồn

**Yêu cầu**: Python 3.12

**Nền tảng gói được hỗ trợ**: Mã nguồn này tạo gói cho Windows x64 và Ubuntu 24.04 x64. Các bản phân
phối Linux khác và macOS vẫn là đường dẫn mã nguồn/phát triển, không thuộc cam kết của gói Ubuntu.

**Cách 1: Sử dụng uv (khuyến nghị)**

Cài đặt [uv](https://docs.astral.sh/uv/getting-started/), sau đó:

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
```

Mã nguồn, kiểm thử và bản dựng DocWen chỉ hỗ trợ tệp khóa trong kho với `uv 0.12.0`; `pip install -e` không được hỗ trợ.

### Khởi chạy chương trình

Với bản đóng gói trên Windows: nhấp đúp `DocWen.exe` để mở GUI. Sau khi cài từ mã nguồn:

```bash
docwen-gui  # Chế độ GUI
docwen      # Chế độ CLI
```

### Ghi chú cho macOS

**Giới hạn hiện tại**: Trên macOS, các chức năng `convert`, `validate`, `number`, `merge`, `split` hiện không khả dụng. Phần dưới chỉ ghi các phụ thuộc tùy chọn cho thử nghiệm phát triển.

**Hỗ trợ LibreOffice (Tùy chọn)**

Nếu cần chuyển đổi các định dạng cũ như `.doc`, `.xls`, hãy cài LibreOffice:  
Tải về: https://www.libreoffice.org/download/

**Hỗ trợ ảnh HEIC (Tùy chọn)**

Để xử lý ảnh HEIC/HEIF:

```bash
brew install libheif
pip install pillow-heif
```

### Yêu cầu cho bản GUI trên Linux

**Đích gói được hỗ trợ**: DocWen hỗ trợ GUI và CLI trong gói Ubuntu 24.04 x64. Các yêu cầu này
không mở rộng cam kết sang bản phân phối hoặc kiến trúc khác.

- Có môi trường desktop (GNOME, KDE, XFCE, ...)
- GUI dùng PySide6 (Qt6) và không còn phụ thuộc vào Python Tk. Nếu khởi động lỗi vì thiếu thư viện hệ thống, hãy cài các thư viện cần thiết để chạy Qt theo thông báo lỗi (thường liên quan OpenGL/X11).
- Với máy chủ headless, hãy ưu tiên entry CLI `docwen` thay vì GUI; bản đóng gói Windows cũng cung cấp `DocWenCLI.exe`.

### Hướng dẫn nhanh

1.  **Chuẩn bị file Markdown**:

    ```markdown
    ---
    title: Tài liệu thử nghiệm
    ---
    
    ## Tiêu đề thử nghiệm
    
    Đây là nội dung dùng để thử nghiệm.
    ```

2.  **Chuyển đổi bằng kéo thả**:
    - Khởi chạy chương trình.
    - Kéo file `.md` vào cửa sổ.
    - Chọn template.
    - Nhấn "Tạo DOCX".

3.  **Nhận kết quả**:
    - Tài liệu Word được lưu trong thư mục kết quả tại vị trí đầu ra đã chọn. Có thể mở vị trí này từ thanh trạng thái.

**Mẹo**: Có thể dùng các file mẫu trong thư mục `samples/` để trải nghiệm nhanh.

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

1.  **Khởi chạy**: Nhấp đúp `DocWen.exe` (Windows đóng gói) hoặc chạy `docwen-gui`.
2.  **Nhập file**:
    -   Cách 1: Kéo thả file vào cửa sổ.
    -   Cách 2: Nhấn nút "Add" trong vùng kéo thả để chọn file.
3.  **Chọn template** (nếu cần): Panel template bên phải tự mở, chọn template phù hợp.
4.  **Chọn tuỳ chọn**: Tick các tuỳ chọn xuất/chuyển đổi trong panel thao tác.
5.  **Thực thi**: Nhấn nút chức năng tương ứng (ví dụ: "Xuất Markdown", "Tạo DOCX", ...).
6.  **Xem kết quả**: Thanh trạng thái hiển thị tiến độ và kết quả; nhấn thao tác "Mở đầu ra" ở bên phải để mở vị trí đầu ra.

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
| Tài liệu Word | Xuất Markdown, Chuyển PDF, Kiểm tra văn bản, OCR |
| Markdown | Tạo DOCX, Chuyển PDF, Kiểm tra văn bản |
| Bảng tính Excel | Xuất Markdown, Chuyển PDF, Tóm tắt bảng |
| PDF | Xuất Markdown, Trộn, Tách, OCR |
| Ảnh | Chuyển đổi định dạng, Nén, OCR |
| HTML/EPUB/PPTX v.v. | Xuất Markdown |

### Màn hình cài đặt

Mở cài đặt bằng biểu tượng ở cuối cửa sổ hoặc **Ctrl+,**. Thanh điều hướng bên trái được chia thành **Cơ bản**, **Loại đầu vào**, **Nội dung** và **Chuyên biệt**. Khi cửa sổ thấp có thể cuộn thanh bên; chỉ khi không đủ chiều rộng mới chuyển sang danh sách chọn trang.

Trang **OCR và hình ảnh** quản lý nhận dạng chung và xuất hình ảnh trong Markdown. OCR mặc định theo từng loại tệp nằm ở trang đầu vào tương ứng; chuyển đổi định dạng và nén ảnh ở **Hình ảnh đầu vào**. **Cú pháp Markdown** tách quy tắc xuất chung, quy tắc đọc đầu vào và ánh xạ dành riêng cho Word. Không phải mọi nguồn hay tối ưu chuyên biệt đều áp dụng mọi tùy chọn cú pháp.

**Công văn** và **Hóa đơn** có trang riêng. **Tài liệu → Markdown** đặt lựa chọn nhận dạng ban đầu cho công văn; **Markdown → Tài liệu** cấu hình định dạng trường mẫu, kèm trợ giúp về phần đính kèm, nơi nhận bản sao, ngày và ghi chú. Hóa đơn có lựa chọn riêng cho **PDF / OFD → Markdown** và **Ảnh → Markdown**. Các giá trị mặc định chỉ thay đổi lựa chọn ban đầu cho tệp tương thích, không tự phân loại tệp. **Áp dụng / OK** để lưu, **Hủy** để bỏ thay đổi chưa áp dụng.

**Phạm vi áp dụng:** Tính năng định dạng trường chuyên biệt hướng đến công văn Trung Quốc, còn tối ưu trích xuất hóa đơn hướng đến các mẫu hóa đơn Trung Quốc. Các tệp khác vẫn dùng được các đường chuyển đổi thông thường.

### Phím tắt

-   **Kéo file ngoài**: Kéo trực tiếp vào cửa sổ để nhập
-   **Mở đầu ra**: Nhấn thao tác "Mở đầu ra" ở bên phải thanh trạng thái để mở vị trí đầu ra.
-   **Chuột phải vào template**: Mở vị trí file template

---

## 🔧 Sử dụng CLI

Giao diện dòng lệnh (CLI) cho phép kiểm tra và chuyển đổi file từ terminal hoặc tự động hóa công việc cục bộ bằng script. Để tích hợp DocWen với ứng dụng khác, hãy dùng giao diện Machine Protocol bên dưới.

### Luồng gọi khuyến nghị cho tự động hóa

Với script tự động hóa cục bộ, có thể kiểm tra và thực hiện chuyển đổi theo thứ tự sau:

1. `inspect <file> [--json]`: trước tiên nhận diện loại tệp thực tế, định dạng và các thao tác được hỗ trợ.
2. `resources list formats --json`: đọc các tuyến thực sự khả dụng trong DocWen và các điều kiện phụ thuộc.
3. `schema convert`: đọc hợp đồng máy đọc được và các ràng buộc điều kiện của `convert`.
4. `convert <file> --to <fmt> (--output-dir <dir> | --output <path>) --dry-run --json`: xem trước quá trình nhận diện, chuẩn hóa và định tuyến mà không ghi tệp đầu ra.
5. `convert <file> --to <fmt> (--output-dir <dir> | --output <path>) ...`: sau khi xác nhận, mới chạy chuyển đổi thật.

Ứng dụng và plugin kết nối qua `serve --stdio` với Machine Protocol v2 để xem thao tác khả dụng, lập kế hoạch rồi thực hiện tác vụ. Đây là giao diện ổn định dành cho tích hợp; CLI `--json` dành cho script terminal. Dùng `gui open|activate|status` để mở file trong ứng dụng máy tính hoặc đưa cửa sổ ra phía trước.

### Ví dụ thường dùng

Chuyển đổi từ hoặc sang Markdown tạo thư mục kết quả bên trong `--output-dir DIR`, kể cả khi chỉ có một tệp. Tên gồm tên nguồn, dấu thời gian và định dạng đầu vào hiện tại. Chỉ dùng `--output PATH` cho các chuyển đổi khác tạo đúng một tệp đích.

```bash
# Bản đóng gói Windows
DocWenCLI.exe inspect document.docx --json

# Xuất hợp đồng convert cho script / agent
DocWenCLI.exe schema convert

# Xem trước cách chuyển đổi sẽ chạy mà không ghi kết quả
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --ocr --dry-run --json

# Xuất Word sang Markdown (trích ảnh + OCR)
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --ocr

# Markdown sang Word (mẫu + chế độ gộp tiêu đề/nội dung)
DocWenCLI.exe convert document.md --to docx --output-dir exports --template template.docx.da28ee624892975bc590fd419880875136f22e0edcd878bca69472e81297c0bc --heading-merge-mode punct_required

# Markdown sang Word và soát lỗi DOCX đã tạo trước khi xuất bản
DocWenCLI.exe convert document.md --to docx --output-dir exports --proofread --check typo --check punct

# Điều khiển chế độ ảnh và vị trí văn bản OCR trong Markdown
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --image-mode file --ocr --ocr-placement image_md

# Kiểm tra khả năng chạy và cổng phụ thuộc
DocWenCLI.exe doctor --json
DocWenCLI.exe resources list formats --json

# Soát lỗi tài liệu
DocWenCLI.exe validate document.docx --check typo --check punct
DocWenCLI.exe validate input.md --check typo --check punct

# Từ mã nguồn / uv
# inspect -> schema -> dry-run -> convert
# docwen inspect document.docx --json
# docwen schema convert
# docwen convert document.docx --to md --output-dir exports --dry-run --json
# docwen convert document.docx --to md --output-dir exports
```

### Lệnh và tùy chọn thông dụng

Bảng dưới đây chỉ liệt kê các lệnh thông dụng. Để xem đầy đủ toàn bộ lệnh, dùng `docwen --help` (mã nguồn / uv) hoặc `DocWenCLI --help` (bản đóng gói).

| Lệnh / tùy chọn | Mô tả |
| --- | --- |
| `convert <file> --to <fmt> (--output-dir <dir> \| --output <path>)` | Điểm vào thống nhất cho chuyển đổi. |
| `convert <markdown> --to docx --output-dir <dir> --proofread [--check ...]` | Soát lỗi DOCX được tạo sau Markdown→DOCX; `--check` trong `convert` yêu cầu `--proofread`. Nếu bỏ `--check`, các mặc định soát lỗi đã cấu hình sẽ được dùng. |
| `convert <file> --to <fmt> (--output-dir <dir> \| --output <path>) --dry-run --json` | Xem trước nhận diện, chuẩn hóa, định tuyến và các tùy chọn hiệu lực mà không thực hiện chuyển đổi thật. |
| `schema convert` | Xuất hợp đồng máy đọc được, giá trị mặc định, điều kiện và khóa chuẩn của `convert`. |
| `validate <file> --check ...` | Soát lỗi tài liệu (`typo/punct/symbol/sensitive/all/none`). Dùng `--json` cho envelope của CLI; `--report` là đường dẫn tệp báo cáo tùy chọn. |
| `inspect <file> [--json]` | Kiểm tra loại/định dạng tệp, hành động gợi ý và cảnh báo khi phần mở rộng không khớp nội dung. |
| `doctor --json` | Trả về chẩn đoán cùng với phần tóm tắt khả năng chạy và cổng phụ thuộc. |
| `resources list formats --json` | Liệt kê định dạng đích theo loại nguồn và kèm cổng phụ thuộc / bản tóm tắt giới hạn. |
| `resources list optimizations --json` | Liệt kê các tài nguyên tối ưu hóa có kiểu và liên kết tuyến chuẩn của chúng. |
| `resources list templates` | Liệt kê các mẫu có sẵn. |
| `resources list numbering-schemes` | Liệt kê các sơ đồ đánh số có sẵn. |
| `gui open [ABSOLUTE_FILE] [--timeout <seconds>] [--json]` | Khởi chạy hoặc kích hoạt DocWen qua đường điều khiển GUI cục bộ riêng và tùy chọn mở một tệp bằng đường dẫn tuyệt đối. |
| `--template <id>` | ID tài nguyên chuẩn chính xác từ `resources list templates`; tên hiển thị, tên tệp và đường dẫn đều bị từ chối. ID DOCX dùng cho `docx/doc/odt/rtf/wps/pdf`, ID XLSX cho `xlsx/xls/ods/csv`. |
| `--extract-img` / `--no-extract-img` / `--ocr` | Tùy chọn trích ảnh và OCR cho `convert --to md`. |
| `--image-mode file\|base64\|embed\|omit` | Kiểm soát cách ảnh được xuất ra khi xuất Markdown. |
| `--ocr-placement image_md\|main_md` | Kiểm soát việc ghi văn bản OCR vào Markdown phụ của ảnh hay Markdown chính. |
| `--heading-merge-mode punct_required\|always\|never` | Kiểm soát chiến lược gộp "tiêu đề + nội dung" cho `convert --to docx`. |
| `--optimization <id>` | Bật rõ ràng một hồ sơ tối ưu hóa (xem `resources list optimizations`). |
| `batch convert\|validate ... --jobs <n> [--continue-on-error]` | Điều khiển xử lý hàng loạt. |
| `--json` / `--quiet` / `--timing` | Xuất JSON, giảm nhật ký hoặc hiển thị thời gian chạy trong script terminal. |

Trong chế độ `punct_required`, danh sách mặc định chính xác là `。：！？.:!?`. Có thể chỉnh sửa danh sách này trong phần cài đặt định dạng; để trống sẽ tắt việc gộp trong chế độ này. Dấu phẩy, dấu chấm phẩy, dấu liệt kê, dấu gạch ngang và dấu chấm lửng không được bật mặc định.


## 📝 Quy ước Markdown

**Bạn dùng phần mở rộng Markdown của Obsidian?** Bốn phần mở rộng tùy chọn **mặc định tắt khi cài DocWen mới**: `structural_tables` (ô gộp và tiêu đề bảng nhiều hàng), `captions_references` (chú thích và tham chiếu Number Suite), `extended_headings` (H7–H9) và `typed_endnotes` (chú thích cuối có kiểu).

- **Markdown → Word:** vào **Cài đặt → Cú pháp Markdown → Đầu vào**, bật các phần mở rộng cần thiết hoặc chọn cấu hình **Phần mở rộng Obsidian** cho đầu vào. Nếu tắt, một số dấu đặc biệt có thể giữ nguyên thành văn bản.
- **Word → Markdown:** bật riêng **phần mở rộng đầu ra** tương ứng hoặc cấu hình đầu ra. Nếu không, một số cấu trúc có thể bị đơn giản hóa hay chuẩn hóa kèm cảnh báo.

Hãy giữ Markdown gốc và kiểm tra DOCX tạo ra. Đánh số trong Word do các tùy chọn chuyển đổi DocWen quyết định; không đảm bảo phục hồi đúng từng dấu hay khoảng trắng ban đầu. Xem [tài liệu tương thích Markdown](../specs/markdown-compatibility.md).

### Ánh xạ cấp tiêu đề

Để dễ ghi nhớ, tiêu đề Markdown trong phần mềm tương ứng **1:1** với tiêu đề Word:
- Tiêu đề và phụ đề của tài liệu đặt trong YAML metadata.
- Markdown `# Heading 1` tương ứng Word "Heading 1".
- Markdown `## Heading 2` tương ứng Word "Heading 2".
- Và tiếp tục như vậy, hỗ trợ tối đa 9 cấp.

**Mẹo**: Nếu bạn muốn dùng `#` làm tiêu đề tài liệu và dùng `##` trở đi cho tiêu đề nội dung, hãy chỉnh style "Heading 1" trong template Word để trông giống tiêu đề (ví dụ: căn giữa, in đậm, cỡ chữ lớn) và chọn scheme đánh số bỏ qua cấp 1 trong phần cài đặt.

### Xuống dòng và đoạn văn

**Quy tắc cơ bản**: Xuống dòng mềm trong Markdown thông thường mặc định **không** tạo đoạn Word riêng. Bật **Mỗi dòng là một đoạn** tại **Cài đặt → Cú pháp Markdown** nếu muốn tách các dòng văn bản thường. Danh sách, trích dẫn, khối mã và ngắt dòng rõ ràng vẫn theo quy tắc riêng.

**Đoạn trộn**: Khi một tiêu đề phụ cần trộn với nội dung trong cùng một đoạn (mặc định: chế độ "Cần dấu kết câu"), phải thỏa các điều kiện:
1.  Tiêu đề phụ kết thúc bằng dấu kết câu (hỗ trợ dấu kết câu đa ngôn ngữ).
2.  Nội dung nằm ở **dòng ngay bên dưới** tiêu đề phụ.
3.  Dòng nội dung không được là phần tử Markdown đặc biệt (tiêu đề, code block, bảng, danh sách, trích dẫn, khối công thức, dấu phân cách, ...).

**Ví dụ**:
```markdown
## I. Yêu cầu công việc.
Cuộc họp yêu cầu các đơn vị nghiêm túc triển khai...
```
Hai dòng trên sẽ được gộp thành một đoạn: "I. Yêu cầu công việc." giữ định dạng tiêu đề phụ và "Cuộc họp..." giữ định dạng nội dung.

**Lưu ý**:
- Không được có dòng trống giữa tiêu đề phụ và nội dung; nếu có sẽ bị nhận diện thành hai đoạn riêng.
- Theo mặc định (chế độ "Cần dấu kết câu"), nếu tiêu đề phụ không kết thúc bằng dấu kết câu, sẽ không gộp với dòng kế tiếp dù không có dòng trống.
- Đổi chế độ và các ký tự kích hoạt gộp tại **Cài đặt → Văn bản đầu vào → Markdown → Tài liệu — Xử lý định dạng → Chế độ gộp tiêu đề và nội dung**.

### Chuyển đổi dấu phân cách hai chiều

Hỗ trợ chuyển đổi hai chiều giữa dấu phân cách Markdown và ngắt trang/ngắt mục/dòng kẻ trong Word:

-   **DOCX → MD**: Tự động chuyển ngắt trang, ngắt mục và dòng kẻ của Word thành dấu phân cách Markdown.
-   **MD → DOCX**: Tự động chuyển `---`, `***`, `___` thành phần tử Word tương ứng.
-   **Có thể cấu hình**: Quan hệ ánh xạ có thể tuỳ chỉnh trong phần cài đặt.

### Danh sách công việc

Hỗ trợ chuyển đổi hai chiều danh sách công việc GFM:

```markdown
- [ ] Cần làm
- [x] Hoàn thành
```

-   **MD → DOCX**: Hiển thị dưới dạng danh sách có dấu đầu dòng với tiền tố `☐` / `☑`.
-   **DOCX → MD**: Chuyển đổi các mục danh sách bắt đầu bằng `☐` / `☑` / `☒` thành `- [ ]` / `- [x]`.
-   **Lưu ý về phông chữ**: `☐`/`☑` có thể không hiển thị trên một số phông chữ. Nếu cần, hãy sử dụng phông chữ như "Segoe UI Symbol" trong mẫu Word.

### Nhúng ảnh và kích thước

Hỗ trợ nhúng ảnh kiểu Obsidian/Wiki và Markdown chuẩn, có thể chỉ định kích thước (px):

```markdown
![[image.png]]
![[image.png|300]]
![[image.png\|300]]
![alt](image.png =300x200)
![alt](image.png =300x)
![alt|300](image.png)
```

- Không chỉ định: dùng kích thước gốc, giới hạn bởi chiều rộng khả dụng (trang/ô)
- Có chỉ định: cho phép phóng to, nhưng vẫn giới hạn bởi chiều rộng khả dụng
- Đoạn chỉ có ảnh: dùng style đoạn “Image” (căn giữa, giãn dòng đơn)

### Xử lý liên kết

Hỗ trợ liên kết có thể bấm trong Markdown -> DOCX:

```markdown
[Docwen](https://example.com)
[[Target]]
[[Target|Open target]]
<https://example.com>
<user@example.com>
```

- Liên kết Markdown và Wiki mặc định được ghi thành siêu liên kết Word
- Liên kết Wiki được phân giải thành liên kết cục bộ `file:///` khi tìm thấy tệp đích
- Liên kết tự động trong dấu `< >` hỗ trợ `https://...` và email `mailto:...`
- Tự động liên kết URL trần được áp dụng theo từng yêu cầu Markdown -> DOCX, mặc định tắt và được bật bằng `[non_embed_links].auto_link_bare_url` trong `configs/link.toml`
- Markdown -> XLSX không tạo placeholder siêu liên kết cho DOCX và giữ nguyên cú pháp liên kết gốc

## 📖 Hướng dẫn sử dụng chi tiết

### Word sang Markdown

1.  Kéo file `.docx` vào cửa sổ chương trình.
2. Chọn các tùy chọn xuất rồi nhấn **Xuất Markdown**.
3.  Tạo file `.md` có chứa YAML metadata.

**Định dạng hỗ trợ**:
-   `.docx` - Tài liệu Word chuẩn.
-   `.doc` - Tự chuyển sang DOCX để xử lý.
-   `.wps` - Tự chuyển tài liệu WPS để xử lý.

**Tuỳ chọn xuất**:

| Tuỳ chọn | Mô tả |
| :--- | :--- |
| **Trích xuất ảnh** | Nếu bật, ảnh trong tài liệu sẽ được xuất ra thư mục và chèn link ảnh vào file MD. |
| **OCR ảnh** | Nhận dạng chữ trong ảnh. Tùy cài đặt vị trí đầu ra OCR, lưu vào file Markdown riêng cho mỗi ảnh hoặc file Markdown chính. |
| **Tối ưu theo loại tài liệu** | Áp dụng tối ưu cho loại tài liệu đã chọn, chẳng hạn trích xuất trường từ công văn tiếng Trung. Tắt tùy chọn này để dùng chuyển đổi tiêu chuẩn. |
| **Xóa số tiêu đề gốc** | Nếu bật, xoá số trước tiêu đề phụ (ví dụ: "一、", "（一）", "1.", ...). |
| **Thêm số tiêu đề mới** | Nếu bật, tự thêm số theo cấp tiêu đề (có thể cấu hình). |

Số tiêu đề từ danh sách đa cấp của Word/WPS được giữ lại khi xuất. Bật **Xóa số tiêu đề gốc** để xóa số nhưng vẫn giữ nguyên cấp tiêu đề.

### Markdown sang Word

1.  Chuẩn bị file `.md` có YAML header.
2.  Kéo vào cửa sổ và chọn template Word tương ứng.
3. Chọn DOCX làm định dạng đích rồi nhấn **Tạo DOCX** để điền mẫu và lưu tài liệu.

**Tuỳ chọn chuyển đổi**:

| Tuỳ chọn | Mô tả |
| :--- | :--- |
| **Xóa số tiêu đề gốc** | Nếu bật, xoá số trước tiêu đề phụ. |
| **Thêm số tiêu đề mới** | Nếu bật, tự thêm số theo cấp tiêu đề. |

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
    -   Tên style quản lý được tạo mới theo **ngôn ngữ đầu ra tác vụ** trong cài đặt chung, không theo ngôn ngữ giao diện. Đổi giao diện không đổi tên style có sẵn trong mẫu.

**Gợi ý**: Sau khi bạn tuỳ chỉnh style trong template, chương trình sẽ ưu tiên dùng style đó; nếu template không có thì dùng preset mặc định.

### Xử lý file bảng tính

1. **Excel/CSV sang Markdown**: Thêm file `.xlsx` hoặc `.csv` rồi nhấn **Xuất Markdown**.
2.  **Markdown sang Excel**: Bảng Markdown có thể xuất sang XLSX. Template hỗ trợ trường YAML, placeholder cột và ảnh, cùng ô gộp hoặc ô được bảo vệ.

**Định dạng hỗ trợ**:
-   `.xlsx` - Excel chuẩn.
-   `.xls` - Tự chuyển sang XLSX để xử lý.
-   `.et` - Tự chuyển bảng tính WPS để xử lý.
-   `.csv` - Bảng văn bản CSV.
-   `.tsv` - Bảng phân tách bằng tab TSV.


### Chức năng kiểm tra lỗi văn bản

Chương trình cung cấp 4 quy tắc kiểm tra có thể tuỳ chỉnh:

1.  **Kiểm tra cặp dấu câu** - Kiểm tra ngoặc kép, ngoặc đơn, ... có khớp cặp không.
2.  **Kiểm tra ký hiệu** - Phát hiện dùng lẫn dấu câu tiếng Trung/tiếng Anh.
3.  **Kiểm tra lỗi chính tả** - Dựa trên từ điển tuỳ chỉnh.
4.  **Phát hiện từ nhạy cảm** - Dựa trên từ điển tuỳ chỉnh.

**Từ điển tuỳ chỉnh**: Chỉnh sửa trực quan từ điển lỗi chính tả và từ nhạy cảm trong "Cài đặt".

**Cách dùng**:
1.  Kéo tài liệu Word hoặc file Markdown cần kiểm tra vào chương trình.
2.  Chọn các quy tắc cần dùng.
3.  Nhấn nút "Kiểm tra".
4.  Kết quả hiển thị dưới dạng comment trong tài liệu. Đối với file Markdown, kết quả được xuất dưới dạng báo cáo JSON.

Ghi chú (báo cáo JSON khi kiểm tra Markdown):
- Engine: `text_rules` + adapter Markdown `md_spell`
- Đầu ra: lối vào soát lỗi hiện tại của CLI là `validate`; dùng `--json` cho envelope của CLI. `--report` là đường dẫn tệp báo cáo tùy chọn.

- Khác với `--json` (lớp bao JSON của CLI)

## 🛠️ Hệ thống template

### Dùng template có sẵn

Quản lý mẫu DOCX/XLSX trong **Cài đặt → Mẫu**: bật, tắt, sắp xếp và chọn mặc định. Mẫu tích hợp chỉ được đọc; mẫu tùy chỉnh nằm trong thư mục người dùng có quyền ghi. Với MSIX, không cần sửa WindowsApps.

### Template tuỳ chỉnh

1. Sao chép và chỉnh sửa mẫu tích hợp, hoặc nhập tệp DOCX/XLSX có sẵn.
2. Sửa nội dung, kiểu và chỗ giữ chỗ của bản sao bằng Word, Excel hoặc WPS; xem quy tắc bên dưới.
3. Lưu tệp rồi làm mới danh sách trong DocWen. Không cần khởi động lại.

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

Template XLSX hỗ trợ trường YAML, placeholder cột dọc `{{↓Field}}` và ngang `{{→Field}}`, placeholder ảnh, cùng ô gộp hoặc ô được bảo vệ.

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

**Xử lý ô gộp**:

- Markdown -> Excel tiếp tục giữ nguyên các merged ranges có sẵn của mẫu.
- Với các vùng mẫu dạng cột đã biết, được tạo bởi các placeholder `{{↓Field Name}}` liên tiếp, chương trình có thể khôi phục gộp hình chữ nhật từ dấu gộp `<` / `^` tường minh trong bảng Markdown.
- Chỉ những ô có nội dung sau khi bỏ khoảng trắng đầu/cuối chính xác là `<` hoặc `^` mới tham gia nhận diện gộp; `\<` và `\^` được giữ lại như văn bản literal.
- Hình chữ nhật không hợp lệ hoặc xung đột với merged ranges có sẵn của mẫu sẽ bị hạ cấp thành văn bản thường kèm cảnh báo, thay vì cưỡng ép ghi đè cấu trúc mẫu.

**Gộp dữ liệu nhiều bảng**: Nếu Markdown có nhiều bảng dùng cùng tiêu đề cột, dữ liệu sẽ được gộp theo thứ tự và điền liên tục.

## 🔌 Plugin Obsidian

Plugin Obsidian đồng hành được phát hành ở repo riêng và hoạt động cùng bộ chuyển đổi:

### Tính năng cốt lõi

-   **🚀 Khởi chạy 1 lần nhấn** - Icon ở sidebar để mở nhanh bộ chuyển đổi.
-   **📂 Bàn giao tự động** - Tự truyền đường dẫn file đang mở.
-   **🔄 Dùng cửa sổ đang mở** - Mở file khác trong bộ chuyển đổi mà không cần khởi động lại.

### Cách dùng plugin

DocWen Assistant 3.3 hỗ trợ máy tính để bàn Windows và Linux. Bản Microsoft Store trên Windows có thể dùng tự động phát hiện; trên Linux người dùng chọn thủ công gói tương thích đã giải nén. Assistant hiện chưa hỗ trợ macOS.

1.  **Nhấn lần đầu** → Khởi chạy bộ chuyển đổi và mở file hiện tại.
2.  **Nhấn lại khi có file đang mở** → Thay file đầu vào hiện tại bằng file đó (chế độ một file).
3.  **Nhấn lại khi không có file đang mở** → Đưa cửa sổ bộ chuyển đổi lên trước.

### Cài đặt

Cài một phiên bản DocWen tương thích trước, sau đó cài [DocWen Assistant](https://github.com/ZHYX91/obsidian-docwen-assistant). Assistant 3.3 yêu cầu DocWen 0.17.0 trở lên; plugin không tự cài chương trình DocWen. Xem ghi chú phát hành để chọn các phiên bản tương thích và tải các file cài đặt đã được công bố.

## 🔌 OpenClaw (Plugin + Skill)

[docwen-openclaw](https://github.com/ZHYX91/docwen-openclaw) cho phép OpenClaw dùng DocWen trên máy tính của bạn để chuyển đổi và soát lỗi file. Cài DocWen riêng và chọn phiên bản plugin có ghi rõ phiên bản DocWen tương thích trong ghi chú phát hành. Tải file từ trang phát hành đã công bố; số phiên bản trong mã nguồn không có nghĩa là phiên bản đó đã có sẵn để tải xuống.

Để phát triển tính năng tích hợp với plugin 3.0, xem [Machine Protocol v2 và Artifact Bundle v3](../specs/machine-protocol-v2.md).

## ❓ Câu hỏi thường gặp (FAQ)

### Nếu chuyển đổi thất bại thì sao?

-   Kiểm tra file có đang bị ứng dụng khác mở/chiếm dụng không.
-   Xác nhận định dạng file đúng.
-   Xem mục "Đường dẫn file log thực tế hiện tại" trong phần cài đặt, hoặc kiểm tra lỗi trong thư mục log người dùng của hệ thống; nếu bước kiểm thử gói dùng `DOCWEN_LOG_DIR` thì hãy kiểm tra thư mục bị ghi đè đó.

### Template không hiển thị?

- Trong **Cài đặt → Mẫu**, kiểm tra mẫu đã bật và phù hợp với định dạng đầu ra.
- Kiểm tra tệp DOCX/XLSX có mở bình thường không.
- Làm mới sau khi sửa bên ngoài; nhập các tệp mới.

### Chức năng kiểm tra lỗi không hoạt động?

-   Xác nhận tài liệu là `.docx` hoặc `.md`.
-   Kiểm tra tài liệu có chứa text có thể chỉnh sửa không.
-   Xác nhận các quy tắc kiểm tra đã bật trong cài đặt.

### Định dạng đầu ra không như mong đợi?

- Mở **Cài đặt → Mẫu**. Sao chép mẫu tích hợp trước khi sửa; có thể sửa trực tiếp mẫu tùy chỉnh.
- Điều chỉnh kiểu trong Word, Excel hoặc WPS, lưu rồi làm mới danh sách.
- Chọn mẫu tùy chỉnh trước khi tạo tệp mới.

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

## 🔒 Bảo mật và quyền riêng tư

DocWen xử lý file trên máy tính của bạn và không cung cấp tính năng tải file lên, thu thập dữ liệu đo lường, tải xuống trực tuyến hay dịch vụ mạng. Chương trình hạn chế truy cập mạng của các thành phần Python trong chính DocWen. Nếu không bật được lớp bảo vệ này, DocWen sẽ không tiếp tục chạy.

Các chương trình được khởi chạy riêng như Word/Office, WPS và LibreOffice, kể cả tiến trình hỗ trợ, vẫn tuân theo cài đặt mạng của từng chương trình và hệ thống. Hạn chế này không áp dụng cho chúng. Khi xử lý file nhạy cảm, hãy kiểm tra cả các cài đặt đó.

Xem [tài liệu kiến trúc](../architecture.md) để biết cách bảo vệ hoạt động và giới hạn của nó. Nếu chương trình không khởi động hoặc chuyển đổi thất bại, hãy xem [hướng dẫn khắc phục sự cố](../maintenance/troubleshooting.md).

## 📜 Giấy phép

Dự án dùng giấy phép **GNU Affero General Public License v3.0 (AGPL-3.0)**.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   Dự án sử dụng PyMuPDF (AGPL-3.0), nên toàn bộ dự án cũng theo AGPL-3.0.
- GUI hiện tại có thể sử dụng `PySide6-Fluent-Widgets` (QFluentWidgets) trên các đường host được hỗ trợ; phụ thuộc này dùng mô hình cấp phép kép `GPLv3 / thương mại`, còn kho này vẫn tiếp tục được phân phối theo AGPL.
-   Bạn có thể tự do sử dụng, sửa đổi và phân phối phần mềm.
-   Nếu bạn sửa phần mềm và cung cấp dịch vụ qua mạng, bạn phải cung cấp mã nguồn đã sửa cho người dùng.
-   Xem thêm trong [LICENSE](../../LICENSE).
- Thong bao ve thanh phan ben thu ba nam trong [LICENSE_THIRD_PARTY.txt](../../LICENSE_THIRD_PARTY.txt); tom tat phan phoi nam trong [NOTICE.txt](../../NOTICE.txt).

### Liên hệ

-   **GitHub**: https://github.com/ZHYX91/docwen
-   **Email**: zhengyx91@hotmail.com

---

**Tác giả**: ZhengYX
