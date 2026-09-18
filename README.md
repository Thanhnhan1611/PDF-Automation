# Tự động tạo phiếu lương PDF

Python automation system for reading employee salary data from Excel, calculating salary information, generating employee salary slips as PDF files, and recording the processing results in a log file.

## Giới thiệu

Đây là bài tập A3 với đề tài **Tự động tạo báo cáo PDF**, trong đó hệ thống được xây dựng để tự động tạo **phiếu lương cho nhân viên** từ dữ liệu trong file Excel.

Thay vì tạo từng phiếu lương thủ công, chương trình thực hiện toàn bộ quy trình:

```text
employees.xlsx
      │
      ▼
Đọc dữ liệu bằng Pandas
      │
      ▼
Kiểm tra dữ liệu
      │
      ├── Dữ liệu lỗi ──► Bỏ qua + ghi Log
      │
      ▼
Tính toán tiền lương
      │
      ▼
Tạo HTML/CSS
      │
      ▼
WeasyPrint
      │
      ▼
salary/[Mã NV].pdf
      │
      ▼
process_log.txt
```

## Mục tiêu

* Tự động hóa quá trình tạo phiếu lương.
* Đọc dữ liệu nhân viên từ Excel bằng Pandas.
* Kiểm tra và xử lý dữ liệu không hợp lệ.
* Tính toán các thành phần của bảng lương.
* Tạo PDF động dựa trên dữ liệu từng nhân viên.
* Ghi nhận kết quả xử lý bằng hệ thống logging.
* Không làm dừng toàn bộ chương trình khi một dòng dữ liệu bị lỗi.

## Chức năng chính

* Đọc file `employees.xlsx`.
* Kiểm tra các cột bắt buộc.
* Kiểm tra dữ liệu từng nhân viên.
* Tính lương theo ngày công.
* Tính BHXH, BHYT và BHTN.
* Tính thu nhập tính thuế và thuế TNCN theo mô hình của bài.
* Tính tổng khấu trừ và thực lãnh.
* Chuyển số tiền thành chữ tiếng Việt.
* Tạo phiếu lương bằng HTML/CSS.
* Chuyển HTML thành PDF bằng WeasyPrint.
* Tự động tạo thư mục `salary/`.
* Đặt tên PDF theo mã nhân viên.
* Ghi kết quả thành công/thất bại vào `process_log.txt`.

## Công nghệ sử dụng

| Công nghệ | Mục đích |
|---|---|
| Python | Ngôn ngữ lập trình chính |
| Pandas | Đọc và xử lý dữ liệu Excel |
| OpenPyXL | Đọc file `.xlsx` |
| WeasyPrint | Chuyển HTML/CSS thành PDF |
| HTML | Xây dựng nội dung phiếu lương |
| CSS | Định dạng phiếu lương |
| Logging | Ghi nhật ký quá trình xử lý |

## Cấu trúc project

```text
PDF-Automation/
│
├── employees.xlsx
├── salary_report.py
├── README.md
├── process_log.txt
│
└── salary/
    ├── NV001.pdf
    ├── NV002.pdf
    └── NV003.pdf
```

`salary/` và `process_log.txt` được chương trình tự động tạo trong quá trình chạy.

## Dữ liệu đầu vào

File đầu vào:

```text
employees.xlsx
```

Các trường chính được sử dụng:

| Cột | Ý nghĩa |
|---|---|
| `Mã NV` | Mã nhân viên |
| `Tên NV` | Họ và tên |
| `Chức vụ` | Chức vụ |
| `Lương cơ bản` | Mức lương cơ bản |
| `Lương đóng BHXH` | Mức lương dùng để tính bảo hiểm |
| `Ngày công chuẩn` | Số ngày công chuẩn trong tháng |
| `Ngày công đi làm` | Số ngày nhân viên thực tế đi làm |
| `Thưởng` | Khoản thưởng |
| `Phụ cấp` | Khoản phụ cấp |
| `Số người phụ thuộc` | Số người phụ thuộc |
| `Phạt` | Khoản phạt |

Một số trường như `Lương đóng BHXH`, `Ngày công chuẩn`, `Ngày công đi làm`, `Thưởng`, `Phụ cấp` và `Số người phụ thuộc` có giá trị mặc định trong chương trình khi không được cung cấp.

## Mô hình tính lương

### Lương theo ngày

```text
Lương theo ngày =
Lương cơ bản / Ngày công chuẩn
```

### Lương thực tế

```text
Lương thực tế =
Lương theo ngày × Ngày công đi làm
```

### Tổng thu nhập

```text
Tổng thu nhập =
Lương thực tế + Thưởng + Phụ cấp
```

### Các khoản bảo hiểm

```text
BHXH = Lương đóng BHXH × 8%

BHYT = Lương đóng BHXH × 1,5%

BHTN = Lương đóng BHXH × 1%
```

### Thu nhập tính thuế

Thu nhập tính thuế được xác định dựa trên tổng thu nhập, các khoản bảo hiểm và mức giảm trừ được cấu hình trong chương trình.

### Thực lãnh

```text
Thực lãnh =
Tổng thu nhập - Tổng khấu trừ
```

Chương trình cũng kiểm tra dữ liệu ngày công, trong đó `Ngày công đi làm` không được lớn hơn `Ngày công chuẩn`.

## Tạo phiếu lương PDF

Mỗi nhân viên hợp lệ sẽ được tạo một file PDF riêng.

Thư mục đầu ra:

```text
salary/
```

Tên file:

```text
[Mã NV].pdf
```

Ví dụ:

```text
salary/
├── NV001.pdf
├── NV002.pdf
└── NV003.pdf
```

Phiếu lương chứa các thông tin:

* Thông tin công ty.
* Mã nhân viên.
* Họ và tên.
* Chức vụ.
* Người phụ thuộc.
* Ngày công chuẩn.
* Ngày công đi làm.
* Lương cơ bản.
* Lương đóng BHXH.
* Lương thực tế.
* Thưởng.
* Phụ cấp.
* BHXH.
* BHYT.
* BHTN.
* Thu nhập tính thuế.
* Thuế TNCN.
* Phạt.
* Tổng khấu trừ.
* Thực lãnh.
* Số tiền bằng chữ.

## Xử lý lỗi và Logging

Chương trình sử dụng `try...except` để xử lý lỗi trong quá trình đọc dữ liệu và tạo PDF.

Nếu một dòng dữ liệu không hợp lệ, chương trình sẽ:

```text
Dòng dữ liệu lỗi
      │
      ▼
try...except
      │
      ▼
Ghi lỗi vào process_log.txt
      │
      ▼
Bỏ qua dòng lỗi
      │
      ▼
Tiếp tục xử lý nhân viên tiếp theo
```

File log:

```text
process_log.txt
```

## Cài đặt

Khuyến nghị sử dụng:

```text
Python 3.12
```

Cài đặt các thư viện:

```cmd
pip install pandas openpyxl weasyprint
```

### Windows

Nếu WeasyPrint gặp lỗi liên quan đến DLL như `libgobject-2.0-0`, cần cài đặt thư viện GTK/MSYS2 phù hợp và cấu hình thư mục DLL.

Trong chương trình hiện tại, đường dẫn Windows được kiểm tra là:

```text
C:\msys64\ucrt64\bin
```

## Chạy chương trình

Đặt `employees.xlsx` và `salary_report.py` trong cùng thư mục:

```text
PDF-Automation/
├── employees.xlsx
└── salary_report.py
```

Chạy chương trình:

```cmd
python salary_report.py
```

Sau khi chạy xong, các file PDF được lưu trong `salary/` và log được lưu trong `process_log.txt`.

## Quy trình hoạt động

```text
1. Kiểm tra employees.xlsx
          │
          ▼
2. Đọc Excel bằng Pandas
          │
          ▼
3. Kiểm tra các cột bắt buộc
          │
          ▼
4. Xử lý từng nhân viên
          │
          ├── Không hợp lệ ──► Ghi lỗi + bỏ qua
          │
          ▼
5. Tính toán bảng lương
          │
          ▼
6. Tạo HTML/CSS
          │
          ▼
7. WeasyPrint tạo PDF
          │
          ▼
8. Kiểm tra file PDF
          │
          ▼
9. Lưu salary/[Mã NV].pdf
          │
          ▼
10. Ghi kết quả vào process_log.txt
```

## Project Status

Current version:

```text
V1.0 — Automated Employee Salary PDF Generation
```

Project tập trung vào việc tự động hóa quá trình tạo phiếu lương PDF từ dữ liệu Excel, có kiểm tra lỗi và ghi log trong quá trình xử lý.

## Project Information

| Thông tin | Nội dung |
|---|---|
| Project | Automated PDF Salary Report Generation |
| Loại bài | A3 - Tự động tạo báo cáo PDF |
| Ngôn ngữ | Python |
| Input | `employees.xlsx` |
| Output | `salary/[Mã NV].pdf` |
| Log | `process_log.txt` |
