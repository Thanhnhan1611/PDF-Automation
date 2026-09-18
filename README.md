# Automated PDF Salary Report Generation

## 1. Giới thiệu

Đề tài **Automated PDF Salary Report Generation** xây dựng một hệ thống
Python nhằm tự động hóa quy trình tạo phiếu lương nhân viên dưới dạng
file PDF từ dữ liệu được lưu trữ trong Excel.

Thay vì thực hiện thủ công từng phiếu lương, hệ thống cho phép đọc dữ
liệu của nhiều nhân viên, kiểm tra dữ liệu đầu vào, thực hiện các phép
tính liên quan đến tiền lương và tự động tạo một file PDF tương ứng
cho từng nhân viên.

Bên cạnh chức năng tạo báo cáo, hệ thống còn tích hợp cơ chế xử lý
ngoại lệ và ghi log nhằm đảm bảo chương trình có thể tiếp tục xử lý
các dữ liệu hợp lệ ngay cả khi một số bản ghi xảy ra lỗi.

---

## 2. Mục tiêu

Mục tiêu của đề tài là xây dựng một quy trình tự động có khả năng:

- Đọc dữ liệu nhân viên từ file Excel.
- Kiểm tra tính hợp lệ của dữ liệu đầu vào.
- Tính toán các thành phần của tiền lương.
- Tính các khoản khấu trừ.
- Tính tiền thực lãnh.
- Tự động tạo phiếu lương dưới dạng PDF.
- Tạo riêng một file PDF cho từng nhân viên.
- Tự động tạo thư mục lưu các file PDF.
- Ghi nhận quá trình xử lý vào file log.
- Xử lý dữ liệu lỗi mà không làm dừng toàn bộ chương trình.

---

## 3. Chức năng chính

### 3.1. Đọc dữ liệu Excel

Hệ thống sử dụng file:

```text
employees.xlsx

để làm nguồn dữ liệu đầu vào.

Dữ liệu được đọc và xử lý bằng Python thông qua Pandas và OpenPyXL.

3.2. Kiểm tra dữ liệu

Trước khi tính toán, chương trình kiểm tra các dữ liệu cần thiết như:

Mã nhân viên.
Tên nhân viên.
Chức vụ.
Lương cơ bản.
Ngày công.
Thưởng.
Phụ cấp.
Các khoản khấu trừ.

Nếu dữ liệu không hợp lệ, bản ghi tương ứng sẽ được ghi nhận vào
file log.

3.3. Tính toán tiền lương

Hệ thống thực hiện tính toán dựa trên các thông tin về:

Lương cơ bản.
Ngày công chuẩn.
Ngày công thực tế.
Ngày nghỉ phép hưởng lương.
Ngày nghỉ không lương.
Thưởng.
Phụ cấp.
Bảo hiểm.
Thuế.
Các khoản phạt hoặc khấu trừ khác.
3.4. Tạo PDF tự động

Sau khi hoàn thành quá trình tính toán, dữ liệu được đưa vào mẫu
HTML/CSS và chuyển đổi thành file PDF.

Mỗi nhân viên được tạo một file riêng theo mã nhân viên.

Ví dụ:

NV001.pdf
NV002.pdf
NV003.pdf
3.5. Logging và xử lý lỗi

Chương trình sử dụng try...except để xử lý các ngoại lệ trong quá
trình thực thi.

Kết quả xử lý được ghi vào:

process_log.txt

Nhờ đó, nếu một nhân viên có dữ liệu không hợp lệ, chương trình có
thể ghi nhận lỗi và tiếp tục xử lý các nhân viên khác.

4. Công nghệ sử dụng
Công nghệ	Vai trò
Python	Ngôn ngữ lập trình chính
Pandas	Đọc và xử lý dữ liệu
OpenPyXL	Xử lý file Excel
WeasyPrint	Chuyển đổi HTML/CSS sang PDF
HTML	Xây dựng cấu trúc phiếu lương
CSS	Thiết kế giao diện phiếu lương
5. Cấu trúc project
salary-report/
│
├── employees.xlsx
├── salary_report.py
├── requirements.txt
├── README.md
├── process_log.txt
│
└── salary/
    ├── NV001.pdf
    ├── NV002.pdf
    ├── NV003.pdf
    └── ...
Mô tả các thành phần
Thành phần	Mô tả
employees.xlsx	Dữ liệu nhân viên đầu vào
salary_report.py	Chương trình Python chính
requirements.txt	Danh sách thư viện cần cài đặt
README.md	Tài liệu mô tả và hướng dẫn project
process_log.txt	Nhật ký quá trình xử lý
salary/	Thư mục chứa các file PDF đầu ra
6. Quy trình hoạt động

Hệ thống được xây dựng theo pipeline:

                employees.xlsx
                       │
                       ▼
              ┌─────────────────┐
              │ Đọc dữ liệu      │
              │ Excel            │
              └────────┬────────┘
                       │
                       ▼
              ┌─────────────────┐
              │ Kiểm tra dữ liệu│
              └────────┬────────┘
                       │
                       ▼
              ┌─────────────────┐
              │ Tính toán        │
              │ tiền lương       │
              └────────┬────────┘
                       │
                       ▼
              ┌─────────────────┐
              │ HTML / CSS       │
              │ Salary Template  │
              └────────┬────────┘
                       │
                       ▼
              ┌─────────────────┐
              │ WeasyPrint       │
              │ PDF Generation   │
              └────────┬────────┘
                       │
              ┌────────┴─────────┐
              ▼                  ▼
        salary/*.pdf       process_log.txt
7. Mô hình tính lương
7.1. Ngày công tính lương

Hệ thống phân biệt giữa ngày đi làm, ngày nghỉ phép hưởng lương
và ngày nghỉ không lương.

Công thức:

Ngày công tính lương
= Ngày công đi làm + Ngày nghỉ phép hưởng lương

Ví dụ:

Ngày công chuẩn:              26
Ngày công đi làm:             25
Ngày nghỉ phép hưởng lương:    1
Ngày nghỉ không lương:         0

Ngày công tính lương = 25 + 1 = 26

Trường hợp nghỉ không lương:

Ngày công chuẩn:              26
Ngày công đi làm:             25
Ngày nghỉ phép hưởng lương:    0
Ngày nghỉ không lương:         1

Ngày công tính lương = 25
7.2. Lương theo ngày
Lương theo ngày
= Lương cơ bản / Ngày công chuẩn
7.3. Lương thực tế
Lương thực tế
= Lương theo ngày × Ngày công tính lương
7.4. Tổng thu nhập
Tổng thu nhập
= Lương thực tế + Thưởng + Phụ cấp
7.5. Tổng khấu trừ

Tổng khấu trừ bao gồm các khoản được cấu hình trong hệ thống, chẳng hạn:

BHXH.
BHYT.
BHTN.
Thuế TNCN.
Phạt hoặc các khoản khấu trừ khác.
7.6. Thực lãnh
Thực lãnh
= Tổng thu nhập - Tổng khấu trừ
8. Cài đặt
8.1. Yêu cầu hệ thống

Project yêu cầu:

Python 3.x
pip
Các thư viện được khai báo trong requirements.txt

Kiểm tra phiên bản Python:

python --version

Kiểm tra pip:

pip --version
8.2. Cài đặt thư viện

Tại thư mục project, chạy:

pip install -r requirements.txt

Hoặc:

pip install pandas openpyxl weasyprint
9. Hướng dẫn sử dụng
Bước 1: Chuẩn bị dữ liệu

Đặt file:

employees.xlsx

vào thư mục gốc của project.

Ví dụ:

salary-report/
├── employees.xlsx
└── salary_report.py
Bước 2: Chạy chương trình

Mở Terminal hoặc Command Prompt tại thư mục project và chạy:

python salary_report.py
Bước 3: Kiểm tra kết quả

Sau khi chương trình hoàn thành, các file PDF được tạo trong:

salary/

Ví dụ:

salary/
├── NV001.pdf
├── NV002.pdf
└── NV003.pdf

Kết quả xử lý được ghi trong:

process_log.txt
10. Dữ liệu đầu vào

File employees.xlsx chứa thông tin phục vụ cho quá trình tính lương.

Một số trường dữ liệu chính:

Tên cột	Ý nghĩa
Mã NV	Mã nhân viên
Tên NV	Tên nhân viên
Chức vụ	Chức vụ
Lương cơ bản	Mức lương cơ bản
Lương đóng BHXH	Mức lương dùng cho tính bảo hiểm
Ngày công chuẩn	Số ngày công chuẩn
Ngày công đi làm	Số ngày thực tế đi làm
Ngày nghỉ phép hưởng lương	Số ngày nghỉ vẫn hưởng lương
Ngày nghỉ không lương	Số ngày nghỉ không hưởng lương
Thưởng	Khoản tiền thưởng
Phụ cấp	Khoản phụ cấp
Số người phụ thuộc	Số người phụ thuộc
Phạt	Khoản phạt hoặc khấu trừ
11. Kết quả đầu ra

Hệ thống tạo một file PDF cho mỗi nhân viên.

Ví dụ:

salary/
├── NV001.pdf
├── NV002.pdf
├── NV003.pdf
└── NV004.pdf

Mỗi phiếu lương PDF bao gồm các thông tin như:

Thông tin công ty.
Mã nhân viên.
Họ tên nhân viên.
Chức vụ.
Kỳ lương.
Ngày công.
Lương cơ bản.
Lương thực tế.
Thưởng.
Phụ cấp.
Các khoản khấu trừ.
Thuế.
Thực lãnh.
Số tiền bằng chữ.
Khu vực xác nhận/chữ ký.
12. Xử lý lỗi và Logging

Hệ thống được thiết kế để xử lý lỗi trong từng bản ghi.

Ví dụ dữ liệu không hợp lệ:

Lương cơ bản = "abc"

Thay vì làm chương trình dừng hoàn toàn, hệ thống sẽ:

Dữ liệu lỗi
    ↓
try...except
    ↓
Ghi lỗi vào process_log.txt
    ↓
Tiếp tục xử lý nhân viên tiếp theo

Ví dụ nội dung log:

[2026-09-18 08:00:01] SUCCESS - NV001 - Generated NV001.pdf
[2026-09-18 08:00:02] SUCCESS - NV002 - Generated NV002.pdf
[2026-09-18 08:00:03] ERROR - NV003 - Invalid basic salary
13. Ưu điểm

Hệ thống mang lại một số ưu điểm:

Tự động hóa quy trình tạo phiếu lương.
Giảm các thao tác thủ công.
Có khả năng xử lý nhiều nhân viên.
Tạo file PDF riêng cho từng nhân viên.
Có kiểm tra dữ liệu đầu vào.
Có cơ chế xử lý ngoại lệ.
Có hệ thống logging.
Mẫu PDF có thể tùy chỉnh bằng HTML/CSS.
Có cấu trúc project rõ ràng và dễ mở rộng.
14. Hạn chế và hướng phát triển
Hạn chế
Dữ liệu đầu vào cần tuân theo cấu trúc Excel đã quy định.
Các quy tắc tính lương hiện được cấu hình trong chương trình.
Người dùng cần chuẩn bị file Excel trước khi chạy.
Việc thay đổi giao diện PDF yêu cầu chỉnh sửa template HTML/CSS.
Hướng phát triển

Trong tương lai, hệ thống có thể được mở rộng với:

Giao diện người dùng để chọn file Excel.
Giao diện nhập và chỉnh sửa dữ liệu nhân viên.
Xuất báo cáo tổng hợp.
Hỗ trợ nhiều mẫu phiếu lương.
Kết nối cơ sở dữ liệu.
Tự động gửi phiếu lương qua email.
Xây dựng giao diện Web/Desktop.
Thêm chức năng thống kê và trực quan hóa dữ liệu.
15. Thành viên nhóm
STT	Họ và tên	MSSV	Vai trò
1	................................	................	Phát triển hệ thống
2	................................	................	Xử lý dữ liệu
3	................................	................	Thiết kế PDF
4	................................	................	Kiểm thử & tài liệu
5	................................	................	Báo cáo & trình bày
16. Kết luận

Đề tài đã xây dựng một quy trình tự động hóa việc tạo phiếu lương
từ dữ liệu Excel sang PDF bằng Python.

Hệ thống kết hợp xử lý dữ liệu, kiểm tra dữ liệu đầu vào, tính toán
tiền lương, tạo tài liệu bằng HTML/CSS, chuyển đổi sang PDF, xử lý
ngoại lệ và ghi log trong một quy trình thống nhất.

Kết quả đầu ra là các phiếu lương PDF được tạo tự động theo từng
nhân viên, giúp quy trình tạo báo cáo trở nên có cấu trúc, dễ quản lý
và có khả năng mở rộng.

Project Information

Project: Automated PDF Salary Report Generation
Language: Python
Input: Excel (.xlsx)
Output: PDF (.pdf)
Data Processing: Pandas, OpenPyXL
PDF Generation: WeasyPrint
Template: HTML/CSS
