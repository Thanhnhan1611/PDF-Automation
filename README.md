Automated PDF Salary Report Generation
Giới thiệu

Đây là đồ án Python xây dựng hệ thống tự động tạo phiếu lương PDF từ dữ liệu nhân viên trong Excel.

Thay vì nhập thông tin và tạo từng phiếu lương thủ công, chương trình tự động:

Đọc dữ liệu nhân viên từ file Excel.
Kiểm tra và xử lý dữ liệu đầu vào.
Tính toán tiền lương.
Tạo phiếu lương bằng HTML/CSS.
Chuyển phiếu lương thành file PDF.
Lưu mỗi phiếu lương thành một file riêng.
Ghi lại quá trình xử lý vào file log.
Mục tiêu

Đề tài hướng đến việc ứng dụng Python để tự động hóa quy trình tạo báo cáo lương.

Các mục tiêu chính:

Tự động đọc dữ liệu nhân viên từ Excel.
Tự động tính toán tiền lương.
Tự động tạo phiếu lương PDF.
Giảm thao tác thủ công.
Xử lý dữ liệu không hợp lệ.
Ghi log quá trình xử lý.
Xây dựng hệ thống có khả năng mở rộng trong tương lai.
Chức năng chính
1. Đọc dữ liệu Excel

Chương trình sử dụng Pandas và OpenPyXL để đọc dữ liệu từ file employees.xlsx.

Dữ liệu nhân viên có thể bao gồm:

Mã nhân viên
Tên nhân viên
Chức vụ
Lương cơ bản
Lương đóng BHXH
Ngày công chuẩn
Ngày công đi làm
Ngày nghỉ phép hưởng lương
Ngày nghỉ không lương
Thưởng
Phụ cấp
Số người phụ thuộc
Phạt
2. Tính toán tiền lương

Hệ thống tự động tính:

Lương theo ngày công.
Tổng thu nhập.
BHXH.
BHYT.
BHTN.
Thu nhập chịu thuế.
Thuế TNCN.
Tổng các khoản khấu trừ.
Tiền lương thực lĩnh.
3. Tạo phiếu lương PDF

Mỗi nhân viên được tạo một phiếu lương PDF riêng.

Tên file được đặt theo mã nhân viên.

Ví dụ:

NV001.pdf
NV002.pdf
NV003.pdf
4. Tự động tạo thư mục

Nếu thư mục salary/ chưa tồn tại, chương trình sẽ tự động tạo thư mục để lưu các phiếu lương.

5. Logging

Chương trình tạo file process_log.txt để ghi lại quá trình xử lý.

Log bao gồm:

Nhân viên xử lý thành công.
Nhân viên xử lý thất bại.
Nguyên nhân lỗi.
Thời gian xử lý.
Công nghệ sử dụng
Công nghệ	Mục đích
Python	Ngôn ngữ lập trình chính
Pandas	Đọc và xử lý dữ liệu
OpenPyXL	Đọc file Excel .xlsx
WeasyPrint	Chuyển HTML/CSS thành PDF
HTML	Xây dựng nội dung phiếu lương
CSS	Thiết kế giao diện phiếu lương
Cấu trúc Project
salary-report/
│
├── salary_report.py
├── employees.xlsx
├── README.md
├── process_log.txt
│
└── salary/
    ├── NV001.pdf
    ├── NV002.pdf
    ├── NV003.pdf
    └── ...

Trong đó:

salary_report.py: chương trình chính.
employees.xlsx: dữ liệu đầu vào.
salary/: thư mục chứa các phiếu lương PDF.
process_log.txt: file ghi log.
README.md: tài liệu mô tả project.
Quy trình hoạt động
employees.xlsx
      ↓
Đọc dữ liệu bằng Pandas
      ↓
Kiểm tra dữ liệu
      ↓
Xử lý từng nhân viên
      ↓
Tính toán tiền lương
      ↓
Tạo HTML + CSS
      ↓
Chuyển HTML → PDF
      ↓
Lưu vào thư mục salary/
      ↓
Ghi process_log.txt
Mô hình tính lương

Hệ thống phân biệt giữa:

Ngày công chuẩn.
Ngày công thực tế.
Ngày nghỉ phép hưởng lương.
Ngày nghỉ không lương.
Tính ngày công

Ngày công tính lương được xác định theo:

Ngày công tính lương = Ngày công thực tế + Ngày nghỉ phép hưởng lương

Ví dụ:

Ngày công chuẩn: 26 ngày
Ngày công thực tế: 25 ngày
Ngày nghỉ phép hưởng lương: 1 ngày

Khi đó:

Ngày công tính lương = 25 + 1 = 26 ngày

Nhân viên được tính đủ lương theo 26 ngày công.

Trong trường hợp nhân viên đi làm 25 ngày và không có ngày nghỉ phép hưởng lương:

Ngày công tính lương = 25 ngày

Tiền lương được tính theo tỷ lệ 25/26 ngày.

Tính lương theo ngày công

Lương thực tế = Lương cơ bản / Ngày công chuẩn × Ngày công tính lương

Các khoản khấu trừ

Hệ thống có thể tính các khoản:

BHXH.
BHYT.
BHTN.
Thuế TNCN.
Các khoản phạt hoặc khấu trừ khác nếu có.
Cài đặt

Cài đặt các thư viện cần thiết:

pip install pandas openpyxl weasyprint

Kiểm tra phiên bản Python:

python --version

Kiểm tra các thư viện:

pip show pandas
pip show openpyxl
pip show weasyprint

Đối với Windows, WeasyPrint có thể yêu cầu thêm thư viện GTK/MSYS2 để hoạt động.

Hướng dẫn sử dụng
Bước 1: Chuẩn bị file Excel

Đặt file employees.xlsx vào cùng thư mục với salary_report.py.

Bước 2: Kiểm tra dữ liệu

Đảm bảo các cột trong Excel được đặt đúng tên và dữ liệu có định dạng phù hợp.

Bước 3: Chạy chương trình
python salary_report.py
Bước 4: Kiểm tra kết quả

Các phiếu lương được lưu trong thư mục salary/.

Ví dụ:

salary/
├── NV001.pdf
├── NV002.pdf
└── NV003.pdf

File process_log.txt cũng được tạo để theo dõi quá trình xử lý.

Dữ liệu đầu vào

File employees.xlsx có thể sử dụng cấu trúc:

Mã NV	Tên NV	Chức vụ	Lương cơ bản	Ngày công chuẩn	Ngày công đi làm	Thưởng	Phụ cấp
NV001	Nguyễn Văn A	Nhân viên	10000000	26	26	1000000	500000
NV002	Trần Văn B	Nhân viên	12000000	26	25	500000	300000

Có thể bổ sung thêm các trường dữ liệu phục vụ cho việc tính lương.

Kết quả đầu ra

Hệ thống tạo ra hai loại kết quả chính.

Phiếu lương PDF

Mỗi nhân viên có một file PDF riêng.

Phiếu lương bao gồm:

Thông tin công ty.
Thông tin nhân viên.
Tháng tính lương.
Ngày công.
Các khoản thu nhập.
Các khoản khấu trừ.
Thuế TNCN.
Số tiền thực lĩnh.
Số tiền bằng chữ.
Khu vực ký xác nhận.
File Log

File process_log.txt ghi lại kết quả xử lý.

Ví dụ:

[2026-09-01 10:30:15] SUCCESS - NV001 - Nguyen Van A - NV001.pdf
[2026-09-01 10:30:16] SUCCESS - NV002 - Tran Van B - NV002.pdf
[2026-09-01 10:30:17] FAILED - NV003 - Invalid salary data
Xử lý lỗi

Chương trình sử dụng try...except để xử lý lỗi trong quá trình chạy.

Một số trường hợp được xử lý:

File Excel không tồn tại.
Thiếu cột bắt buộc.
Lương không phải dữ liệu số.
Ngày công không hợp lệ.
Dữ liệu nhân viên bị thiếu.
Không thể tạo file PDF.
Lỗi trong quá trình xử lý từng nhân viên.

Khi một nhân viên có dữ liệu không hợp lệ, chương trình sẽ ghi nhận lỗi vào process_log.txt và tiếp tục xử lý các nhân viên còn lại.

Ưu điểm
Tự động hóa quy trình tạo phiếu lương.
Giảm thao tác thủ công.
Có thể xử lý nhiều nhân viên.
Tạo PDF riêng cho từng nhân viên.
Có cơ chế xử lý lỗi.
Có hệ thống logging.
Có thể mở rộng công thức tính lương.
Có thể tái sử dụng cho nhiều loại báo cáo PDF khác.
Hạn chế và hướng phát triển
Hạn chế
Chưa có giao diện người dùng.
Một số thông số tính lương được thiết lập trong chương trình.
Chưa kết nối trực tiếp với hệ thống quản lý nhân sự.
Công thức thuế và bảo hiểm cần được cập nhật khi quy định thay đổi.
Hướng phát triển
Xây dựng giao diện bằng Tkinter hoặc PyQt.
Cho phép người dùng chọn file Excel.
Cho phép lựa chọn thư mục xuất PDF.
Xuất báo cáo tổng hợp.
Gửi phiếu lương qua email.
Kết nối cơ sở dữ liệu.
Xây dựng dashboard thống kê.
Cho phép cấu hình công thức tính lương từ giao diện.
Thành viên nhóm
STT	Họ và tên	Vai trò
1	Thành viên 1	Phát triển hệ thống
2	Thành viên 2	Xử lý dữ liệu
3	Thành viên 3	Kiểm thử
4	Thành viên 4	Báo cáo và thuyết trình

Thay thông tin thành viên bằng thông tin thực tế của nhóm.

Kết luận

Đề tài xây dựng hệ thống tự động tạo báo cáo PDF từ dữ liệu Excel bằng Python.

Thông qua project, nhóm áp dụng các kiến thức về:

Python.
Xử lý dữ liệu.
Đọc và xử lý Excel.
Exception Handling.
HTML/CSS.
Tự động hóa.
Tạo báo cáo PDF.
Logging.

Hệ thống giúp giảm các thao tác thủ công trong quá trình tạo phiếu lương và có khả năng mở rộng thành một hệ thống báo cáo tự động hoàn chỉnh.

Project Information
Thông tin	Nội dung
Project	Automated PDF Salary Report Generation
Language	Python
Input	Excel (.xlsx)
Output	PDF (.pdf)
Template	HTML/CSS
