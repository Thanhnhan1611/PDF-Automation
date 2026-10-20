# PDF Automation — Tự động tạo phiếu lương PDF

Ứng dụng Python đọc dữ liệu nhân viên từ Excel, kiểm tra đầu vào, tính các khoản lương theo quy tắc được cấu hình, tạo phiếu lương PDF riêng cho từng nhân viên và ghi nhật ký xử lý.

> **Phạm vi sử dụng:** Đây là đồ án học tập và công cụ tự động hóa tài liệu, không phải phần mềm kế toán hoặc quyết toán thuế đã được chứng nhận. Trước khi dùng cho bảng lương thật, phải xác minh các quy tắc thuế, bảo hiểm và chính sách của doanh nghiệp.

## Chức năng

- Đọc tệp `employees.xlsx`.
- Kiểm tra các cột bắt buộc, dữ liệu số, ngày công và mã nhân viên trùng.
- Chuẩn hóa giá trị số có dấu phân cách hàng nghìn hoặc thập phân.
- Tính lương bằng Python với `Decimal` và quy tắc làm tròn half-up.
- Tính các khoản BHXH, BHYT, BHTN, thu nhập tính thuế, thuế TNCN, tổng khấu trừ và thực lãnh.
- Tạo phiếu lương A4 bằng HTML/CSS và WeasyPrint.
- Ghi nhật ký và tổng kết số dòng thành công/thất bại.
- Kiểm thử tự động bằng pytest; GitHub Actions chạy kiểm thử khi có push hoặc pull request.

## Luồng xử lý

```text
employees.xlsx
    |
    v
Đọc và chuẩn hóa tiêu đề
    |
    v
Kiểm tra cấu trúc và mã nhân viên
    |
    v
Xác thực dữ liệu từng nhân viên
    |
    v
Tính lương trong Python
    |
    v
Tạo HTML/CSS và PDF tạm
    |
    v
Kiểm tra chữ ký PDF -> thay thế file đích
    |
    v
salary/[Mã NV].pdf + process_log.txt
```

Python là nguồn tính toán duy nhất. Các cột kết quả công thức có sẵn trong Excel không được dùng để ghi đè từng phần kết quả Python, nhằm tránh trường hợp trộn giá trị cache cũ với giá trị tính lại.

## Cột dữ liệu đầu vào

### Cột bắt buộc

| Cột | Ý nghĩa |
|---|---|
| `Mã NV` | Mã nhân viên duy nhất; chỉ dùng chữ không dấu, số, dấu gạch ngang và gạch dưới |
| `Tên NV` | Họ và tên |
| `Chức vụ` | Chức vụ |
| `Lương cơ bản` | Mức lương cơ bản |
| `Thưởng` | Khoản thưởng |
| `Phạt` | Khoản khấu trừ bổ sung |

### Cột tùy chọn

| Cột | Mặc định khi không có hoặc để trống |
|---|---|
| `Lương đóng BHXH` | Bằng lương cơ bản |
| `Ngày công chuẩn` | 26 |
| `Ngày công đi làm` | Bằng ngày công chuẩn |
| `Phụ cấp` | 0 |
| `Số người phụ thuộc` | 0 |

Các tiêu đề cần đúng như trên. Mã nhân viên nên được lưu dưới dạng văn bản trong Excel để tránh mất số 0 ở đầu.

## Quy tắc tính toán

```text
Lương theo ngày = Lương cơ bản / Ngày công chuẩn
Lương thực tế = Lương theo ngày × Ngày công đi làm
Tổng thu nhập = Lương thực tế + Thưởng + Phụ cấp

BHXH = Lương đóng BHXH × 8%
BHYT = Lương đóng BHXH × 1,5%
BHTN = Lương đóng BHXH × 1%

Thu nhập tính thuế = max(
    0,
    Tổng thu nhập - BHXH - BHYT - BHTN
    - Giảm trừ bản thân - Giảm trừ người phụ thuộc
)

Tổng khấu trừ = BHXH + BHYT + BHTN + Thuế TNCN + Phạt
Thực lãnh = Tổng thu nhập - Tổng khấu trừ
```

### Mô hình thuế TNCN năm 2026

Hàm dự phòng được cập nhật theo biểu thuế lũy tiến 5 bậc cho thu nhập tính thuế tháng:

| Phần thu nhập tính thuế tháng | Thuế suất |
|---|---:|
| Đến 10 triệu đồng | 5% |
| Trên 10 đến 30 triệu đồng | 10% |
| Trên 30 đến 60 triệu đồng | 20% |
| Trên 60 đến 100 triệu đồng | 30% |
| Trên 100 triệu đồng | 35% |

Chương trình sử dụng mức giảm trừ bản thân 15,5 triệu đồng/tháng và người phụ thuộc 6,2 triệu đồng/tháng. Các quy tắc này được dùng cho mô hình theo kỳ tính thuế năm 2026; hãy kiểm tra văn bản chính thức và các trường hợp áp dụng trước khi dùng trong thực tế.

Nguồn tham khảo:
- [Biểu thuế TNCN lũy tiến từng phần — Cổng thông tin Chính phủ](https://xaydungchinhsach.chinhphu.vn/chi-tiet-bieu-thue-thu-nhap-ca-nhan-luy-tien-tung-phan-119260327091407544.htm)
- [Quy định giảm trừ gia cảnh — Cổng thông tin Chính phủ](https://xaydungchinhsach.chinhphu.vn/quy-dinh-giam-tru-gia-canh-119260703144038062.htm)

## Kết quả đầu ra

- Mỗi nhân viên hợp lệ tạo một tệp `salary/[Mã NV].pdf`.
- PDF gồm thông tin đơn vị, kỳ lương, thông tin nhân viên, ngày công, các khoản thu nhập, bảo hiểm, thuế, phạt, tổng khấu trừ, thực lãnh và số tiền bằng chữ.
- File `process_log.txt` lưu thời gian, cấp độ và chi tiết xử lý.
- Nếu có bản ghi lỗi, chương trình ghi lại dòng Excel, mã nhân viên và lý do để người dùng kiểm tra.

Tệp PDF được tạo tạm trước khi thay thế file đích, giúp tránh dùng một tệp PDF dở dang nếu bước kết xuất gặp lỗi. Mã nhân viên trùng được phát hiện trước khi xử lý để tránh ghi đè phiếu lương giữa các bản ghi trong cùng bảng.

## Cài đặt

Khuyến nghị Python 3.12 trở lên.

```bash
python -m venv .venv
```

Windows CMD:

```cmd
.venv\Scripts\activate
python -m pip install -r requirements.txt
```

macOS/Linux:

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
```

### Lưu ý Windows và WeasyPrint

Nếu WeasyPrint báo thiếu DLL Pango/GLib, cần cài và cấu hình các thư viện hệ thống phù hợp. Script sẽ dùng `C:\\msys64\\ucrt64\\bin` nếu thư mục này tồn tại. Việc cài gói Python thôi có thể chưa đủ trên Windows.

## Chạy chương trình

Đặt `employees.xlsx` cùng thư mục với `salary_report.py`.

```cmd
python salary_report.py
```

Mặc định kỳ lương lấy theo tháng hiện tại của máy. Nếu xử lý kỳ lương khác, hãy chỉ định biến môi trường.

Windows CMD ví dụ:

```cmd
set PAY_MONTH=09/2026
python salary_report.py
```

PowerShell ví dụ:

```powershell
$env:PAY_MONTH = "09/2026"
python salary_report.py
```

Định dạng kỳ lương là `MM/YYYY`.

## Chạy kiểm thử

Cài thêm thư viện phục vụ phát triển:

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Các kiểm thử hiện bao gồm định dạng số, giá trị không hợp lệ, các ngưỡng thuế năm 2026, công thức tính lương, cột người phụ thuộc tùy chọn, mã nhân viên trùng và kiểm tra đường đi tạo PDF bằng bộ giả lập.

## An toàn dữ liệu

- **Không đưa bảng lương thật, thông tin cá nhân thật hoặc log chứa dữ liệu nhạy cảm lên repository công khai.**
- Dùng dữ liệu giả lập để demo.
- Bảo vệ thư mục chứa PDF và chỉ chia sẻ phiếu lương cho người có thẩm quyền.
- Nếu `employees.xlsx` hiện tại chứa dữ liệu thật, hãy thay bằng dữ liệu giả lập và loại bỏ bản ghi nhạy cảm khỏi lịch sử Git trước khi tiếp tục dùng repository công khai.

## Cấu trúc chính

```text
PDF-Automation/
├── salary_report.py
├── employees.xlsx
├── README.md
├── requirements.txt
├── requirements-dev.txt
├── .gitignore
├── tests/
│   └── test_salary_report.py
└── .github/
    └── workflows/
        └── tests.yml
```
