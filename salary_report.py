# ============================================================
# BÀI A3 - TỰ ĐỘNG TẠO PHIẾU LƯƠNG PDF
# File: salary_report.py
#
# Luồng:
# employees.xlsx
#     -> Pandas đọc Excel
#     -> kiểm tra dữ liệu từng dòng
#     -> lấy kết quả công thức Excel
#     -> nếu WPS chưa lưu kết quả công thức thì tính dự phòng
#     -> HTML/CSS
#     -> WeasyPrint
#     -> salary/NVxxx.pdf
#     -> process_log.txt
#
# Cài:
# pip install pandas openpyxl weasyprint
#
# Windows + MSYS2 UCRT64:
# C:\msys64\ucrt64\bin phải tồn tại và chứa DLL của Pango/GLib.
# ============================================================

import os
import sys
import html
import logging
from pathlib import Path
from datetime import datetime

# ------------------------------------------------------------
# 1. Cấu hình DLL cho WeasyPrint trên Windows
# ------------------------------------------------------------

if sys.platform.startswith("win"):
    dll_dir = r"C:\msys64\ucrt64\bin"

    if os.path.isdir(dll_dir):
        os.environ["WEASYPRINT_DLL_DIRECTORIES"] = dll_dir

# ------------------------------------------------------------
# 2. Import thư viện
# ------------------------------------------------------------

import pandas as pd
from weasyprint import HTML


# ------------------------------------------------------------
# 3. Đường dẫn
# ------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "employees.xlsx"
OUTPUT_DIR = BASE_DIR / "salary"
LOG_FILE = BASE_DIR / "process_log.txt"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ------------------------------------------------------------
# 4. Thông tin phiếu
# ------------------------------------------------------------

COMPANY_NAME = "CÔNG TY BTL"

COMPANY_ADDRESS = (
    "Địa chỉ: 20A, 20 Lý Tự Trọng, Phường Sài Gòn, "
    "Thành phố Hồ Chí Minh, Việt Nam"
)

from datetime import datetime

PAY_MONTH = datetime.now().strftime("%m/%Y")

# ------------------------------------------------------------
# 5. Logging
# ------------------------------------------------------------

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8",
)


# ------------------------------------------------------------
# 6. Các tỷ lệ dùng trong bài mô phỏng
# ------------------------------------------------------------
# Người lao động:
# BHXH = 8%
# BHYT  = 1.5%
# BHTN  = 1%
#
# Các tỷ lệ này phù hợp với mức đóng bắt buộc hiện hành đối
# với lao động Việt Nam từ 01/07/2025 theo BHXH Việt Nam.
#
# Thuế TNCN trong file Excel nên được xem là mô hình phục vụ
# bài A3, không phải phần mềm quyết toán thuế pháp lý.
# ------------------------------------------------------------

BHXH_RATE = 0.08
BHYT_RATE = 0.015
BHTN_RATE = 0.01

PERSONAL_DEDUCTION = 15_500_000
DEPENDENT_DEDUCTION = 6_200_000


# ------------------------------------------------------------
# 7. Hàm format tiền
# ------------------------------------------------------------

def format_money(value):
    """Đổi số thành dạng 15.000.000 VNĐ."""

    try:
        number = float(value)

        return (
            f"{number:,.0f}"
            .replace(",", ".")
            + " VNĐ"
        )

    except Exception:
        return "—"


# ------------------------------------------------------------
# 8. Hàm lấy giá trị an toàn
# ------------------------------------------------------------

def get_value(row, column, default=None):
    """Lấy dữ liệu từ Pandas Series."""

    if column not in row.index:
        return default

    value = row[column]

    if pd.isna(value):
        return default

    return value


# ------------------------------------------------------------
# 9. Chuyển dữ liệu tiền về số
# ------------------------------------------------------------

def parse_number(value, field_name):
    """
    Chuyển dữ liệu thành số.
    Cho phép:
        15000000
        "15,000,000"
        "15.000.000"
    """

    if value is None or pd.isna(value):
        raise ValueError(
            f"{field_name} đang bị bỏ trống"
        )

    if isinstance(value, str):
        value = value.strip()

        if not value:
            raise ValueError(
                f"{field_name} đang bị bỏ trống"
            )

        value = (
            value
            .replace(",", "")
            .replace(".", "")
            .replace(" ", "")
            .replace("VNĐ", "")
            .strip()
        )

    try:
        number = float(value)

    except Exception:
        raise ValueError(
            f"{field_name} không phải số: {value}"
        )

    if number < 0:
        raise ValueError(
            f"{field_name} không được âm"
        )

    return number


# ------------------------------------------------------------
# 10. Đọc số thành chữ
# ------------------------------------------------------------

ONES = [
    "không",
    "một",
    "hai",
    "ba",
    "bốn",
    "năm",
    "sáu",
    "bảy",
    "tám",
    "chín",
]

UNITS = [
    "",
    "nghìn",
    "triệu",
    "tỷ",
    "nghìn tỷ",
    "triệu tỷ",
]


def read_three_digits(number, full=False):
    """Đọc một nhóm tối đa 3 chữ số."""

    number = int(number)

    hundreds = number // 100
    remainder = number % 100

    tens = remainder // 10
    units = remainder % 10

    result = []

    if hundreds > 0:
        result.append(
            ONES[hundreds] + " trăm"
        )

    elif full and number > 0:
        result.append("không trăm")

    if tens > 1:
        result.append(
            ONES[tens] + " mươi"
        )

        if units == 1:
            result.append("mốt")
        elif units == 4:
            result.append("tư")
        elif units == 5:
            result.append("lăm")
        elif units > 0:
            result.append(ONES[units])

    elif tens == 1:
        result.append("mười")

        if units == 5:
            result.append("lăm")
        elif units > 0:
            result.append(ONES[units])

    elif units > 0:
        if hundreds > 0 or full:
            result.append("lẻ")

        result.append(ONES[units])

    return " ".join(result)


def number_to_vietnamese(number):
    """Chuyển số tiền thành chữ tiếng Việt."""

    try:
        number = int(round(float(number)))
    except Exception:
        return "Không xác định"

    if number == 0:
        return "Không đồng"

    if number < 0:
        return "Âm " + number_to_vietnamese(-number)

    groups = []

    temp = number

    while temp > 0:
        groups.append(temp % 1000)
        temp //= 1000

    parts = []

    for index in range(len(groups) - 1, -1, -1):
        group = groups[index]

        if group == 0:
            continue

        # Nhóm đầu không cần "không trăm".
        # Các nhóm sau có thể cần "không trăm" nếu < 100.
        full = index < len(groups) - 1

        text = read_three_digits(group, full)

        if index > 0:
            text += " " + UNITS[index]

        parts.append(text)

    result = " ".join(parts)

    return result[:1].upper() + result[1:] + " đồng"


# ------------------------------------------------------------
# 11. Tính thuế TNCN dạng lũy tiến 7 bậc
# ------------------------------------------------------------
# Đây là hàm dự phòng để PDF vẫn chạy nếu WPS/Excel chưa
# lưu cached value của công thức.
#
# Nếu bạn muốn dùng một biểu thuế khác cho đề tài, chỉ cần
# sửa hàm này và phần công thức Excel tương ứng.
# ------------------------------------------------------------

def calculate_pit(taxable_income):
    """Tính thuế TNCN lũy tiến 7 bậc."""

    x = max(float(taxable_income), 0)

    brackets = [
        (5_000_000, 0.05),
        (10_000_000, 0.10),
        (18_000_000, 0.15),
        (32_000_000, 0.20),
        (52_000_000, 0.25),
        (80_000_000, 0.30),
        (float("inf"), 0.35),
    ]

    tax = 0
    previous = 0

    for upper, rate in brackets:
        if x <= previous:
            break

        taxable_part = min(
            x,
            upper
        ) - previous

        tax += taxable_part * rate

        previous = upper

        if x <= upper:
            break

    return round(tax)


# ------------------------------------------------------------
# 12. Tính lương dự phòng
# ------------------------------------------------------------
# Quan trọng:
#
# Excel vẫn là nơi có công thức.
# Hàm này chỉ dùng khi WPS/Excel chưa lưu kết quả công thức.
# ------------------------------------------------------------

def calculate_salary_fallback(row):
    basic = parse_number(
        row["Lương cơ bản"],
        "Lương cơ bản"
    )

    insurance_salary = parse_number(
        get_value(
            row,
            "Lương đóng BHXH",
            basic
        ),
        "Lương đóng BHXH"
    )

    standard_days = parse_number(
        get_value(
            row,
            "Ngày công chuẩn",
            26
        ),
        "Ngày công chuẩn"
    )

    actual_days = parse_number(
        get_value(
            row,
            "Ngày công đi làm",
            standard_days
        ),
        "Ngày công đi làm"
    )

    bonus = parse_number(
        get_value(row, "Thưởng", 0),
        "Thưởng"
    )

    allowance = parse_number(
        get_value(row, "Phụ cấp", 0),
        "Phụ cấp"
    )

    dependents = parse_number(
        get_value(
            row,
            "Số người phụ thuộc",
            0
        ),
        "Số người phụ thuộc"
    )

    penalty = parse_number(
        get_value(row, "Phạt", 0),
        "Phạt"
    )

    if standard_days <= 0:
        raise ValueError(
            "Ngày công chuẩn phải lớn hơn 0"
        )

    if actual_days < 0 or actual_days > standard_days:
        raise ValueError(
            "Ngày công đi làm phải nằm trong khoảng "
            "0 đến ngày công chuẩn"
        )

    daily_salary = (
        basic / standard_days
    )

    actual_salary = round(
        daily_salary * actual_days
    )

    total_income = (
        actual_salary
        + bonus
        + allowance
    )

    bhxh = round(
        insurance_salary * BHXH_RATE
    )

    bhyt = round(
        insurance_salary * BHYT_RATE
    )

    bhtn = round(
        insurance_salary * BHTN_RATE
    )

    taxable_income = max(
        0,
        total_income
        - bhxh
        - bhyt
        - bhtn
        - PERSONAL_DEDUCTION
        - dependents * DEPENDENT_DEDUCTION
    )

    pit = calculate_pit(
        taxable_income
    )

    total_deduction = (
        bhxh
        + bhyt
        + bhtn
        + pit
        + penalty
    )

    net_salary = (
        total_income
        - total_deduction
    )

    return {
        "Lương theo ngày": daily_salary,
        "Lương thực tế": actual_salary,
        "Tổng thu nhập": total_income,
        "BHXH": bhxh,
        "BHYT": bhyt,
        "BHTN": bhtn,
        "Thu nhập tính thuế": taxable_income,
        "Thuế TNCN": pit,
        "Tổng khấu trừ": total_deduction,
        "Thực lãnh": net_salary,
        "Lương đóng BHXH": insurance_salary,
        "Ngày công chuẩn": standard_days,
        "Ngày công đi làm": actual_days,
        "Thưởng": bonus,
        "Phụ cấp": allowance,
        "Phạt": penalty,
    }


# ------------------------------------------------------------
# 13. HTML/CSS
# ------------------------------------------------------------

def create_salary_html(employee):
    """Tạo HTML phiếu lương A4."""

    employee_id = html.escape(
        str(employee["Mã NV"])
    )

    employee_name = html.escape(
        str(employee["Tên NV"])
    )

    position = html.escape(
        str(employee["Chức vụ"])
    )

    insurance_salary = employee["Lương đóng BHXH"]

    standard_days = employee["Ngày công chuẩn"]

    actual_days = employee["Ngày công đi làm"]

    basic_salary = employee["Lương cơ bản"]

    daily_salary = employee["Lương theo ngày"]

    actual_salary = employee["Lương thực tế"]

    bonus = employee["Thưởng"]

    allowance = employee["Phụ cấp"]

    total_income = employee["Tổng thu nhập"]

    bhxh = employee["BHXH"]

    bhyt = employee["BHYT"]

    bhtn = employee["BHTN"]

    taxable_income = employee[
        "Thu nhập tính thuế"
    ]

    pit = employee["Thuế TNCN"]

    penalty = employee["Phạt"]

    total_deduction = employee[
        "Tổng khấu trừ"
    ]

    net_salary = employee["Thực lãnh"]

    dependents = employee[
        "Số người phụ thuộc"
    ]

    amount_in_word = number_to_vietnamese(
        net_salary
    )

    return f"""
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">

<style>

@page {{
    size: A4 portrait;
    margin: 10mm;
}}

body {{
    font-family: "DejaVu Sans", Arial, sans-serif;
    font-size: 10px;
    color: #111;
    margin: 0;
}}

.slip {{
    border: 1.2px solid #222;
    padding: 12px;
}}

.company {{
    font-size: 11px;
    font-weight: bold;
}}

.address {{
    font-size: 8.5px;
    margin-top: 3px;
}}

.title {{
    text-align: center;
    font-size: 19px;
    font-weight: bold;
    margin-top: 14px;
}}

.period {{
    text-align: center;
    margin-top: 4px;
    margin-bottom: 14px;
    font-size: 10px;
}}

.info {{
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 12px;
}}

.info td {{
    border: 1px solid #333;
    padding: 5px 6px;
}}

.label {{
    font-weight: bold;
    background: #eeeeee;
    width: 18%;
}}

.value {{
    width: 32%;
}}

.salary {{
    width: 100%;
    border-collapse: collapse;
}}

.salary th,
.salary td {{
    border: 1px solid #333;
    padding: 5px 6px;
}}

.salary th {{
    text-align: center;
    background: #eeeeee;
    font-weight: bold;
}}

.section {{
    font-weight: bold;
    background: #f2f2f2;
}}

.item {{
    width: 65%;
}}

.amount {{
    width: 35%;
    text-align: right;
    white-space: nowrap;
}}

.total td {{
    font-weight: bold;
    font-size: 12px;
    padding-top: 8px;
    padding-bottom: 8px;
}}

.net td {{
    font-weight: bold;
    font-size: 14px;
    padding-top: 10px;
    padding-bottom: 10px;
}}

.words {{
    margin-top: 12px;
    padding: 8px;
    border: 1px solid #333;
}}

.note {{
    margin-top: 14px;
    font-size: 8.5px;
}}

.signatures {{
    width: 100%;
    margin-top: 35px;
    text-align: center;
}}

.signatures td {{
    width: 50%;
    vertical-align: top;
}}

.signature-space {{
    height: 55px;
}}

</style>
</head>

<body>

<div class="slip">

    <div class="company">
        {COMPANY_NAME}
    </div>

    <div class="address">
        {COMPANY_ADDRESS}
    </div>

    <div class="title">
        PHIẾU LƯƠNG
    </div>

    <div class="period">
        Tháng {PAY_MONTH}
    </div>


    <table class="info">

        <tr>
            <td class="label">Mã nhân viên</td>
            <td class="value">{employee_id}</td>

            <td class="label">Họ và tên</td>
            <td class="value">{employee_name}</td>
        </tr>

        <tr>
            <td class="label">Chức vụ</td>
            <td class="value">{position}</td>

            <td class="label">Người phụ thuộc</td>
            <td class="value">{dependents}</td>
        </tr>

        <tr>
            <td class="label">Ngày công chuẩn</td>
            <td class="value">{standard_days}</td>

            <td class="label">Ngày công đi làm</td>
            <td class="value">{actual_days}</td>
        </tr>

        <tr>
            <td class="label">Lương đóng BHXH</td>
            <td class="value">{format_money(insurance_salary)}</td>

            <td class="label">Lương cơ bản</td>
            <td class="value">{format_money(basic_salary)}</td>
        </tr>

    </table>


    <table class="salary">

        <tr>
            <th colspan="2">
                CHI TIẾT THU NHẬP
            </th>
        </tr>

        <tr>
            <td class="item">Lương theo ngày</td>
            <td class="amount">{format_money(daily_salary)}</td>
        </tr>

        <tr>
            <td class="item">
                Lương thực tế
                ({actual_days}/{standard_days} ngày)
            </td>
            <td class="amount">{format_money(actual_salary)}</td>
        </tr>

        <tr>
            <td class="item">Thưởng</td>
            <td class="amount">{format_money(bonus)}</td>
        </tr>

        <tr>
            <td class="item">Phụ cấp</td>
            <td class="amount">{format_money(allowance)}</td>
        </tr>

        <tr class="total">
            <td class="item">TỔNG THU NHẬP</td>
            <td class="amount">{format_money(total_income)}</td>
        </tr>


        <tr>
            <th colspan="2">
                CÁC KHOẢN KHẤU TRỪ
            </th>
        </tr>

        <tr>
            <td class="item">BHXH - 8%</td>
            <td class="amount">{format_money(bhxh)}</td>
        </tr>

        <tr>
            <td class="item">BHYT - 1,5%</td>
            <td class="amount">{format_money(bhyt)}</td>
        </tr>

        <tr>
            <td class="item">BHTN - 1%</td>
            <td class="amount">{format_money(bhtn)}</td>
        </tr>

        <tr>
            <td class="item">Thu nhập tính thuế</td>
            <td class="amount">{format_money(taxable_income)}</td>
        </tr>

        <tr>
            <td class="item">Thuế TNCN</td>
            <td class="amount">{format_money(pit)}</td>
        </tr>

        <tr>
            <td class="item">Phạt</td>
            <td class="amount">{format_money(penalty)}</td>
        </tr>

        <tr class="total">
            <td class="item">TỔNG KHẤU TRỪ</td>
            <td class="amount">{format_money(total_deduction)}</td>
        </tr>


        <tr class="net">
            <td class="item">THỰC LÃNH</td>
            <td class="amount">{format_money(net_salary)}</td>
        </tr>

    </table>


    <div class="words">
        <strong>Số tiền bằng chữ:</strong>
        {html.escape(amount_in_word)}
    </div>


    <div class="note">
        Ghi chú: Phiếu được tạo tự động từ dữ liệu bảng lương.
        Các khoản bảo hiểm và thuế trong bài là mô hình phục vụ
        mục đích học tập/tự động hóa tài liệu.
    </div>


    <table class="signatures">

        <tr>
            <td>
                <strong>NGƯỜI LẬP PHIẾU</strong>
                <div class="signature-space"></div>
            </td>

            <td>
                <strong>NGƯỜI NHẬN LƯƠNG</strong>
                <div class="signature-space"></div>
            </td>
        </tr>

    </table>

</div>

</body>
</html>
"""


# ------------------------------------------------------------
# 14. Các cột bắt buộc
# ------------------------------------------------------------

def validate_columns(df):
    required = [
        "Mã NV",
        "Tên NV",
        "Chức vụ",
        "Lương cơ bản",
        "Thưởng",
        "Phạt",
    ]

    missing = [
        column
        for column in required
        if column not in df.columns
    ]

    if missing:
        raise ValueError(
            "Thiếu cột bắt buộc: "
            + ", ".join(missing)
        )


# ------------------------------------------------------------
# 15. Xử lý một nhân viên
# ------------------------------------------------------------

def process_employee(row, excel_row_number):
    """Kiểm tra dữ liệu và tạo một PDF."""

    employee_id = str(
        get_value(row, "Mã NV", "")
    ).strip()

    if not employee_id:
        raise ValueError(
            "Mã NV bị bỏ trống"
        )

    employee_name = str(
        get_value(row, "Tên NV", "")
    ).strip()

    if not employee_name:
        raise ValueError(
            "Tên NV bị bỏ trống"
        )

    position = str(
        get_value(row, "Chức vụ", "")
    ).strip()

    if not position:
        raise ValueError(
            "Chức vụ bị bỏ trống"
        )

    # --------------------------------------------------------
    # Lấy dữ liệu đầu vào
    # --------------------------------------------------------

    basic_salary = parse_number(
        row["Lương cơ bản"],
        "Lương cơ bản"
    )

    insurance_salary = parse_number(
        get_value(
            row,
            "Lương đóng BHXH",
            basic_salary
        ),
        "Lương đóng BHXH"
    )

    standard_days = parse_number(
        get_value(
            row,
            "Ngày công chuẩn",
            26
        ),
        "Ngày công chuẩn"
    )

    actual_days = parse_number(
        get_value(
            row,
            "Ngày công đi làm",
            standard_days
        ),
        "Ngày công đi làm"
    )

    bonus = parse_number(
        get_value(row, "Thưởng", 0),
        "Thưởng"
    )

    allowance = parse_number(
        get_value(row, "Phụ cấp", 0),
        "Phụ cấp"
    )

    dependents = parse_number(
        get_value(
            row,
            "Số người phụ thuộc",
            0
        ),
        "Số người phụ thuộc"
    )

    penalty = parse_number(
        get_value(row, "Phạt", 0),
        "Phạt"
    )

    # --------------------------------------------------------
    # Kiểm tra ngày công
    # --------------------------------------------------------

    if standard_days <= 0:
        raise ValueError(
            "Ngày công chuẩn phải > 0"
        )

    if actual_days > standard_days:
        raise ValueError(
            "Ngày công đi làm không được lớn hơn "
            "ngày công chuẩn"
        )

    # --------------------------------------------------------
    # Tính dự phòng.
    #
    # Nếu Excel đã có kết quả công thức thì ưu tiên kết quả
    # từ Excel. Nếu cached value chưa có thì dùng kết quả này.
    # --------------------------------------------------------

    fallback = calculate_salary_fallback(
        row
    )

    employee = row.copy()

    employee["Mã NV"] = employee_id
    employee["Tên NV"] = employee_name
    employee["Chức vụ"] = position

    for key, value in fallback.items():
        employee[key] = value

    # --------------------------------------------------------
    # Nếu Excel có kết quả công thức hợp lệ thì dùng nó.
    # --------------------------------------------------------

    formula_columns = [
        "Lương theo ngày",
        "Lương thực tế",
        "Tổng thu nhập",
        "BHXH",
        "BHYT",
        "BHTN",
        "Thu nhập tính thuế",
        "Thuế TNCN",
        "Tổng khấu trừ",
        "Thực lãnh",
    ]

    for column in formula_columns:

        excel_value = get_value(
            row,
            column,
            None
        )

        if excel_value is None:
            continue

        try:
            number = float(excel_value)

            if number >= 0:
                employee[column] = number

        except Exception:
            pass

    # --------------------------------------------------------
    # Kiểm tra thực lãnh cuối cùng
    # --------------------------------------------------------

    employee["Thực lãnh"] = round(
        float(employee["Thực lãnh"])
    )

    if employee["Thực lãnh"] < 0:
        raise ValueError(
            "Thực lãnh không được âm"
        )

    # --------------------------------------------------------
    # Tạo HTML
    # --------------------------------------------------------

    html_content = create_salary_html(
        employee
    )

    # --------------------------------------------------------
    # Tạo PDF
    # --------------------------------------------------------

    pdf_path = (
        OUTPUT_DIR
        / f"{employee_id}.pdf"
    )

    HTML(
        string=html_content,
        base_url=str(BASE_DIR)
    ).write_pdf(
        str(pdf_path)
    )

    # --------------------------------------------------------
    # Kiểm tra file
    # --------------------------------------------------------

    if not pdf_path.exists():
        raise RuntimeError(
            "Không tạo được file PDF"
        )

    if pdf_path.stat().st_size == 0:
        raise RuntimeError(
            "File PDF có kích thước 0 byte"
        )

    return pdf_path


# ------------------------------------------------------------
# 16. Main
# ------------------------------------------------------------

def main():

    print()
    print("=" * 65)
    print("        CHƯƠNG TRÌNH TẠO PHIẾU LƯƠNG PDF")
    print("=" * 65)
    print()

    logging.info("=" * 65)
    logging.info("BẮT ĐẦU CHƯƠNG TRÌNH")

    # --------------------------------------------------------
    # Kiểm tra Excel
    # --------------------------------------------------------

    print(
        f"Đang đọc file: {EXCEL_FILE.name}"
    )

    if not EXCEL_FILE.exists():

        message = (
            f"Không tìm thấy file "
            f"'{EXCEL_FILE.name}'"
        )

        print(
            f"LỖI: {message}"
        )

        logging.error(message)

        return

    # --------------------------------------------------------
    # Đọc Excel bằng Pandas
    #
    # data_only=False:
    # đọc được công thức, nhưng không phù hợp để lấy kết quả.
    #
    # Pandas/openpyxl sẽ lấy cached value nếu WPS/Excel đã
    # tính và lưu file.
    # --------------------------------------------------------

    try:

        df = pd.read_excel(
            EXCEL_FILE,
            engine="openpyxl"
        )

        print(
            f"Đọc dữ liệu thành công: "
            f"{len(df)} dòng."
        )

        logging.info(
            f"Đọc Excel thành công - {len(df)} dòng"
        )

    except Exception as error:

        print(
            f"LỖI khi đọc Excel: {error}"
        )

        logging.exception(
            f"Lỗi khi đọc Excel: {error}"
        )

        return

    # --------------------------------------------------------
    # Kiểm tra cột
    # --------------------------------------------------------

    try:

        validate_columns(df)

    except Exception as error:

        print(
            f"LỖI cấu trúc Excel: {error}"
        )

        logging.error(
            f"Lỗi cấu trúc Excel: {error}"
        )

        return

    # --------------------------------------------------------
    # Xử lý từng dòng
    # --------------------------------------------------------

    success_count = 0
    failed_count = 0

    for index, row in df.iterrows():

        excel_row = index + 2

        print(
            f"\nĐang xử lý dòng {excel_row}..."
        )

        try:

            pdf_path = process_employee(
                row,
                excel_row
            )

            employee_id = str(
                row["Mã NV"]
            ).strip()

            employee_name = str(
                row["Tên NV"]
            ).strip()

            success_count += 1

            print(
                f"SUCCESS - {employee_id} - "
                f"{employee_name}"
            )

            print(
                f"          {pdf_path}"
            )

            logging.info(
                f"SUCCESS | Dòng {excel_row} | "
                f"{employee_id} | {employee_name} | "
                f"Tạo thành công: {pdf_path.name}"
            )

        except Exception as error:

            failed_count += 1

            try:
                employee_id = str(
                    row["Mã NV"]
                ).strip()
            except Exception:
                employee_id = "UNKNOWN"

            print(
                f"FAILED - Dòng {excel_row} - "
                f"{employee_id}"
            )

            print(
                f"         Lý do: {error}"
            )

            logging.error(
                f"FAILED | Dòng {excel_row} | "
                f"{employee_id} | Lý do: {error}"
            )

    # --------------------------------------------------------
    # Tổng kết
    # --------------------------------------------------------

    print()
    print("=" * 65)
    print("                 KẾT QUẢ XỬ LÝ")
    print("=" * 65)

    print(
        f"Tạo thành công : {success_count} file"
    )

    print(
        f"Thất bại       : {failed_count} dòng"
    )

    print(
        f"Thư mục PDF    : {OUTPUT_DIR}"
    )

    print(
        f"File log       : {LOG_FILE}"
    )

    print("=" * 65)

    logging.info(
        f"KẾT THÚC | Thành công: "
        f"{success_count} | Thất bại: {failed_count}"
    )


# ------------------------------------------------------------
# 17. Chạy chương trình
# ------------------------------------------------------------

if __name__ == "__main__":

    try:
        main()

    except KeyboardInterrupt:

        print(
            "\nChương trình đã được người dùng dừng."
        )

        logging.warning(
            "Chương trình bị dừng bởi người dùng."
        )

    except Exception as error:

        print(
            f"\nLỖI NGHIÊM TRỌNG: {error}"
        )

        logging.exception(
            f"Lỗi nghiêm trọng: {error}"
        )
	