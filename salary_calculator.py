"""Core payroll parsing and calculation functions."""

import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

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

def round_money(value):
    """Làm tròn tiền đến đồng theo quy tắc half-up."""

    number = Decimal(str(value))

    if not number.is_finite():
        raise ValueError("Giá trị tiền phải là số hữu hạn")

    return number.quantize(Decimal("1"), rounding=ROUND_HALF_UP)


def format_money(value):
    """Đổi số thành dạng 15.000.000 VNĐ."""

    try:
        number = round_money(value)
        return f"{number:,.0f}".replace(",", ".") + " VNĐ"
    except (ValueError, TypeError, ArithmeticError):
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
    """Chuẩn hóa số nhập từ Excel, chấp nhận dấu phân cách hàng nghìn và thập phân."""

    if value is None or pd.isna(value):
        raise ValueError(f"{field_name} đang bị bỏ trống")

    if isinstance(value, str):
        normalized = (
            value.strip()
            .replace("\u00a0", "")
            .replace(" ", "")
            .replace("VNĐ", "")
            .replace("vnđ", "")
            .replace("VND", "")
            .replace("vnd", "")
        )

        if not normalized:
            raise ValueError(f"{field_name} đang bị bỏ trống")

        has_dot = "." in normalized
        has_comma = "," in normalized

        if has_dot and has_comma:
            if normalized.rfind(",") > normalized.rfind("."):
                normalized = normalized.replace(".", "").replace(",", ".")
            else:
                normalized = normalized.replace(",", "")
        elif has_dot or has_comma:
            separator = "," if has_comma else "."
            if normalized.count(separator) > 1:
                normalized = normalized.replace(separator, "")
            else:
                left, right = normalized.split(separator)
                if len(right) == 3 and left.lstrip("+-").isdigit():
                    normalized = left + right
                else:
                    normalized = left + "." + right

        value = normalized

    try:
        number = Decimal(str(value))
    except Exception as error:
        raise ValueError(f"{field_name} không phải số hợp lệ: {value}") from error

    if not number.is_finite():
        raise ValueError(f"{field_name} phải là số hữu hạn")

    if number < 0:
        raise ValueError(f"{field_name} không được âm")

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
        number = int(round_money(number))
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
# 11. Tính thuế TNCN lũy tiến theo kỳ tính thuế năm 2026
# ------------------------------------------------------------
# Python là nguồn tính toán duy nhất để tránh phụ thuộc vào cached
# formula có thể cũ hoặc không nhất quán trong tệp Excel.
# Các quy tắc trong chương trình là mô hình học tập, cần được
# đối chiếu với quy định áp dụng trước khi dùng trong thực tế.
# ------------------------------------------------------------

def calculate_pit(taxable_income):
    """Tính thuế TNCN lũy tiến 5 bậc theo mô hình áp dụng cho kỳ tính thuế 2026."""

    x = Decimal(str(taxable_income))
    if not x.is_finite():
        raise ValueError("Thu nhập tính thuế phải là số hữu hạn")
    x = max(x, Decimal("0"))

    brackets = [
        (Decimal("10000000"), Decimal("0.05")),
        (Decimal("30000000"), Decimal("0.10")),
        (Decimal("60000000"), Decimal("0.20")),
        (Decimal("100000000"), Decimal("0.30")),
        (Decimal("Infinity"), Decimal("0.35")),
    ]

    tax = Decimal("0")
    previous = Decimal("0")

    for upper, rate in brackets:
        if x <= previous:
            break

        taxable_part = min(x, upper) - previous
        tax += taxable_part * rate
        previous = upper

        if x <= upper:
            break

    return round_money(tax)


# ------------------------------------------------------------
# 12. Tính lương từ dữ liệu đầu vào
# ------------------------------------------------------------
# Các giá trị đầu vào lấy từ Excel; các khoản tính toán được
# xác định trong Python để giữ một nguồn công thức thống nhất.
# ------------------------------------------------------------

def calculate_salary(row):
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

    if dependents != dependents.to_integral_value():
        raise ValueError("Số người phụ thuộc phải là số nguyên không âm")

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

    actual_salary = round_money(
        daily_salary * actual_days
    )

    total_income = (
        actual_salary
        + bonus
        + allowance
    )

    bhxh = round_money(insurance_salary * Decimal(str(BHXH_RATE)))
    bhyt = round_money(insurance_salary * Decimal(str(BHYT_RATE)))
    bhtn = round_money(insurance_salary * Decimal(str(BHTN_RATE)))

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
        "Số người phụ thuộc": int(dependents),
        "Phạt": penalty,
    }
