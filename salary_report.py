# ============================================================
# BÀI A3 - TỰ ĐỘNG TẠO PHIẾU LƯƠNG PDF
# File: salary_report.py
#
# Luồng:
# employees.xlsx
#     -> Pandas đọc Excel
#     -> kiểm tra dữ liệu từng dòng
#     -> Python tính toán theo một nguồn công thức thống nhất
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
import re
from string import Template

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


# ------------------------------------------------------------
# 3. Đường dẫn
# ------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "employees.xlsx"
OUTPUT_DIR = BASE_DIR / "salary"
LOG_FILE = BASE_DIR / "process_log.txt"
TEMPLATE_FILE = BASE_DIR / "templates" / "salary_template.html"

# Thư mục đầu ra được tạo khi main() bắt đầu, không tạo khi import module.


# ------------------------------------------------------------
# 4. Thông tin phiếu
# ------------------------------------------------------------

COMPANY_NAME = "CÔNG TY BTL"

COMPANY_ADDRESS = (
    "Địa chỉ: 20A, 20 Lý Tự Trọng, Phường Sài Gòn, "
    "Thành phố Hồ Chí Minh, Việt Nam"
)

PAY_MONTH = os.environ.get("PAY_MONTH", datetime.now().strftime("%m/%Y"))

# ------------------------------------------------------------
# 5. Logging
# ------------------------------------------------------------

def configure_logging():
    """Cấu hình nhật ký khi chạy chương trình, không gây tác dụng phụ lúc import."""

    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        encoding="utf-8",
        force=True,
    )


from salary_calculator import (
    calculate_pit,
    calculate_salary,
    format_money,
    get_value,
    number_to_vietnamese,
    parse_number,
    round_money,
)


# ------------------------------------------------------------
# 7. HTML template rendering
# ------------------------------------------------------------


def create_salary_html(employee):
    """Render the salary slip template using validated, escaped values."""

    def safe_text(value):
        return html.escape(str(value))

    net_salary = employee["Thực lãnh"]
    values = {
        "company_name": safe_text(COMPANY_NAME),
        "company_address": safe_text(COMPANY_ADDRESS),
        "pay_month": safe_text(PAY_MONTH),
        "employee_id": safe_text(employee["Mã NV"]),
        "employee_name": safe_text(employee["Tên NV"]),
        "position": safe_text(employee["Chức vụ"]),
        "dependents": safe_text(employee["Số người phụ thuộc"]),
        "standard_days": safe_text(employee["Ngày công chuẩn"]),
        "actual_days": safe_text(employee["Ngày công đi làm"]),
        "insurance_salary": safe_text(format_money(employee["Lương đóng BHXH"])),
        "basic_salary": safe_text(format_money(employee["Lương cơ bản"])),
        "daily_salary": safe_text(format_money(employee["Lương theo ngày"])),
        "actual_salary": safe_text(format_money(employee["Lương thực tế"])),
        "bonus": safe_text(format_money(employee["Thưởng"])),
        "allowance": safe_text(format_money(employee["Phụ cấp"])),
        "total_income": safe_text(format_money(employee["Tổng thu nhập"])),
        "bhxh": safe_text(format_money(employee["BHXH"])),
        "bhyt": safe_text(format_money(employee["BHYT"])),
        "bhtn": safe_text(format_money(employee["BHTN"])),
        "taxable_income": safe_text(format_money(employee["Thu nhập tính thuế"])),
        "pit": safe_text(format_money(employee["Thuế TNCN"])),
        "penalty": safe_text(format_money(employee["Phạt"])),
        "total_deduction": safe_text(format_money(employee["Tổng khấu trừ"])),
        "net_salary": safe_text(format_money(net_salary)),
        "amount_in_word": safe_text(number_to_vietnamese(net_salary)),
    }

    if not TEMPLATE_FILE.is_file():
        raise FileNotFoundError(f"Không tìm thấy mẫu phiếu lương: {TEMPLATE_FILE}")

    return Template(TEMPLATE_FILE.read_text(encoding="utf-8")).substitute(values)


# ------------------------------------------------------------
# 14. Các cột bắt buộc
# ------------------------------------------------------------

def validate_columns(df):
    if df.empty:
        raise ValueError("Tệp Excel không có bản ghi nhân viên")

    required = [
        "Mã NV",
        "Tên NV",
        "Chức vụ",
        "Lương cơ bản",
        "Thưởng",
        "Phạt",
    ]

    missing = [column for column in required if column not in df.columns]

    if missing:
        raise ValueError("Thiếu cột bắt buộc: " + ", ".join(missing))


def validate_employee_ids(df):
    """Không cho phép mã NV trùng vì tên PDF được tạo theo mã NV."""

    ids = df["Mã NV"].fillna("").astype(str).str.strip()
    ids = ids[ids != ""]
    duplicates = ids[ids.duplicated(keep=False)]

    if not duplicates.empty:
        duplicated_ids = sorted(set(duplicates.tolist()))
        raise ValueError(
            "Mã NV bị trùng, có thể làm ghi đè phiếu lương: "
            + ", ".join(duplicated_ids)
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
        raise ValueError("Mã NV bị bỏ trống")

    if not re.fullmatch(r"[A-Za-z0-9_-]{1,64}", employee_id):
        raise ValueError(
            "Mã NV chỉ được chứa chữ cái không dấu, số, dấu gạch ngang "
            "hoặc gạch dưới (tối đa 64 ký tự)"
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
    # Tính tất cả trường dẫn xuất bằng cùng một bộ công thức.
    # Không trộn kết quả cached trong Excel với kết quả Python.
    # --------------------------------------------------------

    salary_result = calculate_salary(row)

    employee = row.copy()

    employee["Mã NV"] = employee_id
    employee["Tên NV"] = employee_name
    employee["Chức vụ"] = position

    for key, value in salary_result.items():
        employee[key] = value

    # Python là nguồn tính toán duy nhất để tránh trộn kết quả
    # cached từ Excel với các giá trị dự phòng và gây lệch tổng.

    # --------------------------------------------------------
    # Kiểm tra thực lãnh cuối cùng
    # --------------------------------------------------------

    employee["Thực lãnh"] = round_money(
        employee["Thực lãnh"]
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

    temporary_pdf = pdf_path.with_name(f".{employee_id}.tmp.pdf")

    try:
        from weasyprint import HTML

        HTML(
            string=html_content,
            base_url=str(BASE_DIR)
        ).write_pdf(str(temporary_pdf))

        if not temporary_pdf.is_file() or temporary_pdf.stat().st_size == 0:
            raise RuntimeError("Không tạo được file PDF hợp lệ")

        with temporary_pdf.open("rb") as pdf_file:
            if not pdf_file.read(5).startswith(b"%PDF-"):
                raise RuntimeError("Tệp đầu ra không có chữ ký PDF hợp lệ")

        temporary_pdf.replace(pdf_path)
    finally:
        if temporary_pdf.exists():
            temporary_pdf.unlink()

    return pdf_path


# ------------------------------------------------------------
# 16. Main
# ------------------------------------------------------------

def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    configure_logging()

    if not re.fullmatch(r"(0[1-9]|1[0-2])/\d{4}", PAY_MONTH):
        message = "PAY_MONTH phải có định dạng MM/YYYY, ví dụ 09/2026"
        print(f"LỖI: {message}")
        logging.error(message)
        return

    print()
    print("=" * 65)
    print("        CHƯƠNG TRÌNH TẠO PHIẾU LƯƠNG PDF")
    print("=" * 65)
    print()

    logging.info("=" * 65)
    logging.info("BẮT ĐẦU CHƯƠNG TRÌNH")

    try:
        from weasyprint import HTML
    except Exception as error:
        message = (
            "Không thể khởi tạo WeasyPrint. Hãy kiểm tra cài đặt "
            "gói Python và thư viện hệ thống cần thiết trên Windows."
        )
        print(f"LỖI: {message}")
        logging.exception("%s Chi tiết: %s", message, error)
        return

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
    # Đọc Excel bằng Pandas.
    # Chỉ dùng dữ liệu đầu vào; các cột kết quả tính sẵn trong
    # bảng tính không được dùng để ghi đè kết quả tính bằng Python.
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

        df.columns = [str(column).strip() for column in df.columns]
        validate_columns(df)
        validate_employee_ids(df)

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
	