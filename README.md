Automated PDF Salary Report Generation

Python automation system for generating employee salary-slip PDF reports from Excel data.

The system reads employee information from an Excel file, calculates salary components, generates dynamic PDF salary slips and records the processing results in a log file.

Features
Read employee salary data from Excel using Pandas
Validate input data before processing
Calculate salary based on working days
Calculate insurance deductions and personal income tax
Generate dynamic HTML/CSS salary slips
Export each employee salary slip as a PDF
Automatically create the salary/ output folder
Generate process_log.txt
Skip invalid employee records without stopping the entire process
Record successful and failed processing results
Technologies
Python
Pandas
OpenPyXL
WeasyPrint
HTML
CSS
Project Structure
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
Requirements

Recommended environment:

Python 3.12

Install the required packages:

pip install pandas openpyxl weasyprint
Input Data

The program reads employee information from:

employees.xlsx

The Excel file contains employee-related information such as:

Column	Description
Mã NV	Employee ID
Tên NV	Employee name
Chức vụ	Position
Lương cơ bản	Basic salary
Lương đóng BHXH	Salary used for social insurance
Ngày công chuẩn	Standard working days
Ngày công đi làm	Actual working days
Thưởng	Bonus
Phụ cấp	Allowance
Số người phụ thuộc	Number of dependents
Phạt	Penalty
Salary Calculation

The system calculates salary based on the employee's actual working days.

Lương ngày
    │
    ▼
Lương cơ bản / Ngày công chuẩn
    │
    ▼
Lương theo ngày công
    │
    ├── Thưởng
    ├── Phụ cấp
    │
    ▼
Tổng thu nhập
    │
    ├── BHXH
    ├── BHYT
    ├── BHTN
    ├── Thuế TNCN
    └── Phạt
    │
    ▼
Thực lĩnh

For example:

Lương thực tế =
Lương cơ bản / Ngày công chuẩn × Ngày công tính lương

The system distinguishes between actual working days and paid leave when calculating the final salary.

Generate PDF

After processing valid employee records, the system automatically generates PDF salary slips.

Output directory:

salary/

Each PDF is named using the employee ID:

[Mã NV].pdf

Example:

salary/
├── NV001.pdf
├── NV002.pdf
└── NV003.pdf
PDF Salary Slip

Each generated PDF contains information such as:

Company information
Employee information
Salary month
Basic salary
Working days
Bonus
Allowance
Social insurance
Health insurance
Unemployment insurance
Personal income tax
Penalty
Net salary
Salary amount in words
Logging

The system automatically creates:

process_log.txt

The log records:

Successfully generated salary slips
Failed records
Invalid input data
Reasons for processing failures

Example:

[2026-09-18 08:30:12] SUCCESS - NV001 - PDF generated successfully
[2026-09-18 08:30:13] SUCCESS - NV002 - PDF generated successfully
[2026-09-18 08:30:13] ERROR - Row 5 - Invalid salary value
Error Handling

The program uses try...except to prevent one invalid employee record from stopping the entire process.

Read Excel
    │
    ▼
Validate employee data
    │
    ├── Invalid ──► Log error ──► Skip row
    │
    ▼
Calculate salary
    │
    ▼
Generate PDF
    │
    ▼
Save PDF
    │
    ▼
Write success log
Run the Program

Place the following files in the same project directory:

PDF-Automation/
├── employees.xlsx
└── salary_report.py

Run:

python salary_report.py

After execution, the generated salary slips will be stored in:

salary/

and the processing log will be stored in:

process_log.txt
Project Workflow

The complete automation workflow is:

employees.xlsx
       │
       ▼
Read Excel Data
       │
       ▼
Validate Data
       │
       ▼
Calculate Salary
       │
       ▼
Generate HTML/CSS
       │
       ▼
Convert HTML to PDF
       │
       ▼
Save [Mã NV].pdf
       │
       ▼
Write process_log.txt
Project Status

Current version:

V1.0 — Automated Employee Salary PDF Generation

The project currently focuses on automatically generating employee salary-slip PDFs from structured Excel data.
