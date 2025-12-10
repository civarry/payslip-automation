"""Configuration constants for payslip automation"""

# Required Excel columns for validation
REQUIRED_COLUMNS = [
    # Employee Info
    'EmployeeNumber', 'Name', 'Position', 'Email', 'PayrollPeriod',

    # Salary
    'BasicSalary', 'MonthlyAllowance', 'Allowance',

    # Regular Work
    'RegularHours', 'RegularAmount',
    'RegularOTHours', 'RegularOTAmount',

    # Holidays
    'LegalHolidayHours', 'LegalHolidayAmount',
    'SpecialHolidayHours', 'SpecialHolidayAmount',

    # Other Earnings
    'NightDiffHours', 'NightDiffAmount',
    'OffsetHours', 'OffsetAmount',
    'PaidLeaveHours', 'PaidLeaveAmount',
    'AdjustmentEarnings', 'ThirteenthMonthPay', 'OthersEarnings',

    # Totals
    'GrossIncome',

    # Deductions
    'SSSContribution', 'PhilhealthContribution', 'PagibigContribution',
    'PagibigLoan', 'SSSLoan', 'WithholdingTax',
    'AdjustmentDeductions', 'OthersDeductions',
    'TotalDeductions', 'NetPay'
]

# SMTP Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Default Company Details (empty for security)
DEFAULT_COMPANY_NAME = ""
DEFAULT_FOOTER_TEXT = ""
DEFAULT_DOCUMENT_ID = ""
DEFAULT_EFFECTIVITY_DATE = ""

# File Upload Settings
MAX_FILE_SIZE_MB = 10
ALLOWED_EXTENSIONS = ['.xlsx']

# PDF Configuration
DEFAULT_LOGO_PATH = "assets/logo.png"
