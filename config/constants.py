"""Configuration constants for payslip automation"""

# Required Excel columns for validation
REQUIRED_COLUMNS = [
    'EmployeeNumber',
    'Name',
    'Email',
    'PayrollPeriod',
    'GrossIncome',
    'TotalDeductions',
    'NetPay'
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
