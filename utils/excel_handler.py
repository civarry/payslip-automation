"""Excel file processing for payroll data"""

import pandas as pd
from io import BytesIO


class ColumnValidationError(Exception):
    """Custom exception for column validation errors"""
    pass


def validate_excel_columns(df: pd.DataFrame) -> tuple[bool, list, list]:
    """
    Validate that the Excel file has all required columns

    Args:
        df: DataFrame to validate

    Returns:
        tuple: (is_valid, missing_columns, extra_columns)
    """
    required_columns = [
        'EmployeeNumber', 'Name', 'Position', 'PayrollPeriod',
        'BasicSalary', 'MonthlyAllowance', 'Allowance',
        'RegularHours', 'RegularAmount',
        'RegularOTHours', 'RegularOTAmount',
        'LegalHolidayHours', 'LegalHolidayAmount',
        'SpecialHolidayHours', 'SpecialHolidayAmount',
        'NightDiffHours', 'NightDiffAmount',
        'OffsetHours', 'OffsetAmount',
        'PaidLeaveHours', 'PaidLeaveAmount',
        'AdjustmentEarnings', 'ThirteenthMonthPay',
        'OthersEarnings', 'GrossIncome',
        'SSSContribution', 'PhilhealthContribution', 'PagibigContribution',
        'PagibigLoan', 'SSSLoan', 'WithholdingTax',
        'AdjustmentDeductions', 'OthersDeductions',
        'TotalDeductions', 'NetPay'
    ]

    # Optional columns that can be in the file but aren't required for PDF
    optional_columns = ['Email']

    actual_columns = set(df.columns)
    required_set = set(required_columns)
    optional_set = set(optional_columns)

    # Find missing required columns
    missing = sorted(list(required_set - actual_columns))

    # Find extra columns (not required and not optional)
    extra = sorted(list(actual_columns - required_set - optional_set))

    is_valid = len(missing) == 0

    return is_valid, missing, extra


def load_excel_file(uploaded_file, sheet_name: str = "Sheet1") -> pd.DataFrame:
    """
    Load Excel file from Streamlit UploadedFile object

    Args:
        uploaded_file: Streamlit UploadedFile object
        sheet_name: Name of the sheet to read (default: "Sheet1")

    Returns:
        pd.DataFrame: Loaded data with cleaned columns and proper data types
    """
    # Read Excel file
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    # Clean column names (remove extra spaces and special characters)
    df.columns = df.columns.str.strip()

    # Convert numeric columns to proper data types
    numeric_columns = [
        'BasicSalary', 'MonthlyAllowance', 'Allowance',
        'RegularHours', 'RegularAmount',
        'RegularOTHours', 'RegularOTAmount',
        'LegalHolidayHours', 'LegalHolidayAmount',
        'SpecialHolidayHours', 'SpecialHolidayAmount',
        'NightDiffHours', 'NightDiffAmount',
        'OffsetHours', 'OffsetAmount',
        'PaidLeaveHours', 'PaidLeaveAmount',
        'AdjustmentEarnings', 'ThirteenthMonthPay', 'OthersEarnings',
        'GrossIncome',
        'SSSContribution', 'PhilhealthContribution', 'PagibigContribution',
        'PagibigLoan', 'SSSLoan',
        'WithholdingTax', 'AdjustmentDeductions', 'OthersDeductions',
        'TotalDeductions', 'NetPay'
    ]

    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Clean string columns (remove extra spaces)
    string_columns = ['EmployeeNumber', 'Name', 'Position', 'Email', 'PayrollPeriod']
    for col in string_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    return df
