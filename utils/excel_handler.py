"""Excel file processing for payroll data"""

import pandas as pd
from io import BytesIO


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
        'BasicSalary', 'Allowance',
        'RegularHours', 'RegularAmount',
        'RegularOTHours', 'RegularOTAmount',
        'LegalHolidayHours', 'LegalHolidayAmount',
        'SpecialHolidayHours', 'SpecialHolidayAmount',
        'NightDiffHours', 'NightDiffAmount',
        'OffsetHours', 'OffsetAmount',
        'PaidLeaveAmount', 'AdjustmentEarnings',
        'ThirteenthMonthPay', 'OthersEarnings',
        'GrossIncome',
        'SSSContribution', 'PhilhealthContribution', 'PagibigContribution',
        'PagibigLoan', 'SSSLoan',
        'WithholdingTax', 'AdjustmentDeductions',
        'OthersDeductions', 'OtherDeductions',
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
