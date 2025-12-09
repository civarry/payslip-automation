"""Validation functions for Excel data and SMTP configuration"""

import pandas as pd
import smtplib
import re
from typing import Tuple, List


def validate_email(email: str) -> bool:
    """
    Validate email format using regex

    Args:
        email: Email address to validate

    Returns:
        bool: True if email format is valid
    """
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, str(email)) is not None


def validate_excel_data(df: pd.DataFrame, required_columns: List[str]) -> Tuple[bool, List[str]]:
    """
    Validate Excel data structure and content

    Args:
        df: pandas DataFrame from uploaded Excel
        required_columns: List of required column names

    Returns:
        Tuple[bool, List[str]]: (is_valid, list of error messages)
    """
    errors = []

    # Check for empty dataframe
    if df.empty:
        errors.append("Excel file contains no data rows")
        return False, errors

    # Check required columns
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        errors.append(f"Missing required columns: {', '.join(missing_cols)}")

    # Check for missing emails
    if 'Email' in df.columns:
        missing_emails = df[df['Email'].isna() | (df['Email'] == '')]
        if not missing_emails.empty:
            employee_names = missing_emails['Name'].astype(str).tolist()[:5]
            errors.append(
                f"Missing email addresses for {len(missing_emails)} employee(s): "
                f"{', '.join(employee_names)}" +
                (f" and {len(missing_emails) - 5} more" if len(missing_emails) > 5 else "")
            )

        # Validate email formats
        invalid_emails = []
        for idx, row in df.iterrows():
            if pd.notna(row.get('Email')) and row['Email'] != '':
                if not validate_email(str(row['Email'])):
                    employee_name = row.get('Name', f'Row {idx + 1}')
                    invalid_emails.append(f"{employee_name} ({row['Email']})")

        if invalid_emails:
            errors.append(
                f"Invalid email format for {len(invalid_emails)} employee(s): "
                f"{', '.join(invalid_emails[:5])}" +
                (f" and {len(invalid_emails) - 5} more" if len(invalid_emails) > 5 else "")
            )

    # Check for missing critical data
    critical_fields = ['EmployeeNumber', 'Name', 'PayrollPeriod', 'NetPay']
    for field in critical_fields:
        if field in df.columns:
            missing = df[df[field].isna() | (df[field] == '')]
            if not missing.empty:
                errors.append(f"Missing {field} for {len(missing)} employee(s)")

    # Check for duplicate employee numbers
    if 'EmployeeNumber' in df.columns:
        duplicates = df[df.duplicated(subset=['EmployeeNumber'], keep=False)]
        if not duplicates.empty:
            dup_numbers = duplicates['EmployeeNumber'].unique().tolist()[:5]
            errors.append(
                f"Duplicate employee numbers found: {', '.join(map(str, dup_numbers))}"
            )

    return len(errors) == 0, errors


def test_smtp_connection(email: str, password: str,
                        smtp_server: str = "smtp.gmail.com",
                        smtp_port: int = 587) -> Tuple[bool, str]:
    """
    Test SMTP connection with provided credentials

    Args:
        email: Email address
        password: App password
        smtp_server: SMTP server address
        smtp_port: SMTP port

    Returns:
        Tuple[bool, str]: (success, message)
    """
    try:
        # Remove spaces from password (Gmail app passwords often have spaces)
        password = password.replace(" ", "")

        # Validate email format first
        if not validate_email(email):
            return False, "Invalid email format"

        # Test SMTP connection
        smtp = smtplib.SMTP(smtp_server, smtp_port, timeout=10)
        smtp.starttls()
        smtp.login(email, password)
        smtp.quit()
        return True, "SMTP connection successful!"

    except smtplib.SMTPAuthenticationError:
        return False, "Authentication failed. Check your email and app password."
    except smtplib.SMTPException as e:
        return False, f"SMTP error: {str(e)}"
    except TimeoutError:
        return False, "Connection timeout. Check your internet connection."
    except Exception as e:
        return False, f"Connection error: {str(e)}"
