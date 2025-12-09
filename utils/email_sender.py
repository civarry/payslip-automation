"""Email sending functionality for payslips"""

import smtplib
from email.message import EmailMessage
from pathlib import Path
from typing import Tuple


class EmailSender:
    """Handle SMTP email sending for payslips"""

    def __init__(self, email: str, password: str,
                 smtp_server: str = "smtp.gmail.com",
                 smtp_port: int = 587):
        """
        Initialize email sender

        Args:
            email: SMTP email address
            password: SMTP password (App Password for Gmail)
            smtp_server: SMTP server address
            smtp_port: SMTP port number
        """
        self.email = email
        self.password = password.replace(" ", "")  # Remove spaces from app password
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.smtp = None

    def connect(self) -> Tuple[bool, str]:
        """
        Establish SMTP connection

        Returns:
            Tuple[bool, str]: (success, message)
        """
        try:
            self.smtp = smtplib.SMTP(self.smtp_server, self.smtp_port)
            self.smtp.starttls()
            self.smtp.login(self.email, self.password)
            return True, "Connected successfully"
        except smtplib.SMTPAuthenticationError:
            return False, "Authentication failed. Check your email and app password."
        except smtplib.SMTPException as e:
            return False, f"SMTP error: {str(e)}"
        except Exception as e:
            return False, f"Connection error: {str(e)}"

    def disconnect(self):
        """Close SMTP connection"""
        if self.smtp:
            try:
                self.smtp.quit()
            except:
                pass  # Ignore errors when closing
            self.smtp = None

    def send_payslip(self, row, pdf_path: str) -> Tuple[bool, str]:
        """
        Send payslip email with PDF attachment

        Args:
            row: pandas Series with employee data (must have Email, Name, PayrollPeriod columns)
            pdf_path: Path to PDF file

        Returns:
            Tuple[bool, str]: (success, message)
        """
        try:
            # Check if PDF exists
            if not Path(pdf_path).exists():
                return False, f"PDF file not found: {pdf_path}"

            # Check if SMTP connection is established
            if not self.smtp:
                return False, "Not connected to SMTP server"

            # Build email message
            to_email = row["Email"]
            name = row["Name"]
            period = row["PayrollPeriod"]

            msg = EmailMessage()
            msg["Subject"] = f"Payslip for {period}"
            msg["From"] = self.email
            msg["To"] = to_email

            body = (
                f"Hi {name},\n\n"
                f"Please find attached your payslip for {period}.\n\n"
                "This is a system-generated email. If you have any questions, "
                "please contact HR.\n\n"
                "Best regards,\n"
                "HR Department"
            )
            msg.set_content(body)

            # Attach PDF
            with open(pdf_path, "rb") as f:
                pdf_data = f.read()

            filename = Path(pdf_path).name
            msg.add_attachment(
                pdf_data,
                maintype="application",
                subtype="pdf",
                filename=filename
            )

            # Send email
            self.smtp.send_message(msg)
            return True, f"Sent to {to_email}"

        except Exception as e:
            return False, f"Error: {str(e)}"
