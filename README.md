# Payslip Automation System

A Streamlit web application that automatically generates professional payslip PDFs and sends them to employees via email.

## Features

- 📤 Upload Excel payroll data via drag-and-drop or file picker
- 👀 Preview employee data before processing
- 🏢 Fully configurable company details (name, logo, footer text)
- 📧 SMTP email configuration with Gmail App Password support
- 🧪 Dry-run mode to test PDF generation without sending emails
- 📊 Real-time progress tracking during processing
- 📥 Download all generated PDFs as ZIP
- 📄 Export results as CSV

## Quick Start

### Prerequisites

- Python 3.8 or higher
- Gmail account with App Password enabled (for sending emails)

### Local Installation

1. **Clone or download this repository**

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   streamlit run app.py
   ```

4. **Open your browser**
   - The app will automatically open at `http://localhost:8501`

## How to Use

### Step 1: Download and Fill Config File

1. In the sidebar, expand "📥 Download Templates"
2. Click "📄 Company Config" to download `company_config.json`
3. Open the file in a text editor and fill in **ALL** fields:
   - **Company Details:**
     - `company_name`: Your company's full name
     - `footer_text`: Payslip disclaimer text
     - `document_id`: Internal reference number (optional)
     - `effectivity_date`: Format effective date (optional)
   - **SMTP Credentials:**
     - `smtp.email`: Your Gmail address
     - `smtp.password`: Your Gmail App Password (see guide below)
4. Save the file locally
5. **IMPORTANT:** Keep this file secure - it contains your email password!

#### How to Get Gmail App Password

1. **Enable 2-Factor Authentication**
   - Go to [Google Account Settings](https://myaccount.google.com/)
   - Navigate to Security
   - Enable 2-Step Verification

2. **Generate App Password**
   - In Security settings, search for "App passwords"
   - Select "Mail" and your device
   - Copy the 16-character password

3. **Use in config file**
   - Paste the app password in `smtp.password` field
   - Spaces are ok (e.g., "xxxx xxxx xxxx xxxx")
   - **Important:** Use the app password, NOT your regular Gmail password!

### Step 2: Upload Config File

1. In the sidebar, expand "⚙️ Settings"
2. Click "Upload Config File"
3. Select your filled `company_config.json`
4. Wait for validation:
   - ✅ Green checkmarks = Config loaded and SMTP validated
   - ❌ Red error = Check your config file (especially SMTP credentials)

### Step 3: Upload Company Logo (Optional)

1. In the Settings section, find "Logo (optional)"
2. Upload PNG/JPG file
3. Logo will appear on all payslips

### Step 4: Upload Payroll Excel File

1. In the main area, click "Choose an Excel file" or drag & drop
2. Download template from sidebar if you need sample format
3. Wait for validation to complete

### Step 5: Preview Employee Data

- Review employee list and totals
- Check that all data looks correct
- Verify email addresses are valid

### Step 6: Choose Processing Mode

- **Normal Mode**: Generate PDFs and send emails
- **Dry-Run Mode**: Generate PDFs only (for testing)

### Step 7: Start Processing

1. Click "🚀 Start Processing"
2. Watch the progress bar
3. Wait for completion message

### Step 8: Download Results

- Download results as CSV
- Download all PDFs as ZIP file
- Review any failures in the detailed results table

## Excel File Format

### Required Columns

Your Excel file must include these columns:

- `EmployeeNumber` - Unique employee identifier
- `Name` - Employee full name
- `Email` - Employee email address
- `PayrollPeriod` - Pay period (e.g., "January 1-15, 2025")
- `GrossIncome` - Total gross income
- `TotalDeductions` - Total deductions
- `NetPay` - Final net pay

### Optional Columns

For detailed payslips, include these columns:

**Basic Info:**
- `Position` - Job title

**Earnings:**
- `BasicSalary`, `Allowance`
- `RegularHours`, `RegularAmount`
- `RegularOTHours`, `RegularOTAmount`
- `LegalHolidayHours`, `LegalHolidayAmount`
- `SpecialHolidayHours`, `SpecialHolidayAmount`
- `NightDiffHours`, `NightDiffAmount`
- `OffsetHours`, `OffsetAmount`
- `PaidLeaveAmount`, `AdjustmentEarnings`
- `ThirteenthMonthPay`, `OthersEarnings`

**Deductions:**
- `SSSContribution`, `PhilhealthContribution`, `PagibigContribution`
- `PagibigLoan`, `SSSLoan`
- `WithholdingTax`, `AdjustmentDeductions`
- `OthersDeductions`, `OtherDeductions`

### Download Template

Click the "Download Template" button in the app to get a sample Excel file with the correct format.

## Project Structure

```
payslip-automation/
├── app.py                      # Main Streamlit application
├── utils/
│   ├── pdf_generator.py        # PDF generation logic
│   ├── email_sender.py         # Email sending functionality
│   ├── validators.py           # Data validation functions
│   └── excel_handler.py        # Excel file processing
├── config/
│   └── constants.py            # Configuration constants
├── assets/
│   └── logo.png                # Company logo
├── templates/
│   ├── payroll_template.xlsx  # Sample Excel template
│   └── company_config.json    # Company config template
├── .streamlit/
│   └── config.toml             # Streamlit configuration
├── requirements.txt            # Python dependencies
├── .gitignore                  # Git ignore rules
└── README.md                   # This file
```

## Security & Privacy

### Config File Security

⚠️ **CRITICAL:** Your `company_config.json` file contains sensitive information:
- Email credentials (SMTP password)
- Company details

**Best Practices:**
- ✅ Store config file in a secure location on your computer
- ✅ Use file permissions to restrict access (chmod 600 on Unix/Mac)
- ✅ Back up config file to secure location

### Data Handling

- **Company config file is stored LOCALLY on your computer** - not uploaded to any server
- **Each user uploads their own config** - no data is shared between users
- Config and SMTP credentials are only stored in session (cleared on browser close)
- Uploaded Excel files are processed in memory (not saved to disk)
- Generated PDFs are stored in temporary directory (auto-cleaned)
- No sensitive data is sent to external servers except:
  - Email sending via your Gmail SMTP
  - If deployed on Streamlit Cloud, files are processed on Streamlit's servers

## Troubleshooting

### "Authentication failed" error

- Make sure you're using a Gmail App Password, not your regular password
- Verify 2-Factor Authentication is enabled
- Check that the email address is correct

### "Missing required columns" error

- Download the template file and compare with your Excel
- Make sure column names match exactly (case-sensitive)
- Check for extra spaces in column headers

### "Invalid email format" error

- Verify all email addresses in the Excel are valid
- Remove any blank or invalid email entries
- Check for typos in email addresses

### PDFs not generating

- Check that all required columns have data
- Verify numeric columns (GrossIncome, NetPay) contain numbers
- Make sure PayrollPeriod is filled for all employees

### Emails not sending

- Test SMTP connection first
- Check your Gmail account for security alerts
- Verify App Password is correct
- Check that you haven't exceeded Gmail's sending limits (500 emails/day)

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Review the Gmail App Password guide in the sidebar
3. Download and review the template Excel file

## License

This project is provided as-is for payroll automation purposes.

## Version History

- **v1.0.0** (2025) - Initial release
  - Excel upload and validation
  - PDF generation with configurable company details
  - Email sending via SMTP
  - Dry-run mode
  - Results export and PDF download
