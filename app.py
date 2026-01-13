"""
Payslip Automation System - Streamlit Application
Generate and email payslips automatically from Excel data
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import tempfile
import shutil
import zipfile
import io
import json
import chardet
import atexit
from pathlib import Path
from datetime import datetime

from utils.pdf_generator import create_payslip_pdf
from utils.email_sender import EmailSender
from utils.validators import validate_excel_data, test_smtp_connection
from utils.excel_handler import load_excel_file
from config.constants import (
    REQUIRED_COLUMNS, SMTP_SERVER, SMTP_PORT,
    DEFAULT_COMPANY_NAME, DEFAULT_FOOTER_TEXT,
    DEFAULT_DOCUMENT_ID, DEFAULT_EFFECTIVITY_DATE,
    DEFAULT_LOGO_PATH
)

# ---------- TEMP FILE CLEANUP ----------

def cleanup_old_temp_dirs():
    """Clean up old temporary directories created by this app"""
    try:
        temp_base = Path(tempfile.gettempdir())
        current_time = datetime.now()

        # Find all temp directories older than 24 hours
        for temp_dir in temp_base.glob("tmp*"):
            if temp_dir.is_dir():
                try:
                    # Check if directory is older than 24 hours
                    dir_mtime = datetime.fromtimestamp(temp_dir.stat().st_mtime)
                    age_hours = (current_time - dir_mtime).total_seconds() / 3600

                    # If older than 24 hours and contains payslip PDFs, clean it up
                    if age_hours > 24:
                        pdf_files = list(temp_dir.glob("payslip_*.pdf"))
                        if pdf_files:  # Only delete if it looks like our temp dir
                            shutil.rmtree(temp_dir, ignore_errors=True)
                except (OSError, PermissionError):
                    # Skip directories we can't access
                    pass
    except Exception:
        # Don't fail app startup if cleanup fails
        pass

def cleanup_temp_dir(temp_dir_path):
    """Safely cleanup a temporary directory"""
    if temp_dir_path and Path(temp_dir_path).exists():
        try:
            # Only cleanup if it's in the system temp directory (safety check)
            if str(Path(temp_dir_path).parent) == tempfile.gettempdir():
                shutil.rmtree(temp_dir_path, ignore_errors=True)
        except Exception:
            pass

# Cleanup old temp directories on app startup
cleanup_old_temp_dirs()

# Register cleanup on exit
@atexit.register
def cleanup_on_exit():
    """Cleanup temp directory when app exits"""
    if hasattr(st.session_state, 'temp_dir') and st.session_state.temp_dir:
        cleanup_temp_dir(st.session_state.temp_dir)

# ---------- PAGE CONFIGURATION ----------

st.set_page_config(
    page_title="Payslip Automation System",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------- HIDE STREAMLIT BRANDING ----------

st.markdown("""
<style>
    /* Hide Fork button and GitHub icon */
    [data-testid="stToolbarActionButton"] {
        display: none !important;
    }

    /* Hide footer */
    footer {
        visibility: hidden !important;
    }
</style>
""", unsafe_allow_html=True)

# ---------- COMPANY CONFIG HELPERS ----------

def load_company_config(uploaded_file):
    """Load company configuration from uploaded JSON file"""
    try:
        # Read raw bytes and auto-detect encoding
        raw_bytes = uploaded_file.read()
        detected = chardet.detect(raw_bytes)
        encoding = detected.get('encoding', 'utf-8') or 'utf-8'

        content = raw_bytes.decode(encoding)
        config_data = json.loads(content)

        # Validate SMTP nested structure
        smtp_config = config_data.get('smtp', {})
        if not smtp_config.get('email') or not smtp_config.get('password'):
            return False, None, "SMTP configuration incomplete. Both email and password are required. Please download the latest config template."

        # Extract valid fields (flatten SMTP for easier session state management)
        valid_config = {
            'company_name': config_data.get('company_name', ''),
            'footer_text': config_data.get('footer_text', ''),
            'document_id': config_data.get('document_id', ''),
            'effectivity_date': config_data.get('effectivity_date', ''),
            'smtp_email': smtp_config.get('email', ''),
            'smtp_password': smtp_config.get('password', '')
        }

        return True, valid_config, "Configuration loaded successfully!"

    except json.JSONDecodeError:
        return False, None, "Invalid JSON file. Please check the file format."
    except UnicodeDecodeError:
        return False, None, "Unable to read file encoding. Please try re-saving the file."
    except Exception as e:
        return False, None, f"Error loading config: {str(e)}"

# ---------- SESSION STATE INITIALIZATION ----------

def init_session_state():
    """Initialize session state variables"""
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'smtp_email' not in st.session_state:
        st.session_state.smtp_email = ""
    if 'smtp_password' not in st.session_state:
        st.session_state.smtp_password = ""
    if 'smtp_validated' not in st.session_state:
        st.session_state.smtp_validated = False
    if 'temp_dir' not in st.session_state:
        st.session_state.temp_dir = None
    if 'processing_results' not in st.session_state:
        st.session_state.processing_results = None
    if 'company_name' not in st.session_state:
        st.session_state.company_name = DEFAULT_COMPANY_NAME
    if 'footer_text' not in st.session_state:
        st.session_state.footer_text = DEFAULT_FOOTER_TEXT
    if 'document_id' not in st.session_state:
        st.session_state.document_id = DEFAULT_DOCUMENT_ID
    if 'effectivity_date' not in st.session_state:
        st.session_state.effectivity_date = DEFAULT_EFFECTIVITY_DATE
    if 'company_logo_path' not in st.session_state:
        st.session_state.company_logo_path = DEFAULT_LOGO_PATH
    if 'config_loaded' not in st.session_state:
        st.session_state.config_loaded = False
    if 'output_directory' not in st.session_state:
        st.session_state.output_directory = ""

init_session_state()

# ---------- SIDEBAR ----------

with st.sidebar:
    st.title("Payslip Automation")

    # Templates Section
    with st.expander("📥 Download Templates", expanded=False):
        # Company config template
        st.markdown("**Company Config Template**")
        config_template_path = "templates/company_config.json"
        if Path(config_template_path).exists():
            with open(config_template_path, "rb") as template_file:
                st.download_button(
                    label="📄 Company Config",
                    data=template_file,
                    file_name="company_config.json",
                    mime="application/json",
                    width='stretch',
                    help="Download template to fill in your company details"
                )

        # Payroll template
        st.markdown("**Payroll Excel Template**")
        payroll_template_path = "templates/payroll_template.xlsx"
        if Path(payroll_template_path).exists():
            with open(payroll_template_path, "rb") as template_file:
                st.download_button(
                    label="📊 Payroll Excel",
                    data=template_file,
                    file_name="payroll_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width='stretch',
                    help="Download Excel template for payroll data"
                )

    # Combined Settings expander
    with st.expander("⚙️ Settings", expanded=False):
        # Configuration Upload
        st.subheader("Configuration")

        # Config file upload
        config_file = st.file_uploader(
            "Upload Config File",
            type=['json'],
            help="Upload your company_config.json file (download template above)",
            key="company_config_uploader"
        )

        # Load config if uploaded
        if config_file is not None:
            # Only process if not already loaded (prevent re-processing on every rerun)
            if not st.session_state.get('config_loaded', False):
                success, config_data, message = load_company_config(config_file)
                if success:
                    # Load company details
                    st.session_state.company_name = config_data['company_name']
                    st.session_state.footer_text = config_data['footer_text']
                    st.session_state.document_id = config_data['document_id']
                    st.session_state.effectivity_date = config_data['effectivity_date']

                    # Load SMTP credentials
                    st.session_state.smtp_email = config_data['smtp_email']
                    st.session_state.smtp_password = config_data['smtp_password']

                    # Auto-validate SMTP connection
                    with st.spinner("Validating SMTP connection..."):
                        smtp_success, smtp_message = test_smtp_connection(
                            config_data['smtp_email'],
                            config_data['smtp_password'],
                            SMTP_SERVER,
                            SMTP_PORT
                        )
                        st.session_state.smtp_validated = smtp_success

                    st.session_state.config_loaded = True

                    # Show one-time message based on SMTP validation
                    if not smtp_success:
                        st.error(f"❌ SMTP validation failed: {smtp_message}")
                        st.info("💡 Check your SMTP credentials in the config file")
                else:
                    st.error(f"❌ {message}")
                    st.session_state.config_loaded = False

        # Show current config if loaded
        if st.session_state.get('config_loaded', False) and st.session_state.company_name:
            st.divider()

            # Display loaded configuration summary
            col1, col2 = st.columns([3, 1])
            with col1:
                st.caption("**Loaded Configuration:**")
                st.caption(f"🏢 {st.session_state.company_name}")
                st.caption(f"📧 {st.session_state.smtp_email}")
            with col2:
                if st.button("Clear", width='stretch', help="Clear configuration"):
                    # Clear all config-related session state
                    st.session_state.company_name = ""
                    st.session_state.footer_text = ""
                    st.session_state.document_id = ""
                    st.session_state.effectivity_date = ""
                    st.session_state.smtp_email = ""
                    st.session_state.smtp_password = ""
                    st.session_state.smtp_validated = False
                    st.session_state.config_loaded = False
                    st.rerun()

        st.divider()

        # Company logo (always visible)
        company_logo = st.file_uploader(
            "Logo (optional)",
            type=['png', 'jpg', 'jpeg'],
            help="Appears at top of payslip"
        )

    # Gmail App Password Guide
    with st.expander("ℹ️ Gmail App Password Guide"):
        st.markdown("""
        **1. Enable 2-Factor Authentication**
        - Go to [Google Account](https://myaccount.google.com/)
        - Navigate to Security
        - Enable 2-Step Verification

        **2. Generate App Password**
        - Search for "App passwords"
        - Select "Mail" and your device
        - Copy the 16-character password

        **3. Use in config file**
        - Paste the app password in the `smtp.password` field of your config file
        - Note: Use app password, not regular password!
        - Spaces are ok (e.g., "xxxx xxxx xxxx xxxx")
        """)


# ---------- MAIN AREA ----------

st.title("📧 Payslip Automation")

# File Upload
uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=['xlsx'],
    help="Upload your payroll Excel file (download template from sidebar)"
)

# Process uploaded file
if uploaded_file is not None:
    try:
        # Load and validate data
        df = load_excel_file(uploaded_file)
        is_valid, validation_errors = validate_excel_data(df, REQUIRED_COLUMNS)

        if is_valid:
            st.session_state.df = df
            st.success(f"✅ File loaded successfully! Found {len(df)} employee(s)")

            # Data Preview (collapsible)
            with st.expander("📊 Data Preview", expanded=True):
                # Display key metrics (reduced to 2)
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Employees", len(df))
                with col2:
                    payroll_period = df['PayrollPeriod'].iloc[0] if 'PayrollPeriod' in df.columns else "N/A"
                    st.metric("Payroll Period", payroll_period)

                # Display all columns
                df_display = df.copy()

                # Format currency columns (common ones)
                currency_columns = ['GrossIncome', 'TotalDeductions', 'NetPay',
                                  'BasicPay', 'Allowance', 'Overtime', 'Deductions']
                for col in currency_columns:
                    if col in df_display.columns:
                        df_display[col] = df_display[col].apply(lambda x: f"₱{x:,.2f}")

                st.dataframe(df_display, width='stretch', height=300)

            # Send Options
            dry_run = st.checkbox(
                "🧪 Dry run (generate PDFs only)",
                value=False,
                help="Test PDF generation without sending emails"
            )

            # Output directory selection for dry run
            if dry_run:
                output_dir = st.text_input(
                    "📁 Output Directory",
                    value=st.session_state.output_directory,
                    placeholder="/path/to/save/pdfs",
                    help="Enter the full path where PDFs will be saved (e.g., /Users/yourname/Desktop/payslips)"
                )
                if output_dir != st.session_state.output_directory:
                    st.session_state.output_directory = output_dir

            # Validation before sending
            can_send = True
            missing_requirements = []

            if not st.session_state.config_loaded:
                missing_requirements.append("Upload config file in sidebar")

            if dry_run:
                # Validate output directory for dry run
                if not st.session_state.output_directory:
                    missing_requirements.append("Enter output directory path for dry run")
                elif not Path(st.session_state.output_directory).exists():
                    missing_requirements.append(f"Output directory does not exist: {st.session_state.output_directory}")
                elif not Path(st.session_state.output_directory).is_dir():
                    missing_requirements.append(f"Output path is not a directory: {st.session_state.output_directory}")

            if not dry_run and not st.session_state.smtp_validated:
                missing_requirements.append("SMTP validation failed - check config file")

            if missing_requirements:
                st.warning("⚠️ **Requirements missing:**")
                for req in missing_requirements:
                    st.warning(f"  • {req}")
                can_send = False

            # Send button
            if st.button("🚀 Start Processing", type="primary", width='stretch', disabled=not can_send and not dry_run):
                # Determine output directory for PDFs
                if dry_run:
                    # Use user-specified directory for dry run
                    st.session_state.temp_dir = st.session_state.output_directory
                else:
                    # Create temp directory for normal mode
                    if st.session_state.temp_dir is None:
                        st.session_state.temp_dir = tempfile.mkdtemp()

                # Determine logo path
                logo_path = DEFAULT_LOGO_PATH
                if company_logo is not None:
                    # Save custom logo to temp
                    custom_logo_path = Path(st.session_state.temp_dir) / "custom_logo.png"
                    with open(custom_logo_path, "wb") as f:
                        f.write(company_logo.getbuffer())
                    logo_path = str(custom_logo_path)

                # Prepare company config
                company_config = {
                    'company_name': st.session_state.company_name,
                    'footer_text': st.session_state.footer_text,
                    'document_id': st.session_state.document_id,
                    'effectivity_date': st.session_state.effectivity_date
                }

                # Process payslips
                results = []

                # Progress tracking
                progress_bar = st.progress(0)
                status_text = st.empty()

                # Initialize email sender if not dry run
                email_sender = None
                if not dry_run:
                    email_sender = EmailSender(
                        st.session_state.smtp_email,
                        st.session_state.smtp_password,
                        SMTP_SERVER,
                        SMTP_PORT
                    )
                    success, message = email_sender.connect()
                    if not success:
                        st.error(f"Failed to connect to SMTP server: {message}")
                        st.stop()

                # Process each employee
                quota_exceeded = False
                for idx, (_, row) in enumerate(df.iterrows()):
                    try:
                        # Update progress
                        progress = (idx + 1) / len(df)
                        progress_bar.progress(progress)
                        status_text.text(f"Processing {row['Name']} ({idx + 1}/{len(df)})")

                        # Generate PDF
                        pdf_path = create_payslip_pdf(
                            row,
                            output_dir=st.session_state.temp_dir,
                            logo_path=logo_path if Path(logo_path).exists() else None,
                            company_config=company_config
                        )

                        # Send email if not dry run
                        if not dry_run and email_sender:
                            success, message, quota_exceeded = email_sender.send_payslip(row, pdf_path)

                            # Check if quota was exceeded
                            if quota_exceeded:
                                results.append({
                                    'Employee': row['Name'],
                                    'Email': row['Email'],
                                    'Status': 'Quota Exceeded',
                                    'Message': 'Gmail daily limit reached - email not sent',
                                    'PDF': pdf_path
                                })
                                # Stop processing immediately
                                status_text.text(f"⚠️ Gmail quota exceeded at employee {idx + 1}/{len(df)}")
                                break

                            results.append({
                                'Employee': row['Name'],
                                'Email': row['Email'],
                                'Status': 'Sent' if success else 'Failed',
                                'Message': message,
                                'PDF': pdf_path
                            })
                        else:
                            results.append({
                                'Employee': row['Name'],
                                'Email': row['Email'],
                                'Status': 'Generated',
                                'Message': 'PDF created (dry run mode)',
                                'PDF': pdf_path
                            })

                    except Exception as e:
                        results.append({
                            'Employee': row.get('Name', 'Unknown'),
                            'Email': row.get('Email', 'Unknown'),
                            'Status': 'Error',
                            'Message': str(e),
                            'PDF': None
                        })

                # Cleanup email connection
                if email_sender:
                    email_sender.disconnect()

                # Clear progress
                progress_bar.empty()
                status_text.empty()

                # Store results in session state
                st.session_state.processing_results = pd.DataFrame(results)

                # Show completion message
                if quota_exceeded:
                    # Count how many were successfully sent
                    sent_count = len([r for r in results if r['Status'] == 'Sent'])
                    remaining_count = len(df) - sent_count
                    st.error(f"⚠️ Gmail daily sending limit reached!")
                    st.warning(f"📊 Sent: {sent_count} | Remaining: {remaining_count}")
                    st.info("💡 Wait 24 hours and re-run to send remaining payslips, or upload an Excel with only the remaining employees.")
                elif dry_run:
                    st.success(f"✅ Dry run completed! PDFs saved to: {st.session_state.temp_dir}")
                else:
                    st.success("✅ Processing completed!")
                st.rerun()

        else:
            st.error("❌ Excel file validation failed!")
            st.write("")  # Add spacing

            # Check if there are missing column errors
            has_missing_columns = any("Missing required columns:" in error for error in validation_errors)

            if has_missing_columns:
                st.warning("**Column Name Mismatch Detected**")
                st.write("Your Excel file is missing required columns. Please ensure column names match EXACTLY (case-sensitive).")
                st.write("")

            # Display errors
            for error in validation_errors:
                if "Missing required columns:" in error:
                    # Extract column names from error message
                    missing_cols = error.replace("Missing required columns: ", "")
                    st.error(f"**Missing columns:** {missing_cols}")
                else:
                    st.error(f"• {error}")

            # Show helpful instructions
            if has_missing_columns:
                st.write("")
                st.info("💡 **How to fix:**\n"
                       "1. Download the Excel template from the sidebar (📥 Download Templates)\n"
                       "2. Compare your column names with the template\n"
                       "3. Rename your columns to match exactly\n"
                       "4. Column names are case-sensitive (e.g., 'EmployeeNumber' not 'Employee Number')")

            st.write("")

    except Exception as e:
        st.error(f"❌ Error loading file: {str(e)}")

# Display results if available
if st.session_state.processing_results is not None:
    results_df = st.session_state.processing_results

    # Results (collapsible)
    with st.expander("📋 Results", expanded=True):
        # Results summary
        col1, col2, col3 = st.columns(3)
        with col1:
            success_count = len(results_df[results_df['Status'].isin(['Sent', 'Generated'])])
            st.metric("✅ Successful", success_count)
        with col2:
            failed_count = len(results_df[results_df['Status'] == 'Failed'])
            st.metric("❌ Failed", failed_count)
        with col3:
            error_count = len(results_df[results_df['Status'] == 'Error'])
            st.metric("⚠️ Errors", error_count)

        # Display results without PDF path column
        display_df = results_df[['Employee', 'Email', 'Status', 'Message']].copy()
        st.dataframe(display_df, width='stretch')

        # Download options
        col1, col2, col3 = st.columns(3)

        with col1:
            # Download results as CSV
            csv = results_df[['Employee', 'Email', 'Status', 'Message']].to_csv(index=False)
            st.download_button(
                label="📄 CSV",
                data=csv,
                file_name="payslip_results.csv",
                mime="text/csv",
                width='stretch'
            )

        with col2:
            # Download all PDFs as ZIP
            if st.button("📦 ZIP", width='stretch'):
                # Create ZIP file
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for _, result in results_df.iterrows():
                        if result['PDF'] and Path(result['PDF']).exists():
                            zip_file.write(
                                result['PDF'],
                                arcname=Path(result['PDF']).name
                            )

                st.download_button(
                    label="💾 Download",
                    data=zip_buffer.getvalue(),
                    file_name="payslips.zip",
                    mime="application/zip",
                    width='stretch'
                )

        with col3:
            # Clear results and temp files
            if st.button("🗑️ Clear", width='stretch'):
                # Clean up temp directory (but not if it's a user-specified dry-run directory)
                if st.session_state.temp_dir:
                    # Only auto-cleanup if it's in system temp (not user's dry-run directory)
                    if str(Path(st.session_state.temp_dir).parent) == tempfile.gettempdir():
                        cleanup_temp_dir(st.session_state.temp_dir)
                    st.session_state.temp_dir = None

                # Clear results
                st.session_state.processing_results = None
                st.rerun()

# ---------- HIDE STREAMLIT CLOUD BRANDING ----------
# Hide profile preview & Streamlit Cloud badge (must be at end of app)
components.html("""
<script>
    const topDoc = window.top.document;

    // Inject CSS into top document
    const css = `
        [class*="_profilePreview_"] { display: none !important; }
        [class*="_profileContainer_"] { display: none !important; }
        a[href*="streamlit.io/cloud"] { display: none !important; }
        [class*="_viewerBadge_"] { display: none !important; }
    `;
    const style = document.createElement('style');
    style.textContent = css;
    try { topDoc.head.appendChild(style); } catch(e) {}

    // Hide elements via JS (backup)
    function hideElements() {
        try {
            topDoc.querySelectorAll('[class*="_profilePreview_"], [class*="_profileContainer_"]').forEach(el => el.style.display = 'none');
            topDoc.querySelectorAll('a[href*="streamlit.io/cloud"]').forEach(el => el.style.display = 'none');
            topDoc.querySelectorAll('[class*="_viewerBadge_"]').forEach(el => el.style.display = 'none');
        } catch(e) {}
    }

    setInterval(hideElements, 500);
</script>
""", height=0, scrolling=False)
