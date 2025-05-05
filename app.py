import os
import urllib.parse
import uuid
import traceback
from flask import Flask, request, jsonify, make_response
from flask_cors import CORS
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
import time
import sys
import openpyxl

app = Flask(__name__)
# Configure CORS to allow requests from the frontend domain
CORS(app, resources={r"/api/*": {
    "origins": ["https://email-frontend-eosin.vercel.app", "http://localhost:3000"],
    "methods": ["GET", "POST", "OPTIONS"],
    "allow_headers": ["Content-Type", "Authorization"],
    "expose_headers": ["Content-Type"],
    "max_age": 86400
}})

# Ensure CORS headers are applied to all responses, including errors
@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = 'https://email-frontend-eosin.vercel.app'
    response.headers['Access-Control-Allow-Methods'] = 'GET,POST,OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type,Authorization'
    return response

# Health check endpoint to verify CORS and backend status
@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({
        'status': 'healthy',
        'message': 'Backend is running',
        'python_version': sys.version,
        'openpyxl_version': openpyxl.__version__,
        'pandas_version': pd.__version__
    })

# Get SMTP credentials from environment variables
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")

if not SMTP_USER or not SMTP_PASSWORD:
    raise ValueError("SMTP_USER and SMTP_PASSWORD must be set as environment variables")

# Temporary in-memory storage for company data
companies_data = []

@app.route('/api/upload-excel', methods=['POST'])
def upload_excel():
    global companies_data
    try:
        print("Received request to upload Excel file")
        if 'file' not in request.files:
            print("No file part in request")
            response = jsonify({'error': 'No file part'})
            response.status_code = 400
            return response

        file = request.files['file']
        if file.filename == '':
            print("No file selected")
            response = jsonify({'error': 'No selected file'})
            response.status_code = 400
            return response

        if not file.filename.endswith('.xlsx'):
            print(f"Invalid file format: {file.filename}")
            response = jsonify({'error': 'File must be an .xlsx file'})
            response.status_code = 400
            return response

        print(f"Received file: {file.filename}")
        # Check file size (Vercel has a 4.5MB limit for Hobby plan)
        file.seek(0, os.SEEK_END)
        file_size = file.tell()
        file.seek(0)  # Reset file pointer to the beginning
        if file_size > 4.5 * 1024 * 1024:  # 4.5MB in bytes
            print(f"File too large: {file_size} bytes")
            response = jsonify({'error': 'File size exceeds 4.5MB limit'})
            response.status_code = 400
            return response

        # Sanitize filename and save to /tmp with a unique name
        unique_filename = f"{uuid.uuid4()}.xlsx"
        temp_file_path = f"/tmp/{unique_filename}"
        try:
            file.save(temp_file_path)
            print(f"Saved file temporarily to {temp_file_path}")
        except Exception as e:
            print(f"Error saving file to {temp_file_path}: {str(e)}")
            print(traceback.format_exc())
            response = jsonify({'error': f'Failed to save file: {str(e)}'})
            response.status_code = 500
            return response

        # Read the Excel file
        try:
            df = pd.read_excel(temp_file_path, engine='openpyxl')
            print(f"Excel file read successfully. Columns: {df.columns.tolist()}")
        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")
            print(traceback.format_exc())
            raise
        finally:
            # Ensure the temporary file is always removed
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
                print(f"Removed temporary file: {temp_file_path}")

        required_columns = ['Company', 'Patent Number', 'Email', 'First Name', 'Response']
        if not all(col in df.columns for col in required_columns):
            print(f"Missing required columns. Found: {df.columns.tolist()}, Required: {required_columns}")
            response = jsonify({'error': 'Missing required columns'})
            response.status_code = 400
            return response

        # Handle potential NaN values in the DataFrame
        df = df.fillna('')

        try:
            grouped = df.groupby('Company').agg({
                'Patent Number': lambda x: list(x),
                'Email': lambda x: list(x),
                'First Name': lambda x: list(x),
                'Response': 'first'
            }).reset_index()
        except Exception as e:
            print(f"Error grouping data: {str(e)}")
            print(traceback.format_exc())
            response = jsonify({'error': f'Failed to process data: {str(e)}'})
            response.status_code = 500
            return response

        companies_data = grouped.to_dict('records')
        print(f"Processed {len(companies_data)} companies: {companies_data}")
        return jsonify({
            'message': 'File processed successfully',
            'total_companies': len(companies_data)
        })
    except ImportError as e:
        print(f"ImportError: {str(e)}")
        print(traceback.format_exc())
        response = jsonify({'error': f'Missing dependency: {str(e)}'})
        response.status_code = 500
        return response
    except pd.errors.EmptyDataError as e:
        print(f"EmptyDataError: {str(e)}")
        print(traceback.format_exc())
        response = jsonify({'error': 'Excel file is empty or invalid'})
        response.status_code = 400
        return response
    except openpyxl.utils.exceptions.InvalidFileException as e:
        print(f"InvalidFileException: {str(e)}")
        print(traceback.format_exc())
        response = jsonify({'error': 'Invalid Excel file format'})
        response.status_code = 400
        return response
    except ValueError as e:
        print(f"ValueError: {str(e)}")
        print(traceback.format_exc())
        response = jsonify({'error': f'Value error in processing: {str(e)}'})
        response.status_code = 400
        return response
    except Exception as e:
        print(f"Unexpected error processing Excel file: {str(e)}")
        print(traceback.format_exc())
        response = jsonify({'error': f'Unexpected error: {str(e)}'})
        response.status_code = 500
        return response

@app.route('/api/send-emails', methods=['POST'])
def send_emails():
    try:
        data = request.json
        start_index = data.get('startIndex')
        end_index = data.get('endIndex')

        if not companies_data:
            response = jsonify({'error': 'No company data available. Upload an Excel file first.'})
            response.status_code = 400
            return response

        if start_index < 0 or end_index >= len(companies_data) or start_index > end_index:
            response = jsonify({'error': 'Invalid index range'})
            response.status_code = 400
            return response

        email_tasks = []
        total_emails = 0

        # Initialize SMTP server connection
        server = smtplib.SMTP("smtp.office365.com", 587)
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)

        for idx in range(start_index, end_index + 1):
            company = companies_data[idx]
            company_name = company['Company']
            emails = company['Email']
            first_names = company['First Name']
            patents = company['Patent Number']
            response = company.get('Response', '')

            valid_emails = [email for email in emails if isinstance(email, str) and '@' in email]
            if not valid_emails:
                email_tasks.append({'company': company_name, 'status': 'skipped', 'reason': 'No valid emails'})
                continue

            valid_first_names = first_names[:len(valid_emails)]
            if not valid_first_names:
                email_tasks.append({'company': company_name, 'status': 'skipped', 'reason': 'No valid names'})
                continue

            if len(valid_first_names) > 1:
                names_list = ', '.join(valid_first_names[:-1]) + ' & ' + valid_first_names[-1]
            else:
                names_list = valid_first_names[0]

            patents = [str(patent) for patent in patents if not pd.isna(patent)]
            patents = patents[:2]
            patents_str = ', '.join(patents) if patents else 'No patent information available'

            if isinstance(response, str) and response.lower() == 'yes':
                email_tasks.append({'company': company_name, 'status': 'skipped', 'reason': 'Response is yes'})
                continue

            follow_up_date = datetime(2024, 11, 27) + timedelta(days=15)
            current_date = datetime.now()

            if pd.isna(response) or response == '':
                subject = f"Patent Monetization Interest for {patents_str} etc."
                html = f"""
                <html>
                <head>
                <meta charset="UTF-8">
                <title>Patent Monetization Interest for {patents_str}</title>
                </head>
                <body>
                <p style="font-size: 10.5pt;">
                Hi {names_list},<br><br>
                Hope all is well at your end.<br><br>
                Our internal framework has identified patents {patents_str} etc. and we think there is a monetization opportunity for them.<br><br>
                We work closely with a network of active buyers who regularly acquire high-quality patents for monetization across various technology sectors.<br><br>
                Could you help facilitate a discussion with your client about this matter?<br><br>
                <p style="font-size: 10.5pt;">Best regards,</p>
                <p style="font-size: 10.5pt;">
                <span style="color: black;">Sarita (Sara) /
                <a href="https://bayslope.com/" style="color: rgb(208, 0, 0); text-decoration: none;">Baysl</span><span style="color: rgb(169, 169, 169);">o</span><span style="color: rgb(208, 0, 0); text-decoration: none;">pe</span>
                </a>
                </span><br>
                <span style="color: black; text-decoration: underline;">
                <a href="https://techreport99.com/" style="color: rgb(208, 0, 0); text-decoration: underline;">Techreport99</a></span> <span style="color: rgb(169, 169, 169);"> | </span>
                <a href="https://bayslope.com/" style="color: rgb(208, 0, 0); text-decoration: underline;">Baysl</span><span style="color: rgb(169, 169, 169);">o</span><span style="color: rgb(208, 0, 0); text-decoration: none;">pe</span>
                </a>
                </span>
                </p>
                e: <a href="mailto:patents@bayslope.com">patents@bayslope.com</a><br>
                p: +91-9811967160 (IN), +1 650 353 7723 (US), +44 1392 58 1535 (UK)
                </p>
                <p style="color: grey; font-size: 8.5pt; font-family: Arial, sans-serif;">
                The content of this email message and any attachments are intended solely for the addressee(s) and may contain confidential and/or privileged information and may be legally protected from disclosure. If you are not the intended recipient of this email, or if this email message has been addressed to you in error, please immediately alert the sender by reply email and then delete this message and any attachments. If you are not the intended recipient, you may not copy, store or deliver this message to anyone, without a written consent of the sender. Thank you!
                </p>
                </body>
                </html>
                """
            elif isinstance(response, str) and response.lower() == 'no' and current_date >= follow_up_date:
                subject = f"Follow-up: Patent Acquisition Interest"
                html = f"""
                <html>
                <head>
                <meta charset="UTF-8">
                <title>Follow-up: Patent Acquisition Interest</title>
                </head>
                <body>
                <p style="font-size: 10.5pt;">Hi {names_list},</p>
                <p style="font-size: 10.5pt;">Hope all is well at your end.</p>
                <p style="font-size: 10.5pt;">We understand your busy schedule so didn’t mean to bother you via this email. Just checking if you could assist in facilitating a discussion with your client.</p>
                <p style="font-size: 10.5pt;">It will be great to hear from you.</p>
                <p style="font-size: 10.5pt;">Best regards,</p>
                <p style="font-size: 10.5pt;" >
                <span style="color: black;">Sarita (Sara) /
                <a href="https://bayslope.com/" style="color: rgb(208, 0, 0); text-decoration: none;">Baysl</span><span style="color: rgb(208, 206, 206);">o</span><span style="color: rgb(208, 0, 0); text-decoration: none;">pe</span>
                </a>
                </span><br>
                <span style="color: black; text-decoration: underline;">
                <a href="https://techreport99.com/" style="color: rgb(208, 0, 0); text-decoration: underline;">Techreport99</a></span> <span style="color: rgb(208, 206, 206);"> | </span>
                <a href="https://bayslope.com/" style="color: rgb(208, 0, 0); text-decoration: underline;">Baysl</span><span style="color: rgb(208, 206, 206);">o</span><span style="color: rgb(208, 0, 0); text-decoration: none;">pe</span>
                </a>
                </span>
                </p>
                <p>
                e: <a href="mailto:patents@bayslope.com">patents@bayslope.com</a><br>
                p: +91-9811967160 (IN), +1 650 353 7723 (US), +44 1392 58 1535 (UK)
                </p>
                <p style="color: grey; font-size: 8.5pt; font-family: Arial, sans-serif;">
                The content of this email message and any attachments are intended solely for the addressee(s) and may contain confidential and/or privileged information and may be legally protected from disclosure. If you are not the intended recipient of this email, or if this email message has been addressed to you in error, please immediately alert the sender by reply email and then delete this message and any attachments. If you are not the intended recipient, you may not copy, store or deliver this message to anyone, without a written consent of the sender. Thank you!
                </p>
                </body>
                </html>
                """
            else:
                email_tasks.append({'company': company_name, 'status': 'skipped', 'reason': 'Response or date condition not met'})
                continue

            # Send the email via SMTP
            message = MIMEMultipart("alternative")
            from_email = SMTP_USER
            message["Subject"] = subject
            message["From"] = formataddr(("Bayslope Business Solutions", from_email))
            message["To"] = ', '.join(valid_emails)
            message.attach(MIMEText(html, "html"))

            try:
                print(f"Attempting to send email to {valid_emails} for company {company_name}...")
                server.sendmail(from_email, valid_emails, message.as_string())
                print(f"Email sent successfully to {valid_emails} for company {company_name}.")
                email_tasks.append({
                    'company': company_name,
                    'status': 'sent',
                    'recipients': valid_emails
                })
                total_emails += len(valid_emails)
                # Short delay to avoid rate limiting
                time.sleep(5)
            except Exception as e:
                print(f"Error sending email to {valid_emails}: {str(e)}")
                email_tasks.append({
                    'company': company_name,
                    'status': 'failed',
                    'reason': str(e)
                })

        # Close the SMTP connection
        server.quit()

        return jsonify({
            'message': f'Processed {len(email_tasks)} email tasks, sent {total_emails} emails',
            'email_tasks': email_tasks,
            'total_emails': total_emails
        })
    except Exception as e:
        print(f"Error sending emails: {str(e)}")
        response = jsonify({'error': str(e)})
        response.status_code = 500
        return response

# Export the app for Vercel
app = app