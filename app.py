import os
import json
from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from datetime import datetime, timedelta

app = Flask(__name__)
# Configure CORS to allow requests from the frontend domain, including POST and OPTIONS methods
CORS(app, resources={r"/api/*": {
    "origins": "https://email-frontend-eosin.vercel.app",
    "methods": ["GET", "POST", "OPTIONS"],
    "allow_headers": ["Content-Type"]
}})

# Temporary in-memory storage for company data
companies_data = []

# Path to store email_task state in /tmp
STATE_FILE = "/tmp/email_task.json"

# Load email_task from file if it exists, otherwise initialize
def load_email_task():
    try:
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, 'r') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading email_task: {str(e)}")
    return {
        "in_progress": False,
        "email": None,
        "password": None,
        "start_index": None,
        "end_index": None,
        "current_index": None,
        "sent_emails": 0,
        "total_emails_to_send": 0,
        "error": None
    }

# Save email_task to file
def save_email_task(email_task):
    try:
        with open(STATE_FILE, 'w') as f:
            json.dump(email_task, f)
    except Exception as e:
        print(f"Error saving email_task: {str(e)}")

# Load initial email_task state
email_task = load_email_task()

@app.route('/api/upload-excel', methods=['POST'])
def upload_excel():
    global companies_data
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400

        print(f"Received file: {file.filename}")  # Debug log
        # Read the Excel file
        df = pd.read_excel(file, engine='openpyxl')
        print(f"Excel file read successfully. Columns: {df.columns.tolist()}")  # Debug log

        # Ensure required columns exist
        required_columns = ['Company', 'Patent Number', 'Email', 'First Name', 'Response']
        if not all(col in df.columns for col in required_columns):
            return jsonify({'error': 'Missing required columns'}), 400

        # Group by company and aggregate data
        grouped = df.groupby('Company').agg({
            'Patent Number': lambda x: list(x),
            'Email': lambda x: list(x),
            'First Name': lambda x: list(x),
            'Response': 'first'
        }).reset_index()

        # Convert to list of dictionaries for easier handling
        companies_data = grouped.to_dict('records')
        print(f"Processed {len(companies_data)} companies: {companies_data}")  # Debug log
        return jsonify({
            'message': 'File processed successfully',
            'total_companies': len(companies_data)
        })
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")  # Debug log
        return jsonify({'error': str(e)}), 500

@app.route('/api/send-emails', methods=['POST'])
def send_emails():
    global email_task
    try:
        email_task = load_email_task()
        if email_task["in_progress"]:
            return jsonify({'error': 'Email sending is already in progress'}), 400

        data = request.json
        email = data.get('email', os.getenv('EMAIL'))
        password = data.get('password', os.getenv('PASSWORD'))
        start_index = data.get('startIndex')
        end_index = data.get('endIndex')

        if not email or not password:
            return jsonify({'error': 'Email and password are required'}), 400

        if not companies_data:
            return jsonify({'error': 'No company data available. Upload an Excel file first.'}), 400

        if start_index < 0 or end_index >= len(companies_data) or start_index > end_index:
            return jsonify({'error': 'Invalid index range'}), 400

        # Calculate total emails to send
        total_emails_to_send = 0
        for idx in range(start_index, end_index + 1):
            company = companies_data[idx]
            emails = company['Email']
            response = company.get('Response', '')
            if isinstance(response, str) and response.lower() == 'yes':
                continue
            valid_emails = [email for email in emails if isinstance(email, str) and '@' in email]
            if not valid_emails:
                continue
            valid_first_names = company['First Name'][:len(valid_emails)]
            if not valid_first_names:
                continue
            total_emails_to_send += len(valid_emails)

        # Initialize the email task
        email_task = {
            "in_progress": True,
            "email": email,
            "password": password,
            "start_index": start_index,
            "end_index": end_index,
            "current_index": start_index,
            "sent_emails": 0,
            "total_emails_to_send": total_emails_to_send,
            "error": None
        }
        save_email_task(email_task)

        # Process the first company immediately
        return process_next_company()
    except Exception as e:
        print(f"Error starting email sending: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/process-next', methods=['POST'])
def process_next():
    return process_next_company()

def process_next_company():
    global email_task
    try:
        email_task = load_email_task()
        if not email_task["in_progress"]:
            return jsonify({'error': 'No email sending task in progress'}), 400

        current_index = email_task["current_index"]
        if current_index > email_task["end_index"]:
            email_task["in_progress"] = False
            save_email_task(email_task)
            return jsonify({
                'message': f'Finished sending {email_task["sent_emails"]} emails',
                'status': 'completed',
                'sent_emails': email_task["sent_emails"],
                'total_emails': email_task["total_emails_to_send"]
            })

        # Process the current company
        company = companies_data[current_index]
        company_name = company['Company']
        emails = company['Email']
        first_names = company['First Name']
        patents = company['Patent Number']
        response = company.get('Response', '')

        # Filter invalid emails (must contain '@')
        valid_emails = [email for email in emails if isinstance(email, str) and '@' in email]
        sent_emails = 0

        if not valid_emails:
            print(f"Skipping company {company_name} due to no valid emails.")
        else:
            # Match first names to valid emails
            valid_first_names = first_names[:len(valid_emails)]

            # Combine names for greeting (e.g., "John, Jane & Alice")
            if len(valid_first_names) > 1:
                names_list = ', '.join(valid_first_names[:-1]) + ' & ' + valid_first_names[-1]
            else:
                names_list = valid_first_names[0] if valid_first_names else ''

            # Skip if no valid names
            if not names_list:
                print(f"Skipping company {company_name} due to no valid names.")
            else:
                # Handle patents: Convert to strings, remove NaN, take top 2
                patents = [str(patent) for patent in patents if not pd.isna(patent)]
                patents = patents[:2]  # Take top 2 patents
                patents_str = ', '.join(patents) if patents else 'No patent information available'

                # Skip if response is 'yes'
                if isinstance(response, str) and response.lower() == 'yes':
                    print(f"Skipping company {company_name} because response is 'yes'.")
                else:
                    # Determine email content based on response
                    follow_up_date = datetime(2024, 11, 27) + timedelta(days=15)
                    current_date = datetime.now()

                    # Set subject and body based on response
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
                        print(f"Skipping company {company_name} due to response or date condition.")
                        email_task["current_index"] += 1
                        save_email_task(email_task)
                        return jsonify({
                            'message': f'Skipped company {company_name} due to response or date condition',
                            'status': 'in_progress',
                            'current_index': email_task["current_index"],
                            'sent_emails': email_task["sent_emails"],
                            'total_emails': email_task["total_emails_to_send"]
                        })

                    # Send the email using smtplib with SMTP_SSL on port 465
                    retries = 2
                    for attempt in range(retries):
                        try:
                            server = smtplib.SMTP_SSL('smtp.office365.com', 465)
                            print(f"Attempting to login with email: {email_task['email']} (Attempt {attempt + 1}/{retries})")
                            server.login(email_task['email'], email_task['password'])
                            print("Login successful")

                            # Create email message
                            msg = MIMEMultipart("alternative")
                            msg['Subject'] = subject
                            msg['From'] = formataddr(("Bayslope Business Solutions", email_task['email']))
                            msg['To'] = ', '.join(valid_emails)
                            msg.attach(MIMEText(html, 'html'))

                            print(f"Sending email to {valid_emails} for company {company_name}")
                            server.sendmail(email_task['email'], valid_emails, msg.as_string())
                            print(f"Email sent successfully to {valid_emails} for company {company_name}.")
                            sent_emails = len(valid_emails)
                            email_task["sent_emails"] += sent_emails
                            break  # Success, exit retry loop
                        except Exception as e:
                            print(f"Error sending email (Attempt {attempt + 1}/{retries}): {str(e)}")
                            if attempt == retries - 1:
                                email_task["error"] = str(e)
                                email_task["in_progress"] = False
                                save_email_task(email_task)
                                return jsonify({'error': str(e)}), 500
                        finally:
                            try:
                                server.quit()
                            except:
                                pass  # Ignore errors during server quit

        # Move to the next company
        email_task["current_index"] += 1
        save_email_task(email_task)
        return jsonify({
            'message': f'Successfully sent {sent_emails} emails for company {company_name}' if sent_emails > 0 else f'Skipped company {company_name}',
            'status': 'in_progress',
            'current_index': email_task["current_index"],
            'sent_emails': email_task["sent_emails"],
            'total_emails': email_task["total_emails_to_send"]
        })
    except Exception as e:
        print(f"Error processing email: {str(e)}")
        email_task["error"] = str(e)
        email_task["in_progress"] = False
        save_email_task(email_task)
        return jsonify({'error': str(e)}), 500

@app.route('/api/email-status', methods=['GET'])
def email_status():
    global email_task
    email_task = load_email_task()
    if not email_task["in_progress"] and email_task["error"]:
        return jsonify({
            'status': 'failed',
            'error': email_task["error"],
            'sent_emails': email_task["sent_emails"],
            'total_emails': email_task["total_emails_to_send"]
        })
    elif not email_task["in_progress"]:
        return jsonify({
            'status': 'completed',
            'sent_emails': email_task["sent_emails"],
            'total_emails': email_task["total_emails_to_send"]
        })
    else:
        return jsonify({
            'status': 'in_progress',
            'current_index': email_task["current_index"],
            'start_index': email_task["start_index"],
            'end_index': email_task["end_index"],
            'sent_emails': email_task["sent_emails"],
            'total_emails': email_task["total_emails_to_send"]
        })

# Export the app for Vercel
app = app