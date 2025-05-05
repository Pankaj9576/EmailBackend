import os
from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from datetime import datetime, timedelta
import time

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Temporary in-memory storage for company data
companies_data = []

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
    try:
        data = request.json
        email = data.get('email', os.getenv('EMAIL'))  # Use environment variable if email not provided
        password = data.get('password', os.getenv('PASSWORD'))  # Use environment variable if password not provided
        start_index = data.get('startIndex')
        end_index = data.get('endIndex')

        if not email or not password:
            return jsonify({'error': 'Email and password are required'}), 400

        if not companies_data:
            return jsonify({'error': 'No company data available. Upload an Excel file first.'}), 400

        if start_index < 0 or end_index >= len(companies_data) or start_index > end_index:
            return jsonify({'error': 'Invalid index range'}), 400

        sent_emails = 0
        for idx in range(start_index, end_index + 1):
            company = companies_data[idx]
            company_name = company['Company']
            emails = company['Email']
            first_names = company['First Name']
            patents = company['Patent Number']
            response = company.get('Response', '')

            # Filter invalid emails (must contain '@')
            valid_emails = [email for email in emails if isinstance(email, str) and '@' in email]
            if not valid_emails:
                print(f"Skipping company {company_name} due to no valid emails.")
                continue

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
                continue

            # Handle patents: Convert to strings, remove NaN, take top 2
            patents = [str(patent) for patent in patents if not pd.isna(patent)]
            patents = patents[:2]  # Take top 2 patents
            patents_str = ', '.join(patents) if patents else 'No patent information available'

            # Skip if response is 'yes'
            if isinstance(response, str) and response.lower() == 'yes':
                print(f"Skipping company {company_name} because response is 'yes'.")
                continue

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
                continue

            # Create a new SMTP connection for each company to avoid timeout
            retries = 2
            for attempt in range(retries):
                try:
                    server = smtplib.SMTP('smtp.office365.com', 587)
                    server.starttls()
                    print(f"Attempting to login with email: {email} (Attempt {attempt + 1}/{retries})")
                    server.login(email, password)
                    print("Login successful")

                    # Create email message
                    msg = MIMEMultipart("alternative")
                    msg['Subject'] = subject
                    msg['From'] = formataddr(("Bayslope Business Solutions", email))
                    msg['To'] = ', '.join(valid_emails)
                    msg.attach(MIMEText(html, 'html'))

                    print(f"Sending email to {valid_emails} for company {company_name}")
                    server.sendmail(email, valid_emails, msg.as_string())
                    print(f"Email sent successfully to {valid_emails} for company {company_name}.")
                    sent_emails += len(valid_emails)
                    break  # Success, exit retry loop
                except Exception as e:
                    print(f"Error sending email (Attempt {attempt + 1}/{retries}): {str(e)}")
                    if attempt == retries - 1:
                        raise e  # Raise the error after all retries fail
                finally:
                    try:
                        server.quit()
                    except:
                        pass  # Ignore errors during server quit

            if idx < end_index:  # Don't sleep after the last email
                print("Waiting for 2 minutes before sending the next emails...")
                time.sleep(120)

        return jsonify({'message': f'Successfully sent {sent_emails} emails'})
    except Exception as e:
        print(f"Error sending emails: {str(e)}")
        return jsonify({'error': str(e)}), 500

# Export the app for Vercel
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT', 5000)))