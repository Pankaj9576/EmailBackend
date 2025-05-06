import os
import urllib.parse
import uuid
from flask import Flask, request, jsonify, make_response
from flask_cors import CORS
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app, resources={r"/api/*": {
    "origins": ["https://email-frontend-eosin.vercel.app", "http://localhost:3000"],
    "methods": ["GET", "POST", "OPTIONS"],
    "allow_headers": ["Content-Type", "Authorization"],
    "expose_headers": ["Content-Type"],
    "max_age": 86400
}})

@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = 'https://email-frontend-eosin.vercel.app'
    response.headers['Access-Control-Allow-Methods'] = 'GET,POST,OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type,Authorization'
    return response

@app.before_request
def handle_options():
    if request.method == "OPTIONS":
        response = make_response()
        response.headers["Access-Control-Allow-Origin"] = request.headers.get("Origin", "https://email-frontend-eosin.vercel.app")
        response.headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type,Authorization"
        response.headers["Access-Control-Max-Age"] = "86400"
        return response, 200

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
        if not file.filename.endswith('.xlsx'):
            return jsonify({'error': 'File must be an .xlsx file'}), 400

        file.seek(0, os.SEEK_END)
        file_size = file.tell()
        file.seek(0)
        if file_size > 4.5 * 1024 * 1024:
            return jsonify({'error': 'File size exceeds 4.5MB limit'}), 400

        unique_filename = f"{uuid.uuid4()}.xlsx"
        temp_file_path = f"/tmp/{unique_filename}"
        file.save(temp_file_path)

        df = pd.read_excel(temp_file_path, engine='openpyxl')
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

        required_columns = ['Company', 'Patent Number', 'Email', 'First Name', 'Response']
        if not all(col in df.columns for col in required_columns):
            return jsonify({'error': 'Missing required columns'}), 400

        grouped = df.groupby('Company').agg({
            'Patent Number': lambda x: list(x),
            'Email': lambda x: list(x),
            'First Name': lambda x: list(x),
            'Response': 'first'
        }).reset_index()

        companies_data = grouped.to_dict('records')

        # Process companies to attach patents from same company name if patents are missing
        for company in companies_data:
            # Clean patent numbers and check if empty
            company['Patent Number'] = [
                str(patent).strip() for patent in company['Patent Number']
                if isinstance(patent, (str, int, float)) and str(patent).strip() and str(patent).lower() != 'nan'
            ]
            if not company['Patent Number']:  # If no patents
                # Look for another company with the same name that has patents
                for other_company in companies_data:
                    if (
                        other_company['Company'] == company['Company'] and
                        other_company['Patent Number'] and
                        other_company != company
                    ):
                        # Attach the first 2 patent numbers from the other company
                        company['Patent Number'] = other_company['Patent Number'][:2]
                        break

        print(f"Processed {len(companies_data)} companies")
        return jsonify({
            'message': 'File processed successfully',
            'total_companies': len(companies_data)
        })
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/send-emails', methods=['POST'])
def send_emails():
    try:
        data = request.json
        email = data.get('email')
        password = data.get('password')
        start_index = data.get('startIndex')
        end_index = data.get('endIndex')
        email_format = data.get('emailFormat', 'General Email')

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

            valid_emails = [str(email).strip() for email in emails if isinstance(email, (str, int, float)) and '@' in str(email)]
            if not valid_emails:
                continue

            valid_first_names = [str(name).strip() for name in first_names if isinstance(name, (str, int, float)) and str(name).strip()]
            valid_first_names = valid_first_names[:len(valid_emails)]
            names_list = ', '.join(valid_first_names[:-1]) + ' & ' + valid_first_names[-1] if len(valid_first_names) > 1 else valid_first_names[0] if valid_first_names else ''

            # Select only the first 2 patent numbers from the list
            selected_patents = patents[:2]  # Already ensured in upload_excel, just confirming
            patents_str = ', '.join(selected_patents) if selected_patents else 'No patent information available'

            if isinstance(response, str) and response.lower() == 'yes':
                continue

            follow_up_date = datetime(2024, 11, 27) + timedelta(days=15)
            current_date = datetime.now()

            if email_format == 'Meeting Email':
                subject = f"Request for Meeting to Discuss Patent Monetization Opportunities"
                html = f"""
                <html>
                <body>
                <p style="font-size: 10.5pt;">
                    Hi {names_list},<br><br>
                    Hope this email finds you well.<br><br>
                    We have identified patents {patents_str} etc. as potential candidates for monetization and believe there is a significant opportunity for your organization.<br><br>
                    We would love to schedule a meeting to discuss this further and explore how we can collaborate with your team. Please let us know a convenient time for a call or virtual meeting.<br><br>
                    Feel free to suggest a date and time, or you can book a slot directly via our calendar: <a href="https://calendly.com/bayslope/meeting">Schedule a Meeting</a>.<br><br>
                    Looking forward to your response.<br><br>
                    Best regards,<br>
                    Sarita (Sara) / Bayslope<br>
                    Techreport99 | Bayslope<br>
                    e: <a href="mailto:patents@bayslope.com">patents@bayslope.com</a><br>
                    p: +91-9811967160 (IN), +1 650 353 7723 (US), +44 1392 58 1535 (UK)<br><br>
                    <span style="color: grey; font-size: 8.5pt;">
                        The content of this email message and any attachments are intended solely for the addressee(s) and may contain confidential and/or privileged information...
                    </span>
                </p>
                </body>
                </html>
                """
            elif (pd.isna(response) or response == '') and email_format == 'General Email':
                subject = f"Patent Monetization Interest for {patents_str} etc."
                html = f"""
                <html>
                <body>
                <p style="font-size: 10.5pt;">
                    Hi {names_list},<br><br>
                    Hope all is well at your end.<br><br>
                    Our internal framework has identified patents {patents_str} etc. and we think there is a monetization opportunity for them.<br><br>
                    We work closely with a network of active buyers who regularly acquire high-quality patents for monetization across various technology sectors.<br><br>
                    Could you help facilitate a discussion with your client about this matter?<br><br>
                    Best regards,<br>
                    Sarita (Sara) / Bayslope<br>
                    Techreport99 | Bayslope<br>
                    e: <a href="mailto:patents@bayslope.com">patents@bayslope.com</a><br>
                    p: +91-9811967160 (IN), +1 650 353 7723 (US), +44 1392 58 1535 (UK)<br><br>
                    <span style="color: grey; font-size: 8.5pt;">
                        The content of this email message and any attachments are intended solely for the addressee(s) and may contain confidential and/or privileged information...
                    </span>
                </p>
                </body>
                </html>
                """
            elif isinstance(response, str) and response.lower() == 'no' and current_date >= follow_up_date and email_format == 'General Email':
                subject = f"Follow-up: Patent Acquisition Interest"
                html = f"""
                <html>
                <body>
                <p style="font-size: 10.5pt;">
                    Hi {names_list},<br><br>
                    Hope all is well at your end.<br><br>
                    We understand your busy schedule so didn’t mean to bother you via this email. Just checking if you could assist in facilitating a discussion with your client.<br><br>
                    It will be great to hear from you.<br><br>
                    Best regards,<br>
                    Sarita (Sara) / Bayslope<br>
                    Techreport99 | Bayslope<br>
                    e: <a href="mailto:patents@bayslope.com">patents@bayslope.com</a><br>
                    p: +91-9811967160 (IN), +1 650 353 7723 (US), +44 1392 58 1535 (UK)<br><br>
                    <span style="color: grey; font-size: 8.5pt;">
                        The content of this email message and any attachments are intended solely for the addressee(s) and may contain confidential and/or privileged information...
                    </span>
                </p>
                </body>
                </html>
                """
            else:
                continue

            server = smtplib.SMTP('smtp.office365.com', 587)
            server.starttls()
            server.login(email, password)
            msg = MIMEMultipart("alternative")
            msg['Subject'] = subject
            msg['From'] = formataddr(("Bayslope Business Solutions", email))
            msg['To'] = ', '.join(valid_emails)
            msg.attach(MIMEText(html, 'html'))
            server.sendmail(email, valid_emails, msg.as_string())
            server.quit()
            sent_emails += len(valid_emails)

        return jsonify({'message': f'Successfully sent {sent_emails} emails'})
    except Exception as e:
        print(f"Error sending emails: {str(e)}")
        return jsonify({'error': str(e)}), 500

app = app