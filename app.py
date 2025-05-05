import os
import urllib.parse
import uuid
from flask import Flask, request, jsonify, make_response
from flask_cors import CORS
import pandas as pd
from datetime import datetime, timedelta

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

# Handle preflight OPTIONS requests manually for better control
@app.before_request
def handle_options():
    if request.method == "OPTIONS":
        print("Handling OPTIONS request for", request.path)
        response = make_response()
        response.headers["Access-Control-Allow-Origin"] = request.headers.get("Origin", "https://email-frontend-eosin.vercel.app")
        response.headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type,Authorization"
        response.headers["Access-Control-Max-Age"] = "86400"
        return response, 200

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

        if not file.filename.endswith('.xlsx'):
            return jsonify({'error': 'File must be an .xlsx file'}), 400

        print(f"Received file: {file.filename}")
        # Check file size (Vercel has a 4.5MB limit for Hobby plan)
        file.seek(0, os.SEEK_END)
        file_size = file.tell()
        file.seek(0)  # Reset file pointer to the beginning
        if file_size > 4.5 * 1024 * 1024:  # 4.5MB in bytes
            return jsonify({'error': 'File size exceeds 4.5MB limit'}), 400

        # Save the file temporarily to read it (Vercel serverless environment workaround)
        unique_filename = f"{uuid.uuid4()}.xlsx"
        temp_file_path = f"/tmp/{unique_filename}"
        file.save(temp_file_path)

        df = pd.read_excel(temp_file_path, engine='openpyxl')
        print(f"Excel file read successfully. Columns: {df.columns.tolist()}")

        # Clean up the temporary file
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
        print(f"Processed {len(companies_data)} companies: {companies_data}")
        return jsonify({
            'message': 'File processed successfully',
            'total_companies': len(companies_data)
        })
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate-emails', methods=['POST'])
def generate_emails():
    try:
        data = request.json
        start_index = data.get('startIndex')
        end_index = data.get('endIndex')

        if not companies_data:
            return jsonify({'error': 'No company data available. Upload an Excel file first.'}), 400

        # Validate start_index and end_index
        if start_index is None or end_index is None:
            return jsonify({'error': 'startIndex and endIndex are required'}), 400

        if not isinstance(start_index, int) or not isinstance(end_index, int):
            return jsonify({'error': 'startIndex and endIndex must be integers'}), 400

        if start_index < 0 or end_index >= len(companies_data) or start_index > end_index:
            return jsonify({'error': 'Invalid index range'}), 400

        email_tasks = []
        total_emails = 0

        for idx in range(start_index, end_index + 1):
            company = companies_data[idx]
            company_name = company['Company']
            emails = company['Email']
            first_names = company['First Name']
            patents = company['Patent Number']
            response = company.get('Response', '')

            # Ensure emails and first_names are lists of strings
            valid_emails = [str(email).strip() for email in emails if isinstance(email, (str, int, float)) and str(email).strip() and '@' in str(email)]
            if not valid_emails:
                email_tasks.append({'company': company_name, 'status': 'skipped', 'reason': 'No valid emails'})
                continue

            valid_first_names = [str(name).strip() for name in first_names if isinstance(name, (str, int, float)) and str(name).strip()]
            if not valid_first_names:
                email_tasks.append({'company': company_name, 'status': 'skipped', 'reason': 'No valid names'})
                continue

            # Match the number of names to emails
            valid_first_names = valid_first_names[:len(valid_emails)]
            if len(valid_first_names) > 1:
                names_list = ', '.join(valid_first_names[:-1]) + ' & ' + valid_first_names[-1]
            else:
                names_list = valid_first_names[0]

            # Handle patents safely
            patents = [str(patent) for patent in patents if isinstance(patent, (str, int, float)) and str(patent).strip()]
            patents = patents[:2]
            patents_str = ', '.join(patents) if patents else 'No patent information available'

            if isinstance(response, str) and response.lower() == 'yes':
                email_tasks.append({'company': company_name, 'status': 'skipped', 'reason': 'Response is yes'})
                continue

            follow_up_date = datetime(2024, 11, 27) + timedelta(days=15)
            current_date = datetime.now()

            if pd.isna(response) or response == '':
                subject = f"Patent Monetization Interest for {patents_str} etc."
                body = f"""Hi {names_list},

Hope all is well at your end.

Our internal framework has identified patents {patents_str} etc. and we think there is a monetization opportunity for them.

We work closely with a network of active buyers who regularly acquire high-quality patents for monetization across various technology sectors.

Could you help facilitate a discussion with your client about this matter?

Best regards,
Sarita (Sara) / Bayslope
Techreport99 | Bayslope
e: patents@bayslope.com
p: +91-9811967160 (IN), +1 650 353 7723 (US), +44 1392 58 1535 (UK)

The content of this email message and any attachments are intended solely for the addressee(s) and may contain confidential and/or privileged information and may be legally protected from disclosure. If you are not the intended recipient of this email, or if this email message has been addressed to you in error, please immediately alert the sender by reply email and then delete this message and any attachments. If you are not the intended recipient, you may not copy, store or deliver this message to anyone, without a written consent of the sender. Thank you!
"""
            elif isinstance(response, str) and response.lower() == 'no' and current_date >= follow_up_date:
                subject = f"Follow-up: Patent Acquisition Interest"
                body = f"""Hi {names_list},

Hope all is well at your end.

We understand your busy schedule so didn’t mean to bother you via this email. Just checking if you could assist in facilitating a discussion with your client.

It will be great to hear from you.

Best regards,
Sarita (Sara) / Bayslope
Techreport99 | Bayslope
e: patents@bayslope.com
p: +91-9811967160 (IN), +1 650 353 7723 (US), +44 1392 58 1535 (UK)

The content of this email message and any attachments are intended solely for the addressee(s) and may contain confidential and/or privileged information and may be legally protected from disclosure. If you are not the intended recipient of this email, or if this email message has been addressed to you in error, please immediately alert the sender by reply email and then delete this message and any attachments. If you are not the intended recipient, you may not copy, store or deliver this message to anyone, without a written consent of the sender. Thank you!
"""
            else:
                email_tasks.append({'company': company_name, 'status': 'skipped', 'reason': 'Response or date condition not met'})
                continue

            # Generate mailto link
            recipients = ','.join(valid_emails)
            subject_encoded = urllib.parse.quote(subject)
            body_encoded = urllib.parse.quote(body)
            mailto_link = f"mailto:{recipients}?subject={subject_encoded}&body={body_encoded}"

            email_tasks.append({
                'company': company_name,
                'status': 'pending',
                'mailto_link': mailto_link,
                'recipients': valid_emails
            })
            total_emails += len(valid_emails)

        return jsonify({
            'message': f'Generated {len(email_tasks)} email tasks',
            'email_tasks': email_tasks,
            'total_emails': total_emails
        })
    except Exception as e:
        print(f"Error generating emails: {str(e)}")
        return jsonify({'error': str(e)}), 500

# Export the app for Vercel
app = app