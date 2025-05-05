Bayslope Email Sender Backend
This is the backend for the Bayslope Email Sender application, built using Flask. It provides API endpoints to process Excel files containing company data and send emails for patent monetization opportunities. The backend is deployed on Vercel and integrates with the frontend to handle email-sending workflows.
Project Overview
The backend processes Excel files uploaded by the user, extracts company data (e.g., company names, patent numbers, emails), and sends emails to the specified recipients using an SMTP server (Office365). It includes features like:

Excel file parsing and data aggregation
Email sending with a delay to avoid rate limits
Response-based email content customization
Error handling and retry logic for email sending

Prerequisites

Python 3.8+: Ensure Python is installed on your system.
pip: Python package manager for installing dependencies.
SMTP Account: An email account (e.g., Office365) for sending emails.
Vercel Account: For deployment (optional for local development).

Setup Instructions
1. Clone the Repository
Clone the project repository and navigate to the backend folder:
git clone https://github.com/your-username/bayslope-email-sender.git
cd bayslope-email-sender/backend

2. Create a Virtual Environment
Set up a virtual environment to manage dependencies:
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

3. Install Dependencies
Install the required Python packages listed in requirements.txt:
pip install -r requirements.txt

The requirements.txt includes:
flask==2.0.1
flask-cors==3.0.10
pandas==1.5.3
openpyxl==3.1.2

4. Set Up Environment Variables
Create a .env file in the backend folder to store sensitive data:
touch .env

Add the following environment variables to the .env file:
EMAIL=your-email@example.com
PASSWORD=your-email-password

Replace your-email@example.com and your-email-password with your actual email and password for the SMTP server.
5. Run the Backend Locally
Start the Flask development server:
python app.py

The server will run on http://localhost:5000. You can test the API endpoints using tools like Postman or the frontend application.
API Endpoints
1. /upload-excel (POST)

Description: Uploads an Excel file, processes it, and stores company data in memory.
Request:
Method: POST
Content-Type: multipart/form-data
Body: { "file": <Excel file> }


Response:
Success: 200 OK{
  "message": "File processed successfully",
  "total_companies": 5
}


Error: 400 Bad Request or 500 Internal Server Error{
  "error": "Error message"
}





2. /send-emails (POST)

Description: Sends emails to companies based on the processed Excel data.
Request:
Method: POST
Content-Type: application/json
Body:{
  "email": "your-email@example.com",
  "password": "your-email-password",
  "startIndex": 0,
  "endIndex": 4
}


Note: If email and password are not provided in the request, the backend uses the EMAIL and PASSWORD environment variables.


Response:
Success: 200 OK{
  "message": "Successfully sent 7 emails"
}


Error: 400 Bad Request or 500 Internal Server Error{
  "error": "Error message"
}





Deployment on Vercel
1. Project Structure
The backend is structured for Vercel deployment:
backend/
├── api/
│   └── app.py
├── requirements.txt
├── vercel.json
└── .env

2. Vercel Configuration
The vercel.json file is set up to deploy the Flask app as a serverless function:
{
  "builds": [
    {
      "src": "api/app.py",
      "use": "@vercel/python"
    }
  ],
  "routes": [
    {
      "src": "/(.*)",
      "dest": "api/app.py"
    }
  ]
}

3. Deploy to Vercel

Push the backend code to a GitHub repository.
Go to the Vercel dashboard and create a new project.
Import the GitHub repository and set the root directory to backend.
Add the following environment variables in Vercel:
EMAIL: Your email address (e.g., pankaj@bayslope.com)
PASSWORD: Your email password


Deploy the project. The deployed backend URL will be something like https://email-backend-beta-two.vercel.app/.

Usage with Frontend
The backend integrates with the frontend application deployed at https://email-frontend-eosin.vercel.app/. The frontend sends requests to the backend for uploading Excel files and sending emails. Ensure the frontend's REACT_APP_API_URL environment variable is set to the backend's Vercel URL.
Troubleshooting

CORS Issues: The backend uses flask-cors to allow cross-origin requests. Ensure the frontend domain is allowed if you encounter CORS errors.
SMTP Errors: Check the EMAIL and PASSWORD environment variables if email sending fails. Ensure the SMTP server (smtp.office365.com:587) is accessible.
Deployment Fails: Verify requirements.txt and vercel.json. Check Vercel logs for detailed error messages.

License
This project is licensed under the MIT License.
