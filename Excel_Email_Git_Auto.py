import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import git
import os

# Configuration
EXCEL_FILE = 'projects.xlsx'  # Path to your Excel file
GITHUB_REPO_PATH = '/path/to/your/repo'  # Path to your local GitHub repository
EMAIL_SENDER = 'your_email@example.com'
EMAIL_RECEIVER = 'recipient_email@example.com'
EMAIL_SUBJECT = 'New Project Added'
SMTP_SERVER = 'smtp.example.com'
SMTP_PORT = 587
SMTP_USERNAME = 'your_email@example.com'
SMTP_PASSWORD = 'your_email_password'

# Function to send email
def send_email(project_name):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_SENDER
    msg['To'] = EMAIL_RECEIVER
    msg['Subject'] = EMAIL_SUBJECT

    body = f"A new project '{project_name}' has been added."
    msg.attach(MIMEText(body, 'plain'))

    # Set up the server
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()  # Secure the connection
    server.login(SMTP_USERNAME, SMTP_PASSWORD)
    server.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, msg.as_string())
    server.quit()

# Function to update GitHub repository
def update_github_repo():
    repo = git.Repo(GITHUB_REPO_PATH)
    repo.git.add(EXCEL_FILE)
    repo.index.commit('Updated Excel file with new project')
    repo.git.push()  # Push to GitHub

# Load the existing Excel file
if os.path.exists(EXCEL_FILE):
    df = pd.read_excel(EXCEL_FILE)

    # Check for the latest entry (assuming the first column is the project name)
    last_project_name = df.iloc[-1, 0]  # Adjust based on your Excel structure

    # Send email notification
    send_email(last_project_name)

    # Update the GitHub repository
    update_github_repo()
else:
    print("Excel file not found.")

