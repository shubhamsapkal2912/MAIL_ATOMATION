import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Load the data from the Excel (.xlsx) file
file_path = 'HR_LIST.xlsx'  # Update this with the correct path
data = pd.read_excel(file_path, header=0)  # Adjust if needed

# Manually rename columns if needed
data.columns = ['name', 'email', 'company']  # Ensure these match your actual headers

# Convert email column to strings and drop rows with missing emails
data['email'] = data['email'].astype(str).str.strip()
data = data.dropna(subset=['email'])

# Email credentials
sender_email = 'your_mail'
password = 'google_app_password'  # Replace with the generated app password

# Path to your resume PDF file
resume_path = 'your_resume.pdf'  # Update this with the correct path

# Set up the server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(sender_email, password)

# Loop through each row in the dataframe
for index, row in data.iterrows():
    name = row['name']
    email = row['email']
    company = row['company']

    # Create the email content
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = email
    msg['Subject'] = f"Intern Position at {company}"

    body = f"Dear {name},\n\nI hope this email finds you well. I am writing to express my interest for the intern position at {company} in technical domain. I believe my skills and experience make me a strong candidate for this role.\n\nPlease find my resume attached for your consideration.\n\nLooking forward to hearing from you.\n\nBest regards,\nYOUR_NAME"
    
    msg.attach(MIMEText(body, 'plain'))

    # Attach the resume
    with open(resume_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={resume_path.split("/")[-1]}')
        msg.attach(part)

    # Send the email
    try:
        server.send_message(msg)
        print(f"Email sent to {name} at {email}")
    except Exception as e:
        print(f"Failed to send email to {name} at {email}: {e}")

# Terminate the server session
server.quit()
