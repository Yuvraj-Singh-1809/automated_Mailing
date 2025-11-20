import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


# Load the CSV file
csv_file = 'mails.csv' # Replace with your actual CSV file path
df = pd.read_csv(csv_file, encoding='latin1')

# Email configuration
smtp_server = 'smtp.gmail.com'
smtp_port = 587  #SMTP port 465 for ssl 587 for tls

sender_email = '' # Replace with your actual email
sender_password = ''  #email passwords s 'javzxdrzscseruow'  b 'qshdjpbtrrkqmwwd' 


# Email content
subject = 'Analyst'
body = """
Dear {Name},

I hope this email finds you well.

I am writing to express my interest in the  Analyst position at {Company} . With a strong background in data analysis and a passion for deriving actionable insights from complex datasets, I am excited about the opportunity to contribute to your team.

In my previous role at EY, I was responsible for Data Analytics,data cleaning, Power BI report . I have proficiency in  SQL, Python, Excel, VBA, and I am skilled in data visualization, statistical analysis, data cleaning.

Please find my resume attached to this email for your review. I would welcome the opportunity to discuss how my skills and experiences align with the needs of your team. Thank you for considering my application. I look forward to the possibility of discussing this exciting opportunity with you.

Best regards,

Shambhavi Singh
"""

# Path to the PDF file
pdf_file_path = 'Shambhavi.pdf'  # Replace with your actual PDF file path

# Function to send email


def send_email(to_email, name, designation, company):
    try:
        # Create the email
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Subject'] = subject

        # Format the email body
        formatted_body = body.format(Name = name, Company = company, Designation = designation)
        msg.attach(MIMEText(formatted_body, 'plain'))

        # Attach the PDF file
        with open(pdf_file_path, 'rb') as attachment:
            part = MIMEApplication(attachment.read(), Name = 'Shambhavi.pdf')
            part['Content-Disposition'] = f'attachment; filename="Shambhavi.pdf"'
            msg.attach(part)

        # Connect to the SMTP server
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Secure the connection
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, to_email, msg.as_string())

        print(f"Email sent to {to_email} successfully!")

    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

# Iterate over the CSV and send emails


for index, row in df.iterrows():
    send_email(row['Email'], row['Name'], row['Designation'], row['Company'])

