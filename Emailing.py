from email.message import EmailMessage
import ssl
import smtplib
import pandas as pd
from decouple import config

email_sender = 'ttorubarov@gmail.com'
email_password = config('EMAIL_KEY')

df = pd.read_excel('C:\\Users\\ttoru\\Programs\\EmailAutomation\\Research Emails.xlsx')

def sendEmail(rowNum):
    email_receiver = df.loc[rowNum, "Email"]

    subject = 'Automation Test'
    name = df.loc[rowNum, "Name"]
    body = f"""
    Dear Dr. {name},
        If you receive this, the code works, and automation of the emails is done.
    We're getting published.
    """
    print(df.loc[rowNum, "Name"])
    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_receiver
    em['Subject'] = subject
    em.set_content(body)

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, email_receiver, em.as_string())

for index, row in df.iterrows():
    email_status = str(row["Email Status (NS, SNR, SR)"])
    if "NS" in email_status:
        sendEmail(index)
