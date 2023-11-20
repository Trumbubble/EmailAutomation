from email.message import EmailMessage
import ssl
import smtplib
import pandas as pd
import openai
from decouple import config

from urllib.request import urlopen
import numpy as np

email_sender = 'ttorubarov@gmail.com'
email_password = config('EMAIL_KEY')
openai.api_key = config('OPENAI_KEY')

topics = np.array([])
df = pd.read_excel('C:\\Users\\ttoru\\Programs\\EmailAutomation\\Research Emails.xlsx')

def findInterests(url):
    global topics
    page = urlopen(url)
    html_bytes = page.read()
    html = html_bytes.decode("utf-8")
    subtext = '<span class="concept">'
    firstText = "concept-badge-large-container dropdown-overflow"

    if subtext in html:
        positions = [index for index in range(len(html)) if html.startswith(subtext, index)]

    if firstText in html:
        positions1 = [index for index in range(len(html)) if html.startswith(firstText, index)]

    realPositions = np.array([])
    for position in positions:
        for position1 in positions1:
            if abs(position-position1)<400:
                realPositions = np.append(realPositions, position)

    for realPosition in realPositions:
        end_index = html.find('<', int(realPosition) + 1)
        topics = np.append(topics, html[int(realPosition) + len('<span class="concept">'):end_index])

def sendChatGPT(interests):
    response = openai.Completion.create(
        model="gpt-3.5-turbo", 
        prompt= "I am writing an email to a professor asking for research opportunities. I want to write a short paragraph explaining how I'm a little interested in their interests or at least a broader version of their research topics. These are their interests: " + interests,
        temperature=0.7,
        max_tokens=300
    )
    return response.choices[0].text.strip()

def sendEmail(rowNum):
    global topics
    # email_receiver = df.loc[rowNum, "Email"]
    email_receiver = "ttorubarov@gmail.com"
    subject = 'Automation Test'
    name = df.loc[rowNum, "Name"]
    chatGPTText = sendChatGPT(findInterests(df.loc[rowNum,"URL"]))
    print(chatGPTText)
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
        # smtp.sendmail(email_sender, email_receiver, em.as_string())

for index, row in df.iterrows():
    email_status = str(row["Email Status (NS, SNR, SR)"])
    if "NS" in email_status:
        sendEmail(index)



