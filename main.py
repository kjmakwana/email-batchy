import os
import base64
import json
from email.message import EmailMessage
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import pandas as pd
import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


SCOPES = ['https://www.googleapis.com/auth/gmail.send']


def read_file():
    df=pd.read_excel('cold_email_contacts.xlsx')
    return df

def authenticate():
    creds = None
    if os.path.exists("token.json"):
        flow = InstalledAppFlow.from_client_secrets_file('token.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

def generate_text(name, company, role):
    body = f"""
Hello {name},<br><br>
I am Kshitij Makwana. I know your time is valuable, so I will leave these points here:<br>
<ol>
<li>I have a <b>Master's in Data Science</b> from the <b>University of Michigan, Ann Arbor</b>. (GPA 4.0)
<li>I have taken courses/have full-time experience in LLMs, GenAI, Financial Modelling, Data Science and Analytics, DevOps
<li>I have <b>2 years of experience</b> using Python, PyTorch, and building production-grade software.
</ol>
I love the work that {company} is doing and would like to contribute to your projects as a {role}. I have graduated, and I'm looking for <b>full-time roles or internships</b>. Please let me know if we can talk about this more.<br><br>

Thanks,<br>
Kshitij Makwana (he/him)<br>
(734) 489-2922<br>
MS in Data Science
    """
    subject = f"Exploring Opportunities at {company}"
    return body,subject

def send_email(creds, email, body, subject, file):
    try:
    # create gmail api client
        service = build("gmail", "v1", credentials=creds)

        message = MIMEMultipart()

        # message.set_content(body)

        message["To"] = email
        message["From"] = "mkshitij@umich.edu"
        message["Subject"] = subject
        msg = MIMEText(body,"html")
        message.attach(msg)
        filename = os.path.basename(file)
        with open(file, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            message.attach(part)
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        create_message = {"raw": encoded_message}
        send_message = (
            service.users()
            .messages()
            .send(userId="me", body=create_message)
            .execute()
        )
        print(f'Message Id: {send_message["id"]}')
    except HttpError as error:
        print(f"An error occurred: {error}")
        send_message = None
    return send_message

#   return draft
if __name__ == "__main__":
    df = read_file()
    creds = authenticate()
    if not creds:
        print("Authentication failed.")
        exit(1)

    for index, row in df.iterrows():
        name = row['First']
        company = row['Company']
        role = row['Role']
        email = row['Email']
        resume = row["Resume"]
        if resume=="MLE":
            file = "Kshitij Makwana - MLE.pdf"
        elif resume=="DS":
            file = "Kshitij Makwana - DS.pdf"
        else:
            file = "Kshitij Makwana - Financial Analyst.pdf"

        body, subject = generate_text(name, company, role)
        print(f"Sending email to {name} at {company} ({email})")
        send_email(creds, email, body, subject, file)