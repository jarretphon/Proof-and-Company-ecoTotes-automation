import os
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.utils import formataddr
from pathlib import Path
from dotenv import load_dotenv
from func_timeout import func_set_timeout
from messages import message_body
import streamlit as st

PORT = 587
EMAIL_SERVER = "smtp-mail.outlook.com"

#load the environment variable
current_dir = Path.cwd()
envdir = os.path.join(current_dir, ".env")
image_path = os.path.join(current_dir, "ecoTotes_QR.png")
file_name = "ecoTotes_QR.png"
load_dotenv(envdir)

# get the environment variables
sender_email = st.secrets["email"]
password = st.secrets["password"]

@func_set_timeout(5)
def send_email(subject, receipients, message, branch_name):
    """Unlike some other classes where you can access attributes directly (e.g., obj.attribute), 
    EmailMessage relies on the dictionary interface for header access.
    EmailMessage does not have a direct Subject attribute that you can access with dot notation."""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"{subject} ({branch_name})"
    msg["From"] = formataddr((f"Proof & Company", f"{sender_email}"))
    msg["To"] = ", ".join(receipients)
    
    body = MIMEText(message, "html")
    msg.attach(body)
    
    with open(image_path, "rb") as fp:
        image_content = fp.read()
        image = MIMEImage(image_content, name="EcoTotesQR") 
        image.add_header("Content-ID", "<ecoTotesQR>")
        msg.attach(image)
        
        attachment = MIMEBase("image", "png", filename=file_name)
        attachment.add_header("Content-Disposition", "attachment", filename=file_name)
        attachment.set_payload(image_content)
        encoders.encode_base64(attachment)
        msg.attach(attachment)
        
    # smtplib.SMTP is used to connect to the specified email server (EMAIL_SERVER) on the specified port (PORT).
    with smtplib.SMTP(EMAIL_SERVER, PORT) as server:
        # STARTTLS is a way to upgrade a plain text connection to a secure connection using TLS for added layer of security 
        server.starttls()
        server.login(f"{sender_email}", f"{password}")
        server.sendmail(f"{sender_email}", receipients, msg.as_string())

         
if __name__ == "__main__":
    send_email(
        subject="Testing",
        receipients=["jadev41272@ikumaru.com"],
        message = message_body("table"),
        branch_name = "ZIGGY BAR"
    )
