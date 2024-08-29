import os
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.utils import formataddr
from pathlib import Path
from dotenv import load_dotenv
from func_timeout import func_set_timeout, FunctionTimedOut
import streamlit as st

from messages import message_body
from util import get_table_content, get_recepients, record_data, handle_exception

PORT = 587
EMAIL_SERVER = "smtp-mail.outlook.com"

#load the environment variable
current_dir = Path.cwd()
envdir = os.path.join(current_dir, ".env")
image_path = os.path.join(current_dir, "ecoTotes_QR.png")
file_name = "ecoTotes_QR.png"
load_dotenv(envdir)

# get the environment variables
sender_email = st.secrets["email"] #os.getenv("email") 
password = st.secrets["password"]  #os.getenv("password") 


@func_set_timeout(5) # Send_email function will time out after 5 seconds to optimise sending process
def send_email(subject, receipients, message, branch_name):
    """Unlike some other classes where you can access attributes directly (e.g., obj.attribute), 
    EmailMessage relies on the dictionary interface for header access.
    EmailMessage does not have a direct Subject attribute that you can access with dot notation."""
    
    # Message Structure
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"{subject} ({branch_name})"
    msg["From"] = formataddr((f"Proof & Company", f"{sender_email}"))
    msg["To"] = ", ".join(receipients)
    
    # Text Information in Email
    body = MIMEText(message, "html")
    msg.attach(body)
    
    with open(image_path, "rb") as fp:
        # Ecototes Image included inline in HTML email
        image_content = fp.read()
        image = MIMEImage(image_content, name="EcoTotesQR") 
        image.add_header("Content-ID", "<ecoTotesQR>")
        msg.attach(image)
        
        # Ecototes Image included as attachment in HTML email
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


num_retry = 0
# Iterates through the branches with unaccounted containers
# populates html email template and sends mail to receipient
# Recursively calls itself 5 times or until all emails are successfully sent 
def mass_mail(filtered_df, email_df, recording_df, email_type, unaccounted_branches, message, recorded_by, sending_prog):
    global num_retry
    
    branches_email_success = []
    failed_logs=[]
    
    # Iterate through the list of unaccounted branches to get their essential information for populating the HTML email template
    for branch in unaccounted_branches:
        
        # Artificially slow down the sending process to prevent rate blocking
        time.sleep(3)
        
        # scope the df containing all unaccounted branches to that of the specific branch 
        branch_info = filtered_df[(filtered_df["Branch"] == branch)]
        
        # Get specific info for that branch
        table_rows = get_table_content(branch_info)
        receipients = get_recepients(branch, email_df)
        
        if receipients is None:
            continue
        
        print(f"Sending email to {branch}")
        
        try:
            # Send mail to extracted receipients with populated template
            send_email("Proof & Company Pte Ltd - Overdue ecoTOTES", receipients, message(table_rows), branch)
            print(f"sent to {branch} succesfully")  
            branches_email_success.append(branch) 
            sending_prog.progress(value=len(branches_email_success)/len(unaccounted_branches), text="Sending...")
            record_data(branch, email_type, recording_df, recorded_by)
     
        except smtplib.SMTPException as e:
            print(e)
            handle_exception(branch, e, failed_logs)
            continue
        
        except Exception as e:
            print(e)
            handle_exception(branch, e, failed_logs)
            continue
        
        except FunctionTimedOut as e:
            print(f"Timed out for {branch}")
            handle_exception(branch, f"Timed out for {branch}", failed_logs)
            continue
    
    unsuccessful_emails = list(set(unaccounted_branches) - set(branches_email_success))
    print(f"Unsuccessful: {unsuccessful_emails}")
    
    if num_retry >= 4:
        return unsuccessful_emails
    
    if len(unsuccessful_emails) > 0:
        num_retry += 1
        return mass_mail(filtered_df, email_df, recording_df, email_type, unsuccessful_emails, message, recorded_by, sending_prog) 
    else:
        return []
    
  
         
if __name__ == "__main__":
    send_email(
        subject="Testing",
        receipients=["wepaw39397@kinsef.com"],
        message = message_body("table"),
        branch_name = "ZIGGY BAR"
    )
