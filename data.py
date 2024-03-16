from cgitb import text
from smtplib import SMTPException
import pandas as pd
import time
import streamlit as st
from pathlib import Path
from datetime import datetime
from func_timeout import FunctionTimedOut
from messages import message_body, message_2
from automation import send_email
from util import get_table_content, get_recepients, record_data

num_retry = 0
def mass_mail(master_df, email_df, recording_df, file, branch_names_mastersheet, message, recorded_by, sending_prog):
    global num_retry
    
    branches_email_success = []
    unsuccessful_emails = []
    for branch in branch_names_mastersheet:
        time.sleep(3)
        
        branch_info = master_df[(master_df["Branch"] == branch) & (master_df["Accounted"].isna())]
        
        if branch_info.empty:
            continue
        
        table_rows = get_table_content(branch_info)
        receipients = get_recepients(branch, email_df)
        
        if receipients == []:
            continue
        print(f"Sending email to {branch}")
        
        try:
            send_email("Proof & Company Pte Ltd - Overdue ecoTOTES", receipients, message(table_rows))
            print(f"sent to {branch} succesfully")  
            branches_email_success.append(branch) 
            sending_prog.progress(value=len(branches_email_success)/len(branch_names_mastersheet), text="Sending...")
            record_data(branch, file, recording_df, recorded_by)
     
        except SMTPException as e:
            print(e)
            continue
        
        except Exception as e:
            print(e)
            continue
        
        except FunctionTimedOut:
            print(f"Timed out for {branch}")
            continue
    
    unsuccessful_emails = [branch for branch in branch_names_mastersheet if branch not in branches_email_success]
    print(f"Unsuccessful: {unsuccessful_emails}")
    
    if num_retry >= 4:
        return unsuccessful_emails
    
    if len(unsuccessful_emails) > 0:
        num_retry += 1
        return mass_mail(master_df, email_df, recording_df, file, unsuccessful_emails, message, recorded_by, sending_prog) 
    

    