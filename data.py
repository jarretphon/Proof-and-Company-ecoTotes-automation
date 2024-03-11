from smtplib import SMTPException
import pandas as pd
import time
import os
from pathlib import Path
from datetime import datetime
from func_timeout import FunctionTimedOut
from messages import message_body, message_2
from automation import send_email
from util import get_table_content, get_recepients, record_data

num_retry = 0
def mass_mail(master_df, email_df, recording_df, file, branch_names_mastersheet, message, recorded_by):
    global num_retry
    if num_retry >= 5:
        return
    
    branches_email_success = []
    for branch in branch_names_mastersheet:
        time.sleep(3)
        branch_info = master_df[master_df["Branch"] == branch]
        
        table_rows = get_table_content(branch_info)
        receipients = get_recepients(branch, email_df)
        
        if receipients == []:
            continue
        print(f"Sending email to {branch}")
        
        try:
            send_email("Proof & Company Pte Ltd - Overdue ecoTOTES", receipients, message(table_rows))
            print(f"sent to {branch} succesfully")  
            record_data(branch, file, recording_df, recorded_by)
            branches_email_success.append(branch) 
    
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
    
    if len(unsuccessful_emails) > 0:
        num_retry += 1
        mass_mail(master_df, email_df, recording_df, file, unsuccessful_emails, message, recorded_by)
        


