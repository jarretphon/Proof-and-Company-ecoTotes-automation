from tempfile import NamedTemporaryFile

import streamlit as st
import pandas as pd
from streamlit_option_menu import option_menu
from messages import app_instructions, message_3, message_body, message_2
from automation import mass_mail
from util import init_state_var, get_recepients, load_data, filter_unaccounted_branches, load_preview

st.set_page_config(page_title="EcoTotes Chaser")   

SHEET_NAMES = ["Mastersheet", "Branch with Emails", "Recording"]
STATE_VARS = ["unsucessful_email","tracker_filepath"]

init_state_var(STATE_VARS)

st.title("Proof and Company's automated ecoTotes Chaser")

selected = option_menu(menu_title=None, options=["Home", "Send Emails"],orientation="horizontal")

if selected == "Home":
    st.write(app_instructions())

elif selected == "Send Emails":

    # Define Page Structure
    
    file_upload_placeholder = st.empty()
    email_config_placeholder = st.empty()
    completion_text_placeholder = st.empty()
    download_btn_placeholder = st.empty()
    unsuccesful_email_summary = st.empty()  
    
    # Takes the users uploaded excel file
    file = file_upload_placeholder.file_uploader("Upload your Excel file here", type="xlsx")
    
    if file is not None:
        # Read the excel sheets into individual dataframes
        master_df, email_df, recording_df = load_data(file, SHEET_NAMES)
        
        # Filter the unaccounted branches for sending 
        unaccounted_branches, unaccounted_branches_df = filter_unaccounted_branches(master_df)
        
        with email_config_placeholder.container(border=True):
        
            #User specified email template
            email_type = st.selectbox("Email Type", options=["1st Email Chaser", "2nd Email Chaser", "Reminder"], placeholder="Choose your email type")
            load_preview(email_type)
            
            # Create a form to take user user details for updating tracking sheet
            with st.form("Credentials", clear_on_submit=True, border=False):
                st.write("Recorded by:")
                col1, col2 = st.columns(2)
                first_name = col1.text_input("First name", placeholder="John")
                last_name = col2.text_input("Last name", placeholder="Doe")  
                submit = st.form_submit_button("Send Mails")

            if submit:
                recorded_by = f"{first_name} {last_name}"
                sending_prog = st.progress(value=0, text="Sending...")
                
                mass_mail_kwargs = {
                    "filtered_df": unaccounted_branches_df,
                    "email_df": email_df,
                    "recording_df": recording_df, 
                    "email_type": email_type,
                    "unaccounted_branches": unaccounted_branches,
                    "message": message_body if email_type == "1st Email Chaser" else message_2 if email_type == "2nd Email Chaser" else message_3,
                    "recorded_by": recorded_by,
                    "sending_prog": sending_prog,
                }
                
                failed_logs = mass_mail(**mass_mail_kwargs)
                print(failed_logs)
                st.session_state["unsucessful_email"]  = failed_logs
                sending_prog.empty()

                # Create temporary file to hold updated xlsx file
                tmpfile = NamedTemporaryFile(delete=False)
                
                with pd.ExcelWriter(tmpfile.name, engine="xlsxwriter", date_format="DD-MM") as writer:
                    master_df.to_excel(writer, sheet_name="Mastersheet", index=False,)
                    email_df.to_excel(writer, sheet_name="Branch with Emails", index=False)
                    recording_df.to_excel(writer, sheet_name="Recording", startrow=0, index=False)
                    st.session_state["tracker_filepath"] = tmpfile.name
                
                
    if st.session_state["tracker_filepath"] is not None:
        with open(st.session_state["tracker_filepath"], "rb") as updated_file:
            file_name = "updated_file"#file.name.split(".")[0]
            st.download_button(label="Download updated excel sheet", data=updated_file, file_name=f"{file_name}.xlsx", mime="application/vnd.ms-excel")
        
        
    if st.session_state["unsucessful_email"] is not None:
        unsuccessful_emails = st.session_state["unsucessful_email"]
        print(unsuccessful_emails)
        if unsuccessful_emails:
            completion_text_placeholder.info("Some emails were not sent successfully. This could be due to invalid email addresses or server connection issues. Download the updated excel sheet for more info!")
                        
            # Display summary of unsuccessful emails
            with unsuccesful_email_summary.container(height=200):
                for branch, exception in unsuccessful_emails:   
                    st.error(f"{branch} ({exception})")
                        
        else:
            completion_text_placeholder.success("All emails sent successfully. Download the updated excel sheet for more info!")    
               