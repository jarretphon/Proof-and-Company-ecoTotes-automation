import streamlit as st
import pandas as pd
from streamlit_option_menu import option_menu
from messages import app_instructions, message_body, message_2, message_body_text, message_2_text
from data import mass_mail

st.set_page_config(page_title="EcoTotes Chaser")   

st.title("Proof and Company's automated ecoTotes Chaser")

selected = option_menu(menu_title=None, options=["Home", "Send Emails"],orientation="horizontal")

if selected == "Home":
    st.write(app_instructions())

elif selected == "Send Emails":

    file = st.file_uploader("Upload your Excel file here", type="xlsx")
    
    if file is not None:
        sheet_names = ["Mastersheet", "Branch with Emails", "Recording"]
        master_df, email_df, recording_df = [pd.read_excel(file, sheet_name = sheet) for sheet in sheet_names]
        branch_names_mastersheet = master_df["Branch"].unique()

        email_type = st.selectbox("Email Type", options=["1st Email Chaser", "2nd Email Chaser"], placeholder="Choose your email type")
        
        if email_type == "1st Email Chaser":
            with st.expander(label="See content"):
                st.write(message_body_text, unsafe_allow_html=True)
        elif email_type == "2nd Email Chaser":
            with st.expander(label="See content"):
                st.write(message_2_text, unsafe_allow_html=True)
        
        with st.form("Credentials", clear_on_submit=True):
            st.write("Recorded by:")
            col1, col2 = st.columns(2)
            first_name = col1.text_input("First name", placeholder="John")
            last_name = col2.text_input("Last name", placeholder="Doe")  
            submit_btn = st.form_submit_button("Send Mails")

        if submit_btn:
            recorded_by = f"{first_name} {last_name}"
            sending_prog = st.progress(value=0, text="Sending...")
            params=[master_df, email_df, recording_df, file, branch_names_mastersheet, recorded_by, sending_prog] 
            
            if email_type == "1st Email Chaser":
                params.insert(5, message_body)
            elif email_type == "2nd Email Chaser":
                params.insert(5, message_2)
            
            unsuccessful_emails = mass_mail(*params)
            sending_prog.empty()
            
            with pd.ExcelWriter("updated.xlsx", engine="xlsxwriter", date_format="DD-MM") as writer:
                master_df.to_excel(writer, sheet_name="Mastersheet", index=False,)
                email_df.to_excel(writer, sheet_name="Branch with Emails", index=False)
                recording_df.to_excel(writer, sheet_name="Recording", startrow=0, index=False)

            old_file_name = file.name
            new_file_name = old_file_name.split(".")[0]

            with open("updated.xlsx", "rb") as updated_file:
                completion_text = st.empty()
                st.download_button(label="Download updated excel sheet", data=updated_file, file_name=f"{new_file_name}_updated.xlsx", mime="application/vnd.ms-excel")
                unsuccesful = st.empty()
                
            if unsuccessful_emails:
                completion_text.info("Some emails were not sent successfully. This could be due to invalid email addresses or server connection issues. Download the updated excel sheet for more info!")
                with unsuccesful.container(height=200):
                    for unsuccessful_email in unsuccessful_emails:   
                        st.error(unsuccessful_email)
            else:
                completion_text.success("All emails sent successfully. Download the updated excel sheet for more info!")