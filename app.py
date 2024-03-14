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
            with st.container(height=475):
                st.write(message_body_text, unsafe_allow_html=True)
        elif email_type == "2nd Email Chaser":
            with st.container(height=475):
                st.write(message_2_text, unsafe_allow_html=True)
        
        with st.form("Credentials", clear_on_submit=True):
            st.write("Recorded by:")
            col1, col2 = st.columns(2)
            first_name = col1.text_input("First name", placeholder="John")
            last_name = col2.text_input("Last name", placeholder="Doe")
                    
            submit_btn = st.form_submit_button("Send Mails")

        if submit_btn:
            recorded_by = f"{first_name} {last_name}"
            
            params=[master_df, email_df, recording_df, file, branch_names_mastersheet, recorded_by] 
            
            if email_type == "1st Email Chaser":
                params.insert(5, message_body)
                
            elif email_type == "2nd Email Chaser":
                params.insert(5, message_2)
                
            mass_mail(*params)

            with pd.ExcelWriter("updated.xlsx", engine="xlsxwriter", date_format="DD-MM") as writer:
                master_df.to_excel(writer, sheet_name="Mastersheet", index=False,)
                email_df.to_excel(writer, sheet_name="Branch with Emails", index=False)
                recording_df.to_excel(writer, sheet_name="Recording", startrow=0, index=False)

            old_file_name = file.name
            new_file_name = old_file_name.split(".")[0]

            with open("updated.xlsx", "rb") as updated_file:
                st.download_button(label="Download updated excel sheet", data=updated_file, file_name=f"{new_file_name}_updated.xlsx", mime="application/vnd.ms-excel")