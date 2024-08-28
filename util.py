from datetime import datetime

from func_timeout import FunctionTimedOut
import streamlit as st
from PIL import Image
import pandas as pd
import html


def init_state_var(vars):
    for var in vars:
        if var not in st.session_state:
            st.session_state[var] = None


def get_table_content(df):
    table_rows = ""

    for info in df.itertuples():
        table_row = f"""
            <tr>
                <td>{info.Branch}</td>
                <td>{info.Description}</td>
                <td>{info.SerialNumber}</td>
                <td>{info.Qty}</td>
            </tr>
        """
        table_rows += table_row 
    return table_rows


def get_recepients(branch, email_df):
    # Scope the email df to that of the specific branch
    branch_df = email_df[email_df["Branch"].isin([branch])]
    
    # Drop missing emails and return list of recipient emails
    receipients = branch_df.filter(like="email").dropna(axis=1).values.flatten()
    return list(receipients) if receipients.size > 0 else None


# update excel sheet with timestamp of when email was sent
def record_data(branch, email_type, df, recorded_by):
    
    #Get the row index of branch and update the time of which the email was being sent
    row_index = df.index[df["Branch"] == branch][0]
    
    # Get the 'sent by' column index for the respestive email type
    column_index = df.columns.get_loc(email_type) + 1
    
    # Update the rows with the relevant information
    df.at[row_index, email_type] = pd.to_datetime(datetime.now(), format="%d-%m-%Y")
    df.at[row_index, df.columns[column_index]] = str(recorded_by)


# append failed branch and information to combined logs
def handle_exception(branch, exception, logs):
    
    new_log = (branch, exception)
    logs.append(new_log)

    
@st.cache_data()
def load_data(xlsx_file, sheet_names):
    
    master_df, email_df, recording_df = [
            pd.read_excel(xlsx_file, sheet_name = sheet) 
            for sheet in sheet_names
    ]
    
    return master_df, email_df, recording_df


def filter_unaccounted_branches(df):
    # Get the rows in the dataframe where Accounted column is NA
    filtered_df = df[df["Accounted"].isna()]
    
    # Get the names of the unaccounted branches
    unaccounted_branches_names = filtered_df["Branch"].unique()
    
    return unaccounted_branches_names, filtered_df


# Creates a preview of the email template selected by the user 
def load_preview(email_type):  
    
    table = pd.DataFrame({
        "Branch": ["Bar X", "Bar X"],
        "Description": ["Alcohol 1 / 25%", "Alcohol 2 / 36%"],
        "Serial Number": ["E123", "E456"],
        "Qty": ["1", "1"]
    }, index=[1,2])
    
    eco_totes_image = Image.open("ecoTotes_QR.png")
    
    
    with st.expander("View template"):
        st.info("Actual table contents will be generated dynamically from the uploaded file.")
        
        if email_type == "1st Email Chaser":
            # Email Contents of 1st email chaser template
            
            st.write(f"""
                <p>Hello Team,</p>
                <p>Thank you for your continued support in your <span style="color: rgb(84, 130, 53);">ecoSPIRITS</span> Program!</p>
                <p>
                    To ensure a steady turnover of ecoTOTES, we have implemented a <strong>3 month Flagging System</strong> for the 
                    return of all empty ecoTOTES. Our ecoTOTES are a valuable commodity and their re-use is central to the 
                    success of our sustainable distribution program. We appreciate your help in our efforts in recouping 
                    them.
                </p>
                <p>
                    The following below is a report of all the ecoTOTES at your venue that have been with you <strong>for 3 months
                    or more.</strong>
                </p>
                <p>Please do account for these totes and get back to us confirming either they are:</p>
                <ol>
                    <li>Still in use</li>
                    <li>Empty, ready for collection</li>
                    <li>Lost</li>
                </ol>      
                
            """, unsafe_allow_html=True)
            
            st.dataframe(table, use_container_width=True)     
                       
            st.write(f"""
                <p>
                    ** 
                    <b>
                        <span style="color: rgb(58, 124, 34);">
                            Do kindly send us a brief reply to let us know as well if the above ecoTOTES have been 
                            accounted for. We aim to send these reminders out every three months.
                        </span>
                    </b>
                </p>
                <p>
                    In addition, please reach out once there are empty ecoTOTES to be collected and we will facilitate the 
                    collection accordingly.
                </p>
                <p style="color: black;">
                    For your reference, here are some images of the ecoTOTES and where you can find the serial numbers on 
                    the assets:
                </p>
            """, unsafe_allow_html=True)
            
            st.image(eco_totes_image)
            
            st.write("""         
                <p>Note:</p>
                <p style="color: black;">
                    <u><i>We'll also follow up on this email in a week's time</i></u> to check on the status should we not hear back 
                    from you by then.  
                </p>
                <p style="color: black;">
                    Thereafter will be <u><i>another grace period of two weeks</i></u> to confirm the status of the mentioned totes. 
                    <u>Failure to do so will result in an automatic billing of SGD200 (Before tax) per ecoTOTE not returned.</u>
                </p>
                <p><span class="regards-msg">Thank you so much for your time on this matter and let us know if you have any questions.</span></p>
                <p><span class="regards-msg">Thanks and best regards.</span></p>
                <p>
                    <span class="regards-msg" style="font-size: 13px;">
                        Proof Team.
                        <br>
                        <br>
                        <u></u>
                        <u></u>
                    </span>   
                </p>
            """, unsafe_allow_html=True)          
            
                        
        elif email_type == "2nd Email Chaser":
            
            # Email Contents of 2nd email chaser template
            
            st.write("""
                <p>Hello Team,</p>
                <p>We hope your day has been going great so far!</p>
                <p>
                    We would like to follow up on the outstanding ecoTOTE(S) at your venue as per our previous email:
                </p>     
            """, unsafe_allow_html=True)
            
            st.dataframe(table, use_container_width=True)
            
            st.write("""
                <p>
                    Do take note that <u>should we not receive any confirmation about the status of the metioned totes</u> within the next two weeks, we will deem these totes "Missing" and move forward with charging the value of these totes to your account. As stated in the previous email, each ecoTote is valued at SGD200 (Before Tax).
                </p>
                <p>
                    Once again, we greatly appreciate the time taken to ensure that the ecoTOTES are properly maintained so that these valuable assets can continue serving their purpose in our sustanability movement.
                </p>
                <p>
                    Kind regards,
                </p>
                <p>
                    Proof Team
                </p>
            """, unsafe_allow_html=True)

        else:
            st.write("""
                    <p>Dear Team,</p>
                    <p>
                        We are writing to remind you that several ecoTOTES have been at your venue for over 3 months. To maintain a smooth operation of our sustainable distribution program, it's crucial to return empty ecoTOTES promptly.
                    </p>     
                    <strong>Please review the following list of ecoTOTES and confirm their status:</strong>
            """, unsafe_allow_html=True)
            
            st.dataframe(table, use_container_width=True)
             
            st.write("""       
                    <strong>Please indicate whether these ecoTOTES are:</strong>
                    <ul>
                        <li><strong>Still in use</strong></li>
                        <li><strong>Empty and ready for collection</strong></li>
                        <li><strong>Lost</strong></li>
                    </ul>
                    <p>
                        <strong>If you have any empty ecoTOTES ready for pickup, please contact us immediately.</strong> We will arrange for collection as soon as possible.
                    </p>
                    <p>
                    Please note that the ecoTOTES on circulation are solely owned by Proof & Company and the company reserves the rights to charge our customers for missing ecoTOTES at SGD 200 each excluding the prevailing GST.
                    </p>
                    <p>
                        Thank you for your immediate attention to this matter.
                    </p>
                    <p>
                        Kind regards,
                    </p>
                    <p>
                        Proof Team
                    </p>
            """, unsafe_allow_html=True)