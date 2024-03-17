import pandas as pd
from datetime import datetime
import io

def get_table_content(branch_info):
    table_rows = ""
    
    for row_num in range(branch_info.shape[0]):
        table_row = f"""
            <tr>
                <td>{branch_info.iloc[row_num].Branch}</td>
                <td>{branch_info.iloc[row_num].Description}</td>
                <td>{branch_info.iloc[row_num].SerialNumber}</td>
                <td>{branch_info.iloc[row_num].Qty}</td>
            </tr>
        """
        table_rows += table_row 
    return table_rows


def get_recepients(branch, email_df):
    receipients = []
    # find a match in main sheet
    if branch in email_df["Branch"].to_list():
        
        # get all the emails related to that specific branch
        branch_with_emails = email_df[email_df["Branch"] == branch]
        num_colums = branch_with_emails.shape[1]
        
        for i in range(1, num_colums):
            email = branch_with_emails.iloc[0, i]
            
            # remove empty cells and create a recepients list
            if pd.notna(email):
                receipients.append(f"{email}")
    return receipients    


def record_data(branch, email_type, recording_df, recorded_by):
# update excel sheet with timestamp of when email was sent
    row_index = recording_df.index[recording_df["Branch"] == branch].to_list()
    recording_df.at[row_index[0], email_type] = pd.to_datetime(datetime.now(), format="%d-%m-%Y")
    
    if email_type == "2nd Email Chaser":
        recording_df.at[row_index[0], recording_df.columns[recording_df.columns.get_loc("Sent By")+2]] = str(recorded_by)
    else:
        recording_df.at[row_index[0], "Sent By"] = str(recorded_by)
    