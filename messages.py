style = """
    <style>
        body{
            color: rgb(34, 34, 34);
            font-family: Georgia, 'Times New Roman', Times, serif;
            font-size: 11.0pt;
        }
        p{
            color: rgb(34, 34, 34);
            margin-top: 1em;
            margin-bottom: 1em;
        }
        table, td, th{
            border: 1px solid;
            border-collapse: collapse;
        }
        .regards-msg{
            color: rgb(34, 34, 34);
        }
    </style>
"""

def message_body(table_rows):
    return f"""
    <html> 
        <head>
            {style}
        </head>
        <body>
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
            
            <table>
                <thead>
                    <tr>
                        <th>Branch</th>
                        <th>Description</th>
                        <th>Serial Number</th>
                        <th>Qty</th>
                    </tr>
                </thead>
                <tbody>
                    {table_rows}
                </tbody>
            </table>
            
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
            <img src="cid:ecoTotesQR">
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
        </body>
    </html>    
"""

message_body_text = """
    Hello Team, \n
    Thank you for your continued support in your ecoSPIRITS Program!\n
    To ensure a steady turnover of ecoTOTES, we have implemented a 3 month Flagging System for the return of all empty ecoTOTES. Our ecoTOTES are a valuable commodity and their re-use is central to the success of our sustainable distribution program. We appreciate your help in our efforts in recouping them.\n
    The following below is a report of all the ecoTOTES at your venue that have been with you for 3 months or more.\n
    Please do account for these totes and get back to us confirming either they are:

    - Still in use
    - Empty, ready for collection
    - Lost
    
    ~ Table Content goes here ~

    Do kindly send us a brief reply to let us know as well if the above ecoTOTES have been accounted for. We aim to send these reminders out every three months.\n
    In addition, please reach out once there are empty ecoTOTES to be collected and we will facilitate the collection accordingly.\n
    For your reference, here are some images of the ecoTOTES and where you can find the serial numbers on the assets:\n
    Note: \n
    We'll also follow up on this email in a week's time to check on the status should we not hear back from you by then. \n
    Thereafter will be another grace period of two weeks to confirm the status of the mentioned totes. Failure to do so will result in an automatic billing of SGD200 (Before tax) per ecoTOTE not returned.\n
    Thank you so much for your time on this matter and let us know if you have any questions.

    Thanks and best regards.

    Proof Team.
"""

def message_2(table_rows):
    return f"""
    <html> 
        <head>
            {style}
        </head>
        <body style="color: black;">
            <p>Hello Team,</p>
            <p>We hope your day has been going great so far!</p>
            <p>
                We would like to follow up on the outstanding ecoTOTE(S) at your venue as per our previous email:
            </p>     
            <table>
                <thead>
                    <tr>
                        <th>Branch</th>
                        <th>Description</th>
                        <th>Serial Number</th>
                        <th>Qty</th>
                    </tr>
                </thead>
                <tbody>
                    {table_rows}
                </tbody>
            </table>
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
        </body>
    </html>    
"""

message_2_text = """
    Hello Team,\n
    We hope your day has been going great so far!\n
    We would like to follow up on the outstanding ecoTOTE(S) at your venue as per our previous email:\n

    ~ *Table Content goes here* ~
 
    Do take note that should we not receive any confirmation about the status of the mentioned totes within the next two weeks, we will deem these totes "Missing" and move forward with charging the value of these totes to your account. As stated in the previous email, each ecoTote is valued at SGD200 (Before Tax).\n
    Once again, we greatly appreciate the time taken to ensure that the ecoTOTES are properly maintained so that these valuable assets can continue serving their purpose in our sustainability movement.\n

    Kind regards,

    Proof Team
"""

def app_instructions():
    return f"""
        *Empower your communication effortlessly with automation. This intuitive platform streamlines the process of sending email reminders, ensuring a seamless experience for tracking.*

        ### Get Started:

        #### 1. Upload Excel File:
        - Upload your Excel file containing Mastersheet, Branch with Emails, and Recording sheets.
        - Ensure all sheets and columns are accurately named and included for seamless operation.

        #### 2. **Mastersheet Requirements:**
        - Confirm the presence and accuracy of essential columns: Branch, Description, Qty, and SerialNumber.
        - Any discrepancies may affect the application's functionality.

        #### 3. **Select Email Type:**
        - Choose between 1st Email Chaser or 2nd Email Chaser to tailor your communication strategy.

        #### 4. **Send Emails:**
        - Initiate the email sending process with a single click.
        - Relax as the application automates delivery, ensuring timely communication with recipients.
        - Note: The application will attempt up to five retries for failed deliveries. Any undelivered emails due to invalid addresses or server disconnection will be recorded in the recording sheet.

        #### 5. **Feedback and Tracking:**
        - Monitor email delivery status and potential errors via the Recording sheet in your Excel file.
        - Regularly review the Recording sheet to address any issues promptly and maintain effective communication.

        #### Need Assistance?

        For any inquiries or assistance regarding the EcoTotes Chaser application, please don't hesitate to reach out to our dedicated support team.

        Thank you for choosing EcoTotes Chaser by Proof and Company. We are committed to optimizing your email communication experience and enhancing your productivity.
    """