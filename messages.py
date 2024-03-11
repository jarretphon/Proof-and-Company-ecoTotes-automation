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


def app_instructions():
    return f"""
        *Empower your communication effortlessly with automation. This intuitive platform streamlines the process of sending email reminders, ensuring a seamless experience for tracking.*

        ## Get Started

        ### 1. Upload Excel File:
        - Begin by effortlessly uploading your Excel file through the designated file uploader.
        - The Excel file should comprise three essential sheets: Mastersheet, Branch with Emails, and Recording.

        ### 2. **Mastersheet Requirements:**
        - Elevate your data integrity by confirming the inclusion of these four critical columns in your Mastersheet: Branch, Description, Qty, and SerialNumber.
        - Note the spelling of the columns' names and case-senstivity.
        
        ### 3. **Select Email Type:**
        - Tailor your communication strategy by selecting the type of email you wish to send – whether it's the impactful first email chaser or to follow-up with the second email chaser.

        ### 4. **Input Credentials:**
        - Securely authenticate and enable the sending of emails by entering your email username and password into the designated fields.

        ### 5. **Preview Email (Optional):**
        - Preview the refined email content before initiating the sending process.
        - Click the "Preview Email" button to review the meticulously crafted email structure and contents.

        ### 6. **Send Mail:**
        - Automate the mail sending process by clicking the "Send Mail" button.
        - Sit back as the application seamlessly processes and update your information, ensuring timely and effective delivery to recipients based on your Mastersheet data.

        *Thank you for choosing the Email Sender App. Should you encounter any inquiries or require support, kindly refer to our comprehensive documentation or reach out to our dedicated support team for assistance.*
"""