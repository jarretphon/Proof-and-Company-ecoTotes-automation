# Proof & Company Email Automation

## Overview

This project improves the email reminder process for Proof & Company, a local alcoholic beverage supplier, by introducing automation and enhancing efficiency. The company's operations involve supplying beverages to bars and restaurants in eco-friendly crates, which are later collected for recycling. Before sending invoices, Proof & Company sends out two email reminders weekly. These emails traditionally required manual input of data from an auto-generated excel sheet, a time-consuming and error-prone process.

### Challenges

- **Manual Data Entry:** Data for the email content, including tables and attachments, was previously inputted manually, leading to inefficiencies and potential errors.
- **Repetitive Quarterly Process:** As the company repeats this process quarterly, the manual effort invested in each cycle became a significant overhead.

### Solution

The project addresses these challenges through a two-fold solution:

1. **Email Automation Script:**
   - **Data Processing:** Utilizes pandas to seamlessly read data from an Excel file into a dataframe.
   - **Dynamic Table Generation:** Dynamically generates data for the email table, eliminating the need for manual input.
   - **Automated Email Sending:** Utilizes Python's email library to automate the process of sending reminder emails.
   - **Enhanced Email Content:** Includes a standard template, inline image, and attachment in each email, providing comprehensive information to recipients.

2. **Streamlit Web Application:**
   - **User-Friendly Interface:** Provides a simple and intuitive interface using Streamlit, catering to users with varying levels of technical expertise.
   - **Excel File Upload:** Allows users to upload Excel files, streamlining the data input process.
   - **Preview Section:** Offers a preview of the email content before sending, ensuring accuracy.
   - **Secure Email Credentials Input:** Includes an authentication service to ensure security and a secure form for users to input email credentials.

### Business Impact

- **Efficiency Gains:** The automation of the email reminder process has significantly reduced the time and effort required, resulting in improved overall efficiency.
- **Error Reduction:** Eliminating manual data entry minimizes the risk of errors in email content, contributing to a more reliable communication process.
- **Resource Savings:** With the streamlined automation, Proof & Company estimates saving the equivalent of two teams' worth of man-hours per quarter.
- **Scalability:** The automated script and web application are designed to handle the repetitive quarterly process with ease, providing scalability for future business needs.

