# Proof & Company Email Automation

## Overview

This project improves the email reminder process for Proof & Company, a local alcoholic beverage supplier, by introducing automation and enhancing efficiency. The company's operations involve supplying beverages to bars and restaurants in eco-friendly crates, which are later collected for recycling. Before sending invoices, Proof & Company sends out two email reminders weekly. These emails traditionally required manual input of data from an auto-generated excel sheet, a time-consuming and error-prone process.

### Challenges

- **Manual Data Entry:** Data for the email content, including tables and attachments, was previously inputted manually, leading to inefficiencies and potential errors.
- **Repetitive Quarterly Process:** As the company repeats this process quarterly, the manual effort invested in each cycle became a significant overhead.

## Solution

The project addresses these challenges through a two-fold solution:

1. **Email Automation Script with Python as Core programming language**
    - **Data Processing with Pandas:**
      Utilizes Pandas to seamlessly read data from an Excel file into a dataframe and dynamically generate data for the email table, eliminating the need for manual input. This dynamic approach ensures that email content is always up-to-date and accurate.
    
    - **Automated Email Sending:**
      Leveraged Python email library's MIME capabilities and SMTP library to automate the process of sending reminder emails programmatically. Each email is enriched with a standard template, inline image, and attachment, providing comprehensive information to recipients.
    
    - **Func_timeout Library:**
      Incorporates the Func_timeout library to add timeouts to function calls, enhancing reliability, and gracefully handling delays or errors. Additionally, a delay of 3 seconds between each email sending is implemented to ensure compliance with Microsoft's rate limit while optimising delivery of automated emails.

2. **Streamlit Web Application**
    - **User-Friendly Interface:**
      Provides a simple and intuitive interface using Streamlit, catering to users with varying levels of technical expertise.
    
    - **Excel File Upload:**
      Enables users to upload Excel files, streamlining the data input process.
    
    - **Content Selector and Preview Section:**
      Offers a select box for users to easily switch bewtween email templates for different stages of ecoTote collection process. A preview of the email content will be provided before sending, ensuring accuracy.
    
    - **XlsxWriter Library:**
      Utilised XlsxWriter to update and write data to Excel sheets, ensuring seamless integration with the existing Excel-based workflow of Proof & Company. The Excel file is made available for users to download for tracking purposes.


### Business Impact

- **Efficiency Gains:** The automation of the email reminder process has significantly reduced the time and effort required, resulting in improved overall efficiency.
- **Error Reduction:** Eliminating manual data entry minimizes the risk of errors in email content, contributing to a more reliable communication process.
- **Resource Savings:** With the streamlined automation, Proof & Company estimates saving the equivalent of two teams' worth of man-hours per quarter.
- **Scalability:** The automated script and web application are designed to handle the repetitive quarterly process with ease, providing scalability for future business needs.

