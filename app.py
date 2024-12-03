import streamlit as st
import pandas as pd
import win32com.client as win32
import pythoncom  # Import the module for COM initialization
import os

# Streamlit App
st.title("Send Emails with Outlook Based on Excel Data")

st.sidebar.header("Instructions")
st.sidebar.write("""
1. Prepare an Excel file with the following columns:
   - `ids`: Recipient email addresses
   - `ccids`: CC email addresses (comma-separated)
   - `subject`: Subject of the email
   - `email message`: Body of the email
2. Upload a common attachment for all emails (optional).
3. Fill in the details and click 'Send Emails'.
""")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
attachment_file = st.file_uploader("Upload Common Attachment", type=["pdf", "docx", "xlsx", "csv", "txt"])

if st.button("Send Emails"):
    if not uploaded_file:
        st.error("Please upload an Excel file.")
    else:
        try:
            # Initialize COM for the thread
            pythoncom.CoInitialize()

            # Read the Excel file
            data = pd.read_excel(uploaded_file)

            # Validate required columns
            required_columns = ["ids", "ccids", "subject", "email message"]
            if not all(column in data.columns for column in required_columns):
                st.error(f"Excel file must have the following columns: {', '.join(required_columns)}")
            else:
                # Initialize Outlook
                outlook = win32.Dispatch("outlook.application")
                st.success("Outlook initialized successfully.")

                # Save attachment if provided
                attachment_path = None
                if attachment_file:
                    attachment_path = os.path.join(os.getcwd(), attachment_file.name)
                    with open(attachment_path, "wb") as f:
                        f.write(attachment_file.read())

                # Send emails for each row
                for _, row in data.iterrows():
                    to_email = row["ids"]
                    cc_emails = row["ccids"] if pd.notna(row["ccids"]) else ""
                    subject = row["subject"]
                    body = row["email message"]

                    # Create email
                    mail = outlook.CreateItem(0)
                    mail.To = to_email
                    mail.CC = cc_emails
                    mail.Subject = subject
                    mail.Body = body

                    # Add attachment
                    if attachment_path:
                        mail.Attachments.Add(attachment_path)

                    # Send email
                    mail.Send()
                    st.success(f"Email sent to {to_email}")

                # Clean up attachment file
                if attachment_path:
                    os.remove(attachment_path)
        except Exception as e:
            st.error(f"An error occurred: {e}")
