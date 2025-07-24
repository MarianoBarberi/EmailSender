import streamlit as st
import pandas as pd
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import make_msgid

st.set_page_config(page_title="Mass Email Sender", layout="centered")
st.title("📧 Mass Email Sender")

# 1. Upload Excel File
st.header("Step 1: Upload Excel File")
excel_file = st.file_uploader("Upload an Excel file with a column named 'Email'", type=["xlsx"])

# 2. Email Composition
st.header("Step 2: Compose Your Email")

sender_email = st.text_input("Your Email", placeholder="you@mail.com")
password = st.text_input("Email Password or App Password", type="password")

subject = st.text_input("Subject")
body = st.text_area("Email Body (HTML supported)", height=200)

attachments = st.file_uploader("Attach Files", accept_multiple_files=True)

# 3. Preview and Send
if st.button("📨 Send Emails"):
    if not excel_file:
        st.error("Please upload an Excel file.")
    elif not sender_email or not password or not subject or not body:
        st.error("Please fill out all email fields.")
    else:
        try:
            df = pd.read_excel(excel_file)
            if "Email" not in df.columns:
                st.error("Excel file must have a column named 'Email'.")
            else:
                st.info("Sending emails... This may take a moment.")
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                    server.login(sender_email, password)

                    for index, row in df.iterrows():
                        recipient = row["Email"]

                        msg = EmailMessage()
                        msg["Subject"] = subject
                        msg["From"] = sender_email
                        msg["To"] = recipient
                        msg.set_content(body, subtype="html")

                        for file in attachments:
                            file_data = file.read()
                            msg.add_attachment(file_data, maintype="application",
                                               subtype="octet-stream", filename=file.name)

                        server.send_message(msg)
                st.success(f"Emails sent successfully to {len(df)} recipients.")
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")