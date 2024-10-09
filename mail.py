import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import streamlit as st
import pandas as pd

def send_custom_emails(sender_email, sender_password, subject_template, body_template, smtp_server, smtp_port, data):
    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(sender_email, sender_password)

        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data[current_time] = "Pending"  # Add a new column with the current date-time

        for index, row in data.iterrows():
            if pd.isnull(row['Name']) or pd.isnull(row['Email']):
                continue  # Skip rows without name or email

            name = row['Name']
            recipient_email = row['Email']

            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email

            personalized_subject = subject_template.replace("{name}", name)
            msg['Subject'] = personalized_subject

            personalized_body = body_template.replace("{name}", name)
            msg.attach(MIMEText(personalized_body, 'html')) 
            
            try:
                server.send_message(msg)
                data.at[index, current_time] = 'Yes'  # Update status in DataFrame
            except Exception as e:
                data.at[index, current_time] = 'No'  # Update status in DataFrame

    except Exception as e:
        st.error(f"Error: {e}")
    finally:
        st.success("Email process completed!")
        server.quit()
        return data

def email_sender_app():
    st.title("Custom Email Sender")

    sender_email = st.text_input("Sender Email")
    sender_password = st.text_input("Sender Password", type="password")

    subject_template = st.text_area("Email Subject (HTML not allowed)", value="Hello {name}")
    body_template = st.text_area("Email Body (HTML allowed)", value="Hi <b>{name}</b>,<br><br>This is a <i>custom message</i> for you.")

    st.write("Tip: Use `{name}` as a placeholder in the email subject and body. It will be replaced with the actual recipient's name.")

    smtp_server = st.text_input("SMTP Server", value="smtp.gmail.com")
    smtp_port = st.number_input("SMTP Port", value=465)

    uploaded_file = st.file_uploader("Upload csv file", type=["csv"])

    if uploaded_file is not None:
        data = pd.read_csv(uploaded_file)  # Load the uploaded CSV file into a DataFrame
        st.write("Data preview:", data.head())  # Display the first few rows of the uploaded sheet

        if st.button("Send Emails"):
            if sender_email and sender_password and subject_template and body_template:
                updated_data = send_custom_emails(sender_email, sender_password, subject_template, body_template, smtp_server, smtp_port, data)
                st.write("Updated Email Status:")
                st.write(updated_data)  # Display the updated DataFrame with email status
            else:
                st.error("Please fill in all the fields.")

if __name__ == "__main__":
    email_sender_app()
