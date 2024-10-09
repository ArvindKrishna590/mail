import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import streamlit as st

def send_custom_emails(sender_email, sender_password, subject_template, body_template, smtp_server, smtp_port, sheet):
    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(sender_email, sender_password)

        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        header = sheet.row_values(1)
        new_col_index = len(header) + 1

        sheet.update_cell(1, new_col_index, current_time)

        for i in range(2, sheet.row_count + 1):
            row = sheet.row_values(i)
            if len(row) < 2:
                break

            name = row[0]
            recipient_email = row[1]

            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            
            personalized_subject = subject_template.replace("{name}", name)
            msg['Subject'] = personalized_subject

            personalized_body = body_template.replace("{name}", name)
            msg.attach(MIMEText(personalized_body, 'html')) 
            
            try:
                server.send_message(msg)
                sheet.update_cell(i, new_col_index, 'Yes')
            except Exception as e:
                sheet.update_cell(i, new_col_index, 'No')  

    except Exception as e:
        st.error(f"Error: {e}")
    finally:
        st.success(f"Email sent")
        server.quit()

def connect_to_google_sheet(sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(r"\google_sheet_mail.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).sheet1 
    return sheet

def email_sender_app():
    st.title("Custom Email Sender")

    sender_email = st.text_input("Sender Email")
    sender_password = st.text_input("Sender Password", type="password")

    subject_template = st.text_area("Email Subject (HTML not allowed)", value="Hello {name}")
    body_template = st.text_area("Email Body (HTML allowed)", value="Hi <b>{name}</b>,<br><br>This is a <i>custom message</i> for you.")

    st.write("Tip: Use `{name}` as a placeholder in the email subject and body. It will be replaced with the actual recipient's name.")

    smtp_server = st.text_input("SMTP Server", value="smtp.gmail.com")
    smtp_port = st.number_input("SMTP Port", value=465)

    sheet_name = st.text_input("Google Sheet Name", value="Mail_List")

    if st.button("Send Emails"):
        if sender_email and sender_password and subject_template and body_template and sheet_name:
            sheet = connect_to_google_sheet(sheet_name)
            send_custom_emails(sender_email, sender_password, subject_template, body_template, smtp_server, smtp_port, sheet)
        else:
            st.error("Please fill in all the fields.")

if __name__ == "__main__":
    email_sender_app()
