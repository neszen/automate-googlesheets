import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.message import EmailMessage
import logging 

logging.basicConfig(filename='email_script.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def authenticate_google_sheets():
    # Define the scope
    scope = ["https://www.googleapis.com/auth/spreadsheets", 
             "https://www.googleapis.com/auth/drive"]
    
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    
    client = gspread.authorize(creds)

    return client


def read_google_sheet(sheet_name):
    """ """

    client = authenticate_google_sheets()
   
    sheet = client.open(sheet_name).sheet1  
  
    records = sheet.get_all_records()
    
    return records


def send_email(to_email, subject, body):
    """"""

    msg = EmailMessage()
    msg['From'] = "ravishverma070@gmail.com"  
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.set_content(body)

    smtp_host = "sandbox.smtp.mailtrap.io"  
    smtp_port = 587           
    smtp_user = "63ba1e87853d16"      
    smtp_pass = "b42a391491cda1"    

    with smtplib.SMTP(smtp_host, smtp_port) as smtp:
        smtp.starttls() 
        smtp.login(smtp_user, smtp_pass)
        smtp.send_message(msg)
        print(f"Email sent to {to_email}")


def load_email_template(file_path, name):
    """Load the email template from a file and replace placeholders."""
    with open(file_path, 'r') as file:
        template = file.read()

    return template.format(name=name)


def update_email_status(sheet_name, row_number, status):
    client = authenticate_google_sheets()
    sheet = client.open(sheet_name).sheet1  
    sheet.update_cell(row_number, 8, status)


sheet_name = "sheet_automation"  
records = read_google_sheet(sheet_name)

for row_number, record in enumerate(records, start=2):
    print(record)
    email = record["email"] 
    email_status = record["status"]

    if email_status == 'NEW':
        try:
            # Send welcome email to the user
            user_email_body = load_email_template("templates/message_email.txt", record["name"])
            send_email(email, "Thank you for your feedback!", user_email_body)

            # Send notification email to admin about new user
            admin_email_body = f"New feedback received from {record['name']} ({email})"
            send_email("admin@example.com", "New Feedback Received", admin_email_body)

            # Update the Google Sheet with email status
            update_email_status(sheet_name, row_number, "email_sent")
            
            logging.info(f"Email sent to {email} and admin notified.")
        
        except Exception as e:
            print(f"Failed to send email to {email}: {e}")
            update_email_status(sheet_name, row_number, "email_failed")
            logging.error(f"Failed to send welcome email to {email}: {e}")
            
    elif email_status == 'email_failed':
            try:
                # Retry sending the welcome email to the user
                user_email_body = load_email_template("templates/message_email.txt", record["name"])
                send_email(email, "Thank you for your feedback!", user_email_body)

                # Update Google Sheet with email status
                update_email_status(sheet_name, row_number, "email_sent")
                
                logging.info(f"Email re-sent successfully to {email} after previous failure.")
            
            except Exception as e:
                print(f"Failed to resend email to {email}: {e}")
                update_email_status(sheet_name, row_number, "email_failed")
                logging.error(f"Failed to resend email to {email}: {e}")
    else:
            logging.info(f"No action required for {email} with status: {email_status}.")

