import re
import imaplib
import email
from openpyxl import load_workbook
import time
import phonenumbers
import logging
from email.header import decode_header

# Constants
RETRY_ATTEMPTS = 3
RETRY_DELAY_SECONDS = 10
WAIT_INTERVAL_SECONDS = 30

# Configuration
config = {
    "server": "md-uk-1.webhostbox.net",
    "username": "jnjenga@africa118.com",
    "password": "Naivasha@118"
}

# Sheet headers
SHEET_HEADERS = ["Name", "Phone", "Email", "Business Name", "Location"]

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def parse_email_body(body: str) -> tuple:
    """
    Extracts relevant information from the body of an email message.

    Args:
        body (str): The body of the email message.

    Returns:
        tuple: A tuple containing the extracted information:
            - name (str): The name extracted from the body.
            - phone (str): The phone number extracted from the body.
            - email (str): The email extracted from the body.
            - business_name (str): The business name extracted from the body.
            - location (str): The location extracted from the body.
    """
    name_match = re.search(r"Name: (.+)", body)
    phone_match = re.search(r"Phone number: (.+)", body)
    email_match = re.search(r"Email: (.+)", body)
    business_location_pattern = re.compile(r"(.+)\n\n.+ requesting to be listed as a manager", re.DOTALL)
    business_location_match = business_location_pattern.search(body)

    if business_location_match:
        business_location_text = business_location_match.group(1).strip()
        business_name, location = business_location_text.split("\n", 1)
    else:
        business_name = "N/A"
        location = "N/A"

    # Additional parsing logic as needed

    return name_match.group(1), phone_match.group(1), email_match.group(1), business_name, location

def process_email(email_msg, subject, sheets):
    if email_msg.is_multipart():
        for part in email_msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                body = part.get_payload(decode=True).decode("utf-8")
                break
    else:
        body = email_msg.get_payload(decode=True).decode("utf-8")

    name, phone, email, business_name, location = parse_email_body(body)

    try:
        parsed_phone = phonenumbers.parse(phone, "US")
        formatted_phone = phonenumbers.format_number(parsed_phone, phonenumbers.PhoneNumberFormat.E164)
    except Exception as e:
        logging.error(f"Error while parsing phone number: {e}")
        return

    sheet = sheets.get(subject)
    if not sheet:
        logging.warning("Unrecognized subject: %s", subject)
        return

    sheet.append([name, formatted_phone, email, business_name, location])

def search_emails(mailbox):
    for retry in range(RETRY_ATTEMPTS):
        try:
            status, email_ids = mailbox.search(None, "UNSEEN")
            return email_ids[0].split() if email_ids else []
        except Exception as e:
            logging.error(f"Error while searching for emails: {e}")
            logging.info(f"Retrying after {RETRY_DELAY_SECONDS} seconds...")
            time.sleep(RETRY_DELAY_SECONDS)
    return []

def main():
    # Connect to email account
    with imaplib.IMAP4_SSL(config["server"]) as mail:
        mail.login(config["username"], config["password"])
        mail.select("inbox")

        # Create Excel workbooks and sheets
        sheets = {}

        for sheet_name in ["Ownership Requests", "Management Requests", "Profile Suspensions"]:
            path = f"{sheet_name.lower().replace(' ', '_')}.xlsx"
            workbook = load_workbook(path)
            worksheet = workbook.active
            worksheet.title = sheet_name
            if worksheet.max_row == 1:
                worksheet.append(SHEET_HEADERS)
            sheets[sheet_name] = worksheet

        try:
            while True:
                email_ids = search_emails(mail)

                for email_id in email_ids:
                    _, data = mail.fetch(email_id, "(BODY.PEEK[])")
                    raw_email = data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    subject = msg["subject"]
                    process_email(msg, subject, sheets)

                    mail.store(email_id, '+FLAGS', '\\Seen')

                for sheet_name, worksheet in sheets.items():
                    workbook = worksheet.parent
                    workbook.save(f"{sheet_name.lower().replace(' ', '_')}.xlsx")
                    workbook.close()

                time.sleep(WAIT_INTERVAL_SECONDS)

        except KeyboardInterrupt:
            logging.info("Email parsing stopped.")

if __name__ == "__main__":
    main()
