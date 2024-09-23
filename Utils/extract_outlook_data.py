import imaplib
import email
from email.header import decode_header
from pathlib import Path
import os
import logging

class ExtractData(object):
    """docstring for ExtractData."""
    def __init__(self,server: str,  email_user: str, email_pass: str, port: str, startdate = None, enddate = None):
        super(ExtractData, self).__init__()
        self.startdate = startdate
        self.enddate = enddate
        self.optflag = self.startdate is not None and self.enddate is not None

        # Connect to the server
        try:
            # Connect to the server
            self.mail = imaplib.IMAP4_SSL(server, port)
            self.mailflag = True
            # Login to the account
            self.mail.login(email_user, email_pass)
        except imaplib.IMAP4.error as e:
            logging.error(f"Error: {str(e)}")
            self.mailflag = False
            self.mail.logout()  # Ensure proper logout before retrying connection

    # Search for emails 
    def search_email(self):
        # Select the mailbox you want to download emails from
        self.mail.select("inbox")
        logging.info("Mailbox inbox folder selected.")

        if self.optflag:
            #try:
            # Search for emails within the date range
            status, messages = self.mail.search(None, f'(SINCE "{self.startdate}" BEFORE "{self.enddate}")')
            #except imaplib.IMAP4.error as e:
                #logging.error(f"Dates must be in 'DD-MMM-YYYY' format. Error: {str(e)}")
                #self.mail.logout()  # Ensure proper logout before retrying connection
        else:
            # Search for all emails in the mailbox
            status, messages = self.mail.search(None, "ALL")

        email_ids = messages[0].split()
        return email_ids

    # Function to decode email subjects
    def decode_subject(self, subject: object):
        decoded_subject, encoding = decode_header(subject)[0]
        if isinstance(decoded_subject, bytes):
            for enc in ["utf-8", "big5", "latin1", "ascii"]:
                try:
                    decoded_subject = decoded_subject.decode(encoding if encoding else enc)
                    break
                except UnicodeDecodeError as e:
                    logging.warn(f"An unexpected warn occurred. Warn: {str(e)}")
                    continue

        if not isinstance(decoded_subject, str):
            decoded_subject = str(decoded_subject)
        return decoded_subject

    # Function to extract the email content
    def get_email_content(self, msg: object):
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                try:
                    body = part.get_payload(decode=True).decode()
                except Exception as e:
                    logging.warn(f"An unexpected warn occurred. Warn: {str(e)}")
                    continue
                if "attachment" not in content_disposition:
                    if content_type == "text/plain":
                        return body, content_type
                    elif content_type == "text/html":
                        return body, content_type
        else:
            try:
                body = msg.get_payload(decode=True).decode()
                content_type = msg.get_content_type()
                return body, content_type
            except UnicodeDecodeError as e:
                logging.warn(f"An unexpected warn occurred. Warn: {str(e)}")
                pass

    # Function to save attachments
    def save_attachment(self, part, attachment_folder: str, subject: object):
        if part.get("Content-Disposition") and "attachment" in part.get("Content-Disposition"):
            filename = part.get_filename()
            if filename:
                filename = decode_header(filename)[0][0]
                if isinstance(filename, bytes):
                    filename = filename.decode()
                filename = f"{subject}_{filename}"

                # work around
                filename = filename.replace(" - ", "_").replace(" ", "_").replace("/", "_").replace("-", "_")
                filepath = os.path.join(attachment_folder, filename)
                
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
            
                logging.info(f"Saved attachment: {filename}.")

    def iterate_emails(self, output_folder, attachment_folder):
        # Iterate through all email IDs and fetch each email
        for email_id in self.search_email():
            status, msg_data = self.mail.fetch(email_id, "(RFC822)")
            msg_content = msg_data[0][1]
            msg = email.message_from_bytes(msg_content)

            # Decode email subject
            try:
                subject = self.decode_subject(msg["Subject"])
            except TypeError as e:
                logging.warn(f"Subject is Null. Warn: {str(e)}")
                continue

            # Create a safe file name
            file_name = ''.join(c if c.isalnum() else '_' for c in subject)

            # Get the email content
            try:
                email_body, content_type = self.get_email_content(msg)
            except TypeError as e:
                logging.warn(f"$subject: Email Content is Nul. Warn: {str(e)}")
                continue

            # Get the email received date
            received_date = email.utils.parsedate_to_datetime(msg["Date"]).strftime('%Y-%m-%d')
            received_mon = email.utils.parsedate_to_datetime(msg["Date"]).strftime('%Y-%b')

            output_mon_folder = output_folder / Path(received_mon)
            output_mon_folder.mkdir(parents=True, exist_ok=True)

            attachment_mon_folder = attachment_folder / Path(received_mon)
            attachment_mon_folder.mkdir(parents=True, exist_ok=True)

            # Save the email body to a file
            if content_type == "text/plain":
                file_path = os.path.join(output_mon_folder, f"{received_date}_{file_name}.txt")
            elif content_type == "text/html":
                file_path = os.path.join(output_mon_folder, f"{received_date}_{file_name}.html")

            with open(file_path, "w", encoding="utf-8") as f:
                f.write(email_body)

            logging.info(f"Saved email: {received_date}_{subject}.")

            # Save attachments
            if msg.is_multipart():
                for part in msg.walk():
                    self.save_attachment(part, attachment_mon_folder, subject)