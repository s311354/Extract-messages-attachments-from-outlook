import imaplib
import email
from email.header import decode_header
import os
import logging

class ExtractData(object):
    """docstring for ExtractData."""
    def __init__(self, email_user: str, email_pass: str, startdate = None, enddate = None):
        super(ExtractData, self).__init__()
        self.startdate = startdate
        self.enddate = enddate
        self.optflag = self.startdate is not None and self.enddate is not None

        # Connect to the server
        try:
            # Connect to the server
            self.mail = imaplib.IMAP4_SSL("outlook.office365.com")
            # Login to the account
            self.mail.login(email_user, email_pass)
        except imaplib.IMAP4.error:
            raise ValueError("Failed to login, check your email and password")

    # Search for emails 
    def search_email(self):
        # Select the mailbox you want to download emails from
        self.mail.select("inbox")

        if self.optflag:
            try:
                # Search for emails within the date range
                status, messages = self.mail.search(None, f'(SINCE "{self.startdate}" BEFORE "{self.enddate}")')
            except imaplib.IMAP4.error:
                raise ValueError("Dates must be in 'DD-MMM-YYYY' format")
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
                except UnicodeDecodeError:
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
                except:
                    continue
                if "attachment" not in content_disposition:
                    if content_type == "text/plain":
                        return body, content_type
                    elif content_type == "text/html":
                        return body, content_type
        else:
            body = msg.get_payload(decode=True).decode()
            content_type = msg.get_content_type()
            return body, content_type

    # Function to save attachments
    def save_attachment(self, part, attachment_folder: str, subject: object):
        if part.get("Content-Disposition") and "attachment" in part.get("Content-Disposition"):
            filename = part.get_filename()
            if filename:
                filename = decode_header(filename)[0][0]
                if isinstance(filename, bytes):
                    filename = filename.decode()
                filename = f"{subject}_{filename}"
                filepath = os.path.join(attachment_folder, filename)
                
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                
                print(f"Saved attachment: {filename}")

    def iterate_emails(self, output_folder, attachment_folder):
        # Iterate through all email IDs and fetch each email
        for email_id in self.search_email():
            status, msg_data = self.mail.fetch(email_id, "(RFC822)")
            msg_content = msg_data[0][1]
            msg = email.message_from_bytes(msg_content)

            # Decode email subject
            try:
                subject = self.decode_subject(msg["Subject"])
            except TypeError:
                logging.info('Subject is Null')
                continue

            # Create a safe file name
            file_name = ''.join(c if c.isalnum() else '_' for c in subject)

            # Get the email content
            try:
                email_body, content_type = self.get_email_content(msg)
            except TypeError:
                logging.info('Email Content is Null')
                continue

            # Get the email received date
            received_date = email.utils.parsedate_to_datetime(msg["Date"]).strftime('%Y-%m-%d')

            # Save the email body to a file
            if content_type == "text/plain":
                file_path = os.path.join(output_folder, f"{received_date}_{file_name}.txt")
            elif content_type == "text/html":
                file_path = os.path.join(output_folder, f"{received_date}_{file_name}.html")

            with open(file_path, "w", encoding="utf-8") as f:
                f.write(email_body)

            print(f"Saved email: {received_date}_{subject}")

            # Save attachments
            if msg.is_multipart():
                for part in msg.walk():
                    self.save_attachment(part, attachment_folder, subject)