import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
from collections import defaultdict

# Zoho Mail IMAP server configuration
imap_server = 'imap.zoho.in'
username = ''
password = ''

def get_sent_emails(imap_server, username, password):
    try:
        # Connect to the IMAP server
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(username, password)

        # Select the 'Sent' mailbox
        mail.select('"Sent"')

        # Search for all emails in the 'Sent' mailbox
        status, data = mail.search(None, 'ALL')
        mail_ids = data[0].split()

        # Dictionary for counting and storing emails sent to each recipient
        recipient_emails = defaultdict(lambda: {'count': 0, 'emails': []})

        # Process each email
        for num in mail_ids:
            status, data = mail.fetch(num, '(RFC822)')
            msg = email.message_from_bytes(data[0][1])

            # Get email subject
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()

            # Get email recipient
            recipient = msg["To"]

            # Get time of email sent
            date_tuple = parsedate_to_datetime(msg["Date"])
            send_time = date_tuple.strftime("%Y-%m-%d %H:%M:%S")

            # Storing email details
            recipient_emails[recipient]['count'] += 1
            recipient_emails[recipient]['emails'].append({'subject': subject, 'send_time': send_time})

        total_sent_emails = sum(info['count'] for info in recipient_emails.values())
        print(f"Total Emails Sent: {total_sent_emails}")

        # Displaying emails sent to each recipient
        for recipient, details in recipient_emails.items():
            print(f"\nRecipient: {recipient}")
            print(f"Emails Sent: {details['count']}")
            for email_info in details['emails']:
                print(f"  Subject: {email_info['subject']}, Sent Time: {email_info['send_time']}")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        if 'mail' in locals():
            # Close the connection if it was successfully created
            mail.close()
            mail.logout()

# Replace with actual email and password
get_sent_emails(imap_server, username, password)
