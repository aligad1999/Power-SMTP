import streamlit as st
import pandas as pd
import smtplib
import time
import random
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import queue
import threading
from typing import Dict
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class EmailSender:
    def __init__(self, smtp_account: Dict[str, str]):
        self.smtp_account = smtp_account
        self.account_lock = threading.Lock()
        self.last_send_time = 0
        self.email_queue = queue.Queue()
        self.results = {'success': 0, 'failed': 0}
        self.failed_emails = []
        
    def wait_for_cooldown(self):
        """Implements rate limiting for the email account"""
        with self.account_lock:
            current_time = time.time()
            time_since_last = current_time - self.last_send_time
            if time_since_last < 2:  # Minimum 2 seconds between emails
                time.sleep(2 - time_since_last)
            self.last_send_time = time.time()

    def create_message(self, recipient: str, subject: str, body: str) -> MIMEMultipart:
        msg = MIMEMultipart()
        msg['From'] = self.smtp_account['email']
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        return msg

    def send_single_email(self, recipient: str, subject: str, body: str, max_retries: int = 3) -> bool:
        for attempt in range(max_retries):
            try:
                self.wait_for_cooldown()  # Ensure we don't send too quickly
                
                with smtplib.SMTP("smtp.office365.com", 587, timeout=30) as server:
                    server.starttls()
                    server.login(self.smtp_account['email'], self.smtp_account['password'])
                    msg = self.create_message(recipient, subject, body)
                    server.send_message(msg)
                
                logger.info(f"Successfully sent email to {recipient}")
                return True
                
            except Exception as e:
                logger.error(f"Failed to send email to {recipient} (attempt {attempt + 1}): {str(e)}")
                if attempt == max_retries - 1:
                    self.failed_emails.append((recipient, str(e)))
                    return False
                # Add increasingly longer delays between retries
                time.sleep(5 * (attempt + 1))
        return False

    def process_queue(self):
        while True:
            try:
                email_data = self.email_queue.get_nowait()
                success = self.send_single_email(**email_data)
                with threading.Lock():
                    if success:
                        self.results['success'] += 1
                    else:
                        self.results['failed'] += 1
                self.email_queue.task_done()
            except queue.Empty:
                break

def main():
    st.set_page_config(page_title="Bulk Email Sender", page_icon="📧", layout="wide")
    
    # Custom Styling
    st.markdown(
        """
        <style>
            .main {background-color: #f4f4f4;}
            .stButton>button {background-color: #0078D4; color: white; font-size: 16px; padding: 10px 24px; border-radius: 8px;}
            .error-message {color: red; font-weight: bold;}
            .success-message {color: green; font-weight: bold;}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("📧 Power-SMTP Bulk Email Sender")
    st.markdown("### Upload an Excel file containing email addresses and compose your message below.")

    # File uploader with validation
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"], 
                                    help="Ensure the file contains a column named 'Email'")

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            if 'Email' not in df.columns:
                st.error("The Excel file must contain a column named 'Email'")
                return

            # Clean and validate email list
            df.drop_duplicates(subset=['Email'], inplace=True)
            df.dropna(subset=['Email'], inplace=True)
            email_list = df['Email'].tolist()
            
            st.write(f"**Valid Emails Found:** {len(email_list)}")

            # Email composition
            subject = st.text_input("Email Subject", "Your Subject Here")
            message_body = st.text_area("Email Body", "Type your message here...")
            
            # SMTP account configuration
            smtp_account = {
                "email": st.secrets["email1"],
                "password": st.secrets["password1"]
            }

            if st.button("Send Emails"):
                if not email_list:
                    st.error("No valid emails found in the uploaded file.")
                elif not subject or not message_body:
                    st.error("Please enter both a subject and message body.")
                else:
                    # Initialize sender and progress tracking
                    sender = EmailSender(smtp_account)
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Queue all emails
                    for email in email_list:
                        sender.email_queue.put({
                            'recipient': email,
                            'subject': subject,
                            'body': message_body
                        })

                    # Create worker threads (using fewer threads for single account)
                    num_threads = 2  # Using 2 threads for single account to maintain rate limits
                    threads = []
                    for _ in range(num_threads):
                        thread = threading.Thread(target=sender.process_queue)
                        thread.start()
                        threads.append(thread)

                    # Monitor progress
                    total_emails = len(email_list)
                    while not sender.email_queue.empty():
                        completed = sender.results['success'] + sender.results['failed']
                        progress = completed / total_emails
                        progress_bar.progress(progress)
                        status_text.text(f"Processed: {completed}/{total_emails} emails")
                        time.sleep(0.1)

                    # Wait for all threads to complete
                    for thread in threads:
                        thread.join()

                    # Final status
                    st.success(f"Process completed!\nSuccessfully sent: {sender.results['success']}")
                    if sender.results['failed'] > 0:
                        st.warning(f"Failed to send: {sender.results['failed']}")
                        st.error("Failed emails and reasons:")
                        for email, error in sender.failed_emails:
                            st.write(f"- {email}: {error}")

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            logger.error(f"Application error: {str(e)}", exc_info=True)

if __name__ == "__main__":
    main()
