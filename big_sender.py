import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QLineEdit, QPushButton, QTextEdit, QPlainTextEdit, QFileDialog
)
from PyQt6.QtGui import QDesktopServices
from PyQt6.QtCore import QUrl, Qt
import pandas as pd 

import os
import fitz
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
###~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~START OF FRONTEND~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~###

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Big Sender")
        self.setGeometry(100, 100, 600, 600)

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Main layout
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # colour scheme
        with open("style.css", "r") as file:
            self.setStyleSheet(file.read())

        # Gmail adress input
        self.email_label = QLabel("Gmail Address:")
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("@gmail.com")
        layout.addWidget(self.email_label)
        layout.addWidget(self.email_input)

        # Password input
        def open_GpassHelp_url():
            QDesktopServices.openUrl(QUrl('https://support.google.com/accounts/answer/185833?hl=en'))

        self.password_label = QLabel("Google App Password:")
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("xxxx xxxx xxxx xxxx")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_button = QPushButton("  ?  ")
        self.password_button.clicked.connect(open_GpassHelp_url)

        # Password layout
        password_layout = QHBoxLayout()
        password_layout.addWidget(self.password_input)
        password_layout.addWidget(self.password_button)
        layout.addWidget(self.password_label)
        layout.addLayout(password_layout)

        # CSV File Selector
        self.csv_label = QLabel("To:")
        self.csv_path = QLineEdit()
        self.csv_path.setPlaceholderText("Select CSV File...")
        self.csv_button = QPushButton("Browse")
        self.csv_button.clicked.connect(self.load_csv_file)

        csv_layout = QHBoxLayout()
        csv_layout.addWidget(self.csv_path)
        csv_layout.addWidget(self.csv_button)
        layout.addWidget(self.csv_label)
        layout.addLayout(csv_layout)

        # Subject
        self.subject_label = QLabel("Email Subject:")
        self.subject_input = QLineEdit()
        layout.addWidget(self.subject_label)
        layout.addWidget(self.subject_input)

        # Email Body Input
        self.body_label = QLabel("Email Body:")
        self.body_opener_label = QLabel("Hello [recipient] ,")
        self.body_input = QTextEdit()
        self.body_input.setPlaceholderText("'Hello' and recipents name added automatically. \nPlease enter the Email Body here...")
        layout.addWidget(self.body_label)
        layout.addWidget(self.body_opener_label)
        layout.addWidget(self.body_input)

        # Cover Letter Input
        self.cover_letter_label = QLabel("Cover Letter:")
        self.cover_letter_opener_label = QLabel("Hello [recipient] ,")
        self.cover_letter_input = QTextEdit()
        self.cover_letter_input.setPlaceholderText("'Hello' and recipents name added automatically. \nPlease enter the Cover Letter body here...")
        layout.addWidget(self.cover_letter_label)
        layout.addWidget(self.cover_letter_opener_label)
        layout.addWidget(self.cover_letter_input)

        # CV File Selector
        self.cv_label = QLabel("CV PDF File:")
        self.cv_path = QLineEdit()
        self.cv_path.setPlaceholderText("Select CV PDF File...")
        self.cv_button = QPushButton("Browse")
        self.cv_button.clicked.connect(self.load_cv_file)

        cv_layout = QHBoxLayout()
        cv_layout.addWidget(self.cv_path)
        cv_layout.addWidget(self.cv_button)
        layout.addWidget(self.cv_label)
        layout.addLayout(cv_layout)
    
        # Terminal Output Box
        self.terminal_label = QLabel("Terminal Output:")
        self.terminal_output = QPlainTextEdit()
        self.terminal_output.setReadOnly(True)
        layout.addWidget(self.terminal_label, alignment=Qt.AlignmentFlag.AlignBottom)
        layout.addWidget(self.terminal_output, alignment=Qt.AlignmentFlag.AlignBottom)

        # Send Button
        self.send_button = QPushButton("Send Emails")
        self.send_button.clicked.connect(self.send_emails)
        layout.addWidget(self.send_button)

        # set layout
        central_widget.setLayout(layout)

###~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~END OF FRONT-END~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~###

###~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~START OF BACKEND~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~###

    # Load CSV file
    def load_csv_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select CSV File", "", "CSV Files (*.csv)")
        if file_path:
            self.csv_path.setText(file_path)

    # Load Cover Letter
    def load_cv_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select CV PDF File", "", "PDF Files (*.pdf)")
        if file_path:
            self.cv_path.setText(file_path)

    def send_emails(self):
        # Collect user inputs
        gmail_user = self.email_input.text()
        gmail_password = self.password_input.text()
        csv_file = self.csv_path.text()
        subject = self.subject_input.text()
        email_body = self.body_input.toPlainText()
        cover_letter_template = self.cover_letter_input.toPlainText()
        cv_pdf = self.cv_path.text()

        # Validate inputs
        if not all([gmail_user, gmail_password, subject, csv_file, cover_letter_template, cv_pdf, email_body]):
            self.terminal_output.appendPlainText("[Error] All fields must be filled out.")
            return

        # Load email addresses from the CSV file
        try:
            email_data = pd.read_csv(csv_file)
        except Exception as e:
            self.terminal_output.appendPlainText(f"[Error] Failed to load CSV file: {e}")
            return

        # Validate Gmail credentials
        if "@gmail.com" not in gmail_user:
            self.terminal_output.appendPlainText("[Error] Please provide a valid Gmail address.")
            return
        
        # validate column names in the CSV file
        valid_columns = ['email', 'recipient', 'company']
        for col in email_data.columns:
            if col not in valid_columns:
                self.terminal_output.appendPlainText(f"[Error] Invalid column name in CSV file: {col} \nPlease use 'email', 'recipient' and 'company' as column names.")
                return

        # Initialise report data
        report_data = []
        email_count_success = 0
        email_count_failed = 0

        # Iterate over email data
        self.terminal_output.appendPlainText("+----------------------~!~-----------------------+")

        for _, row in email_data.iterrows():
            email = str(row.get('email')) if pd.notna(row.get('email')) else None
            recipient = row.get('recipient').title() if pd.notna(row.get('recipient')) else row.get('company').title() if pd.notna(row.get('company')) else 'There'
            company = row.get('company').title() if pd.notna(row.get('company')) else row.get('recipient').title() if pd.notna(row.get('recipient')) else 'There'

            # Initialise default status
            status = "--INVALID EMAIL--"

            # Only proceed if email is valid
            if email and "@" in email:
                # Create the email body with recipient name
                modified_email_body = f"Hello {recipient.title()},\n\n{email_body}"

                # Create a custom cover letter for the recipient
                custom_cover_letter_path = self.create_custom_cover_letter(company, cover_letter_template)

                # Send the email with the customized cover letter
                status = self.send_email(gmail_user, gmail_password, email, subject, modified_email_body, cv_pdf, custom_cover_letter_path)

                # Update email count
                if status == "Sent":
                    email_count_success += 1
                else:
                    email_count_failed += 1

            # Append status to report data
            report_data.append((status, email, recipient, company))

        # Save the report to CSV
        try:
            self.save_report(report_data)
        except Exception as e:
            self.terminal_output.appendPlainText(f"[Error] Failed to save email report: {e}")

        # Final output summary
        self.terminal_output.appendPlainText(f"\nNumber of successful emails sent: {email_count_success}")
        self.terminal_output.appendPlainText(f"Number of emails failed to send: {email_count_failed}")
        self.terminal_output.appendPlainText("+----------------------~!~-----------------------+")

    def send_email(self, gmail_user, gmail_password, recipient_email, subject, email_body, cv_pdf, cover_letter_path):
        try:
            # Set up the email server and login
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(gmail_user, gmail_password)

            # Create the email
            msg = MIMEMultipart()
            msg['From'] = gmail_user
            msg['To'] = recipient_email
            msg['Subject'] = subject

            # Attach the email body
            msg.attach(MIMEText(email_body, 'plain'))

            # Attach the CV PDF
            with open(cv_pdf, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(cv_pdf)}')
                msg.attach(part)

            # Attach the custom cover letter
            with open(cover_letter_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(cover_letter_path)}')
                msg.attach(part)

            # Send the email
            server.send_message(msg)
            server.quit()

            return "Sent"
        except Exception as e:
            return f"Failed: {e}"

    # Save report
    def save_report(self, report_data):
        # Get the current date and time in the format YYYY-MM-DD and HH-MM-SS
        Datestamp = datetime.now().strftime("%Y%m%d")
        Timestamp = datetime.now().strftime("%H-%M-%S")
        report_filename = f"Email_reports/{Datestamp}_email_report_{Timestamp}.csv"  # Filename with timestamp

        # Ensure the directory exists
        os.makedirs("Email_reports", exist_ok=True)

        # Save the report data to CSV with the new filename
        report_df = pd.DataFrame(report_data, columns=["status", "email", "recipient", "company"])
        report_df.to_csv(report_filename, index=False)
        self.terminal_output.appendPlainText(f"Email report created and saved to: \n    {report_filename}")

    # Create a custom cover letter for the company
    def create_custom_cover_letter(self, company, cover_letter_template):
        cover_letter_template = f"Dear {company.title()},\n\n{cover_letter_template}"
    
        output_path = f"Created_Coverletters/{company}_PF_Coverletter.pdf"
        os.makedirs("Created_Coverletters", exist_ok=True)
    
        # Setting up the PDF text box
        pdf = fitz.open()
        page = pdf.new_page()
        text_box = fitz.Rect(72, 72, 540, 720)
        page.insert_textbox(
            text_box,
            cover_letter_template,
            fontsize=12,
            fontname="helv",
            color=(0, 0, 0),
            align=0
        )
        # Export the new coverletter PDF
        pdf.save(output_path)
        pdf.close()
        return output_path

###~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~END OF BACKEND~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~###

# Run the application
def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()