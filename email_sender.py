import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import Optional
import os

class EmailConfig:
    """Email configuration settings"""
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    SENDER_EMAIL = os.getenv("SENDER_EMAIL")
    SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
    
    @classmethod
    def validate_config(cls):
        """Validate that required environment variables are set"""
        if not cls.SENDER_EMAIL:
            raise ValueError("SENDER_EMAIL environment variable is required")
        if not cls.SENDER_PASSWORD:
            raise ValueError("SENDER_PASSWORD environment variable is required")

class EmailSender:
    """Handles email sending operations"""
    
    def __init__(self, config: EmailConfig = EmailConfig):
        self.config = config
        # Validate configuration on initialization
        self.config.validate_config()

    def validate_email(self, email: str) -> bool:
        """Basic email format validation"""
        return '@' in email and '.' in email.split('@')[1]

    def send_email(self, recipient: str, subject: str, body: str, attachment_path: Optional[str] = None) -> tuple[bool, Optional[str]]:
        """
        Send an email using Gmail SMTP
        
        Args:
            recipient: Recipient email address
            subject: Email subject
            body: Email body content
            attachment_path: Optional path to PDF attachment
            
        Returns:
            tuple: (success_status, error_message if any)
        """
        # Validate recipient email
        if not self.validate_email(recipient):
            return False, "Invalid recipient email address"

        # Create message
        msg = MIMEMultipart('alternative')
        msg["From"] = self.config.SENDER_EMAIL
        msg["To"] = recipient
        msg["Subject"] = subject
        # Add body as HTML
        msg.attach(MIMEText(body, "html"))

        # Attach PDF if provided
        if attachment_path and os.path.exists(attachment_path):
            try:
                with open(attachment_path, "rb") as f:
                    pdf = MIMEApplication(f.read(), _subtype="pdf")
                    pdf.add_header('Content-Disposition', 'attachment', 
                                 filename=os.path.basename(attachment_path))
                    msg.attach(pdf)
            except Exception as e:
                return False, f"Error attaching PDF: {str(e)}"

        try:
            # Connect to Gmail SMTP Server
            server = smtplib.SMTP(self.config.SMTP_SERVER, self.config.SMTP_PORT)
            server.starttls()  # Enable TLS
            
            # Login and send
            server.login(self.config.SENDER_EMAIL, self.config.SENDER_PASSWORD)
            server.sendmail(self.config.SENDER_EMAIL, recipient, msg.as_string())
            server.quit()
            
            return True, None

        except smtplib.SMTPAuthenticationError:
            return False, "Authentication failed. Check your email and app password."
        except smtplib.SMTPException as e:
            return False, f"SMTP error occurred: {str(e)}"
        except Exception as e:
            return False, f"An unexpected error occurred: {str(e)}"

def main():
    """Example usage"""
    sender = EmailSender()
    recipient = "jeelkakadiya12@gmail.com"
    subject = "Test Email"
    body = "Hello, this is a test email from SMTP."
    
    success, error = sender.send_email(recipient, subject, body)
    
    if success:
        print("✅ Email sent successfully!")
    else:
        print(f"❌ Failed to send email: {error}")

if __name__ == "__main__":
    main() 