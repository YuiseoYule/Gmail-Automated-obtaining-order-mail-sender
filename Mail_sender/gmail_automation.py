import smtplib
import ssl
import os
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

# 1. Get the absolute path of the directory containing this script (gmail_automation.py)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 2. Join that directory with the filename to get the full path
DOTENV_PATH = os.path.join(BASE_DIR, "googleid.env")

# Usage example (explicitly passing the path is often safer)
load_dotenv(dotenv_path=DOTENV_PATH)


def load_email_config(env_path: str = DOTENV_PATH) -> dict:
    """
    Loads email configuration from the specified .env file.

    Args:
        env_path (str): Path to the .env file.

    Returns:
        dict: A dictionary containing SMTP configuration.

    Raises:
        FileNotFoundError: If the .env file does not exist.
        ValueError: If required environment variables are missing.
    """
    if not os.path.exists(env_path):
        raise FileNotFoundError(f"The configuration file {env_path} was not found.")

    # Load the specific env file
    load_dotenv(dotenv_path=env_path)

    config = {
        "smtp_server": os.getenv("SMTP_SERVER"),
        "smtp_port": int(os.getenv("SMTP_PORT", 587)),  # Default to 587 if not set
        "sender_email": os.getenv("SENDER_EMAIL"),
        "app_password": os.getenv("APP_PASSWORD")
    }

    # Validate critical config
    if not config["sender_email"] or not config["app_password"]:
        raise ValueError("Missing SENDER_EMAIL or APP_PASSWORD in configuration.")

    return config


def send_gmail(message: MIMEMultipart):
    """
    Connects to the SMTP server and sends a fully constructed email message.
    The message object can contain text, HTML, images, or files.

    Args:
        message (MIMEMultipart): The email object created by your external script.
                                 It must have 'To', 'Subject', and content already set.
    """
    try:
        # 1. Load Configuration
        config = load_email_config()

        # 2. Validate Recipient in Message
        if not message["To"]:
            raise ValueError("The message object must have a 'To' header set.")

        # 3. Secure Connection Context
        context = ssl.create_default_context()

        # 4. Connect and Send
        with smtplib.SMTP(config["smtp_server"], config["smtp_port"]) as server:
            server.starttls(context=context)  # Secure the connection
            server.login(config["sender_email"], config["app_password"])

            # Send the email
            # We use the sender from config, and recipient from the message object
            server.sendmail(
                config["sender_email"],
                message["To"],
                message.as_string()
            )

        print(f"Email successfully sent to {message['To']}")

    except Exception as e:
        print(f"Failed to send email: {e}")