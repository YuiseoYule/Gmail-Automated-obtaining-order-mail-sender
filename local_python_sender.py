import sys
import time
from pathlib import Path

# --- 1. PATH CONFIGURATION (Root Environment) ---
# Resolves the directory where this execution script is located (which is now the root folder)
root_path = Path(__file__).resolve().parent

# Ensure the root path is in sys.path to allow module imports from subdirectories
if str(root_path) not in sys.path:
    sys.path.append(str(root_path))

# Explicitly import from the respective sub-folders
from Mail_sender import gmail_automation
from PJ_mail import PJ_mail_writer

# --- 2. FILE PATHS ---
# All paths now correctly branch out directly from the root_path
DISTRIBUTION_LIST_PATH = root_path / "distribution_list.xlsx"
HTML_TEMPLATE_PATH = root_path / "Mail_form" / "PJ_template.html"
PDF_ATTACHMENT_PATH = root_path / "Mail_form" / "서울대학교 전략컨설팅학회 BICS 소개서.pdf"
IMAGE_1_PATH = root_path / "Mail_form" / "bics_intro_1.png"
IMAGE_2_PATH = root_path / "Mail_form" / "bics_intro_2.png"

def main():
    print("--- Starting BICS Email Automation (Local) ---")

    try:
        config = gmail_automation.load_email_config()
        sender_email = config["sender_email"]
        print(f"Logged in as: {sender_email}")
    except Exception as e:
        print(f"Error loading email configuration: {e}")
        return

    try:
        recipients = PJ_mail_writer.load_distribution_list(DISTRIBUTION_LIST_PATH)
        print(f"Found {len(recipients)} recipients in {DISTRIBUTION_LIST_PATH.name}")
    except Exception as e:
        print(f"Error reading distribution list: {e}")
        return

    success_count = 0
    fail_count = 0

    for i, recipient in enumerate(recipients, 1):
        company = recipient['company']
        email_addr = recipient['email']
        contents = recipient['contents']

        print(f"\n[{i}/{len(recipients)}] Preparing email for {company} ({email_addr})...")

        try:
            email_message = PJ_mail_writer.create_project_email(
                sender_email=sender_email,
                recipient_email=email_addr,
                company_name=company,
                custom_contents=contents,
                html_template_path=str(HTML_TEMPLATE_PATH),
                pdf_path=str(PDF_ATTACHMENT_PATH),
                img1_path=str(IMAGE_1_PATH),
                img2_path=str(IMAGE_2_PATH)
            )

            gmail_automation.send_gmail(email_message)
            success_count += 1
            time.sleep(25)

        except Exception as e:
            print(f"Failed to send to {company}: {e}")
            fail_count += 1

    print("\n--- Batch Processing Complete ---")
    print(f"Success: {success_count}")
    print(f"Failed:  {fail_count}")

if __name__ == "__main__":
    main()