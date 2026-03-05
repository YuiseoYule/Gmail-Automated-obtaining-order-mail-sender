import sys
import os
import time

# Get the current script's directory, then go one level up (parent)
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.append(parent_dir)

# Now you can import from the sibling folder
from Mail_sender import gmail_automation
import PJ_mail_writer

# The subsequent import lines in the original request (import PJ_mail_writer)
# are redundant as the module is already loaded above.



# --- Configuration: File Paths ---
# Update these filenames if they differ on your system
DISTRIBUTION_LIST_PATH = "distribution_list.xlsx"
HTML_TEMPLATE_PATH = "PJ_template.html"
PDF_ATTACHMENT_PATH = "서울대학교 전략컨설팅학회 BICS 소개서.pdf"  # The PDF file name
IMAGE_1_PATH = "bics_intro_1.png"  # Left image
IMAGE_2_PATH = "bics_intro_2.png"  # Right image


def main():
    print("--- Starting BICS Email Automation ---")

    # 1. Load Email Configuration (to get the Sender Email Address)
    try:
        # We access the internal function just to get the sender email for the "From" header
        config = gmail_automation.load_email_config()
        sender_email = config["sender_email"]
        print(f"Logged in as: {sender_email}")
    except Exception as e:
        print(f"Error loading email configuration: {e}")
        return

    # 2. Load Distribution List
    try:
        recipients = PJ_mail_writer.load_distribution_list(DISTRIBUTION_LIST_PATH)
        print(f"Found {len(recipients)} recipients in {DISTRIBUTION_LIST_PATH}")
    except Exception as e:
        print(f"Error reading distribution list: {e}")
        return

    # 3. Iterate and Send Emails
    success_count = 0
    fail_count = 0

    for i, recipient in enumerate(recipients, 1):
        company = recipient['company']
        email_addr = recipient['email']
        contents = recipient['contents']

        print(f"\n[{i}/{len(recipients)}] Preparing email for {company} ({email_addr})...")

        try:
            # A. Create the full email object (HTML + Images + PDF)
            email_message = PJ_mail_writer.create_project_email(
                sender_email=sender_email,
                recipient_email=email_addr,
                company_name=company,
                custom_contents=contents,
                html_template_path=HTML_TEMPLATE_PATH,
                pdf_path=PDF_ATTACHMENT_PATH,
                img1_path=IMAGE_1_PATH,
                img2_path=IMAGE_2_PATH
            )

            # B. Send the email using the automation module
            gmail_automation.send_gmail(email_message)

            success_count += 1

            # C. Rate Limiting (Important to avoid spam filters)
            time.sleep(2)

        except Exception as e:
            print(f"Failed to send to {company}: {e}")
            fail_count += 1

    # 4. Final Summary
    print("\n--- Batch Processing Complete ---")
    print(f"Success: {success_count}")
    print(f"Failed:  {fail_count}")


if __name__ == "__main__":
    main()