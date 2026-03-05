import pandas as pd
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication


def load_distribution_list(filepath: str) -> list:
    """
    Reads the distribution list from an Excel file.
    Returns a list of dictionaries containing 'company', 'email', and 'contents'.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Distribution list not found at: {filepath}")

    # Read Excel (assumes columns 0=Company, 1=Email, 2=Contents)
    df = pd.read_excel(filepath)

    data = []
    for _, row in df.iterrows():
        item = {
            "company": str(row.iloc[0]),
            "email": str(row.iloc[1]),
            "contents": str(row.iloc[2])
        }
        data.append(item)

    return data


def load_and_format_html_template(template_path: str, company_name: str, custom_contents: str) -> str:
    """
    Reads the external HTML file and replaces {company} and {contents} placeholders.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"HTML template not found at: {template_path}")

    with open(template_path, 'r', encoding='utf-8') as f:
        html_content = f.read()

    formatted_html = html_content.replace("{company}", company_name)
    formatted_html = formatted_html.replace("{contents}", custom_contents)

    return formatted_html


def attach_inline_image(msg: MIMEMultipart, filepath: str, cid: str):
    """
    Helper to attach an image with a Content-ID for inline display.
    """
    if not os.path.exists(filepath):
        print(f"Warning: Image file not found: {filepath}")
        return

    with open(filepath, 'rb') as f:
        img = MIMEImage(f.read())

    img.add_header('Content-ID', f'<{cid}>')
    img.add_header('Content-Disposition', 'inline', filename=os.path.basename(filepath))
    msg.attach(img)


def create_project_email(
        sender_email: str,
        recipient_email: str,
        company_name: str,
        custom_contents: str,
        html_template_path: str,
        pdf_path: str,
        img1_path: str,
        img2_path: str
) -> MIMEMultipart:
    """
    Creates the full email object using data from Excel and an external HTML template.
    Structure: Multipart/Mixed (Root + PDF) -> Multipart/Related (HTML + Images).
    """
    # 1. Create the Root Container (multipart/mixed)
    msg_root = MIMEMultipart('mixed')
    msg_root['Subject'] = f"[서울대학교 산학협력] {company_name}-BICS 전략 컨설팅 프로젝트 제안"
    msg_root['From'] = sender_email
    msg_root['To'] = recipient_email

    # 2. Create the Related Container (multipart/related) for HTML + Images
    msg_related = MIMEMultipart('related')
    msg_root.attach(msg_related)

    # 3. Load and format HTML
    try:
        html_body = load_and_format_html_template(html_template_path, company_name, custom_contents)
        msg_html = MIMEText(html_body, 'html')
        msg_related.attach(msg_html)
    except Exception as e:
        print(f"Error loading HTML template: {e}")
        msg_related.attach(MIMEText("Error loading email template.", "plain"))

    # 4. Attach Inline Images (Referenced by cid in your HTML file)
    attach_inline_image(msg_related, img1_path, 'bics_img_1')
    attach_inline_image(msg_related, img2_path, 'bics_img_2')

    # 5. Attach PDF File to the Root
    if os.path.exists(pdf_path):
        with open(pdf_path, "rb") as f:
            pdf_attachment = MIMEApplication(f.read(), _subtype="pdf")

        pdf_attachment.add_header(
            'Content-Disposition',
            'attachment',
            filename=os.path.basename(pdf_path)
        )
        msg_root.attach(pdf_attachment)
    else:
        print(f"Warning: PDF file not found at {pdf_path}")

    return msg_root