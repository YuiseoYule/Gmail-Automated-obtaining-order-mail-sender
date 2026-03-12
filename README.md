# 📧 BICS Email Automation System

An automated, template-based email dispatch system designed to streamline **B2B client outreach** and project proposals for the Seoul National University Strategic Consulting Group (BICS).

---

## 🚀 Overview
This project automates the process of sending personalized business proposals to multiple companies. It integrates Excel-based distribution lists with dynamic HTML templates and secure SMTP protocols to ensure professional and efficient communication.

### **Key Significance**
* **Scalable Client Contact**: Replaces manual drafting with a technical solution that manages bulk outreach while maintaining a personalized touch for each firm.
* **Professional Delivery**: Supports high-resolution inline images and PDF attachments (Company Introduction) for a polished brand image.

---

## 🛠️ Architecture & Features

### **1. Modular Design**
- `gmail_automation.py`: Core SMTP logic, `.env` configuration, and secure SSL/TLS connection handling.
- `PJ_mail_writer.py`: Logic for parsing Excel files, formatting HTML templates with `{company}` placeholders, and constructing `MIMEMultipart` messages.
- `PJ_trigger.py`: The execution script that orchestrates the workflow and manages rate-limiting.

### **2. Technical Specifications**
* **Data Handling**: Uses `pandas` to read recipient data (Company, Email, Custom Contents).
* **Security**: Sensitive credentials (APP_PASSWORD) are managed via `.env` files using `python-dotenv`.
* **Stability**: Implements **Rate Limiting** (25s intervals) to prevent Gmail API blocks and spam filtering.
* **Environment**: Optimized for **Google Colab** with automated Drive mounting and path configuration. 

---

## 📂 Project Structure
```text
Mail_automation_program/
├── Mail_sender/
│   ├── gmail_automation.py
│   └── googleid.env           # Excluded from Git (Sensitive)
├── Mail_form/
│   ├── PJ_template.html       # HTML Email Body
│   ├── bics_intro_1.png       # Inline Image 1
│   ├── bics_intro_2.png       # Inline Image 2
│   └── BICS_Intro.pdf         # Proposal Attachment
├── PJ_mail/
│    ├── PJ_mail_writer.py      # Email construction logic
│    └── PJ_trigger.py          # Execution script (Main)
├── Colab_notebook_sender.ipynb
│
└── Local_python_sender
```

---

## 📧 Email Service Configuration

This system requires an SMTP configuration to send notifications. 

### 1. Environment Setup
To protect sensitive credentials, the `.env` file is excluded from this repository. 

1.  **Copy the template file**:
    ```bash
    cp .env.example .env
    ```
2.  **Edit the `.env` file** and provide your credentials:
    - `SMTP_SERVER`: The SMTP host (e.g., smtp.gmail.com).
    - `SMTP_PORT`: Usually 587 for TLS.
    - `SENDER_EMAIL`: The email address sending the alerts.
    - `APP_PASSWORD`: Your 16-character Google App Password.
    - `SENDER_NAME`: The display name for the sender.
