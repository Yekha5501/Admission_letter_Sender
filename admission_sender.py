import pandas as pd
from docx import Document
from docx2pdf import convert
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time

# ---------- CONFIG ----------
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GMAIL_USER = "yekhapmandindi@gmail.com"  # FIXED: Removed @mail
GMAIL_PASS = "cjpgbpainjcersvu"          # Your App Password

# ---------- PATHS ----------
TEMPLATE_PATH = "admission_template.docx"
OUTPUT_DIR = "admission_letters"

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# ---------- FUNCTION: Generate PDF from Template ----------
def generate_admission_letter(student_data):
    name = student_data['Name']
    
    doc = Document(TEMPLATE_PATH)
    
    replacements = {
        '{Name}': student_data['Name'],
        '{Program}': student_data['Program'],
        '{Faculty}': student_data['Faculty'],
        '{Duration}': student_data['Duration'],
        '{StartDate}': student_data['StartDate'],
        '{Date}': time.strftime("%B %d, %Y")
    }
    
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, value)
    
    temp_docx = os.path.join(OUTPUT_DIR, f"{name}_admission.docx")
    doc.save(temp_docx)
    print(f"   📄 Word doc created: {name}_admission.docx")
    
    pdf_path = os.path.join(OUTPUT_DIR, f"{name}_admission.pdf")
    
    try:
        convert(temp_docx, pdf_path)
        print(f"   ✅ PDF generated: {name}_admission.pdf")
        return pdf_path
    except Exception as e:
        print(f"   ❌ PDF conversion failed: {e}")
        return temp_docx

# ---------- FUNCTION: Send Email ----------
def send_admission_email(student_data, pdf_path):
    name = student_data['Name']
    email = student_data['Email']
    
    msg = MIMEMultipart()
    msg['From'] = GMAIL_USER
    msg['To'] = email
    msg['Subject'] = f"Admission Letter - {name}"
    
    body = f"""
    Dear {name},
    
    Please find attached your admission letter from our institution.
    We are delighted to welcome you to our academic community.
    
    For any queries, please contact the Registry Office.
    
    Regards,
    Registry Office
    """
    msg.attach(MIMEText(body, 'plain'))
    
    try:
        with open(pdf_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(pdf_path)}")
            msg.attach(part)
        print(f"   📎 Attached: {os.path.basename(pdf_path)}")
    except Exception as e:
        print(f"   ❌ Attachment error: {e}")
        return False
    
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(GMAIL_USER, GMAIL_PASS)
        server.send_message(msg)
        server.quit()
        print(f"   ✅ EMAIL SENT to {email}")
        return True
    except Exception as e:
        print(f"   ❌ Email failed: {e}")
        return False

# ---------- MAIN PROCESS ----------
def main():
    df = pd.read_excel("students.xlsx")
    
    # DEBUG: Show what columns exist
    print(f"📊 Loaded {len(df)} students")
    print(f"📋 Columns found: {df.columns.tolist()}")
    print(f"📋 First row sample: {df.iloc[0].to_dict()}")
    print("\n")
    
    print("🚀 Starting admission letter generation...\n")
    
    for index, row in df.iterrows():
        print(f"📝 Processing: {row['Name']}")
        
        pdf_path = generate_admission_letter(row)
        send_admission_email(row, pdf_path)
        
        print("   " + "-"*40 + "\n")
        time.sleep(2)
    
    print("✅ All admission letters have been generated and sent!")

if __name__ == "__main__":
    main()