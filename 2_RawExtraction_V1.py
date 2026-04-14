import os
import email
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader

# -----------------------------
# Processors for each file type
# -----------------------------

def process_eml(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            msg = email.message_from_file(f)
        subject = msg.get("Subject", "")
        sender = msg.get("From", "")
        recipient = msg.get("To", "")
        date = msg.get("Date", "")
        body = ""

        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body += part.get_payload(decode=True).decode(errors="ignore")
        else:
            body = msg.get_payload(decode=True)
            if body:
                body = body.decode(errors="ignore")

        return {
            "file": file_path,
            "type": "eml",
            "subject": subject,
            "from": sender,
            "to": recipient,
            "date": date,
            "content": body
        }
    except Exception as e:
        return {"file": file_path, "type": "eml", "error": str(e)}

def process_docx(file_path):
    try:
        doc = Document(file_path)
        text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
        return {
            "file": file_path,
            "type": "docx",
            "subject": None,
            "from": None,
            "to": None,
            "date": None,
            "content": text
        }
    except Exception as e:
        return {"file": file_path, "type": "docx", "error": str(e)}

def process_pdf(file_path):
    try:
        with open(file_path, 'rb') as f:
            pdf_reader = PdfReader(f)
            text = ""
            for page in pdf_reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        return {
            "file": file_path,
            "type": "pdf",
            "subject": None,
            "from": None,
            "to": None,
            "date": None,
            "content": text
        }
    except Exception as e:
        return {"file": file_path, "type": "pdf", "error": str(e)}

def process_xlsx(file_path):
    try:
        df = pd.read_excel(file_path)
        rows = df.to_dict(orient="records")
        return [
            {
                "file": file_path,
                "type": "xlsx",
                "subject": None,
                "from": None,
                "to": None,
                "date": None,
                "content": str(row)
            }
            for row in rows
        ]
    except Exception as e:
        return [{"file": file_path, "type": "xlsx", "error": str(e)}]

# -----------------------------
# Dispatcher
# -----------------------------

def process_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".eml":
        return process_eml(file_path)
    elif ext == ".docx":
        return process_docx(file_path)
    elif ext == ".pdf":
        return process_pdf(file_path)
    elif ext == ".xlsx":
        return process_xlsx(file_path)
    else:
        return {"file": file_path, "type": "unknown", "error": "unsupported file type"}

# -----------------------------
# Main ETL Loop
# -----------------------------

root_dir = r"D:\Local_Git_repo\python-projects\Project-2-Ediscovery-ChatGPT-Project1\Data\Loose_Files\Sample_Data_20"
ETL_Ready = []

for dirpath, _, filenames in os.walk(root_dir):
    for filename in filenames:
        file_path = os.path.join(dirpath, filename)
        result = process_file(file_path)

        if isinstance(result, list):
            ETL_Ready.extend(result)
        elif isinstance(result, dict):
            ETL_Ready.append(result)
        else:
            ETL_Ready.append({"file": file_path, "type": "unknown", "error": "Invalid result"})

# -----------------------------
# Convert to DataFrame
# -----------------------------

df = pd.DataFrame(ETL_Ready)
print(df.head())

# Save to CSV for easy viewing
output_path = r"D:\Local_Git_repo\python-projects\Project-2-Ediscovery-ChatGPT-Project1\Output\etl_ready_output.csv"
df.to_csv(output_path, index=False)
print("DataFrame saved to:", output_path)