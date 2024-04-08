import os
import pdfplumber
import re
import pandas as pd
from docx import Document
from flask import Flask, render_template, request, send_file

app = Flask(__name__)
UPLOAD_FOLDER = 'Sample2'  # Change the directory name here
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def extract_info_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text()
    return text

def extract_info_from_docx(docx_path):
    doc = Document(docx_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text
    return text

def extract_info_from_cv(cv_folder):
    emails = []
    phones = []
    texts = []

    for filename in os.listdir(cv_folder):
        if filename.endswith('.pdf'):
            text = extract_info_from_pdf(os.path.join(cv_folder, filename))
        elif filename.endswith('.docx'):
            text = extract_info_from_docx(os.path.join(cv_folder, filename))
        else:
            continue

        # Regular expressions for finding email IDs and phone numbers
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        phone_pattern = r'\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b'

        emails.extend(re.findall(email_pattern, text))
        phones.extend(re.findall(phone_pattern, text))
        texts.append(text)

    return emails, phones, texts

def save_to_excel(emails, phones, texts, output_file):
    df = pd.DataFrame({'Email': emails, 'Phone': phones, 'Text': texts})
    df.to_excel(output_file, index=False)

@app.route('/', methods=['GET', 'POST'])
def upload_cv():
    if request.method == 'POST':
        cv_files = request.files.getlist('cv_files')
        for cv_file in cv_files:
            cv_file.save(os.path.join(app.config['UPLOAD_FOLDER'], cv_file.filename))
        cv_folder = app.config['UPLOAD_FOLDER']
        emails, phones, texts = extract_info_from_cv(cv_folder)
        output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'cv_info.xlsx')
        save_to_excel(emails, phones, texts, output_file)
        return send_file(output_file, as_attachment=True)
    return render_template('upload_cv.html')

if __name__ == '__main__':
    app.run(debug=True)
