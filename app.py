# app.py
from flask import Flask, render_template, request, send_from_directory
import os
import openpyxl
import re
import fitz
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def extract_pages_with_mpn(pdf_file, mpn, output_dir, naming_convention, page_numbers=None):
    pdf_document = fitz.open(pdf_file)
    new_pdf = fitz.open()

    if page_numbers is None:
        page_numbers = []

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        page_text = page.get_text("text")

        if re.search(str(mpn), page_text) or (page_numbers and (page_num + 1) in page_numbers):
            new_pdf.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)

    if new_pdf.page_count > 0:
        output_file = os.path.join(output_dir, f"{naming_convention}.pdf")
        new_pdf.save(output_file)

    new_pdf.close()
    pdf_document.close()

def process_excel(excel_path, pdf_path, output_dir):
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active

    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if len(row) < 3:
                continue
            naming_convention, mpn, pages_str = row[:3]

            if isinstance(pages_str, int):
                pages_str = str(pages_str)
            elif not isinstance(pages_str, str):
                pages_str = ''

            page_numbers = list(map(int, pages_str.split(','))) if pages_str else None

            futures.append(executor.submit(
                extract_pages_with_mpn,
                pdf_path, mpn, output_dir, naming_convention, page_numbers
            ))

        for f in futures:
            f.result()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/extract', methods=['POST'])
def extract():
    excel = request.files['excel_file']
    pdf = request.files['pdf_file']

    excel_filename = secure_filename(excel.filename)
    pdf_filename = secure_filename(pdf.filename)

    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)

    excel.save(excel_path)
    pdf.save(pdf_path)

    process_excel(excel_path, pdf_path, app.config['OUTPUT_FOLDER'])

    extracted_files = os.listdir(app.config['OUTPUT_FOLDER'])
    return render_template('success.html', files=extracted_files)

@app.route('/download/<filename>')
def download(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
