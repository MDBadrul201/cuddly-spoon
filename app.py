from flask import Flask, render_template, request, redirect, url_for
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/extract', methods=['POST'])
def extract():
    # Get uploaded files
    excel_file = request.files['excel_file']
    pdf_file = request.files['pdf_file']
    output_dir = request.form['output_dir']

    # Save uploaded files to a temp folder
    excel_path = os.path.join("uploads", excel_file.filename)
    pdf_path = os.path.join("uploads", pdf_file.filename)
    excel_file.save(excel_path)
    pdf_file.save(pdf_path)

    # Call your processing logic here (e.g. extract_pages_worker)
    # You'll need to adapt your current logic into a callable function

    return "PDF extraction complete!"

if __name__ == '__main__':
    app.run(debug=True)
