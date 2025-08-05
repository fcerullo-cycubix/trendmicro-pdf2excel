import os
import re
import glob
import pandas as pd
from flask import Flask, request, send_file, render_template, redirect, flash, url_for
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader

app = Flask(__name__)
app.secret_key = 'supersecretkey'
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def extract_server_info_from_pdf(pdf_path, output_excel_path):
    reader = PdfReader(pdf_path)
    pdf_text = "\n".join(page.extract_text() for page in reader.pages if page.extract_text())

    pattern = re.compile(
        r"Group:\s*(?P<Group>.*?)\n.*?\((?P<ComputerName>.*?)\)\n"
        r"Platform:\s*(?P<Platform>.*?)\n"
        r"Status:\s*(?P<Status>.*?)\n"
        r"State:\s*(?P<State>.*?)\n"
        r"Computer Created:\s*(?P<ComputerCreated>.*?)\n"
        r"Last Update Required:\s*(?P<LastUpdateRequired>.*?)\n"
        r"Last Successful Update:\s*(?P<LastSuccessfulUpdate>.*?)\n",
        re.DOTALL
    )

    data = [match.groupdict() for match in pattern.finditer(pdf_text)]
    df = pd.DataFrame(data)
    df.to_excel(output_excel_path, index=False)
    return output_excel_path

@app.route('/', methods=['GET', 'POST'])
def upload_single_pdf():
    download_url = None
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith('.pdf'):
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)

            output_path = os.path.join(OUTPUT_FOLDER, filename.replace('.pdf', '.xlsx'))
            extract_server_info_from_pdf(filepath, output_path)
            download_url = f'/download/{os.path.basename(output_path)}'

    return render_template('layout.html', download_url=download_url)

@app.route('/combine-pdfs', methods=['GET', 'POST'])
def combine_multiple_pdfs():
    download_url = None
    if request.method == 'POST':
        files = request.files.getlist('files')
        dfs = []

        for file in files:
            if file and file.filename.endswith('.pdf'):
                filename = secure_filename(file.filename)
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                file.save(filepath)
                temp_xlsx = filepath.replace('.pdf', '.xlsx')
                extract_server_info_from_pdf(filepath, temp_xlsx)
                df = pd.read_excel(temp_xlsx)
                df['Source PDF'] = filename
                dfs.append(df)

        if dfs:
            combined_df = pd.concat(dfs, ignore_index=True)
            output_path = os.path.join(OUTPUT_FOLDER, 'combined_from_multiple_pdfs.xlsx')
            combined_df.to_excel(output_path, index=False)
            download_url = f'/download/{os.path.basename(output_path)}'

    return render_template('layout.html', download_url=download_url)

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

@app.route('/delete-files', methods=['POST'])
def delete_files():
    try:
        for f in glob.glob(os.path.join(UPLOAD_FOLDER, '*.pdf')):
            os.remove(f)
        for f in glob.glob(os.path.join(OUTPUT_FOLDER, '*.xlsx')):
            os.remove(f)
        flash('All uploaded and generated files have been deleted.', 'success')
    except Exception as e:
        flash(f'Error deleting files: {str(e)}', 'error')
    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5001)