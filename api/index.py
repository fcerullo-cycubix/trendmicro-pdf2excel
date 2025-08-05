import os
import re
import tempfile
import pandas as pd
from flask import Flask, request, send_file, render_template, redirect, flash, url_for, jsonify
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader
import io

app = Flask(__name__)
app.secret_key = 'supersecretkey'

def extract_server_info_from_pdf(pdf_content, filename):
    reader = PdfReader(io.BytesIO(pdf_content))
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
    
    # Create Excel file in memory
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output.getvalue()

@app.route('/', methods=['GET', 'POST'])
def upload_single_pdf():
    if request.method == 'GET':
        return render_template('layout.html')
    
    if request.method == 'POST':
        file = request.files.get('file')
        if file and file.filename.endswith('.pdf'):
            pdf_content = file.read()
            excel_data = extract_server_info_from_pdf(pdf_content, file.filename)
            
            filename = secure_filename(file.filename).replace('.pdf', '.xlsx')
            
            return send_file(
                io.BytesIO(excel_data),
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    
    return jsonify({'error': 'Invalid request'}), 400

@app.route('/combine-pdfs', methods=['GET', 'POST'])
def combine_multiple_pdfs():
    if request.method == 'GET':
        return render_template('layout.html')
    
    if request.method == 'POST':
        files = request.files.getlist('files')
        dfs = []

        for file in files:
            if file and file.filename.endswith('.pdf'):
                pdf_content = file.read()
                excel_data = extract_server_info_from_pdf(pdf_content, file.filename)
                df = pd.read_excel(io.BytesIO(excel_data))
                df['Source PDF'] = secure_filename(file.filename)
                dfs.append(df)

        if dfs:
            combined_df = pd.concat(dfs, ignore_index=True)
            output = io.BytesIO()
            combined_df.to_excel(output, index=False)
            output.seek(0)
            
            return send_file(
                output,
                as_attachment=True,
                download_name='combined_from_multiple_pdfs.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    
    return jsonify({'error': 'Invalid request'}), 400

@app.route('/static/<path:filename>')
def static_files(filename):
    static_path = os.path.join(os.path.dirname(__file__), '..', 'static', filename)
    if os.path.exists(static_path):
        return send_file(static_path)
    return '', 404

# Vercel serverless function handler
def handler(request):
    return app(request.environ, lambda status, headers: None)