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
        html_content = '''<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Trend Micro Report to Excel Converter</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-100 text-gray-800">
  <div class="max-w-2xl mx-auto p-6 bg-white shadow-lg rounded-xl mt-10">
    <div class="flex justify-center mb-6">
      <div class="h-12 w-32 bg-red-600 flex items-center justify-center rounded text-white font-bold">TREND MICRO</div>
    </div>
    <h1 class="text-3xl font-bold mb-2 text-center">Trend Micro Deep Security Report Converter</h1>
    <p class="text-gray-700 text-center mb-6 max-w-xl mx-auto">
      This tool extracts key server update details from <strong>Trend Micro Deep Security PDF reports</strong>
      and converts them into Excel spreadsheets.
    </p>
    <h2 class="text-xl font-semibold mb-2">Convert Single Report</h2>
    <form method="post" enctype="multipart/form-data" action="/" class="mb-8 space-y-4">
      <label for="file-upload" class="block border-2 border-dashed border-gray-300 p-4 rounded-md text-center hover:border-blue-500 transition cursor-pointer">
        <p class="text-gray-600">Click or drag a Trend Micro PDF report</p>
        <input type="file" name="file" id="file-upload" class="hidden" accept=".pdf">
      </label>
      <button type="submit" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded w-full">
        Convert Report to Excel
      </button>
    </form>
    <div class="border-t border-gray-300 my-8"></div>
    <h2 class="text-xl font-semibold mb-2">Convert & Combine Multiple Reports</h2>
    <form method="post" enctype="multipart/form-data" action="/combine-pdfs" class="space-y-4">
      <label for="multi-upload" class="block border-2 border-dashed border-gray-300 p-4 rounded-md text-center hover:border-yellow-500 transition cursor-pointer">
        <p class="text-gray-600">Click or drag multiple Trend Micro PDF reports</p>
        <input type="file" name="files" id="multi-upload" class="hidden" accept=".pdf" multiple>
      </label>
      <button type="submit" class="bg-yellow-500 hover:bg-yellow-600 text-white font-bold py-2 px-4 rounded w-full">
        Combine Reports into Excel
      </button>
    </form>
  </div>
</body>
</html>'''
        return html_content
    
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
        return redirect('/')
    
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

# Vercel expects 'app' to be available at module level
# The Flask app is already defined above as 'app'