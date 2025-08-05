import os
import re
import pandas as pd
from flask import Flask, request, send_file
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader
import io

app = Flask(__name__)

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
    
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output.getvalue()

@app.route('/', methods=['GET', 'POST'])
def upload_single_pdf():
    if request.method == 'GET':
        return '''<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Trend Micro Report to Excel Converter</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    .progress-bar { transition: width 0.3s ease; }
    .hidden { display: none; }
  </style>
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
    <form method="post" enctype="multipart/form-data" action="/" id="single-form" class="mb-8 space-y-4">
      <label for="file-upload" class="block border-2 border-dashed border-gray-300 p-4 rounded-md text-center hover:border-blue-500 transition cursor-pointer">
        <p class="text-gray-600">Click or drag a Trend Micro PDF report</p>
        <input type="file" name="file" id="file-upload" class="hidden" accept=".pdf">
      </label>
      <div id="single-progress" class="hidden">
        <div class="w-full bg-gray-200 rounded-full h-2.5 mb-2">
          <div id="single-progress-bar" class="bg-blue-600 h-2.5 rounded-full progress-bar" style="width: 0%"></div>
        </div>
        <p id="single-progress-text" class="text-sm text-gray-600 text-center">Uploading...</p>
      </div>
      <button type="submit" id="single-submit" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded w-full">
        Convert Report to Excel
      </button>
    </form>
    <div class="border-t border-gray-300 my-8"></div>
    <h2 class="text-xl font-semibold mb-2">Convert & Combine Multiple Reports</h2>
    <form method="post" enctype="multipart/form-data" action="/combine-pdfs" id="multi-form" class="space-y-4">
      <label for="multi-upload" class="block border-2 border-dashed border-gray-300 p-4 rounded-md text-center hover:border-yellow-500 transition cursor-pointer">
        <p class="text-gray-600">Click or drag multiple Trend Micro PDF reports</p>
        <input type="file" name="files" id="multi-upload" class="hidden" accept=".pdf" multiple>
      </label>
      <div id="multi-progress" class="hidden">
        <div class="w-full bg-gray-200 rounded-full h-2.5 mb-2">
          <div id="multi-progress-bar" class="bg-yellow-500 h-2.5 rounded-full progress-bar" style="width: 0%"></div>
        </div>
        <p id="multi-progress-text" class="text-sm text-gray-600 text-center">Uploading...</p>
      </div>
      <button type="submit" id="multi-submit" class="bg-yellow-500 hover:bg-yellow-600 text-white font-bold py-2 px-4 rounded w-full">
        Combine Reports into Excel
      </button>
    </form>
  </div>
  
  <script>
    function setupFormProgress(formId, progressId, progressBarId, progressTextId, submitId) {
      const form = document.getElementById(formId);
      const progress = document.getElementById(progressId);
      const progressBar = document.getElementById(progressBarId);
      const progressText = document.getElementById(progressTextId);
      const submitBtn = document.getElementById(submitId);
      
      form.addEventListener('submit', function(e) {
        e.preventDefault();
        
        const formData = new FormData(form);
        const xhr = new XMLHttpRequest();
        
        // Show progress bar
        progress.classList.remove('hidden');
        submitBtn.disabled = true;
        submitBtn.textContent = 'Processing...';
        
        // Track upload progress
        xhr.upload.addEventListener('progress', function(e) {
          if (e.lengthComputable) {
            const percentComplete = (e.loaded / e.total) * 100;
            progressBar.style.width = percentComplete + '%';
            progressText.textContent = `Uploading... ${Math.round(percentComplete)}%`;
          }
        });
        
        // Handle completion
        xhr.addEventListener('load', function() {
          if (xhr.status === 200) {
            progressBar.style.width = '100%';
            progressText.textContent = 'Processing complete!';
            
            // Create download link
            const blob = new Blob([xhr.response], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = xhr.getResponseHeader('Content-Disposition')?.split('filename=')[1] || 'report.xlsx';
            a.click();
            window.URL.revokeObjectURL(url);
          } else {
            progressText.textContent = 'Upload failed!';
            progressBar.classList.add('bg-red-500');
          }
          
          // Reset form
          setTimeout(() => {
            progress.classList.add('hidden');
            submitBtn.disabled = false;
            submitBtn.textContent = formId === 'single-form' ? 'Convert Report to Excel' : 'Combine Reports into Excel';
            progressBar.style.width = '0%';
            progressBar.classList.remove('bg-red-500');
          }, 2000);
        });
        
        xhr.responseType = 'blob';
        xhr.open('POST', form.action);
        xhr.send(formData);
      });
    }
    
    // Setup both forms
    setupFormProgress('single-form', 'single-progress', 'single-progress-bar', 'single-progress-text', 'single-submit');
    setupFormProgress('multi-form', 'multi-progress', 'multi-progress-bar', 'multi-progress-text', 'multi-submit');
  </script>
</body>
</html>'''
    
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
    
    return 'Invalid request', 400

@app.route('/combine-pdfs', methods=['POST'])
def combine_multiple_pdfs():
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
    
    return 'No valid files', 400

if __name__ == '__main__':
    app.run()