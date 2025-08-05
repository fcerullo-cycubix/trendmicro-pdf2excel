from flask import Flask

app = Flask(__name__)

@app.route('/')
def hello():
    return '''<!doctype html>
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
  </div>
</body>
</html>'''

if __name__ == '__main__':
    app.run()