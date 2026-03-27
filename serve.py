#!/usr/bin/env python3
"""
Excel to draw.io Converter - GUI Server
Run: python3 serve.py [port]
Then open http://localhost:[port] in your browser
"""

import os
import sys
import json
import base64
import tempfile
from http.server import HTTPServer, BaseHTTPRequestHandler

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from converter import ExcelReader, DrawioWriter

TEMP_FILE = None


def get_html():
    return '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel to draw.io Converter</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f5f5f5; min-height: 100vh; padding: 20px; }
        .container { max-width: 700px; margin: 0 auto; }
        h1 { text-align: center; color: #333; margin-bottom: 20px; }
        .card { background: white; border-radius: 12px; padding: 25px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .drop { border: 2px dashed #ccc; border-radius: 8px; padding: 50px; text-align: center; cursor: pointer; margin-bottom: 15px; transition: all 0.3s; }
        .drop:hover { border-color: #667eea; background: #f8f9ff; }
        .drop.drag { border-color: #667eea; background: #e8f5ff; }
        .drop-icon { font-size: 48px; margin-bottom: 10px; }
        .drop-text { color: #666; }
        .drop-hint { font-size: 12px; color: #999; margin-top: 5px; }
        input[type=file] { display: none; }
        .info { background: #e3f2fd; padding: 15px; border-radius: 8px; margin-bottom: 15px; display: none; }
        .info.show { display: block; }
        .info-name { font-weight: bold; word-break: break-all; }
        .info-size { font-size: 12px; color: #666; margin-top: 5px; }
        .clear-btn { float: right; background: #f44336; color: white; border: none; padding: 5px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; }
        .sheets { display: none; margin-bottom: 15px; }
        .sheets.show { display: block; }
        .sheets-title { font-weight: bold; margin-bottom: 10px; color: #333; }
        .sheet-list { display: flex; flex-wrap: wrap; gap: 8px; max-height: 250px; overflow-y: auto; padding: 10px; background: #f8f8f8; border-radius: 6px; }
        .sheet { display: flex; align-items: center; background: white; padding: 8px 14px; border-radius: 6px; border: 1px solid #eee; cursor: pointer; transition: all 0.2s; font-size: 13px; }
        .sheet:hover { background: #e8f5e9; border-color: #667eea; }
        .sheet input { margin-right: 8px; }
        .sheets-actions { margin-top: 10px; font-size: 12px; color: #666; }
        .actions { text-align: center; display: none; margin-bottom: 15px; }
        .actions.show { display: block; }
        .btn { padding: 12px 30px; background: #667eea; color: white; border: none; border-radius: 8px; cursor: pointer; font-size: 16px; font-weight: bold; transition: all 0.3s; }
        .btn:hover { background: #5568d3; transform: translateY(-1px); }
        .btn:disabled { background: #ccc; cursor: not-allowed; transform: none; }
        .progress { text-align: center; padding: 30px; display: none; }
        .progress.show { display: block; }
        .spinner { width: 40px; height: 40px; border: 4px solid #f3f3f3; border-top: 4px solid #667eea; border-radius: 50%; animation: spin 1s linear infinite; margin: 0 auto 15px; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .result { text-align: center; padding: 20px; background: #e8f5e9; border-radius: 8px; margin-top: 15px; display: none; }
        .result.show { display: block; }
        .result-title { font-size: 18px; margin-bottom: 10px; }
        .result-size { font-size: 14px; color: #666; }
        .download-btn { display: inline-block; margin-top: 15px; padding: 12px 30px; background: #43a047; color: white; text-decoration: none; border-radius: 6px; font-weight: bold; }
        .download-btn:hover { background: #388e3c; }
        .error { padding: 15px; background: #ffebee; color: #c62828; border-radius: 8px; margin-top: 15px; display: none; }
        .error.show { display: block; }
        .footer { text-align: center; margin-top: 20px; font-size: 12px; color: #999; }
    </style>
</head>
<body>
<div class="container">
    <h1>Excel to draw.io Converter</h1>
    <div class="card">
        <div class="drop" id="drop">
            <div class="drop-icon">&#128193;</div>
            <div class="drop-text">Drop Excel file or click to browse</div>
            <div class="drop-hint">Supports .xlsx, .xls</div>
        </div>
        <input type="file" id="fileInput" accept=".xlsx,.xls">

        <div class="info" id="info">
            <button class="clear-btn" onclick="clearFile()">Clear</button>
            <div class="info-name" id="fileName"></div>
            <div class="info-size" id="fileSize"></div>
        </div>

        <div class="sheets" id="sheets">
            <div class="sheets-title">Select sheets to convert:</div>
            <div class="sheet-list" id="sheetList"></div>
            <div class="sheets-actions">
                <label><input type="checkbox" id="selectAll" checked onchange="toggleAll()"> Select All / Deselect All</label>
            </div>
        </div>

        <div class="actions" id="actions">
            <button class="btn" id="convertBtn" onclick="convert()">Convert to draw.io</button>
        </div>

        <div class="progress" id="progress">
            <div class="spinner"></div>
            <div>Converting...</div>
        </div>

        <div class="result" id="result">
            <div class="result-title">Conversion Complete!</div>
            <div class="result-size" id="resultSize"></div>
            <a href="#" id="downloadBtn" class="download-btn" download>Download draw.io</a>
        </div>

        <div class="error" id="error"></div>
    </div>
    <div class="footer">Excel to draw.io Converter - Works on Windows, Mac, Linux</div>
</div>

<script>
let currentFile = null;
let currentFileName = '';

const drop = document.getElementById('drop');
const fileInput = document.getElementById('fileInput');

drop.onclick = () => fileInput.click();

drop.ondragover = e => { e.preventDefault(); drop.classList.add('drag'); };
drop.ondragleave = () => drop.classList.remove('drag');
drop.ondrop = e => {
    e.preventDefault();
    drop.classList.remove('drag');
    if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
};

fileInput.onchange = e => { if (e.target.files[0]) loadFile(e.target.files[0]); };

function loadFile(file) {
    if (!file.name.match(/\\.xlsx?$/i)) {
        showError('Please select an Excel file (.xlsx or .xls)');
        return;
    }
    currentFile = file;
    currentFileName = file.name;
    document.getElementById('fileName').textContent = file.name;
    document.getElementById('fileSize').textContent = formatSize(file.size);
    document.getElementById('info').classList.add('show');
    drop.style.display = 'none';
    loadSheets(file);
}

async function loadSheets(file) {
    const reader = new FileReader();
    reader.onload = async function(e) {
        const base64 = e.target.result.split(',')[1];
        try {
            const response = await fetch('/sheets', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({data: base64})
            });
            const data = await response.json();
            if (data.success) showSheets(data.sheets);
            else showError(data.error);
        } catch(err) { showError(err.message); }
    };
    reader.readAsDataURL(file);
}

function showSheets(sheets) {
    const list = document.getElementById('sheetList');
    list.innerHTML = '';
    sheets.forEach(s => {
        const label = document.createElement('label');
        label.className = 'sheet';
        label.innerHTML = '<input type="checkbox" name="sheets" value="' + s + '" checked> ' + s;
        list.appendChild(label);
    });
    document.getElementById('sheets').classList.add('show');
    document.getElementById('actions').classList.add('show');
}

function toggleAll() {
    const checked = document.getElementById('selectAll').checked;
    document.querySelectorAll('[name=sheets]').forEach(c => c.checked = checked);
}

function clearFile() {
    currentFile = null;
    currentFileName = '';
    fileInput.value = '';
    ['info','sheets','actions','result','error'].forEach(id => document.getElementById(id).classList.remove('show'));
    drop.style.display = 'block';
}

async function convert() {
    if (!currentFile) return;
    const sheets = [...document.querySelectorAll('[name=sheets]:checked')].map(c => c.value);
    if (!sheets.length) { showError('Please select at least one sheet'); return; }

    document.getElementById('progress').classList.add('show');
    document.getElementById('convertBtn').disabled = true;

    const reader = new FileReader();
    reader.onload = async function(e) {
        const base64 = e.target.result.split(',')[1];
        try {
            const response = await fetch('/convert?sheets=' + encodeURIComponent(sheets.join(',')), {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({data: base64})
            });
            const data = await response.json();
            document.getElementById('progress').classList.remove('show');
            document.getElementById('convertBtn').disabled = false;

            if (data.success) {
                document.getElementById('resultSize').textContent = 'Size: ' + formatSize(data.size);
                const downloadName = currentFileName.replace(/\\.[^.]+$/, '') + '.drawio';
                document.getElementById('downloadBtn').href = 'data:application/octet-stream;base64,' + data.file;
                document.getElementById('downloadBtn').download = downloadName;
                document.getElementById('result').classList.add('show');
            } else {
                showError(data.error);
            }
        } catch(err) {
            document.getElementById('progress').classList.remove('show');
            document.getElementById('convertBtn').disabled = false;
            showError(err.message);
        }
    };
    reader.readAsDataURL(currentFile);
}

function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
}

function showError(msg) {
    document.getElementById('error').textContent = 'Error: ' + msg;
    document.getElementById('error').classList.add('show');
}
</script>
</body>
</html>
'''


class Handler(BaseHTTPRequestHandler):
    global TEMP_FILE

    def do_GET(self):
        if self.path == '/' or self.path == '/index.html':
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(get_html().encode('utf-8'))
        else:
            self.send_error(404)

    def do_POST(self):
        global TEMP_FILE

        if self.path == '/sheets':
            try:
                length = int(self.headers.get('Content-Length', 0))
                body = json.loads(self.rfile.read(length).decode())
                data = base64.b64decode(body.get('data', ''))

                # Save temp file
                if TEMP_FILE and os.path.exists(TEMP_FILE):
                    os.unlink(TEMP_FILE)
                TEMP_FILE = os.path.join(tempfile.gettempdir(), 'temp_upload.xlsx')
                with open(TEMP_FILE, 'wb') as f:
                    f.write(data)

                # Get sheets
                reader = ExcelReader(TEMP_FILE)
                sheets = list(reader.read_all().keys())

                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'success': True, 'sheets': sheets}).encode())
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

        elif self.path.startswith('/convert'):
            try:
                # Parse sheets from query
                sheets_param = ''
                if '?' in self.path:
                    query = self.path.split('?')[1]
                    for p in query.split('&'):
                        if p.startswith('sheets='):
                            sheets_param = p.split('=')[1]
                selected_sheets = [s for s in urllib.parse.unquote(sheets_param).split(',') if s]

                # Read body
                length = int(self.headers.get('Content-Length', 0))
                body = json.loads(self.rfile.read(length).decode())
                data = base64.b64decode(body.get('data', ''))

                # Save temp file if not already saved
                if not TEMP_FILE or not os.path.exists(TEMP_FILE):
                    TEMP_FILE = os.path.join(tempfile.gettempdir(), 'temp_upload.xlsx')
                    with open(TEMP_FILE, 'wb') as f:
                        f.write(data)

                # Convert
                reader = ExcelReader(TEMP_FILE, sheet_names=selected_sheets if selected_sheets else None)
                sheets_data = reader.read_all()

                output_path = os.path.join(tempfile.gettempdir(), 'output.drawio')
                writer = DrawioWriter(sheets_data)
                writer.write(output_path)

                # Read output
                with open(output_path, 'rb') as f:
                    output_data = f.read()
                b64 = base64.b64encode(output_data).decode()

                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'success': True, 'file': b64, 'size': len(output_data)}).encode())
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

        else:
            self.send_error(404)

    def log_message(self, fmt, *args):
        pass  # Suppress logs


def run(port=8765):
    server = HTTPServer(('localhost', port), Handler)
    print(f'========================================')
    print(f'Excel to draw.io Converter')
    print(f'========================================')
    print(f'Server running at: http://localhost:{port}')
    print(f'')
    print(f'Open the URL above in your browser')
    print(f'Press Ctrl+C to stop the server')
    print(f'========================================')
    server.serve_forever()


if __name__ == '__main__':
    import urllib.parse
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
    run(port)
