#!/usr/bin/env python3
"""
Excel to draw.io Converter - GUI Server
Run: python3 serve.py [port]
Then open http://localhost:[port] in your browser

Phase 1: Progress bar via SSE, options panel, SVG output
Phase 2: Conversion preview (shape/connector counts), settings persistence
Phase 3: Dark mode toggle, conversion history (last 10)
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
ORIGINAL_FILENAME = None


def get_html():
    return '''


<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel to draw.io Converter</title>
    <style>
        :root {
            --bg:#f5f5f5;--card-bg:#ffffff;--text:#333;--text-muted:#666;--accent:#667eea;--accent-hover:#5568d3;
            --border:#e0e0e0;--drop-border:#ccc;--info-bg:#e3f2fd;--info-text:#0d47a1;
            --success-bg:#e8f5e9;--success-text:#2e7d32;--error-bg:#ffebee;--error-text:#c62828;
            --options-bg:#f8f9ff;--options-border:#e0e4ff;--preview-bg:#fffde7;--preview-border:#fff176;
            --overlay-bg:rgba(0,0,0,0.5);--shadow:0 2px 10px rgba(0,0,0,0.1);
        }
        [data-theme="dark"] {
            --bg:#1a1a2e;--card-bg:#16213e;--text:#e0e0e0;--text-muted:#a0a0a0;--accent:#667eea;--accent-hover:#7c8ff5;
            --border:#2a3a5e;--drop-border:#3a4a6e;--info-bg:#1a2a4a;--info-text:#90caf9;
            --success-bg:#1b3a2b;--success-text:#81c784;--error-bg:#3a1a1a;--error-text:#ef9a9a;
            --options-bg:#1a2540;--options-border:#2a3a6e;--preview-bg:#2a2a1a;--preview-border:#8a7a20;
            --overlay-bg:rgba(0,0,0,0.7);--shadow:0 2px 10px rgba(0,0,0,0.4);
        }
        *{box-sizing:border-box;margin:0;padding:0}
        body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;padding:20px;transition:background .3s,color .3s}
        .container{max-width:700px;margin:0 auto;position:relative}
        .header{display:flex;justify-content:space-between;align-items:center;margin-bottom:20px}
        h1{color:var(--text);margin:0;font-size:22px}
        .header-right{display:flex;gap:10px;align-items:center}
        .theme-btn{background:var(--card-bg);border:1px solid var(--border);color:var(--text);width:38px;height:38px;border-radius:8px;cursor:pointer;font-size:18px;display:flex;align-items:center;justify-content:center;transition:all .2s;box-shadow:var(--shadow)}
        .theme-btn:hover{background:var(--accent);color:#fff;border-color:var(--accent)}
        .history-btn{background:var(--card-bg);border:1px solid var(--border);color:var(--text);padding:6px 14px;border-radius:8px;cursor:pointer;font-size:13px;display:none;align-items:center;gap:5px;transition:all .2s;box-shadow:var(--shadow)}
        .history-btn:hover{background:var(--accent);color:#fff;border-color:var(--accent)}
        .history-btn.show{display:flex}
        .history-btn .badge{background:var(--accent);color:#fff;border-radius:10px;padding:1px 6px;font-size:11px;font-weight:bold}
        .card{background:var(--card-bg);border-radius:12px;padding:25px;box-shadow:var(--shadow);transition:background .3s,box-shadow .3s}
        .drop{border:2px dashed var(--drop-border);border-radius:8px;padding:50px;text-align:center;cursor:pointer;margin-bottom:15px;transition:all .3s}
        .drop:hover{border-color:var(--accent);background:var(--info-bg)}
        .drop.drag{border-color:var(--accent);background:var(--info-bg)}
        .drop-icon{font-size:48px;margin-bottom:10px}.drop-text{color:var(--text-muted)}.drop-hint{font-size:12px;color:var(--text-muted);margin-top:5px;opacity:.7}
        input[type=file]{display:none}
        .info{background:var(--info-bg);color:var(--info-text);padding:15px;border-radius:8px;margin-bottom:15px;display:none}
        .info.show{display:block}.info-name{font-weight:bold;word-break:break-all;font-size:14px}
        .info-size{font-size:12px;margin-top:5px;opacity:.8}
        .clear-btn{float:right;background:var(--error-text);color:#fff;border:none;padding:5px 12px;border-radius:4px;cursor:pointer;font-size:12px}
        .sheets{display:none;margin-bottom:15px}.sheets.show{display:block}
        .sheets-title{font-weight:bold;margin-bottom:10px;color:var(--text)}
        .sheet-list{display:flex;flex-wrap:wrap;gap:8px;max-height:250px;overflow-y:auto;padding:10px;background:var(--bg);border-radius:6px}
        .sheet{display:flex;align-items:center;background:var(--card-bg);padding:8px 14px;border-radius:6px;border:1px solid var(--border);cursor:pointer;transition:all .2s;font-size:13px;color:var(--text)}
        .sheet:hover{background:var(--success-bg);border-color:var(--accent)}
        .sheet input{margin-right:8px;accent-color:var(--accent)}
        .sheets-actions{margin-top:10px;font-size:12px;color:var(--text-muted)}
        .preview{background:var(--preview-bg);border:1px solid var(--preview-border);border-radius:8px;padding:15px;margin-bottom:15px;display:none}
        .preview.show{display:block}.preview-title{font-weight:bold;margin-bottom:10px;color:var(--text);font-size:14px}
        .preview-row{display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid var(--border);font-size:13px;color:var(--text-muted)}
        .preview-row:last-child{border-bottom:none}
        .preview-row .sheet-name{font-weight:600;color:var(--text)}
        .preview-row .counts{display:flex;gap:12px}
        .preview-row .counts span{background:var(--card-bg);padding:2px 8px;border-radius:10px;font-size:12px}
        .preview-total{margin-top:8px;padding-top:8px;border-top:2px solid var(--preview-border);display:flex;justify-content:space-between;font-weight:bold;font-size:13px;color:var(--text)}
        .preview-total .est-size{color:var(--accent);font-size:12px}
        .options{background:var(--options-bg);border:1px solid var(--options-border);border-radius:8px;padding:15px;margin-bottom:15px;display:none}
        .options.show{display:block}.options-title{font-weight:bold;margin-bottom:12px;color:var(--text);font-size:14px}
        .options-row{display:flex;flex-wrap:wrap;gap:20px;align-items:center}
        .option-group{display:flex;align-items:center;gap:8px}
        .option-group label{font-size:13px;color:var(--text-muted)}
        .option-group select{padding:6px 10px;border:1px solid var(--border);border-radius:5px;font-size:13px;background:var(--card-bg);color:var(--text);cursor:pointer}
        .option-group select:focus{outline:none;border-color:var(--accent)}
        .option-check{display:flex;align-items:center;gap:6px;cursor:pointer}
        .option-check input[type=checkbox]{width:16px;height:16px;cursor:pointer;accent-color:var(--accent)}
        .option-check span{font-size:13px;color:var(--text-muted)}
        .actions{text-align:center;display:none;margin-bottom:15px}.actions.show{display:block}
        .btn{padding:12px 30px;background:var(--accent);color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:16px;font-weight:bold;transition:all .3s}
        .btn:hover{background:var(--accent-hover);transform:translateY(-1px)}
        .btn:disabled{background:var(--border);cursor:not-allowed;transform:none}
        .progress{text-align:center;padding:30px;display:none}.progress.show{display:block}
        .progress-text{margin-bottom:10px;font-size:14px;color:var(--text-muted)}
        .progress-text span{color:var(--accent);font-weight:bold}
        .progress-bar-wrap{background:var(--border);border-radius:10px;height:20px;width:100%;max-width:400px;margin:0 auto;overflow:hidden}
        .progress-bar{background:linear-gradient(90deg,var(--accent) 0%,#764ba2 100%);height:100%;width:0%;border-radius:10px;transition:width .3s ease}
        .progress-percent{font-size:12px;color:var(--text-muted);margin-top:5px}
        .result{text-align:center;padding:20px;background:var(--success-bg);border-radius:8px;margin-top:15px;display:none;color:var(--success-text)}
        .result.show{display:block}.result-title{font-size:18px;margin-bottom:10px}
        .result-size{font-size:14px;opacity:.8}
        .download-btn{display:inline-block;margin-top:15px;padding:12px 30px;background:var(--success-text);color:#fff;text-decoration:none;border-radius:6px;font-weight:bold}
        .download-btn:hover{opacity:.9}
        .error{padding:15px;background:var(--error-bg);color:var(--error-text);border-radius:8px;margin-top:15px;display:none}
        .error.show{display:block}.footer{text-align:center;margin-top:20px;font-size:12px;color:var(--text-muted);opacity:.6}
        .modal-overlay{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:var(--overlay-bg);z-index:1000;justify-content:center;align-items:center}
        .modal-overlay.show{display:flex}
        .modal{background:var(--card-bg);border-radius:12px;padding:25px;width:90%;max-width:480px;max-height:80vh;overflow-y:auto;box-shadow:0 4px 30px rgba(0,0,0,.3)}
        .modal-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:15px}
        .modal-title{font-size:18px;font-weight:bold;color:var(--text)}
        .modal-close{background:none;border:none;color:var(--text-muted);font-size:24px;cursor:pointer;line-height:1;padding:0 5px}
        .modal-close:hover{color:var(--error-text)}
        .history-empty{text-align:center;color:var(--text-muted);padding:20px;font-size:14px}
        .history-item{display:flex;justify-content:space-between;align-items:center;padding:10px 12px;border-radius:8px;margin-bottom:8px;cursor:pointer;border:1px solid var(--border);transition:all .2s;background:var(--bg)}
        .history-item:hover{border-color:var(--accent);background:var(--options-bg)}
        .history-item-info{flex:1;min-width:0}.history-item-name{font-weight:600;font-size:13px;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
        .history-item-sheets{font-size:11px;color:var(--text-muted);margin-top:3px}
        .history-item-meta{text-align:right;flex-shrink:0;margin-left:10px}
        .history-item-time{font-size:11px;color:var(--text-muted)}
        .history-item-size{font-size:11px;color:var(--accent);font-weight:bold}
        .history-clear{text-align:center;margin-top:10px}
        .history-clear-btn{background:var(--error-bg);color:var(--error-text);border:1px solid var(--error-text);padding:6px 16px;border-radius:5px;cursor:pointer;font-size:12px}
        .history-clear-btn:hover{opacity:.8}
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>Excel to draw.io Converter</h1>
        <div class="header-right">
            <button class="history-btn" id="historyBtn" onclick="showHistory()">
                <span>&#128203;</span> History <span class="badge" id="historyBadge">0</span>
            </button>
            <button class="theme-btn" id="themeBtn" onclick="toggleTheme()" title="Toggle dark mode">&#9788;</button>
        </div>
    </div>
    <div class="card">
        <div class="drop" id="drop"><div class="drop-icon">&#128193;</div><div class="drop-text">Drop Excel file or click to browse</div><div class="drop-hint">Supports .xlsx, .xls, .xlsm</div></div>
        <input type="file" id="fileInput" accept=".xlsx,.xls,.xlsm">
        <div class="info" id="info"><button class="clear-btn" onclick="clearFile()">Clear</button><div class="info-name" id="fileName"></div><div class="info-size" id="fileSize"></div></div>
        <div class="sheets" id="sheets"><div class="sheets-title">Select sheets to convert:</div><div class="sheet-list" id="sheetList"></div><div class="sheets-actions"><label><input type="checkbox" id="selectAll" checked onchange="toggleAll()"> Select All / Deselect All</label></div></div>
        <div class="preview" id="preview"><div class="preview-title">&#128202; Conversion Preview</div><div id="previewBody"></div><div class="preview-total"><span>Total</span><span><span id="totalShapes">0</span> shapes &middot; <span id="totalConnectors">0</span> connectors &nbsp;|&nbsp; <span class="est-size" id="estSize">~0 KB</span></span></div></div>
        <div class="options" id="options"><div class="options-title">Conversion Options:</div><div class="options-row"><div class="option-group"><label for="outputFormat">Output Format:</label><select id="outputFormat"><option value="drawio">draw.io (.drawio)</option><option value="svg">SVG (.svg)</option></select></div><label class="option-check"><input type="checkbox" id="includeConnectors" checked><span>Include connectors/lines</span></label><label class="option-check"><input type="checkbox" id="includeCellColors" checked><span>Include cell background colors</span></label></div></div>
        <div class="actions" id="actions"><button class="btn" id="convertBtn" onclick="convert()">Convert to draw.io</button></div>
        <div class="progress" id="progress"><div class="progress-text" id="progressText">Preparing conversion...</div><div class="progress-bar-wrap"><div class="progress-bar" id="progressBar"></div></div><div class="progress-percent" id="progressPercent">0%</div></div>
        <div class="result" id="result"><div class="result-title">Conversion Complete!</div><div class="result-size" id="resultSize"></div><a href="#" id="downloadBtn" class="download-btn" download>Download draw.io</a></div>
        <div class="error" id="error"></div>
    </div>
    <div class="footer">Excel to draw.io Converter - Works on Windows, Mac, Linux</div>
</div>
<div class="modal-overlay" id="historyModal"><div class="modal"><div class="modal-header"><div class="modal-title">&#128203; Conversion History</div><button class="modal-close" onclick="closeHistory()">&times;</button></div><div id="historyList"></div><div class="history-clear"><button class="history-clear-btn" onclick="clearHistory()">Clear All History</button></div></div></div>
<script>
var currentFile=null,currentFileName="",sheetPreviewData={};
var THEME_KEY="excel2drawio_dark_mode",OPTIONS_KEY="excel2drawio_options",LAST_SHEETS_KEY="excel2drawio_last_sheets",HISTORY_KEY="excel2drawio_history";
function initTheme(){var s=localStorage.getItem(THEME_KEY);if(s==="dark")applyTheme(true);else if(s===null&&window.matchMedia&&window.matchMedia("(prefers-color-scheme: dark)").matches)applyTheme(true);}
function applyTheme(d){document.documentElement.setAttribute("data-theme",d?"dark":"");document.getElementById("themeBtn").textContent=d?"☾":"☀";try{localStorage.setItem(THEME_KEY,d?"dark":"light");}catch(e){}}
function toggleTheme(){applyTheme(document.documentElement.getAttribute("data-theme")!=="dark");}
initTheme();
function loadOptions(){try{var s=localStorage.getItem(OPTIONS_KEY);if(s){var o=JSON.parse(s);document.getElementById("outputFormat").value=o.format||"drawio";document.getElementById("includeConnectors").checked=o.connectors!==false;document.getElementById("includeCellColors").checked=o.cellColors!==false;}}catch(e){}}
function saveOptions(){try{localStorage.setItem(OPTIONS_KEY,JSON.stringify({format:document.getElementById("outputFormat").value,connectors:document.getElementById("includeConnectors").checked,cellColors:document.getElementById("includeCellColors").checked}));}catch(e){}}
function loadLastSheets(fn){try{var s=localStorage.getItem(LAST_SHEETS_KEY);if(s){var d=JSON.parse(s);if(d.filename===fn)return d.sheets||[];}}catch(e){}return null;}
function saveLastSheets(fn,sheets){try{localStorage.setItem(LAST_SHEETS_KEY,JSON.stringify({filename:fn,sheets:sheets}));}catch(e){}}
function getHistory(){try{return JSON.parse(localStorage.getItem(HISTORY_KEY)||"[]");}catch(e){return[];}}
function saveHistory(h){try{localStorage.setItem(HISTORY_KEY,JSON.stringify(h.slice(-10)));}catch(e){}}
function addHistory(e){var h=getHistory();h.push(e);saveHistory(h);updateHistoryBadge();}
function updateHistoryBadge(){var n=getHistory().length;document.getElementById("historyBadge").textContent=n;document.getElementById("historyBtn").className="history-btn"+(n>0?" show":"");}
function showHistory(){var h=getHistory();var l=document.getElementById("historyList");if(!h.length)l.innerHTML="<div class="history-empty">No conversion history yet.</div>";else l.innerHTML=h.map(function(e,i){return"<div class="history-item" onclick="restoreFromHistory("+i+")"><div class="history-item-info"><div class="history-item-name">"+escHtml(e.filename)+"</div><div class="history-item-sheets">Sheets: "+escHtml(e.sheets.join(", "))+"</div></div><div class="history-item-meta"><div class="history-item-size">"+formatSize(e.outputSize)+"</div><div class="history-item-time">"+formatTime(e.timestamp)+"</div></div></div>";}).join("");document.getElementById("historyModal").classList.add("show");}
function closeHistory(){document.getElementById("historyModal").classList.remove("show");}
function clearHistory(){localStorage.removeItem(HISTORY_KEY);updateHistoryBadge();closeHistory();}
function restoreFromHistory(i){var h=getHistory();if(!h[i])return;closeHistory();alert("Please load the file \u201c"+h[i].filename+"\u201d again to auto-select sheets: "+h[i].sheets.join(", "));}
function escHtml(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function formatTime(ts){if(!ts)return"";var d=new Date(ts),now=new Date();var m=Math.floor((now-d)/60000);if(m<1)return"Just now";if(m<60)return m+"m ago";var h=Math.floor(m/60);if(h<24)return h+"h ago";return Math.floor(h/24)+"d ago";}
document.getElementById("historyModal").addEventListener("click",function(e){if(e.target===this)closeHistory();});
updateHistoryBadge();
var drop=document.getElementById("drop"),fileInput=document.getElementById("fileInput");
drop.onclick=function(){fileInput.click();};
drop.ondragover=function(e){e.preventDefault();drop.classList.add("drag");};
drop.ondragleave=function(){drop.classList.remove("drag");};
drop.ondrop=function(e){e.preventDefault();drop.classList.remove("drag");if(e.dataTransfer.files[0])loadFile(e.dataTransfer.files[0]);};
fileInput.onchange=function(e){if(e.target.files[0])loadFile(e.target.files[0]);};
function loadFile(file){if(!file.name.match(/\\.(xlsx|xls|xlsm)$/i)){showError("Please select an Excel file (.xlsx, .xls, or .xlsm)");return;}currentFile=file;currentFileName=file.name;sheetPreviewData={};document.getElementById("fileName").textContent=file.name;document.getElementById("fileSize").textContent=formatSize(file.size);document.getElementById("info").classList.add("show");document.getElementById("preview").classList.remove("show");drop.style.display="none";loadSheets(file);}
async function loadSheets(file){var reader=new FileReader();reader.onload=async function(e){var b64=e.target.result.split(",")[1];try{var resp=await fetch("/sheets?filename="+encodeURIComponent(currentFileName),{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({data:b64})});var data=await resp.json();if(data.success)showSheets(data.sheets);else showError(data.error);}catch(err){showError(err.message);}};reader.readAsDataURL(file);}
function showSheets(sheets){var list=document.getElementById("sheetList");list.innerHTML="";var lastSheets=loadLastSheets(currentFileName);sheets.forEach(function(s){var checked=lastSheets?lastSheets.indexOf(s)!==-1:true;var label=document.createElement("label");label.className="sheet";label.innerHTML="<input type="checkbox" name="sheets" value=""+s+"""+(checked?" checked":"")+"> "+escHtml(s);list.appendChild(label);});document.getElementById("sheets").classList.add("show");document.getElementById("options").classList.add("show");document.getElementById("actions").classList.add("show");updateConvertBtnText();loadOptions();document.querySelectorAll("[name=sheets]").forEach(function(c){c.addEventListener("change",onSheetChange);});var sel=[].slice.call(document.querySelectorAll("[name=sheets]:checked")).map(function(c){return c.value;});if(sel.length>0)fetchPreview();}
var previewTimeout=null;
function onSheetChange(){clearTimeout(previewTimeout);previewTimeout=setTimeout(function(){var sel=[].slice.call(document.querySelectorAll("[name=sheets]:checked")).map(function(c){return c.value;});if(sel.length>0)fetchPreview();else document.getElementById("preview").classList.remove("show");},300);}
async function fetchPreview(){var sel=[].slice.call(document.querySelectorAll("[name=sheets]:checked")).map(function(c){return c.value;});if(!sel.length||!currentFile)return;var reader=new FileReader();reader.onload=async function(e){var b64=e.target.result.split(",")[1];try{var resp=await fetch("/preview?filename="+encodeURIComponent(currentFileName),{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({data:b64,sheets:sel})});var data=await resp.json();if(data.success){sheetPreviewData=data.preview||{};renderPreview(sel);}}catch(err){}};reader.readAsDataURL(currentFile);}
function renderPreview(sel){var body=document.getElementById("previewBody");var ts=0,tc=0;body.innerHTML=sel.map(function(sn){var info=sheetPreviewData[sn]||{shapes:0,connectors:0};ts+=info.shapes||0;tc+=info.connectors||0;return"<div class="preview-row"><span class="sheet-name">"+escHtml(sn)+"</span><span class="counts"><span>"+(info.shapes||0)+" shapes</span><span>"+(info.connectors||0)+" connectors</span></span></div>";}).join("");document.getElementById("totalShapes").textContent=ts;document.getElementById("totalConnectors").textContent=tc;document.getElementById("estSize").textContent="~"+formatSize(2048+ts*200+tc*100);document.getElementById("preview").classList.add("show");}
function updateConvertBtnText(){var fmt=document.getElementById("outputFormat").value;document.getElementById("convertBtn").textContent="Convert to "+(fmt==="svg"?"SVG":"draw.io");}
document.getElementById("outputFormat").addEventListener("change",function(){saveOptions();updateConvertBtnText();});
document.getElementById("includeConnectors").addEventListener("change",saveOptions);
document.getElementById("includeCellColors").addEventListener("change",saveOptions);
function toggleAll(){var checked=document.getElementById("selectAll").checked;document.querySelectorAll("[name=sheets]").forEach(function(c){c.checked=checked;});setTimeout(function(){var sel=[].slice.call(document.querySelectorAll("[name=sheets]:checked")).map(function(c){return c.value;});if(sel.length)fetchPreview();else document.getElementById("preview").classList.remove("show");},50);}
function clearFile(){currentFile=null;currentFileName="";sheetPreviewData={};fileInput.value="";["info","sheets","options","actions","result","error","progress","preview"].forEach(function(id){document.getElementById(id).classList.remove("show");});resetProgress();drop.style.display="block";}
function resetProgress(){setProgress(0,"Preparing conversion...");}
function setProgress(percent,text){document.getElementById("progressBar").style.width=Math.min(100,Math.max(0,percent))+"%";if(text)document.getElementById("progressText").innerHTML=text;document.getElementById("progressPercent").textContent=Math.round(percent)+"%";}
function showProgress(){document.getElementById("progress").classList.add("show");}
function hideProgress(){document.getElementById("progress").classList.remove("show");resetProgress();}
function hideResult(){document.getElementById("result").classList.remove("show");}
function hideError(){document.getElementById("error").classList.remove("show");}
async function convert(){if(!currentFile)return;var sheets=[].slice.call(document.querySelectorAll("[name=sheets]:checked")).map(function(c){return c.value;});if(!sheets.length){showError("Please select at least one sheet");return;}saveOptions();saveLastSheets(currentFileName,sheets);var format=document.getElementById("outputFormat").value;var includeConnectors=document.getElementById("includeConnectors").checked;var includeCellColors=document.getElementById("includeCellColors").checked;showProgress();setProgress(5,"Reading file...");document.getElementById("convertBtn").disabled=true;hideResult();hideError();var reader=new FileReader();reader.onload=async function(e){var b64=e.target.result.split(",")[1];var params=new URLSearchParams({sheets:sheets.join(","),format:format,connectors:includeConnectors?"1":"0",cellColors:includeCellColors?"1":"0"});try{var response=await fetch("/convert-stream?filename="+encodeURIComponent(currentFileName)?"+params.toString(),{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({data:b64})});if(!response.ok){var err=await response.json().catch(function(){return{};});throw new Error(err.error||"Conversion failed (HTTP "+response.status+")");}var reader2=response.body.getReader();var decoder=new TextDecoder();var buffer="",done=false,resultData=null;while(!done){var r=await reader2.read();done=r.done;if(r.value){buffer+=decoder.decode(r.value,{stream:!done});var lines=buffer.split("
");buffer=lines.pop()||"";for(var i=0;i<lines.length;i++){if(lines[i].indexOf("data: ")===0){try{var msg=JSON.parse(lines[i].slice(6));handleStreamMessage(msg);if(msg.type==="complete"||msg.type==="error"){done=true;if(msg.type==="complete")resultData=msg;}}catch(e){}}}}}if(buffer.indexOf("data: ")===0){try{var msg=JSON.parse(buffer.slice(6));handleStreamMessage(msg);if(msg.type==="complete")resultData=msg;}catch(e){}}hideProgress();document.getElementById("convertBtn").disabled=false;if(resultData){addHistory({filename:currentFileName,sheets:sheets,timestamp:Date.now(),outputSize:resultData.size||0});showResult(resultData,format);}}catch(err){hideProgress();document.getElementById("convertBtn").disabled=false;if(err.name!=="AbortError")showError(err.message);}};reader.readAsDataURL(currentFile);}
function handleStreamMessage(msg){switch(msg.type){case"progress":setProgress(msg.percent||0,msg.text||"Converting...");break;case"complete":setProgress(100,"Done!");break;case"error":hideProgress();document.getElementById("convertBtn").disabled=false;showError(msg.error||"Conversion failed");break;}}
function showResult(data,format){var size=data.size||0;document.getElementById("resultSize").textContent="Size: "+formatSize(size);var baseName=currentFileName.replace(/\\.[^.]+$/,"");var ext=format==="svg"?".svg":".drawio";var mimeType=format==="svg"?"image/svg+xml":"application/octet-stream";var dlBtn=document.getElementById("downloadBtn");dlBtn.href="data:"+mimeType+";base64,"+data.file;dlBtn.download=baseName+ext;dlBtn.textContent=format==="svg"?"Download SVG":"Download draw.io";document.getElementById("result").classList.add("show");}
function formatSize(bytes){if(!bytes||bytes<1024)return(bytes||0)+" B";if(bytes<1048576)return(bytes/1024).toFixed(1)+" KB";return(bytes/1048576).toFixed(1)+" MB";}
function showError(msg){document.getElementById("error").textContent="Error: "+msg;document.getElementById("error").classList.add("show");}
<\/script>
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
                TEMP_FILE = os.path.join(tempfile.gettempdir(), f'temp_upload.{ext}')
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

        elif self.path.startswith('/preview'):
            self._handle_preview()
        elif self.path.startswith('/convert-stream'):
            self._handle_convert_stream()
        elif self.path.startswith('/convert'):
            self._handle_convert_legacy()

        else:
            self.send_error(404)

    def _handle_preview(self):
        """Return shape and connector counts for selected sheets."""
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length).decode())
            data_bytes = base64.b64decode(body.get('data', ''))
            selected_sheets = body.get('sheets', [])
            global TEMP_FILE
            if not TEMP_FILE or not os.path.exists(TEMP_FILE):
                TEMP_FILE = os.path.join(tempfile.gettempdir(), f'temp_upload.{ext}')
                with open(TEMP_FILE, 'wb') as f:
                    f.write(data_bytes)
            reader = ExcelReader(TEMP_FILE, sheet_names=selected_sheets if selected_sheets else None)
            all_data = reader.read_all()
            reader.close()
            preview = {}
            for sheet_name in (selected_sheets or list(all_data.keys())):
                data = dict(all_data.get(sheet_name, {}))
                preview[sheet_name] = {
                    'shapes': len(data.get('shapes', [])),
                    'connectors': len(data.get('connectors', []))
                }
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True, 'preview': preview}).encode())
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    def _handle_convert_stream(self):
        """Streaming convert with SSE progress updates."""
        try:
            global TEMP_FILE
            query = ''
            if '?' in self.path:
                query = self.path.split('?', 1)[1]
            params = {}
            for p in query.split('&'):
                if '=' in p:
                    k, v = p.split('=', 1)
                    params[urllib.parse.unquote(k)] = urllib.parse.unquote(v)
            selected_sheets = [s for s in params.get('sheets', '').split(',') if s]
            fmt = params.get('format', 'drawio')
            include_connectors = params.get('connectors', '1') == '1'
            include_cells = params.get('cellColors', '1') == '1'

            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length).decode())
            data_bytes = base64.b64decode(body.get('data', ''))

            if not TEMP_FILE or not os.path.exists(TEMP_FILE):
                TEMP_FILE = os.path.join(tempfile.gettempdir(), f'temp_upload.{ext}')
                with open(TEMP_FILE, 'wb') as f:
                    f.write(data_bytes)

            self.send_response(200)
            self.send_header('Content-Type', 'text/event-stream')
            self.send_header('Cache-Control', 'no-cache')
            self.send_header('Connection', 'keep-alive')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('X-Accel-Buffering', 'no')
            self.end_headers()

            def progress_callback(percent, text):
                self.send_sse({'type': 'progress', 'percent': percent, 'text': text})

            progress_callback(5, 'Reading workbook...')
            reader = ExcelReader(TEMP_FILE, sheet_names=selected_sheets if selected_sheets else None, include_cells=include_cells)
            all_sheets_data = reader.read_all()
            sheet_names = list(all_sheets_data.keys())
            n_sheets = len(sheet_names)
            progress_callback(15, f'Loaded {n_sheets} sheet(s), extracting shapes...')

            sheets_data = {}
            for i, sheet_name in enumerate(sheet_names):
                pct = 15 + int(55 * (i / max(n_sheets, 1)))
                progress_callback(pct, f'Processing sheet {i+1}/{n_sheets}: {sheet_name}...')
                ws_data = dict(all_sheets_data[sheet_name])
                if not include_connectors:
                    ws_data['connectors'] = []
                sheets_data[sheet_name] = ws_data

            progress_callback(75, 'Writing output file...')
            output_path = os.path.join(tempfile.gettempdir(), 'output.drawio')
            if fmt == 'svg':
                self._write_svg(sheets_data, output_path, include_connectors, progress_callback)
            else:
                writer = DrawioWriter(sheets_data)
                writer.write(output_path)

            progress_callback(95, 'Preparing download...')
            with open(output_path, 'rb') as f:
                output_data_bytes = f.read()
            b64 = base64.b64encode(output_data_bytes).decode()
            progress_callback(100, 'Conversion complete!')
            self.send_sse({'type': 'complete', 'file': b64, 'size': len(output_data_bytes)})
        except Exception as e:
            try:
                self.send_sse({'type': 'error', 'error': str(e)})
            except:
                pass

    def _write_svg(self, sheets_data, output_path, include_connectors, progress_callback):
        """Write sheets data as SVG."""
        try:
            from converter.shape_mapper import ShapeMapper
            import xml.etree.ElementTree as ET
        except ImportError:
            writer = DrawioWriter(sheets_data)
            writer.write(output_path)
            return
        mapper = ShapeMapper()
        first_sheet_name = list(sheets_data.keys())[0] if sheets_data else 'Sheet1'
        data = sheets_data.get(first_sheet_name, {})
        shapes = data.get('shapes', [])
        connectors = data.get('connectors', []) if include_connectors else []
        if not shapes:
            svg_content = '<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" width="800" height="600" viewBox="0 0 800 600"><rect width="800" height="600" fill="white"/><text x="400" y="300" text-anchor="middle" fill="#999" font-family="sans-serif" font-size="14">No shapes found</text></svg>'
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(svg_content)
            return
        min_x = min(s.x for s in shapes)
        min_y = min(s.y for s in shapes)
        max_x = max(s.x + s.width for s in shapes)
        max_y = max(s.y + s.height for s in shapes)
        padding = 20
        width_px = max(800, (max_x - min_x) / 914400 * 96 + padding * 2)
        height_px = max(600, (max_y - min_y) / 914400 * 96 + padding * 2)
        svg = ET.Element('svg', {'xmlns': 'http://www.w3.org/2000/svg', 'width': str(int(width_px)), 'height': str(int(height_px)), 'viewBox': f'0 0 {int(width_px)} {int(height_px)}'})
        ET.SubElement(svg, 'rect', {'width': '100%', 'height': '100%', 'fill': 'white'})
        for i, shape in enumerate(shapes):
            pct = 75 + int(20 * i / max(len(shapes), 1))
            progress_callback(pct, f'Writing shape {i+1}/{len(shapes)}...')
            x_px = shape.x / 914400 * 96 - min_x / 914400 * 96 + padding
            y_px = shape.y / 914400 * 96 - min_y / 914400 * 96 + padding
            w_px = shape.width / 914400 * 96
            h_px = shape.height / 914400 * 96
            style = shape.style or {}
            fill = style.get('fillColor', '#ffffff')
            stroke = style.get('strokeColor', '#000000')
            stroke_w = style.get('strokeWidth', 1)
            shape_type = mapper.map_type(shape.type)
            if shape_type == 'ellipse':
                cx = x_px + w_px / 2; cy = y_px + h_px / 2
                ET.SubElement(svg, 'ellipse', {'cx': str(cx), 'cy': str(cy), 'rx': str(max(w_px / 2, 1)), 'ry': str(max(h_px / 2, 1)), 'fill': fill, 'stroke': stroke, 'stroke-width': str(stroke_w)})
            elif shape_type == 'diamond':
                cx = x_px + w_px / 2; cy = y_px + h_px / 2
                points = f'{cx},{y_px} {x_px+w_px},{cy} {cx},{y_px+h_px} {x_px},{cy}'
                ET.SubElement(svg, 'polygon', {'points': points, 'fill': fill, 'stroke': stroke, 'stroke-width': str(stroke_w)})
            else:
                ET.SubElement(svg, 'rect', {'x': str(x_px), 'y': str(y_px), 'width': str(max(w_px, 1)), 'height': str(max(h_px, 1)), 'fill': fill, 'stroke': stroke, 'stroke-width': str(stroke_w)})
            if shape.text:
                font_size = int(style.get('fontSize', 12))
                font_color = style.get('fontColor', '#000000')
                txt_elem = ET.SubElement(svg, 'text', {'x': str(x_px + 4), 'y': str(y_px + h_px / 2 + font_size / 3), 'fill': font_color, 'font-size': str(font_size), 'font-family': 'sans-serif'})
                txt_elem.text = shape.text
        for conn in connectors:
            if not conn.points or len(conn.points) < 2:
                continue
            x1 = conn.points[0][0] / 914400 * 96 - min_x / 914400 * 96 + padding
            y1 = conn.points[0][1] / 914400 * 96 - min_y / 914400 * 96 + padding
            x2 = conn.points[-1][0] / 914400 * 96 - min_x / 914400 * 96 + padding
            y2 = conn.points[-1][1] / 914400 * 96 - min_y / 914400 * 96 + padding
            style = conn.style or {}
            stroke = style.get('strokeColor', '#000000')
            stroke_w = style.get('strokeWidth', 1)
            ET.SubElement(svg, 'line', {'x1': str(x1), 'y1': str(y1), 'x2': str(x2), 'y2': str(y2), 'stroke': stroke, 'stroke-width': str(stroke_w)})
        ET.indent(svg)
        tree = ET.ElementTree(svg)
        tree.write(output_path, encoding='utf-8', xml_declaration=True)

    def _handle_convert_legacy(self):
        """Legacy non-streaming convert for backward compatibility."""
        try:
            global TEMP_FILE
            sheets_param = ''
            if '?' in self.path:
                query = self.path.split('?')[1]
                for p in query.split('&'):
                    if p.startswith('sheets='):
                        sheets_param = p.split('=')[1]
            selected_sheets = [s for s in urllib.parse.unquote(sheets_param).split(',') if s]
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length).decode())
            data_bytes = base64.b64decode(body.get('data', ''))
            if not TEMP_FILE or not os.path.exists(TEMP_FILE):
                TEMP_FILE = os.path.join(tempfile.gettempdir(), f'temp_upload.{ext}')
                with open(TEMP_FILE, 'wb') as f:
                    f.write(data_bytes)
            reader = ExcelReader(TEMP_FILE, sheet_names=selected_sheets if selected_sheets else None)
            sheets_data = reader.read_all()
            output_path = os.path.join(tempfile.gettempdir(), 'output.drawio')
            writer = DrawioWriter(sheets_data)
            writer.write(output_path)
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

    def send_sse(self, data):
        try:
            self.wfile.write(("data: " + json.dumps(data) + "\n\n").encode('utf-8'))
        except:
            pass

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
