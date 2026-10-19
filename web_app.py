"""
Pub2Doc Web Service - FastAPI Backend
Provides web UI and REST API for converting .pub files to .docx over HTTP.
"""

import os
import sys
import shutil
import uuid
from typing import Optional
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from pub2doc_engine import PubConverter, check_publisher_available

app = FastAPI(title="Pub2Doc Web Service", version="1.0.0")

UPLOAD_DIR = os.path.abspath("uploads")
OUTPUT_DIR = os.path.abspath("outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

HTML_PAGE = """
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Pub2Doc Cloud - แปลงไฟล์ Publisher เป็น Word</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Sarabun', sans-serif; }
    </style>
</head>
<body class="bg-slate-50 min-h-screen text-slate-800">
    <div class="max-w-3xl mx-auto px-4 py-12">
        <!-- Header -->
        <div class="text-center mb-8">
            <div class="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-blue-600 text-white text-3xl font-bold mb-4 shadow-lg shadow-blue-500/30">
                P
            </div>
            <h1 class="text-3xl font-bold text-slate-900">Pub2Doc Cloud Service</h1>
            <p class="text-slate-600 mt-2">บริการแปลงไฟล์ Microsoft Publisher (.pub) เป็น Word (.docx) ออนไลน์</p>
            <p class="text-sm text-slate-500">คงรูปภาษาไทย สระ-วรรณยุกต์ ตาราง และรูปภาพสมบูรณ์แบบ</p>
        </div>

        <!-- Main Card -->
        <div class="bg-white rounded-2xl shadow-xl shadow-slate-200/60 p-8 border border-slate-100">
            <!-- Dropzone -->
            <div id="dropzone" class="border-2 border-dashed border-slate-300 rounded-xl p-10 text-center hover:border-blue-500 transition-colors cursor-pointer bg-slate-50/50">
                <svg class="mx-auto h-12 w-12 text-slate-400 mb-3" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                </svg>
                <div class="text-base font-semibold text-slate-700">ลากไฟล์ .PUB มาวางที่นี่</div>
                <div class="text-sm text-slate-500 mt-1">หรือคลิกเพื่อเลือกไฟล์จากอุปกรณ์ของคุณ</div>
                <input type="file" id="fileInput" accept=".pub" class="hidden">
            </div>

            <!-- Selected File Info -->
            <div id="fileInfo" class="hidden mt-6 p-4 rounded-xl bg-blue-50 border border-blue-100 flex items-center justify-between">
                <div class="flex items-center space-x-3 overflow-hidden">
                    <span class="text-2xl">📄</span>
                    <div class="truncate">
                        <div id="fileName" class="font-semibold text-slate-800 truncate">document.pub</div>
                        <div id="fileSize" class="text-xs text-slate-500">0 KB</div>
                    </div>
                </div>
                <button id="removeBtn" class="text-slate-400 hover:text-red-500 text-sm font-semibold">ลบ</button>
            </div>

            <!-- Action Button -->
            <div class="mt-6">
                <button id="convertBtn" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3.5 px-6 rounded-xl shadow-lg shadow-blue-500/25 transition-all disabled:opacity-50 disabled:cursor-not-allowed">
                    เริ่มแปลงเป็น Word (.docx)
                </button>
            </div>

            <!-- Progress / Loading -->
            <div id="loading" class="hidden mt-6 text-center">
                <div class="inline-block animate-spin rounded-full h-8 w-8 border-4 border-blue-600 border-t-transparent mb-2"></div>
                <div class="text-sm font-semibold text-slate-700">กำลังประมวลผลและสร้างเอกสาร Word...</div>
                <div class="text-xs text-slate-500 mt-1">โปรดรอสักครู่ ระบบกำลังจัดฟอร์มและซ่อมสระภาษาไทย</div>
            </div>

            <!-- Download Result -->
            <div id="resultBox" class="hidden mt-6 p-5 rounded-xl bg-emerald-50 border border-emerald-200 text-center">
                <div class="text-3xl mb-2">🎉</div>
                <div class="font-bold text-emerald-800 text-lg">แปลงไฟล์สำเร็จเรียบร้อย!</div>
                <p class="text-sm text-emerald-600 mt-1">เอกสารของคุณพร้อมให้ดาวน์โหลดแล้ว</p>
                <a id="downloadLink" href="#" class="inline-block mt-4 bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-2.5 px-8 rounded-xl shadow-md transition-colors">
                    ⬇️ ดาวน์โหลดไฟล์ .DOCX
                </a>
            </div>
        </div>

        <!-- Footer -->
        <div class="text-center text-xs text-slate-400 mt-8">
            Pub2Doc Cloud Edition • ออกแบบสำหรับรองรับการปิดตัวของ Microsoft Publisher
        </div>
    </div>

    <script>
        const dropzone = document.getElementById('dropzone');
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const fileSize = document.getElementById('fileSize');
        const removeBtn = document.getElementById('removeBtn');
        const convertBtn = document.getElementById('convertBtn');
        const loading = document.getElementById('loading');
        const resultBox = document.getElementById('resultBox');
        const downloadLink = document.getElementById('downloadLink');

        let selectedFile = null;

        dropzone.addEventListener('click', () => fileInput.click());

        dropzone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropzone.classList.add('border-blue-500', 'bg-blue-50/30');
        });

        dropzone.addEventListener('dragleave', () => {
            dropzone.classList.remove('border-blue-500', 'bg-blue-50/30');
        });

        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropzone.classList.remove('border-blue-500', 'bg-blue-50/30');
            if (e.dataTransfer.files.length > 0) {
                handleFile(e.dataTransfer.files[0]);
            }
        });

        fileInput.addEventListener('change', () => {
            if (fileInput.files.length > 0) {
                handleFile(fileInput.files[0]);
            }
        });

        function handleFile(file) {
            if (!file.name.toLowerCase().endsWith('.pub')) {
                alert('กรุณาเลือกไฟล์นามสกุล .pub เท่านั้น');
                return;
            }
            selectedFile = file;
            fileName.textContent = file.name;
            fileSize.textContent = (file.size / 1024).toFixed(1) + ' KB';
            fileInfo.classList.remove('hidden');
            resultBox.classList.add('hidden');
        }

        removeBtn.addEventListener('click', () => {
            selectedFile = null;
            fileInput.value = '';
            fileInfo.classList.add('hidden');
        });

        convertBtn.addEventListener('click', async () => {
            if (!selectedFile) {
                alert('กรุณาเลือกไฟล์ .pub ก่อน');
                return;
            }

            const formData = new FormData();
            formData.append('file', selectedFile);

            convertBtn.disabled = true;
            loading.classList.remove('hidden');
            resultBox.classList.add('hidden');

            try {
                const response = await fetch('/api/convert', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.detail || 'เกิดข้อผิดพลาดในการแปลงไฟล์');
                }

                const data = await response.json();
                downloadLink.href = data.download_url;
                resultBox.classList.remove('hidden');
            } catch (err) {
                alert('เกิดข้อผิดพลาด: ' + err.message);
            } finally {
                convertBtn.disabled = false;
                loading.classList.add('hidden');
            }
        });
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
async def serve_home():
    return HTMLResponse(content=HTML_PAGE)

@app.post("/api/convert")
async def convert_file(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pub"):
        raise HTTPException(status_code=400, detail="อนุญาตเฉพาะไฟล์ .pub เท่านั้น")

    job_id = str(uuid.uuid4())[:8]
    orig_name = os.path.splitext(file.filename)[0]
    
    saved_pub = os.path.join(UPLOAD_DIR, f"{job_id}_{file.filename}")
    out_docx = os.path.join(OUTPUT_DIR, f"{orig_name}_{job_id}.docx")

    try:
        with open(saved_pub, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Run conversion engine
        converter = PubConverter(saved_pub, out_docx)
        converter.convert()

        download_url = f"/api/download/{os.path.basename(out_docx)}"
        return {
            "status": "success",
            "filename": f"{orig_name}.docx",
            "download_url": download_url
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="ไม่พบไฟล์ที่ระบุ")
    return FileResponse(file_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("web_app:app", host="0.0.0.0", port=8000, reload=False)
