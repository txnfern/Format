#!/usr/bin/env python3
"""
Excel Color Extractor - FastAPI Web Service
Complete web service with API endpoints and file handling
"""

import os
import re
import math
import uuid
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, List
import pandas as pd
from openpyxl import load_workbook

from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn

# Initialize FastAPI app
app = FastAPI(
    title="Excel Color Extractor API",
    description="Extract colors from Excel matrices and generate Price/Type files",
    version="1.0.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Create directories
UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
STATIC_DIR = Path("static")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)

# Serve static files
app.mount("/static", StaticFiles(directory="static"), name="static")

class ProcessingResult(BaseModel):
    job_id: str
    status: str
    message: str
    price_file: Optional[str] = None
    type_file: Optional[str] = None
    processing_time: Optional[float] = None
    total_records: Optional[int] = None

class ColorExtractor:
    def __init__(self, job_id: str):
        self.job_id = job_id
        
    def to_number(self, val):
        """Convert value to number, removing commas"""
        try:
            if val is None:
                return None
            
            str_val = str(val).strip()
            # Remove comma, space, and special characters
            clean_val = re.sub(r'[,\s]', '', str_val)
            clean_val = re.sub(r'[^\d.-]', '', clean_val)
            
            if not clean_val or clean_val in ['', '-', '.']:
                return None
                
            f = float(clean_val)
            if math.isnan(f):
                return None
            return int(f) if f.is_integer() else f
        except:
            return None

    def normalize_rgb(self, fill):
        """Convert ARGB color to RGB hex format"""
        if not fill:
            return "FFFFFF"
        
        color_found = ""
        
        # Check fgColor
        if hasattr(fill, 'fgColor') and fill.fgColor:
            if hasattr(fill.fgColor, 'rgb') and fill.fgColor.rgb:
                color_str = str(fill.fgColor.rgb).upper()
                if color_str == "00000000":
                    return "FFFFFF"
                elif len(color_str) == 8:
                    color_found = color_str[2:]
                elif len(color_str) == 6:
                    color_found = color_str
        
        # Check bgColor
        if not color_found and hasattr(fill, 'bgColor') and fill.bgColor:
            if hasattr(fill.bgColor, 'rgb') and fill.bgColor.rgb:
                color_str = str(fill.bgColor.rgb).upper()
                if color_str == "00000000":
                    return "FFFFFF"
                elif len(color_str) == 8:
                    color_found = color_str[2:]
                elif len(color_str) == 6:
                    color_found = color_str
        
        return color_found if color_found else "FFFFFF"

    def find_thickness_matrix(self, ws, raw, thickness_mm):
        """Find matrix with specific thickness label"""
        thickness_patterns = [
            rf"Thk\.{thickness_mm}\s*mm",
            rf"{thickness_mm}\s*mm",
            rf"Thickness\s*{thickness_mm}",
            rf"หนา\s*{thickness_mm}"
        ]
        
        thickness_row = thickness_col = None
        for r in range(raw.shape[0]):
            for c in range(raw.shape[1]):
                cell_val = str(raw.iat[r, c]).strip() if raw.iat[r, c] is not None else ""
                for pattern in thickness_patterns:
                    if re.search(pattern, cell_val, re.IGNORECASE):
                        thickness_row, thickness_col = r, c
                        break
                if thickness_row is not None:
                    break
            if thickness_row is not None:
                break
        
        if thickness_row is None:
            return None, None
        
        # Search for h/w header
        search_range = 15
        for r in range(max(0, thickness_row - search_range), min(raw.shape[0], thickness_row + search_range + 1)):
            for c in range(max(0, thickness_col - search_range), min(raw.shape[1], thickness_col + search_range + 1)):
                cell_val = str(raw.iat[r, c]).strip() if raw.iat[r, c] is not None else ""
                if re.search(r"\bh\s*/\s*w\b", cell_val, re.IGNORECASE):
                    return r, c
        
        return None, None

    def read_color_matrix(self, ws, raw, hr, hc, widths, heights):
        """Read colors from matrix"""
        color_map = {}
        
        for i_h, h in enumerate(heights):
            for i_w, w in enumerate(widths):
                try:
                    excel_row = hr + 2 + i_h
                    excel_col = hc + 2 + i_w
                    
                    cell = ws.cell(row=excel_row, column=excel_col)
                    color = self.normalize_rgb(cell.fill)
                    color_map[(h, w)] = color
                except Exception:
                    continue
        
        return color_map

    def process_file(self, input_file: str, original_filename: str = None):
        """Process the Excel file"""
        try:
            if original_filename:
                base_name = os.path.splitext(original_filename)[0]
            else:
                base_name = os.path.splitext(os.path.basename(input_file))[0]
                # ลบ UUID ออกจากชื่อไฟล์ (UUID format: 8-4-4-4-12 characters)
                uuid_pattern = r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}_'
                base_name = re.sub(uuid_pattern, '', base_name)
            
            xls = pd.ExcelFile(input_file, engine="openpyxl")
            wb = load_workbook(input_file, data_only=True)
            
            price_rows = []
            type_rows = []
            price_id = 1
            type_id = 1
            
            for sheet in xls.sheet_names:
                if sheet.strip().lower() == "สารบัญ":
                    continue
                
                raw = pd.read_excel(xls, sheet_name=sheet, header=None, engine="openpyxl")
                ws = wb[sheet]
                
                # Find Glass_QTY and Description
                sheet_glass_qty = 1
                sheet_description = ""
                
                for r in range(raw.shape[0]):
                    for c in range(raw.shape[1] - 1):
                        if raw.iat[r, c] is None:
                            continue
                        cell = str(raw.iat[r, c]).strip()
                        low = cell.lower()
                        
                        if low in ("glass_qty", "glass qty"):
                            next_cell = raw.iat[r, c + 1]
                            qty = self.to_number(next_cell)
                            if qty is not None:
                                sheet_glass_qty = qty
                            
                        elif low == "description":
                            desc = raw.iat[r, c + 1]
                            if desc is not None:
                                sheet_description = str(desc).strip()
                
                # Find h/w header
                locs = []
                for r in range(raw.shape[0]):
                    for c in range(raw.shape[1]):
                        if raw.iat[r, c] is None:
                            continue
                        if isinstance(raw.iat[r, c], str):
                            if re.search(r"\bh\s*/\s*w\b", raw.iat[r, c], re.IGNORECASE):
                                locs.append((r, c))
                
                if not locs:
                    continue
                
                hr, hc = locs[0]
                
                # Read widths and heights
                widths = []
                for c in range(hc + 1, raw.shape[1]):
                    v = self.to_number(raw.iat[hr, c])
                    if v is None:
                        break
                    widths.append(v)
                
                heights = []
                for r in range(hr + 1, raw.shape[0]):
                    h_val = self.to_number(raw.iat[r, hc])
                    if h_val is None:
                        break
                    heights.append(h_val)
                
                if not widths or not heights:
                    continue
                
                # Find thickness matrices
                color_5mm = {}
                color_6mm = {}
                color_8mm = {}
                
                for thickness in [5, 6, 8]:
                    hr_thick, hc_thick = self.find_thickness_matrix(ws, raw, thickness)
                    if hr_thick is not None:
                        # Read dimensions for thickness matrix
                        widths_thick = []
                        for c in range(hc_thick + 1, raw.shape[1]):
                            v = self.to_number(raw.iat[hr_thick, c])
                            if v is None:
                                break
                            widths_thick.append(v)
                        
                        heights_thick = []
                        for r in range(hr_thick + 1, raw.shape[0]):
                            h_val = self.to_number(raw.iat[r, hc_thick])
                            if h_val is None:
                                break
                            heights_thick.append(h_val)
                        
                        if widths_thick and heights_thick:
                            colors = self.read_color_matrix(ws, raw, hr_thick, hc_thick, widths_thick, heights_thick)
                            if thickness == 5:
                                color_5mm = colors
                            elif thickness == 6:
                                color_6mm = colors
                            elif thickness == 8:
                                color_8mm = colors
                
                # Create Type record
                type_rows.append({
                    "ID": type_id,
                    "Serie": base_name,
                    "Type": sheet.strip(),
                    "Description": sheet_description,
                    "width_min": min(widths),
                    "width_max": max(widths),
                    "height_min": min(heights),
                    "height_max": max(heights),
                })
                type_id += 1
                
                # Create Price records
                for i_h, h in enumerate(heights):
                    for i_w, w in enumerate(widths):
                        raw_price = raw.iat[hr + 1 + i_h, hc + 1 + i_w]
                        p = self.to_number(raw_price)
                        if p is None:
                            continue
                        
                        color_5 = color_5mm.get((h, w), "FFFFFF")
                        color_6 = color_6mm.get((h, w), "FFFFFF")
                        color_8 = color_8mm.get((h, w), "FFFFFF")
                        
                        price_rows.append({
                            "ID": price_id,
                            "Serie": base_name,
                            "Type": sheet.strip(),
                            "Width": w,
                            "Height": h,
                            "Price": p,
                            "Glass_QTY": sheet_glass_qty,
                            "5mm_Color": color_5,
                            "6mm_Color": color_6,
                            "8mm_Color": color_8
                        })
                        price_id += 1
            
            # Save output files
            price_file = OUTPUT_DIR / f"Price_{self.job_id}.xlsx"
            type_file = OUTPUT_DIR / f"Type_{self.job_id}.xlsx"
            
            pd.DataFrame(price_rows).to_excel(price_file, index=False)
            pd.DataFrame(type_rows).to_excel(type_file, index=False)
            
            return {
                "price_file": str(price_file),
                "type_file": str(type_file),
                "total_records": len(price_rows)
            }
            
        except Exception as e:
            raise Exception(f"Processing failed: {str(e)}")

# API Endpoints
@app.get("/", response_class=HTMLResponse)
async def root():
    """Serve the main HTML interface"""
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Format Tostem</title>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
            body {
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                max-width: 800px;
                margin: 0 auto;
                padding: 20px;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                min-height: 100vh;
            }
            .container {
                background: white;
                padding: 30px;
                border-radius: 15px;
                box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            }
            h1 {
                color: #333;
                text-align: center;
                margin-bottom: 10px;
                font-size: 2.5em;
            }
            .subtitle {
                text-align: center;
                color: #666;
                margin-bottom: 30px;
                font-size: 1.1em;
            }
            .upload-area {
                border: 3px dashed #007bff;
                border-radius: 15px;
                padding: 50px;
                text-align: center;
                margin: 30px 0;
                cursor: pointer;
                transition: all 0.3s ease;
                background: #f8f9fa;
            }
            .upload-area:hover {
                background: #e3f2fd;
                border-color: #0056b3;
                transform: translateY(-2px);
            }
            .upload-area.dragover {
                background: #e3f2fd;
                border-color: #1976d2;
                transform: scale(1.02);
            }
            input[type="file"] {
                display: none;
            }
            .btn {
                background: linear-gradient(45deg, #007bff, #0056b3);
                color: white;
                padding: 15px 30px;
                border: none;
                border-radius: 25px;
                cursor: pointer;
                font-size: 16px;
                margin: 10px;
                transition: all 0.3s ease;
                font-weight: bold;
            }
            .btn:hover {
                transform: translateY(-2px);
                box-shadow: 0 5px 15px rgba(0,123,255,0.3);
            }
            .btn:disabled {
                background: #6c757d;
                cursor: not-allowed;
                transform: none;
            }
            .progress {
                display: none;
                margin: 30px 0;
            }
            .progress-bar {
                width: 100%;
                height: 25px;
                background: #e9ecef;
                border-radius: 15px;
                overflow: hidden;
                position: relative;
            }
            .progress-fill {
                height: 100%;
                background: linear-gradient(45deg, #28a745, #20c997);
                width: 0%;
                transition: width 0.3s ease;
                position: relative;
            }
            .progress-text {
                position: absolute;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                color: white;
                font-weight: bold;
                text-shadow: 1px 1px 2px rgba(0,0,0,0.3);
            }
            .result {
                margin: 30px 0;
                padding: 20px;
                border-radius: 10px;
                display: none;
                animation: slideIn 0.5s ease;
            }
            @keyframes slideIn {
                from { opacity: 0; transform: translateY(-20px); }
                to { opacity: 1; transform: translateY(0); }
            }
            .result.success {
                background: linear-gradient(45deg, #d4edda, #c3e6cb);
                color: #155724;
                border: 1px solid #c3e6cb;
            }
            .result.error {
                background: linear-gradient(45deg, #f8d7da, #f5c6cb);
                color: #721c24;
                border: 1px solid #f5c6cb;
            }
            .download-links {
                text-align: center;
                margin-top: 20px;
            }
            .download-link {
                display: inline-block;
                margin: 10px;
                padding: 12px 25px;
                background: linear-gradient(45deg, #28a745, #20c997);
                color: white;
                text-decoration: none;
                border-radius: 20px;
                transition: all 0.3s ease;
                font-weight: bold;
            }
            .download-link:hover {
                transform: translateY(-2px);
                box-shadow: 0 5px 15px rgba(40,167,69,0.3);
                color: white;
                text-decoration: none;
            }
            .features {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 20px;
                margin: 30px 0;
            }
            .feature {
                text-align: center;
                padding: 20px;
                background: #f8f9fa;
                border-radius: 10px;
                border: 1px solid #dee2e6;
            }
            .feature-icon {
                font-size: 2em;
                margin-bottom: 10px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>🎨 Format Tostem </h1>
            <p class="subtitle">
                แปลงไฟล์ Excel และสร้างไฟล์ Price/Type อัตโนมัติ
            </p>
            
            <div class="features">
                <div class="feature">
                    <div class="feature-icon">📊</div>
                    <strong>อ่านเมทริกซ์</strong><br>
                    <small>h/w, 5mm, 6mm, 8mm</small>
                </div>
                <div class="feature">
                    <div class="feature-icon">🎨</div>
                    <strong>สกัดสี</strong><br>
                    <small>RGB Hex codes</small>
                </div>
                <div class="feature">
                    <div class="feature-icon">💾</div>
                    <strong>ส่งออกไฟล์</strong><br>
                    <small>Price.xlsx & Type.xlsx</small>
                </div>
            </div>
            
            <div class="upload-area" onclick="document.getElementById('fileInput').click()">
                <div>
                    <strong style="font-size: 1.2em;">📁 คลิกหรือลากไฟล์ Excel มาวางที่นี่</strong><br><br>
                    <small style="color: #666;">รองรับไฟล์ .xlsx เท่านั้น (ขนาดไม่เกิน 25MB)</small>
                </div>
            </div>
            
            <input type="file" id="fileInput" accept=".xlsx" />
            
            <div style="text-align: center;">
                <button class="btn" onclick="uploadFile()" id="uploadBtn" disabled>
                    🚀 ประมวลผลไฟล์
                </button>
            </div>
            
            <div class="progress" id="progress">
                <div class="progress-bar">
                    <div class="progress-fill" id="progressFill"></div>
                    <div class="progress-text" id="progressText">กำลังประมวลผล... 0%</div>
                </div>
                <div style="text-align: center; margin-top: 15px;">
                    <span id="statusText">กำลังอัพโหลดไฟล์...</span>
                </div>
            </div>
            
            <div class="result" id="result">
                <div id="resultContent"></div>
                <div class="download-links" id="downloadLinks"></div>
            </div>
        </div>

        <script>
            let selectedFile = null;
            
            // File input handling
            document.getElementById('fileInput').addEventListener('change', function(e) {
                selectedFile = e.target.files[0];
                updateUploadButton();
            });
            
            // Drag and drop
            const uploadArea = document.querySelector('.upload-area');
            
            uploadArea.addEventListener('dragover', function(e) {
                e.preventDefault();
                uploadArea.classList.add('dragover');
            });
            
            uploadArea.addEventListener('dragleave', function(e) {
                e.preventDefault();
                uploadArea.classList.remove('dragover');
            });
            
            uploadArea.addEventListener('drop', function(e) {
                e.preventDefault();
                uploadArea.classList.remove('dragover');
                
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    selectedFile = files[0];
                    document.getElementById('fileInput').files = files;
                    updateUploadButton();
                }
            });
            
            function updateUploadButton() {
                const btn = document.getElementById('uploadBtn');
                if (selectedFile && selectedFile.name.endsWith('.xlsx')) {
                    btn.disabled = false;
                    btn.textContent = `🚀 ประมวลผล: ${selectedFile.name}`;
                } else {
                    btn.disabled = true;
                    btn.textContent = '🚀 ประมวลผลไฟล์';
                }
            }
            
            async function uploadFile() {
                if (!selectedFile) return;
                
                const formData = new FormData();
                formData.append('file', selectedFile);
                
                // Show progress
                document.getElementById('progress').style.display = 'block';
                document.getElementById('result').style.display = 'none';
                document.getElementById('uploadBtn').disabled = true;
                
                // Simulate progress
                let progress = 0;
                const progressInterval = setInterval(() => {
                    progress += Math.random() * 10;
                    if (progress > 90) progress = 90;
                    
                    const progressFill = document.getElementById('progressFill');
                    const progressText = document.getElementById('progressText');
                    
                    progressFill.style.width = progress + '%';
                    progressText.textContent = `กำลังประมวลผล... ${Math.round(progress)}%`;
                }, 300);
                
                try {
                    const response = await fetch('/api/process', {
                        method: 'POST',
                        body: formData
                    });
                    
                    const result = await response.json();
                    
                    clearInterval(progressInterval);
                    document.getElementById('progressFill').style.width = '100%';
                    document.getElementById('progressText').textContent = 'เสร็จสิ้น! 100%';
                    
                    setTimeout(() => {
                        document.getElementById('progress').style.display = 'none';
                        showResult(result, response.ok);
                    }, 1000);
                    
                } catch (error) {
                    clearInterval(progressInterval);
                    document.getElementById('progress').style.display = 'none';
                    showResult({message: 'เกิดข้อผิดพลาด: ' + error.message}, false);
                }
                
                document.getElementById('uploadBtn').disabled = false;
            }
            
            function showResult(result, success) {
                const resultDiv = document.getElementById('result');
                const contentDiv = document.getElementById('resultContent');
                const linksDiv = document.getElementById('downloadLinks');
                
                resultDiv.className = 'result ' + (success ? 'success' : 'error');
                resultDiv.style.display = 'block';
                
                if (success) {
                    contentDiv.innerHTML = `
                        <div style="text-align: center;">
                            <h3>✅ ประมวลผลสำเร็จ!</h3>
                            <p><strong>📊 จำนวนรายการ:</strong> ${result.total_records}</p>
                            <p><strong>⏱️ เวลาประมวลผล:</strong> ${result.processing_time?.toFixed(2)} วินาที</p>
                        </div>
                    `;
                    
                    linksDiv.innerHTML = `
                        <a href="/api/download/${result.job_id}/price" class="download-link">
                            📊 ดาวน์โหลด Price.xlsx
                        </a>
                        <a href="/api/download/${result.job_id}/type" class="download-link">
                            📋 ดาวน์โหลด Type.xlsx
                        </a>
                    `;
                } else {
                    contentDiv.innerHTML = `
                        <div style="text-align: center;">
                            <h3>❌ เกิดข้อผิดพลาด</h3>
                            <p>${result.message}</p>
                        </div>
                    `;
                    linksDiv.innerHTML = '';
                }
            }
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

@app.post("/api/process", response_model=ProcessingResult)
async def process_excel(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """Process uploaded Excel file"""
    
    # Validate file
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="ไฟล์ต้องเป็น .xlsx เท่านั้น")
    
    # Generate unique job ID
    job_id = str(uuid.uuid4())
    upload_path = UPLOAD_DIR / f"{job_id}_{file.filename}"
    
    try:
        # Save uploaded file
        with open(upload_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Process file
        start_time = datetime.now()
        extractor = ColorExtractor(job_id)
        result = extractor.process_file(str(upload_path))
        end_time = datetime.now()
        
        processing_time = (end_time - start_time).total_seconds()
        
        # Schedule cleanup
        background_tasks.add_task(cleanup_files, upload_path, delay_hours=1)
        
        return ProcessingResult(
            job_id=job_id,
            status="success",
            message="ประมวลผลสำเร็จ",
            price_file=f"Price_{job_id}.xlsx",
            type_file=f"Type_{job_id}.xlsx",
            processing_time=processing_time,
            total_records=result["total_records"]
        )
        
    except Exception as e:
        # Cleanup on error
        if upload_path.exists():
            upload_path.unlink()
        
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาด: {str(e)}")

@app.get("/api/download/{job_id}/{file_type}")
async def download_file(job_id: str, file_type: str):
    """Download processed files"""
    
    if file_type == "price":
        file_path = OUTPUT_DIR / f"Price_{job_id}.xlsx"
        filename = "Price.xlsx"
    elif file_type == "type":
        file_path = OUTPUT_DIR / f"Type_{job_id}.xlsx"
        filename = "Type.xlsx"
    else:
        raise HTTPException(status_code=400, detail="ประเภทไฟล์ไม่ถูกต้อง")
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="ไม่พบไฟล์")
    
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

async def cleanup_files(file_path: Path, delay_hours: int = 1):
    """Background task to cleanup files after delay"""
    import asyncio
    await asyncio.sleep(delay_hours * 3600)
    
    if file_path.exists():
        file_path.unlink()

@app.get("/api/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info"
    )
