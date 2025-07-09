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
from datetime import datetime
from pathlib import Path
from typing import Optional
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
    processed_sheets: Optional[int] = None
    skipped_sheets: Optional[list] = None
    warnings: Optional[list] = None

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
        """Convert ARGB color to RGB hex format - แก้ไขให้อ่านสีที่ถูกต้อง"""
        if not fill:
            return "FFFFFF"
        
        # ตรวจสอบ patternType ก่อน - เฉพาะ solid fill เท่านั้น
        if hasattr(fill, 'patternType') and fill.patternType:
            pattern_value = fill.patternType.value if hasattr(fill.patternType, 'value') else str(fill.patternType)
            # ถ้าไม่ใช่ solid pattern ให้ถือว่าไม่มีสี
            if pattern_value != 'solid':
                return "FFFFFF"
        else:
            # ถ้าไม่มี patternType ให้ถือว่าไม่มีสี
            return "FFFFFF"
        
        # รายการสีที่ไม่ต้องการ (Excel theme colors) - ไม่รวม 92CDDC
        excluded_colors = [
            "00000000",  # สีใส
            "DCE6F1",    # Excel theme light blue
            "B4C6E7",    # Excel theme blue
            "A9D08E",    # Excel theme green
            "FFE699",    # Excel theme yellow
            "F4B183",    # Excel theme orange
            "F2F2F2",    # เทาอ่อน
            "E6E6E6",    # เทาอ่อน
            "D9D9D9",    # เทากลาง
        ]
        
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
        
        # ตรวจสอบว่าเป็นสีที่ไม่ต้องการหรือไม่
        if color_found in excluded_colors:
            return "FFFFFF"
        
        return color_found if color_found else "FFFFFF"

    def find_thickness_matrix(self, ws, raw, thickness_mm):
        """Find matrix with specific thickness label and its own header"""
        thickness_patterns = [
            rf"Thk\.{thickness_mm}\s*mm",
            rf"{thickness_mm}\s*mm",
            rf"Thickness\s*{thickness_mm}",
            rf"หนา\s*{thickness_mm}"
        ]
        
        # หา thickness header โดยตรง
        for r in range(raw.shape[0]):
            for c in range(raw.shape[1]):
                cell_val = str(raw.iat[r, c]).strip() if raw.iat[r, c] is not None else ""
                for pattern in thickness_patterns:
                    if re.search(pattern, cell_val, re.IGNORECASE):
                        print(f"   ✅ พบ {thickness_mm}mm matrix ที่ row={r+1}, col={c+1}")
                        return r, c
        
        return None, None

    def find_5mm_matrix(self, ws, raw):
        """Find 5mm matrix as the main reference matrix"""
        # หาจาก 5mm header
        for r in range(raw.shape[0]):
            for c in range(raw.shape[1]):
                cell_val = str(raw.iat[r, c]).strip() if raw.iat[r, c] is not None else ""
                # หา 5mm header
                if re.search(r"\b5\s*mm\b", cell_val, re.IGNORECASE):
                    print(f"   ✅ พบ 5mm matrix (main) ที่ row={r+1}, col={c+1}")
                    return r, c
        
        # ถ้าไม่พบ 5mm header ให้หา h/w header แทน (backward compatibility)
        for r in range(raw.shape[0]):
            for c in range(raw.shape[1]):
                if raw.iat[r, c] is None:
                    continue
                if isinstance(raw.iat[r, c], str):
                    if re.search(r"\bh\s*/\s*w\b", raw.iat[r, c], re.IGNORECASE):
                        print(f"   ✅ พบ h/w matrix (fallback) ที่ row={r+1}, col={c+1}")
                        return r, c
        
        return None, None

    def read_color_matrix_with_auto_offset(self, ws, raw, hr, hc, widths, heights, matrix_name=""):
        """อ่านสีโดยลอง offset หลายแบบและเลือกที่ดีที่สุด"""
        print(f"     🔍 {matrix_name}: ลอง offset หลายแบบ...")
        
        best_colors = {}
        max_valid_colors = 0
        best_offset = (2, 2)
        
        # ลอง offset ต่างๆ
        for row_offset in [1, 2, 3]:
            for col_offset in [1, 2, 3]:
                test_colors = {}
                valid_count = 0
                
                # ทดสอบเฉพาะ 4 เซลล์แรก
                for i_h in range(min(2, len(heights))):
                    for i_w in range(min(2, len(widths))):
                        h, w = heights[i_h], widths[i_w]
                        try:
                            excel_row = hr + row_offset + i_h
                            excel_col = hc + col_offset + i_w
                            
                            # ตรวจสอบว่าอยู่ในขอบเขต
                            if excel_row <= ws.max_row and excel_col <= ws.max_column:
                                cell = ws.cell(row=excel_row, column=excel_col)
                                color = self.normalize_rgb(cell.fill)
                                test_colors[(h, w)] = color
                                if color != "FFFFFF":
                                    valid_count += 1
                            else:
                                test_colors[(h, w)] = "FFFFFF"
                        except:
                            test_colors[(h, w)] = "FFFFFF"
                
                # ถ้า offset นี้ให้ผลดีกว่า
                if valid_count > max_valid_colors:
                    max_valid_colors = valid_count
                    best_offset = (row_offset, col_offset)
                    print(f"       🎯 offset +{row_offset},+{col_offset}: {valid_count} สี")
        
        # ใช้ offset ที่ดีที่สุดเพื่ออ่านทั้ง matrix
        row_offset, col_offset = best_offset
        print(f"     ✅ ใช้ offset: +{row_offset},+{col_offset}")
        
        for i_h, h in enumerate(heights):
            for i_w, w in enumerate(widths):
                try:
                    excel_row = hr + row_offset + i_h
                    excel_col = hc + col_offset + i_w
                    
                    if excel_row <= ws.max_row and excel_col <= ws.max_column:
                        cell = ws.cell(row=excel_row, column=excel_col)
                        color = self.normalize_rgb(cell.fill)
                        best_colors[(h, w)] = color
                    else:
                        best_colors[(h, w)] = "FFFFFF"
                except:
                    best_colors[(h, w)] = "FFFFFF"
        
        return best_colors

    def read_color_matrix(self, ws, raw, hr, hc, widths, heights):
        """Read colors from matrix - ใช้ offset มาตรฐาน"""
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
                    color_map[(h, w)] = "FFFFFF"
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
            
            # Track processing results
            processed_sheets = 0
            skipped_sheets = []
            warnings = []
            
            for sheet in xls.sheet_names:
                # ตรวจสอบ Sheet สารบัญ
                if sheet.strip().lower() == "สารบัญ":
                    skipped_sheets.append({"sheet": sheet, "reason": "ข้าม Sheet สารบัญ"})
                    print(f"   ⚠️ ข้าม Sheet: {sheet} (สารบัญ)")
                    continue
                
                print(f"\n🔍 ประมวลผล Sheet: {sheet}")
                
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
                
                # Find 5mm matrix as main reference
                hr, hc = self.find_5mm_matrix(ws, raw)
                
                if hr is None or hc is None:
                    error_msg = "ไม่พบ 5mm matrix หรือ h/w header"
                    print(f"   ❌ {error_msg} ใน {sheet}")
                    skipped_sheets.append({"sheet": sheet, "reason": error_msg})
                    continue
                
                # Read widths and heights from 5mm matrix
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
                    error_msg = "ไม่พบ dimensions (ความกว้าง/ความสูง)"
                    print(f"   ❌ {error_msg} ใน {sheet}")
                    skipped_sheets.append({"sheet": sheet, "reason": error_msg})
                    continue
                
                print(f"   📊 Dimensions: {len(heights)} heights x {len(widths)} widths")
                
                # Read colors from matrices
                color_5mm = {}
                color_6mm = {}
                color_8mm = {}
                
                # อ่าน 5mm จาก main matrix (ใช้ offset มาตรฐาน)
                color_5mm = self.read_color_matrix(ws, raw, hr, hc, widths, heights)
                print(f"   🎨 5mm (main matrix): {len(color_5mm)} colors")
                
                # หา 6mm และ 8mm matrix โดยหา header ของแต่ละ thickness เอง
                thickness_warnings = []
                for thickness in [6, 8]:
                    hr_thick, hc_thick = self.find_thickness_matrix(ws, raw, thickness)
                    if hr_thick is not None:
                        # อ่าน dimensions สำหรับ thickness matrix
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
                        
                        # ตรวจสอบว่า dimensions ตรงกับ main matrix หรือไม่
                        if widths_thick == widths and heights_thick == heights:
                            print(f"     ✅ {thickness}mm dimensions ตรงกับ main matrix")
                            # ใช้ auto-offset detection
                            colors = self.read_color_matrix_with_auto_offset(
                                ws, raw, hr_thick, hc_thick, widths, heights, f"{thickness}mm"
                            )
                        elif widths_thick and heights_thick:
                            warning_msg = f"Sheet {sheet}: {thickness}mm dimensions ต่างจาก main matrix"
                            print(f"     ⚠️ {warning_msg}")
                            print(f"       Main: {len(heights)}x{len(widths)}")
                            print(f"       {thickness}mm: {len(heights_thick)}x{len(widths_thick)}")
                            thickness_warnings.append(warning_msg)
                            # ใช้ dimensions ของ thickness matrix เอง
                            colors = self.read_color_matrix_with_auto_offset(
                                ws, raw, hr_thick, hc_thick, widths_thick, heights_thick, f"{thickness}mm"
                            )
                        else:
                            warning_msg = f"Sheet {sheet}: ไม่พบ dimensions สำหรับ {thickness}mm"
                            print(f"     ❌ {warning_msg}")
                            thickness_warnings.append(warning_msg)
                            colors = {}
                        
                        if thickness == 6:
                            color_6mm = colors
                        elif thickness == 8:
                            color_8mm = colors
                        
                        print(f"   🎨 {thickness}mm: {len(colors)} colors อ่านได้")
                    else:
                        warning_msg = f"Sheet {sheet}: ไม่พบ {thickness}mm matrix"
                        print(f"   ❌ {warning_msg}")
                        thickness_warnings.append(warning_msg)
                        if thickness == 6:
                            color_6mm = {}
                        elif thickness == 8:
                            color_8mm = {}
                
                # เพิ่ม warnings สำหรับ thickness matrices
                warnings.extend(thickness_warnings)
                
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
                sheet_price_count = 0
                for i_h, h in enumerate(heights):
                    for i_w, w in enumerate(widths):
                        # อ่านราคาจาก 5mm matrix
                        raw_price = raw.iat[hr + 1 + i_h, hc + 1 + i_w]
                        p = self.to_number(raw_price)
                        if p is None:
                            continue
                        
                        # ตรวจสอบว่ามีข้อมูล thickness หรือไม่
                        has_thickness_data = bool(color_5mm or color_6mm or color_8mm)
                        
                        if has_thickness_data:
                            # ถ้ามีข้อมูล thickness ให้ใช้สีจาก thickness matrix
                            color_5 = color_5mm.get((h, w), "FFFFFF")
                            color_6 = color_6mm.get((h, w), "FFFFFF")
                            color_8 = color_8mm.get((h, w), "FFFFFF")
                        else:
                            # ถ้าไม่มีข้อมูล thickness ให้อ่านสีจาก main matrix และแสดงที่ช่อง 5mm
                            try:
                                excel_row = hr + 1 + i_h
                                excel_col = hc + 1 + i_w
                                cell = ws.cell(row=excel_row, column=excel_col)
                                main_color = self.normalize_rgb(cell.fill)
                                color_5 = main_color
                                color_6 = "FFFFFF"  # ไม่มีข้อมูล
                                color_8 = "FFFFFF"  # ไม่มีข้อมูล
                            except Exception:
                                color_5 = "FFFFFF"
                                color_6 = "FFFFFF"
                                color_8 = "FFFFFF"
                        
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
                        sheet_price_count += 1
                
                processed_sheets += 1
                print(f"   ✅ สร้าง {sheet_price_count} price records สำหรับ {sheet}")
            
            # Save output files
            price_file = OUTPUT_DIR / f"Price_{self.job_id}.xlsx"
            type_file = OUTPUT_DIR / f"Type_{self.job_id}.xlsx"
            
            pd.DataFrame(price_rows).to_excel(price_file, index=False)
            pd.DataFrame(type_rows).to_excel(type_file, index=False)
            
            print(f"\n✅ เสร็จสิ้น: {len(price_rows)} price records, {len(type_rows)} type records")
            
            return {
                "price_file": str(price_file),
                "type_file": str(type_file),
                "total_records": len(price_rows),
                "processed_sheets": processed_sheets,
                "skipped_sheets": skipped_sheets,
                "warnings": warnings
            }
            
        except Exception as e:
            print(f"❌ Error: {str(e)}")
            raise Exception(f"Processing failed: {str(e)}")

# API Endpoints
@app.get("/", response_class=HTMLResponse)
async def root():
    """Serve the main HTML interface from external file"""
    html_file = Path("index.html")
    if html_file.exists():
        return FileResponse("index.html", media_type="text/html")
    else:
        # Fallback error message
        return HTMLResponse(
            content="""
            <html>
                <body style="font-family: Arial; text-align: center; padding: 50px;">
                    <h1>❌ ไม่พบไฟล์ index.html</h1>
                    <p>กรุณาตรวจสอบว่าไฟล์ <code>index.html</code> อยู่ในไดเรกทอรีเดียวกันกับ <code>main.py</code></p>
                    <p>โครงสร้างไฟล์ที่ถูกต้อง:</p>
                    <pre style="text-align: left; background: #f5f5f5; padding: 15px; border-radius: 5px; display: inline-block;">
project/
├── main.py
├── index.html
├── uploads/
└── outputs/
                    </pre>
                </body>
            </html>
            """,
            status_code=404
        )

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
        
        # Process file - ส่ง original filename ไปด้วย
        start_time = datetime.now()
        extractor = ColorExtractor(job_id)
        result = extractor.process_file(str(upload_path), file.filename)
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
            total_records=result["total_records"],
            processed_sheets=result.get("processed_sheets", 0),
            skipped_sheets=result.get("skipped_sheets", []),
            warnings=result.get("warnings", [])
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