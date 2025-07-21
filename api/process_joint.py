from flask import Flask, request, jsonify
import os
import time
import uuid
import shutil
from datetime import datetime
from werkzeug.utils import secure_filename
import logging
from .utils import allowed_file, MAX_FILE_SIZE, get_temp_file_path, get_output_file_path
from .processors.joint_processor import JointProcessor

app = Flask(__name__)
logger = logging.getLogger(__name__)

def handler(request_data, context=None):
    """Vercel serverless handler for joint processing"""
    try:
        # Parse request data
        if hasattr(request_data, 'files'):
            files = request_data.files
        elif isinstance(request_data, dict) and 'files' in request_data:
            files = request_data['files']
        else:
            return jsonify({'message': 'ไม่พบไฟล์'}), 400
        
        # Check if file was uploaded
        if 'file' not in files:
            return jsonify({'message': 'ไม่พบไฟล์'}), 400
        
        file = files['file']
        if not file or file.filename == '':
            return jsonify({'message': 'ไม่ได้เลือกไฟล์'}), 400
        
        # Validate file
        if not allowed_file(file.filename):
            return jsonify({'message': 'ประเภทไฟล์ไม่ถูกต้อง กรุณาอัพโหลดไฟล์ .xlsx'}), 400
        
        # Check file size
        file_content = file.read()
        if len(file_content) > MAX_FILE_SIZE:
            return jsonify({'message': 'ไฟล์ใหญ่เกินไป (สูงสุด 25MB)'}), 400
        file.seek(0)  # Reset file pointer
        
        # Generate job ID with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        random_suffix = str(uuid.uuid4())[:8]
        job_id = f"{timestamp}_{random_suffix}"
        
        # Save uploaded file to temp location
        filename = secure_filename(file.filename)
        input_path = get_temp_file_path(job_id, filename)
        
        # Ensure temp directory exists
        os.makedirs(os.path.dirname(input_path), exist_ok=True)
        file.save(input_path)
        
        logger.info(f"Processing Joint file: {filename} with job_id: {job_id}")
        
        # Process the file with timeout protection
        start_time = time.time()
        
        try:
            # Use the joint processor
            processor = JointProcessor(input_path, file.filename)
            success = processor.process(job_id)
            
            # Calculate processing time
            processing_time = time.time() - start_time
            
            # Clean up input file
            try:
                os.remove(input_path)
            except:
                pass
            
            # Check timeout (8 seconds to allow for response time)
            if processing_time > 8:
                return jsonify({
                    'message': 'การประมวลผลใช้เวลานานเกินไป กรุณาลองใหม่หรือใช้ไฟล์ที่เล็กกว่า'
                }), 408
            
            if not success:
                return jsonify({'message': 'เกิดข้อผิดพลาดในการประมวลผล'}), 500
            
            # Move files to output directory with job_id
            price_count = 0
            type_count = 0
            
            try:
                price_temp = 'Price.xlsx'
                type_temp = 'Type.xlsx'
                
                if os.path.exists(price_temp):
                    import pandas as pd
                    price_count = len(pd.read_excel(price_temp))
                    shutil.move(price_temp, get_output_file_path(job_id, 'price'))
                    
                if os.path.exists(type_temp):
                    import pandas as pd
                    type_count = len(pd.read_excel(type_temp))
                    shutil.move(type_temp, get_output_file_path(job_id, 'type'))
                    
            except Exception as e:
                logger.error(f"Error moving files: {e}")
                return jsonify({'message': f'เกิดข้อผิดพลาดในการจัดการไฟล์: {str(e)}'}), 500
            
            response_data = {
                'job_id': job_id,
                'total_records': price_count + type_count,
                'price_records': price_count,
                'type_records': type_count,
                'processed_sheets': 1,
                'processing_time': processing_time,
                'message': 'ประมวลผลสำเร็จ'
            }
            
            logger.info(f"Joint processing completed successfully for job_id: {job_id}")
            return jsonify(response_data)
            
        except TimeoutError:
            return jsonify({
                'message': 'การประมวลผลใช้เวลานานเกินไป กรุณาลองใหม่หรือใช้ไฟล์ที่เล็กกว่า'
            }), 408
            
    except Exception as e:
        logger.error(f"Unexpected error in joint processing: {e}")
        return jsonify({'message': f'เกิดข้อผิดพลาดที่ไม่คาดคิด: {str(e)}'}), 500