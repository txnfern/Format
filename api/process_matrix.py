from flask import Flask, request, jsonify
import os
import time
import uuid
import json
from datetime import datetime
from werkzeug.utils import secure_filename
import logging
from .utils import allowed_file, MAX_FILE_SIZE, get_temp_file_path, get_output_file_path
from .processors.matrix_processor import MatrixProcessor

app = Flask(__name__)
logger = logging.getLogger(__name__)

def handler(request_data, context=None):
    """Vercel serverless handler for matrix processing"""
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
        
        logger.info(f"Processing Matrix file: {filename} with job_id: {job_id}")
        
        # Process the file with timeout protection
        start_time = time.time()
        
        try:
            # Use the matrix processor
            processor = MatrixProcessor(job_id)
            result = processor.process_file(
                input_file=input_path,
                output_dir='/tmp/outputs',
                original_filename=file.filename
            )
            
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
            
            # Verify output files exist
            price_file = get_output_file_path(job_id, 'price')
            type_file = get_output_file_path(job_id, 'type')
            
            if not os.path.exists(price_file):
                return jsonify({'message': 'ไม่พบไฟล์ Price ที่สร้างขึ้น'}), 500
            if not os.path.exists(type_file):
                return jsonify({'message': 'ไม่พบไฟล์ Type ที่สร้างขึ้น'}), 500
            
            response_data = {
                'job_id': job_id,
                'total_records': result.get('total_records', 0),
                'price_records': result.get('total_records', 0),
                'type_records': result.get('processed_sheets', 0),
                'processed_sheets': result.get('processed_sheets', 0),
                'processing_time': processing_time,
                'message': 'ประมวลผลสำเร็จ',
                'skipped_sheets': result.get('skipped_sheets', []),
                'warnings': result.get('warnings', [])
            }
            
            logger.info(f"Matrix processing completed successfully for job_id: {job_id}")
            return jsonify(response_data)
            
        except TimeoutError:
            return jsonify({
                'message': 'การประมวลผลใช้เวลานานเกินไป กรุณาลองใหม่หรือใช้ไฟล์ที่เล็กกว่า'
            }), 408
            
    except Exception as e:
        logger.error(f"Unexpected error in matrix processing: {e}")
        return jsonify({'message': f'เกิดข้อผิดพลาดที่ไม่คาดคิด: {str(e)}'}), 500