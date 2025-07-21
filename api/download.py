from flask import Flask, send_file, jsonify
import os
import logging
from .utils import get_output_file_path

app = Flask(__name__)
logger = logging.getLogger(__name__)

def handler(request, context=None):
    """Vercel serverless handler for file downloads"""
    try:
        # Parse URL parameters from Vercel
        if hasattr(request, 'args'):
            params = request.args.get('params', '')
        elif 'params' in request:
            params = request['params']
        else:
            return jsonify({'message': 'ไม่พบพารามิเตอร์'}), 400
        
        # Parse job_id and file_type from params
        if '/' in params:
            parts = params.split('/')
            if len(parts) >= 2:
                job_id = parts[0]
                file_type = parts[1]
            else:
                return jsonify({'message': 'พารามิเตอร์ไม่ถูกต้อง'}), 400
        else:
            return jsonify({'message': 'พารามิเตอร์ไม่ถูกต้อง'}), 400
        
        if file_type not in ['price', 'type']:
            return jsonify({'message': 'ประเภทไฟล์ไม่ถูกต้อง'}), 400
        
        # Get file path
        file_path = get_output_file_path(job_id, file_type)
        
        if not os.path.exists(file_path):
            return jsonify({'message': 'ไม่พบไฟล์'}), 404
        
        # Set download name to simple format
        download_name = 'Price.xlsx' if file_type == 'price' else 'Type.xlsx'
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        logger.error(f"Download error: {e}")
        return jsonify({'message': f'เกิดข้อผิดพลาดในการดาวน์โหลด: {str(e)}'}), 500