import os
import time
import logging
import tempfile
from pathlib import Path

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Use /tmp for Vercel's temporary storage
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'
MAX_FILE_SIZE = 25 * 1024 * 1024  # 25MB
ALLOWED_EXTENSIONS = {'xlsx'}

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def cleanup_old_files():
    """Clean up files older than 1 hour - optimized for serverless"""
    try:
        current_time = time.time()
        for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
            if os.path.exists(folder):
                for filename in os.listdir(folder):
                    file_path = os.path.join(folder, filename)
                    if os.path.isfile(file_path):
                        if current_time - os.path.getctime(file_path) > 3600:  # 1 hour
                            os.remove(file_path)
                            logger.info(f"Cleaned up old file: {file_path}")
    except Exception as e:
        logger.error(f"Error during cleanup: {e}")

def load_html_template(template_name='original'):
    """Load HTML template based on template name"""
    
    template_files = {
        'original': 'templates/index.html', 
        'joint': 'templates/index2.html'
    }
    
    try:
        filename = template_files.get(template_name)
        if filename and os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8') as f:
                return f.read()
        else:
            return f"""
            <html><body>
            <h1>Error: {filename} not found</h1>
            <p>Please make sure {filename} is in the templates directory</p>
            <p><a href="/">← กลับหน้าหลัก</a></p>
            </body></html>
            """
    except Exception as e:
        return f"<html><body><h1>Error loading template: {e}</h1></body></html>"

def get_temp_file_path(job_id, filename):
    """Get temporary file path for processing"""
    return os.path.join(UPLOAD_FOLDER, f'{job_id}_{filename}')

def get_output_file_path(job_id, file_type):
    """Get output file path"""
    filename = f'{file_type.title()}_{job_id}.xlsx'
    return os.path.join(OUTPUT_FOLDER, filename)