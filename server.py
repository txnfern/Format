from flask import Flask, request, jsonify, send_file, render_template_string
import os
import subprocess
import time
import uuid
import shutil
import logging
import json
from datetime import datetime
from werkzeug.utils import secure_filename
import sys
from pathlib import Path

# -------------------- Config & Globals --------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
MAX_FILE_SIZE = 25 * 1024 * 1024  # 25MB
ALLOWED_EXTENSIONS = {'xlsx', 'pdf'}

BASE_DIR = Path(__file__).resolve().parent
PYTHON = sys.executable                  # ใช้ python ของ .venv แน่นอน

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# -------------------- Helpers --------------------
def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def cleanup_old_files(hours: int = 1) -> None:
    """Clean up files older than `hours` hours"""
    try:
        current_time = time.time()
        expire = hours * 3600
        for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                if os.path.isfile(file_path):
                    if current_time - os.path.getctime(file_path) > expire:
                        os.remove(file_path)
                        logger.info(f"Cleaned up old file: {file_path}")
    except Exception as e:
        logger.error(f"Error during cleanup: {e}")

def load_html_template(template_name='original') -> str:
    template_files = {
        'original': 'index.html',
        'joint': 'index2.html',
        'format': 'index3.html'
    }
    try:
        filename = template_files.get(template_name)
        if filename and os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8') as f:
                return f.read()
        return f"""
        <html><body>
        <h1>Error: {filename} not found</h1>
        <p>Please make sure {filename} is in the same directory as server.py</p>
        <p><a href="/">← กลับหน้าหลัก</a></p>
        </body></html>
        """
    except Exception as e:
        return f"<html><body><h1>Error loading template: {e}</h1></body></html>"

# -------------------- Subprocess wrappers --------------------
def run_subprocess(cmd: list[str]) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    env["PYTHONNOUSERSITE"] = "1"  # กันไม่ให้ไปดึง package จาก user-site
    result = subprocess.run(
        cmd,
        cwd=BASE_DIR,
        env=env,
        capture_output=True,
        text=True
    )
    return result

# -------------------- Matrix Mode --------------------
def process_matrix_file_with_main_py(input_path: str, job_id: str, original_filename: str | None):
    try:
        start_time = time.time()

        cmd = [
            PYTHON, str(BASE_DIR / 'main.py'),
            '--input', input_path,
            '--job-id', job_id,
            '--output-dir', OUTPUT_FOLDER
        ]
        if original_filename:
            cmd += ['--original-filename', original_filename]

        result = run_subprocess(cmd)
        processing_time = time.time() - start_time

        # Clean input
        try:
            os.remove(input_path)
        except Exception:
            pass

        if result.returncode != 0:
            logger.error("Processing failed with main.py: %s", result.stderr)
            return None, f'เกิดข้อผิดพลาดในการประมวลผล: {result.stderr}'

        # หา JSON จาก stdout
        output_lines = result.stdout.strip().split('\n')
        json_output = None
        for line in reversed(output_lines):
            line = line.strip()
            if line.startswith('{') and line.endswith('}'):
                try:
                    json_output = json.loads(line)
                    break
                except json.JSONDecodeError:
                    pass

        if not json_output:
            return None, 'ไม่พบผลลัพธ์จาก main.py'

        price_file = os.path.join(OUTPUT_FOLDER, f'Price_{job_id}.xlsx')
        type_file = os.path.join(OUTPUT_FOLDER, f'Type_{job_id}.xlsx')

        if not os.path.exists(price_file):
            return None, 'ไม่พบไฟล์ Price ที่สร้างขึ้น'
        if not os.path.exists(type_file):
            return None, 'ไม่พบไฟล์ Type ที่สร้างขึ้น'

        return {
            'job_id': job_id,
            'total_records': json_output.get('total_records', 0),
            'price_records': json_output.get('total_records', 0),
            'type_records': json_output.get('processed_sheets', 0),
            'processed_sheets': json_output.get('processed_sheets', 0),
            'processing_time': processing_time,
            'message': 'ประมวลผลสำเร็จ',
            'skipped_sheets': json_output.get('skipped_sheets', []),
            'warnings': json_output.get('warnings', [])
        }, None

    except Exception as e:
        logger.exception("Unexpected error with main.py")
        return None, f'เกิดข้อผิดพลาดที่ไม่คาดคิด: {str(e)}'

# -------------------- Joint Mode --------------------
def process_joint_file_with_main2_py(input_path: str, job_id: str):
    try:
        start_time = time.time()

        cmd = [PYTHON, str(BASE_DIR / 'main2.py'), input_path, job_id]
        result = run_subprocess(cmd)
        processing_time = time.time() - start_time

        try:
            os.remove(input_path)
        except Exception:
            pass

        if result.returncode != 0:
            logger.error("Processing failed with main2.py: %s", result.stderr)
            return None, f'เกิดข้อผิดพลาดในการประมวลผล: {result.stderr}'

        output_lines = result.stdout.strip().split('\n')
        price_file = None
        type_file = None
        price_count = 0
        type_count = 0

        for line in output_lines:
            if line.startswith('MOVED_PRICE:'):
                price_file = line.split(':', 1)[1]
            elif line.startswith('MOVED_TYPE:'):
                type_file = line.split(':', 1)[1]
            elif line.startswith('PRICE_COUNT:'):
                price_count = int(line.split(':', 1)[1])
            elif line.startswith('TYPE_COUNT:'):
                type_count = int(line.split(':', 1)[1])

        if price_file and os.path.exists(price_file):
            shutil.move(price_file, os.path.join(OUTPUT_FOLDER, f'Price_{job_id}.xlsx'))
        if type_file and os.path.exists(type_file):
            shutil.move(type_file, os.path.join(OUTPUT_FOLDER, f'Type_{job_id}.xlsx'))

        return {
            'job_id': job_id,
            'total_records': price_count + type_count,
            'price_records': price_count,
            'type_records': type_count,
            'processed_sheets': 1,
            'processing_time': processing_time,
            'message': 'ประมวลผลสำเร็จ'
        }, None

    except Exception as e:
        logger.exception("Unexpected error with main2.py")
        return None, f'เกิดข้อผิดพลาดที่ไม่คาดคิด: {str(e)}'

# -------------------- PDF Format Mode --------------------
def process_pdf_file_with_main3_py(input_path: str, start_page: int, job_id: str):
    try:
        start_time = time.time()

        cmd = [PYTHON, str(BASE_DIR / 'main3.py'), input_path, str(start_page), job_id]
        result = run_subprocess(cmd)
        processing_time = time.time() - start_time

        try:
            os.remove(input_path)
        except Exception:
            pass

        if result.returncode != 0:
            logger.error("Processing failed with main3.py: %s", result.stderr)
            return None, f'เกิดข้อผิดพลาดในการประมวลผล: {result.stderr}'

        output_lines = result.stdout.strip().split('\n')
        json_output = None
        for line in reversed(output_lines):
            line = line.strip()
            if line.startswith('{') and line.endswith('}'):
                try:
                    json_output = json.loads(line)
                    break
                except json.JSONDecodeError:
                    pass

        if not json_output:
            return None, 'ไม่พบผลลัพธ์จาก main3.py'
        if 'error' in json_output:
            return None, json_output['error']

        return {
            'success': True,
            'data': json_output,
            'processing_time': processing_time,
            'message': f"ประมวลผลสำเร็จ: พบ {json_output.get('total_references', 0)} Reference Code และ {json_output.get('total_glass', 0)} GLASS"
        }, None

    except Exception as e:
        logger.exception("Unexpected error with main3.py")
        return None, f'เกิดข้อผิดพลาดที่ไม่คาดคิด: {str(e)}'

# -------------------- Routes --------------------
@app.route('/')
def index():
    cleanup_old_files()
    html_template = load_html_template('original')
    return render_template_string(html_template)

@app.route('/original')
@app.route('/matrix')
def original():
    cleanup_old_files()
    html_template = load_html_template('original')
    return render_template_string(html_template)

@app.route('/joint')
def joint():
    cleanup_old_files()
    html_template = load_html_template('joint')
    return render_template_string(html_template)

@app.route('/format')
def format_page():
    cleanup_old_files()
    html_template = load_html_template('format')
    return render_template_string(html_template)

@app.route('/api/process-matrix', methods=['POST'])
def process_matrix_file():
    try:
        if 'file' not in request.files:
            return jsonify({'message': 'ไม่พบไฟล์'}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({'message': 'ไม่ได้เลือกไฟล์'}), 400
        if not file.filename.lower().endswith('.xlsx'):
            return jsonify({'message': 'ประเภทไฟล์ไม่ถูกต้อง กรุณาอัพโหลดไฟล์ .xlsx'}), 400

        file_content = file.read()
        if len(file_content) > MAX_FILE_SIZE:
            return jsonify({'message': 'ไฟล์ใหญ่เกินไป (สูงสุด 25MB)'}), 400
        file.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        random_suffix = str(uuid.uuid4())[:8]
        job_id = f"{timestamp}_{random_suffix}"

        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, f'{job_id}_{filename}')
        file.save(input_path)

        logger.info(f"Processing Matrix file: {filename} with job_id: {job_id}")

        if not os.path.exists(BASE_DIR / 'main.py'):
            return jsonify({'message': 'ไม่พบไฟล์ main.py สำหรับ Matrix mode'}), 500

        result, error = process_matrix_file_with_main_py(input_path, job_id, file.filename)
        if error:
            return jsonify({'message': error}), 500

        logger.info(f"Matrix processing completed successfully for job_id: {job_id}")
        return jsonify(result)

    except Exception as e:
        logger.exception("Unexpected error in matrix processing")
        return jsonify({'message': f'เกิดข้อผิดพลาดที่ไม่คาดคิด: {str(e)}'}), 500

@app.route('/api/process-joint', methods=['POST'])
def process_joint_file():
    try:
        if 'file' not in request.files:
            return jsonify({'message': 'ไม่พบไฟล์'}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({'message': 'ไม่ได้เลือกไฟล์'}), 400
        if not file.filename.lower().endswith('.xlsx'):
            return jsonify({'message': 'ประเภทไฟล์ไม่ถูกต้อง กรุณาอัพโหลดไฟล์ .xlsx'}), 400

        file_content = file.read()
        if len(file_content) > MAX_FILE_SIZE:
            return jsonify({'message': 'ไฟล์ใหญ่เกินไป (สูงสุด 25MB)'}), 400
        file.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        random_suffix = str(uuid.uuid4())[:8]
        job_id = f"{timestamp}_{random_suffix}"

        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, f'{job_id}_{filename}')
        file.save(input_path)

        logger.info(f"Processing Joint file: {filename} with job_id: {job_id}")

        if not os.path.exists(BASE_DIR / 'main2.py'):
            return jsonify({'message': 'ไม่พบไฟล์ main2.py สำหรับ Joint mode'}), 500

        result, error = process_joint_file_with_main2_py(input_path, job_id)
        if error:
            return jsonify({'message': error}), 500

        logger.info(f"Joint processing completed successfully for job_id: {job_id}")
        return jsonify(result)

    except Exception as e:
        logger.exception("Unexpected error in joint processing")
        return jsonify({'message': f'เกิดข้อผิดพลาดที่ไม่คาดคิด: {str(e)}'}), 500

@app.route('/upload', methods=['POST'])
def upload_pdf():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'ไม่พบไฟล์'}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'ไม่ได้เลือกไฟล์'}), 400
        if not file.filename.lower().endswith('.pdf'):
            return jsonify({'error': 'กรุณาเลือกไฟล์ PDF เท่านั้น'}), 400

        file_content = file.read()
        if len(file_content) > MAX_FILE_SIZE:
            return jsonify({'error': 'ไฟล์ใหญ่เกินไป (สูงสุด 25MB)'}), 400
        file.seek(0)

        start_page = int(request.form.get('start_page', 3))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        random_suffix = str(uuid.uuid4())[:8]
        job_id = f"{timestamp}_{random_suffix}"

        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, f'{job_id}_{filename}')
        file.save(input_path)

        logger.info(f"Processing PDF file: {filename} with job_id: {job_id}, start_page: {start_page}")

        if not os.path.exists(BASE_DIR / 'main3.py'):
            return jsonify({'error': 'ไม่พบไฟล์ main3.py สำหรับ Format mode'}), 500

        result, error = process_pdf_file_with_main3_py(input_path, start_page, job_id)
        if error:
            return jsonify({'error': error}), 500

        logger.info(f"PDF processing completed successfully for job_id: {job_id}")
        return jsonify(result)

    except Exception as e:
        logger.exception("Unexpected error in PDF processing")
        return jsonify({'error': f'เกิดข้อผิดพลาดที่ไม่คาดคิด: {str(e)}'}), 500

@app.route('/download/<format>')
def download_pdf_results(format: str):
    try:
        if format == 'txt':
            txt_file = os.path.join(OUTPUT_FOLDER, 'pdf_results.txt')
            if not os.path.exists(txt_file):
                return jsonify({'error': 'ไม่พบไฟล์ผลลัพธ์'}), 404
            return send_file(txt_file, as_attachment=True, download_name='pdf_extraction_results.txt')
        elif format == 'json':
            json_file = os.path.join(OUTPUT_FOLDER, 'pdf_results.json')
            if not os.path.exists(json_file):
                return jsonify({'error': 'ไม่พบไฟล์ผลลัพธ์'}), 404
            return send_file(json_file, as_attachment=True, download_name='pdf_extraction_results.json')
        else:
            return jsonify({'error': 'รูปแบบไฟล์ไม่ถูกต้อง'}), 400
    except Exception as e:
        return jsonify({'error': f'เกิดข้อผิดพลาดในการดาวน์โหลด: {str(e)}'}), 500

@app.route('/api/download/<job_id>/<file_type>')
def download_file(job_id: str, file_type: str):
    try:
        if file_type == 'price':
            filename = f'Price_{job_id}.xlsx'
        elif file_type == 'type':
            filename = f'Type_{job_id}.xlsx'
        else:
            return jsonify({'message': 'ประเภทไฟล์ไม่ถูกต้อง'}), 400

        file_path = os.path.join(OUTPUT_FOLDER, filename)
        if not os.path.exists(file_path):
            return jsonify({'message': 'ไม่พบไฟล์'}), 404

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

@app.errorhandler(413)
def too_large(e):
    return jsonify({'message': 'ไฟล์ใหญ่เกินไป (สูงสุด 25MB)'}), 413

@app.route('/health')
def health_check():
    return jsonify({
        'status': 'healthy',
        'available_scripts': {
            'main.py': os.path.exists(BASE_DIR / 'main.py'),
            'main2.py': os.path.exists(BASE_DIR / 'main2.py'),
            'main3.py': os.path.exists(BASE_DIR / 'main3.py')
        },
        'available_templates': {
            'index.html': os.path.exists(BASE_DIR / 'index.html'),
            'index2.html': os.path.exists(BASE_DIR / 'index2.html'),
            'index3.html': os.path.exists(BASE_DIR / 'index3.html')
        }
    })

# -------------------- Run --------------------
if __name__ == '__main__':
    print("🚀 Starting Format Tostem Unified Server...")
    print("📁 Upload folder:", UPLOAD_FOLDER)
    print("📁 Output folder:", OUTPUT_FOLDER)
    print()
    print("🌐 Available routes:")
    print("   http://localhost:5000/          → Matrix Mode (index.html)")
    print("   http://localhost:5000/original  → Matrix Mode (index.html)")
    print("   http://localhost:5000/matrix    → Matrix Mode (index.html)")
    print("   http://localhost:5000/joint     → Joint Mode (index2.html)")
    print("   http://localhost:5000/format    → Format Mode - PDF Processing (index3.html)")
    print("   http://localhost:5000/health    → Health Check")
    print()
    print("📱 You can also access from other devices at: http://[your-ip]:5000")
    print("⚠️  Press Ctrl+C to stop the server")
    print()

    required_files = ['main.py', 'main2.py', 'main3.py', 'index.html', 'index2.html', 'index3.html']
    missing_files = [f for f in required_files if not os.path.exists(BASE_DIR / f)]
    if missing_files:
        print("⚠️  Warning: Missing files:")
        for f in missing_files:
            print(f"   - {f}")
        print()

    try:
        import flask  # noqa
        import pandas  # noqa
        import openpyxl  # noqa
        print("✅ Required packages for Matrix/Joint modes are installed")
        try:
            import pdfplumber  # noqa
            print("✅ pdfplumber is installed - PDF processing available")
        except ImportError:
            print("⚠️  pdfplumber not installed - PDF processing will not work")
            print("   Install with: pip install pdfplumber")
    except ImportError as e:
        print(f"❌ Missing required package: {e}")
        print("💡 Please install required packages:")
        print("   pip install flask pandas openpyxl pdfplumber")
        sys.exit(1)

    app.run(debug=True, host='0.0.0.0', port=5000)
