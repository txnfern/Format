from flask import Flask, request, jsonify
from main import ColorExtractor  # import logic จาก main.py (ย้ายมาวางด้วย)
import io, uuid

app = Flask(__name__)

@app.route('/', methods=['POST'])
def handler():
    # อ่านไฟล์จาก request
    f = request.files.get('file')
    if not f or not f.filename.endswith('.xlsx'):
        return jsonify({'message': 'ต้องอัพโหลด .xlsx'}), 400

    data = f.read()                   # อ่านเป็น bytes
    job_id = uuid.uuid4().hex[:8]
    # แปลงเป็น BytesIO เพื่อให้ openpyxl อ่านได้
    from io import BytesIO
    bio = BytesIO(data)

    # เรียกใช้ logic เดิม
    extractor = ColorExtractor(job_id)
    success = extractor.process_file(bio, '/tmp', original_filename=f.filename)

    if not success:
        return jsonify({'message': 'ประมวลผลไม่สำเร็จ'}), 500

    # อ่านผลลัพธ์จาก /tmp/Price_{job_id}.xlsx, Type_{job_id}.xlsx
    # แล้วส่งกลับเป็นลิงก์ดาวน์โหลด หรือเป็นไฟล์แนบ
    return jsonify({
      'price_url': f'/api/download/{job_id}/price',
      'type_url' : f'/api/download/{job_id}/type'
    })
