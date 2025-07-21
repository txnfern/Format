from flask import Flask, jsonify
import os

app = Flask(__name__)

def handler(request, context=None):
    """Vercel serverless handler for health check"""
    return jsonify({
        'status': 'healthy',
        'environment': 'vercel',
        'available_processors': {
            'matrix_processor': os.path.exists('api/processors/matrix_processor.py'),
            'joint_processor': os.path.exists('api/processors/joint_processor.py')
        },
        'available_templates': {
            'index.html': os.path.exists('templates/index.html'),
            'index2.html': os.path.exists('templates/index2.html')
        }
    })