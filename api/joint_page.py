from flask import Flask, render_template_string
from .utils import load_html_template, cleanup_old_files

app = Flask(__name__)

def handler(request, context=None):
    """Vercel serverless handler for joint page"""
    if request.method == 'GET':
        cleanup_old_files()
        html_template = load_html_template('joint')
        return render_template_string(html_template)
    
    return {'error': 'Method not allowed'}, 405