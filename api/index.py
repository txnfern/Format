from flask import Flask, render_template_string
import os
from .utils import load_html_template, cleanup_old_files

app = Flask(__name__)

def handler(request, context=None):
    """Vercel serverless handler for index page"""
    if request.method == 'GET':
        cleanup_old_files()
        html_template = load_html_template('original')
        return render_template_string(html_template)
    
    return {'error': 'Method not allowed'}, 405

# For local testing
if __name__ == "__main__":
    app.run(debug=True)