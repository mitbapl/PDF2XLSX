import os
import subprocess
import secrets
from flask import Flask, render_template, request, send_file, flash
from werkzeug.utils import secure_filename
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Initialize Flask application
app = Flask(__name__)
app.secret_key = secrets.token_hex(16)  # Generate a secret key for flash messages

# Function to handle file uploads and conversion
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if file is present in the request
        if 'file' not in request.files:
            flash('No file part', 'error')
            return render_template('index.html')
        
        file = request.files['file']

        # If user does not select file, browser also submit an empty part without filename
        if file.filename == '':
            flash('No selected file', 'error')
            return render_template('index.html')

        # Securely save file to a temporary directory
        upload_dir = '/tmp'
        if not os.path.exists(upload_dir):
            os.makedirs(upload_dir)

        filename = secure_filename(file.filename)
        file_path = os.path.join(upload_dir, filename)
        file.save(file_path)
        logging.info(f"File saved to: {file_path}")

        # Call external script for conversion
        try:
            subprocess.run(['python', 'CIBILPDFXLSXv3.py', file_path], check=True)
        except subprocess.CalledProcessError as e:
            logging.error(f"Error running script: {e}")
            flash('Conversion failed', 'error')
            return render_template('index.html')

        # Check if conversion was successful and retrieve the converted file
        converted_file_path = os.path.join(upload_dir, 'extracted_data.xlsx')

        if not os.path.isfile(converted_file_path):
            logging.error(f"Converted file not found at: {converted_file_path}")
            flash('Conversion failed', 'error')
            return render_template('index.html')

        # Provide the converted file for download
        try:
            return send_file(converted_file_path, as_attachment=True,
                             cache_timeout=0)  # Disable caching
        except Exception as e:
            logging.error(f"Error sending file: {e}")
            flash('Error sending file', 'error')
            return render_template('index.html')

    # Render initial upload form for GET requests
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
