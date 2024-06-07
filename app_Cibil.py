import os
import subprocess
import flask
import secrets
import logging
from flask import Flask, render_template, request, send_file, flash

secret_key = secrets.token_hex(26)
app = flask.Flask(__name__)
app.secret_key = secret_key

# Configure logging
logging.basicConfig(filename='app.log', level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

@app.route('/')
def index():
    return render_template('/index.html')

@app.route('/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return "No file uploaded"
       
        file = request.files['file']
        if file.filename == '':
            return "No file selected"

        # Create temp dir
        upload_dir = '/tmp'
        if not os.path.exists(upload_dir):
            os.makedirs(upload_dir)

        # Save the file temporarily
        file_path = os.path.join(upload_dir, file.filename)
        file.save(file_path)

        # Invoke your script with the file path as an argument
        subprocess.run(['python', 'CIBILPDFXLSXv3.py', file_path])

        # Path to converted excel file
        converted_file_path = 'extracted_data.xlsx'

        # Check xlsx file
        if not os.path.isfile(converted_file_path):
            logging.error(f"Converted file not found at: {converted_file_path}")
            return "Conversion failed"

        # Download xlsx
        response = send_file(converted_file_path, as_attachment=True)
        response.headers['Cache-control'] = 'no-cache, no-store, must-revalidate'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'

        flash(f"Download successful: {converted_file_path}", "Success")

        return response
    except Exception as e:
        logging.error(f"An error occurred during conversion: {e}")
        return "Conversion failed. Please check the logs for details."

if __name__ == '__main__':
    app.run(debug=True)
