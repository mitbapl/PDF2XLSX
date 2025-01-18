import os
import subprocess
import flask
import time
import secrets
from flask import Flask, render_template, request, send_file, flash

secret_key = secrets.token_hex(26)
app = flask.Flask(__name__)
app.secret_key = secret_key

@app.route('/')
def index():
    return render_template('/index.html')

@app.route('/convert', methods=['POST'])
def convert():
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
    print(file_path)

    timeout_seconds = 9000
    max_attempts = 3
    attempt = 0
    returncode = None

    while attempt < max_attempts and returncode is None:
        try:
            result = subprocess.run(['python', 'CIBILPDFXLSXv3.py', file_path], timeout=timeout_seconds, check=True, capture_output=True)
            returncode = result.returncode
            if returncode == 0:
                converted_file_path = 'extracted_data.xlsx'
                if os.path.exists(converted_file_path):
                    print(f"File successfully converted and saved at: {converted_file_path}")
                else:
                    print(f"Converted file not found at: {converted_file_path}. Return 'Conversion failed'.")
            else:
                print(f"Conversion failed with return code {returncode}.")
        except subprocess.TimeoutExpired:
            print(f"Conversion process timed out after {timeout_seconds} seconds. Retrying attempt {attempt + 1}...")
        except subprocess.CalledProcessError as e:
            returncode = e.returncode
            print(f"Conversion failed with error: {e}")
        except Exception as e:
            print(f"Unexpected error: {e}")
    
        attempt += 1
        if returncode is None:
            print(f"Retrying attempt {attempt}...")
            time.sleep(1)  # Optional: Add a short delay before retrying

    if returncode is None:
        print(f"Exceeded maximum retry attempts. Conversion failed.")

    #download xlsx
    response = send_file(converted_file_path, as_attachment=True)
    response.headers['Cache-control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'

    flash(f"Download successful: {converted_file_path}","Success")

    return response

if __name__ == '__main__':
    app.run(debug=True)
