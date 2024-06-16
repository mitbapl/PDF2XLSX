import os
import subprocess
import flask
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
    
    # Invoke your script with the file path as an argument
    subprocess.run(['python', 'CIBILPDFXLSXv3.py', file_path], timeout=9000)

    #Path to converted excel file
    converted_file_path = 'extracted_data.xlsx'

    # check xlsx file
    if not os.path.isfile(converted_file_path):
        print(f"Converted file not found at: {converted_file_path}")
        return "Conversion failed"

    #download xlsx
    response = send_file(converted_file_path, as_attachment=True)
    response.headers['Cache-control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'

    flash(f"Download successful: {converted_file_path}","Success")

    return response

if __name__ == '__main__':
    app.run(debug=True)
