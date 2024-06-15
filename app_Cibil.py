import os
import subprocess
import secrets
from flask import Flask, render_template, request, send_file, flash

app = Flask(__name__)
secret_key = secrets.token_hex(26)
app.secret_key = secret_key

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file uploaded"
   
    file = request.files['file']
    if file.filename == '':
        return "No file selected"

    # Create a secure temporary directory
    import tempfile
    upload_dir = tempfile.mkdtemp()

    try:
        # Save the file temporarily
        file_path = os.path.join(upload_dir, file.filename)
        file.save(file_path)
        print(f"File saved to: {file_path}")

        # Invoke your script with the file path as an argument
        subprocess.run(['python', 'CIBILPDFXLSXv3.py', file_path], check=True)
        print("Conversion script executed successfully.")

        # Path to converted excel file
        converted_file_path = os.path.join(upload_dir, 'extracted_data.xlsx')
        print(f"Converted file path: {converted_file_path}")

        # Check if xlsx file exists
        if not os.path.isfile(converted_file_path):
            print(f"Converted file not found at: {converted_file_path}")
            return "Conversion failed: No converted file found"

        # Download xlsx
        response = send_file(converted_file_path, as_attachment=True)
        response.headers['Cache-control'] = 'no-cache, no-store, must-revalidate'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'

        flash(f"Download successful: {file.filename}", "success")
        return response

    except subprocess.CalledProcessError as e:
        print(f"Error executing conversion script: {e}")
        return f"Conversion failed: {e}"

    except Exception as e:
        print(f"Error: {e}")
        return f"Error: {e}"

    finally:
        # Clean up temporary directory
        if os.path.exists(upload_dir):
            import shutil
            shutil.rmtree(upload_dir)

if __name__ == '__main__':
    app.run(debug=True)
