"""
PDF to Excel Converter Web Application

This Flask application provides a web interface for converting PDF files
to Excel spreadsheets. Users can upload PDF files through a web form,
and the application will process them and provide a download link for
the converted Excel file.
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, make_response
import os
import secrets
from io import BytesIO
from werkzeug.utils import secure_filename
from convert_clean import convert_pdf_to_excel

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

# Configuration
# Use /tmp for serverless environments like Vercel
UPLOAD_FOLDER = os.path.join('/tmp', 'uploads') if os.path.exists('/tmp') else 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50 MB

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    """
    Check if the uploaded file has an allowed extension.

    Args:
        filename (str): The name of the file to check.

    Returns:
        bool: True if the file extension is allowed, False otherwise.
    """
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    """
    Main route that handles both displaying the upload form and processing uploads.

    GET: Displays the upload form
    POST: Processes the uploaded PDF and converts it to Excel
    """
    if request.method == 'POST':
        # Check if file is present in request
        if 'file' not in request.files:
            flash('No file selected', 'error')
            return redirect(request.url)

        file = request.files['file']

        # Check if user selected a file
        if file.filename == '':
            flash('No file selected', 'error')
            return redirect(request.url)

        # Validate and process the file
        if file and allowed_file(file.filename):
            # Secure the filename and save it
            filename = secure_filename(file.filename)

            # Generate unique filenames to avoid conflicts
            base_name = os.path.splitext(filename)[0]
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{base_name}_{secrets.token_hex(8)}.pdf")
            excel_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{base_name}_{secrets.token_hex(8)}.xlsx")

            try:
                # Save the uploaded file
                file.save(pdf_path)

                # Convert PDF to Excel
                success, message = convert_pdf_to_excel(pdf_path, excel_path)

                # Clean up the uploaded PDF file
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)

                if success:
                    # Read the Excel file into memory for serverless environments
                    with open(excel_path, 'rb') as f:
                        excel_data = f.read()

                    # Clean up the Excel file
                    if os.path.exists(excel_path):
                        os.remove(excel_path)

                    # Create response with file data
                    response = make_response(send_file(
                        BytesIO(excel_data),
                        as_attachment=True,
                        download_name=f"{base_name}.xlsx",
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    ))
                    # Set a cookie to signal download completion
                    response.set_cookie('download_complete', 'true', max_age=10)
                    return response
                else:
                    flash(f'Conversion failed: {message}', 'error')
                    return redirect(request.url)

            except Exception as e:
                # Clean up files in case of error
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                if os.path.exists(excel_path):
                    os.remove(excel_path)

                flash(f'An error occurred: {str(e)}', 'error')
                return redirect(request.url)
        else:
            flash('Invalid file type. Please upload a PDF file.', 'error')
            return redirect(request.url)

    return render_template('index.html')

@app.route('/health')
def health():
    """
    Health check endpoint for monitoring.

    Returns:
        dict: Status information
    """
    return {'status': 'healthy', 'service': 'PDF to Excel Converter'}

@app.errorhandler(413)
def request_entity_too_large(error):
    """
    Handle file size exceeded error.

    Args:
        error: The error object

    Returns:
        Rendered template with error message
    """
    flash('File size exceeds the maximum limit of 50 MB', 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    # Run the application
    # For production, use a proper WSGI server like gunicorn
    app.run(debug=True, host='0.0.0.0', port=5000)
