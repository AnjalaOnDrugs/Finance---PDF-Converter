# PDF to Excel Converter Web Application

A web-based application that converts PDF files with hierarchical data structure into structured Excel spreadsheets.

## Features

- Simple and intuitive web interface
- Drag-and-drop file upload
- Automatic PDF to Excel conversion
- Direct download of converted files
- Support for files up to 50MB
- Responsive design

## Prerequisites

- Python 3.8 or higher
- pip (Python package installer)

## Installation

### 1. Clone or Download the Project

```bash
cd /path/to/your/project
```

### 2. Create a Virtual Environment (Recommended)

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

**Linux/Mac:**
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

## Running Locally

### Development Mode

To run the application in development mode:

```bash
python app.py
```

The application will be available at: `http://localhost:5000`

### Production Mode

For production deployment, use a WSGI server like Gunicorn:

**Linux/Mac:**
```bash
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

**Windows:**
```bash
waitress-serve --host=0.0.0.0 --port=5000 app:app
```

For Windows, you'll need to install waitress first:
```bash
pip install waitress
```

## Project Structure

```
.
├── app.py                  # Main Flask application
├── convert_clean.py        # PDF conversion logic
├── requirements.txt        # Python dependencies
├── templates/
│   └── index.html         # Web interface template
├── static/                # Static files (if needed)
└── uploads/               # Temporary upload directory
```

## Usage

1. Open your web browser and navigate to the application URL
2. Click the upload area or drag and drop a PDF file
3. Click "Convert to Excel" button
4. The converted Excel file will automatically download

## Deployment Options

### Option 1: Deploy on Heroku

1. Create a `Procfile`:
```
web: gunicorn app:app
```

2. Initialize git and push to Heroku:
```bash
git init
git add .
git commit -m "Initial commit"
heroku create your-app-name
git push heroku main
```

### Option 2: Deploy on PythonAnywhere

1. Upload your files to PythonAnywhere
2. Set up a web app with manual configuration
3. Set the source code directory
4. Edit the WSGI configuration file to point to your `app.py`
5. Reload the web app

### Option 3: Deploy on AWS/DigitalOcean/VPS

1. Set up your server with Python 3.8+
2. Clone your repository
3. Install dependencies
4. Set up Nginx as a reverse proxy
5. Use Gunicorn with systemd for process management

Example Nginx configuration:
```nginx
server {
    listen 80;
    server_name yourdomain.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

### Option 4: Deploy on Render

1. Create a new Web Service on Render
2. Connect your GitHub repository
3. Set the build command: `pip install -r requirements.txt`
4. Set the start command: `gunicorn app:app`
5. Deploy

### Option 5: Deploy with Docker

Create a `Dockerfile`:
```dockerfile
FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000

CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "app:app"]
```

Build and run:
```bash
docker build -t pdf-converter .
docker run -p 5000:5000 pdf-converter
```

## Configuration

You can modify the following settings in `app.py`:

- `MAX_FILE_SIZE`: Maximum upload file size (default: 50MB)
- `UPLOAD_FOLDER`: Directory for temporary file storage
- `ALLOWED_EXTENSIONS`: Allowed file extensions

## Security Considerations

- The application automatically cleans up uploaded files after processing
- Filenames are sanitized using `secure_filename()`
- File size limits are enforced
- Only PDF files are accepted

## Troubleshooting

### Issue: Module not found errors
**Solution:** Make sure you've activated your virtual environment and installed all dependencies:
```bash
pip install -r requirements.txt
```

### Issue: Permission denied when creating uploads folder
**Solution:** Ensure the application has write permissions in its directory:
```bash
chmod 755 uploads/
```

### Issue: Application not accessible from other machines
**Solution:** Make sure you're binding to `0.0.0.0` instead of `127.0.0.1`

## API Endpoints

- `GET /` - Main page with upload form
- `POST /` - Upload and convert PDF file
- `GET /health` - Health check endpoint

## License

This project is provided as-is for use in your organization.

## Support

For issues or questions, please contact your system administrator.
