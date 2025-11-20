from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from pathlib import Path
import os
import zipfile
import tempfile
from datetime import datetime
from email_extractor import extract_emails_from_uploaded_files, extract_emails_from_text

app = Flask(__name__)

# Configuration
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max file size
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'temp_uploads')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'doc', 'zip'}

# Create upload folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """Render the main page"""
    return render_template('index.html')

@app.route('/api/extract', methods=['POST'])
def extract():
    """Handle file upload and email extraction"""
    try:
        # Check if files are in request
        if 'files' not in request.files:
            return jsonify({'success': False, 'error': 'No files provided'}), 400
        
        files = request.files.getlist('files')
        
        if not files or all(f.filename == '' for f in files):
            return jsonify({'success': False, 'error': 'No files selected'}), 400
        
        # Validate and save uploaded files
        uploaded_file_paths = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                # Add timestamp to avoid conflicts
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_")
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], timestamp + filename)
                file.save(filepath)
                uploaded_file_paths.append(filepath)
            else:
                return jsonify({
                    'success': False, 
                    'error': f'Invalid file type: {file.filename}. Allowed types: PDF, DOCX, DOC, ZIP'
                }), 400
        
        if not uploaded_file_paths:
            return jsonify({'success': False, 'error': 'No valid files to process'}), 400
        
        # Extract emails from uploaded files
        emails, stats, file_mapping = extract_emails_from_uploaded_files(uploaded_file_paths)
        
        # Clean up uploaded files
        for filepath in uploaded_file_paths:
            try:
                os.remove(filepath)
            except:
                pass
        
        # Remove duplicates and sort
        unique_emails = sorted(list(set(emails)))
        
        return jsonify({
            'success': True,
            'emails': unique_emails,
            'stats': stats,
            'file_mapping': file_mapping,
            'total_unique_emails': len(unique_emails)
        })
    
    except Exception as e:
        return jsonify({'success': False, 'error': f'Error: {str(e)}'}), 500

@app.route('/api/download', methods=['POST'])
def download():
    """Download extracted emails as text file"""
    try:
        data = request.get_json()
        emails = data.get('emails', [])
        
        if not emails:
            return jsonify({'success': False, 'error': 'No emails to download'}), 400
        
        # Create temporary text file
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt', encoding='utf-8') as f:
            for email in sorted(emails):
                f.write(email + '\n')
            temp_path = f.name
        
        # Send file
        filename = f"extracted_emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        response = send_file(temp_path, as_attachment=True, download_name=filename)
        
        # Clean up temp file after sending
        def cleanup(response):
            try:
                os.remove(temp_path)
            except:
                pass
            return response
        
        return cleanup(response)
    
    except Exception as e:
        return jsonify({'success': False, 'error': f'Error: {str(e)}'}), 500

@app.route('/api/download-detailed', methods=['POST'])
def download_detailed():
    """Download detailed extraction results with file mapping"""
    try:
        data = request.get_json()
        emails = data.get('emails', [])
        file_mapping = data.get('file_mapping', [])
        stats = data.get('stats', {})
        
        if not emails:
            return jsonify({'success': False, 'error': 'No emails to download'}), 400
        
        # Create detailed output
        content = f"# Email Extraction Results - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        content += f"# Total files processed: {stats.get('processed', 0)}\n"
        content += f"# Files with emails: {stats.get('with_emails', 0)}\n"
        content += f"# Total emails found: {len(emails)}\n"
        content += f"# Unique emails: {len(set(emails))}\n\n"
        
        # Email list
        content += "EXTRACTED EMAILS:\n"
        content += "=" * 50 + "\n"
        for email in sorted(set(emails)):
            content += f"{email}\n"
        
        # Detailed mapping
        if file_mapping:
            content += "\n\nDETAILED MAPPING (Email -> Source Files):\n"
            content += "=" * 50 + "\n"
            
            # Group emails by source file
            email_to_files = {}
            for mapping in file_mapping:
                email = mapping.get('email')
                if email not in email_to_files:
                    email_to_files[email] = []
                email_to_files[email].append(mapping)
            
            for email in sorted(email_to_files.keys()):
                content += f"\n{email}:\n"
                for file_info in email_to_files[email]:
                    content += f"  - {file_info.get('filename', 'Unknown')} ({file_info.get('file_type', 'Unknown')})\n"
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt', encoding='utf-8') as f:
            f.write(content)
            temp_path = f.name
        
        # Send file
        filename = f"email_extraction_detailed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        response = send_file(temp_path, as_attachment=True, download_name=filename)
        
        # Clean up temp file after sending
        def cleanup(response):
            try:
                os.remove(temp_path)
            except:
                pass
            return response
        
        return cleanup(response)
    
    except Exception as e:
        return jsonify({'success': False, 'error': f'Error: {str(e)}'}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
