# CV Email Extractor - Web UI

A modern web-based application for extracting email addresses from CV files (PDF, DOCX, DOC) with an intuitive user interface.

## 🎯 Features

- **Upload Multiple Files**: Drag and drop or click to upload CV files
- **Supported Formats**: PDF, DOCX, DOC, and ZIP archives
- **Real-time Extraction**: Extract emails instantly from uploaded files
- **Beautiful UI**: Modern, responsive interface with real-time feedback
- **Results Display**: View extracted emails directly in the browser
- **Download Options**: 
  - Simple email list (one per line)
  - Detailed report with file mapping and statistics
- **Email Deduplication**: Automatically removes duplicate emails
- **Statistics**: View extraction statistics (files processed, success rate, etc.)

## 📋 Requirements

- Python 3.7+
- Flask 3.0.0
- pdfplumber (for PDF extraction)
- python-docx (for DOCX extraction)
- pywin32 (for DOC extraction on Windows)

## 🚀 Setup Instructions

### 1. Create Virtual Environment

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the Application

```bash
python app.py
```

The application will start on `http://localhost:5000`

## 📖 Usage

1. **Open Browser**: Navigate to `http://localhost:5000`

2. **Upload Files**:
   - Click the upload area or drag & drop CV files
   - Select PDF, DOCX, DOC, or ZIP files
   - View selected files before extraction

3. **Extract Emails**:
   - Click "🔍 Extract Emails" button
   - Wait for extraction to complete
   - View results with statistics

4. **Download Results**:
   - **Simple Format**: Download cleaned email list (one per line)
   - **Detailed Report**: Download with file mapping and statistics

5. **Clear**: Click "Clear" to start a new extraction

## 📁 Project Structure

```
extract_cv/
├── app.py                 # Flask application with routes
├── email_extractor.py     # Email extraction logic
├── requirements.txt       # Python dependencies
├── README.md             # This file
└── templates/
    └── index.html        # Web UI
```

## 🔧 Configuration

### Upload Folder
- Temporary files are stored in `temp_uploads/` folder
- Files are automatically cleaned up after extraction

### File Size Limit
- Maximum file size: 500MB (adjustable in `app.py`)

### Allowed Extensions
- PDF, DOCX, DOC, ZIP

## 💡 How It Works

### Email Extraction Process

1. **File Upload**: Files are temporarily saved to `temp_uploads/`
2. **Text Extraction**: 
   - **PDF**: Uses `pdfplumber` for text extraction
   - **DOCX**: Uses `python-docx` for paragraph and table extraction
   - **DOC**: Uses `pywin32` for Word COM interface (Windows only)
3. **Email Detection**: Uses multiple regex patterns to find valid emails
4. **Cleaning**: Removes markdown, normalizes whitespace, validates format
5. **Deduplication**: Removes duplicate emails automatically
6. **Results**: Displays results and allows download

### Supported Email Patterns

- Standard emails: `john.doe@company.com`
- With underscores: `john_doe@company.com`
- With numbers: `john123@company.com`
- Malformed/split emails: `john @ company . com`

## ⚙️ Advanced Features

### ZIP File Support
- Extract ZIP files and process all CV files inside
- Recursive folder structure support

### File Mapping
- Track which emails came from which files
- Detailed report shows source file for each email

### Statistics
- Files processed
- Files with extractable text
- Files with @ symbols
- Files with valid emails
- Success rate by file type

## 🐛 Troubleshooting

### Port Already in Use
```bash
python app.py --port 5001
```

### pywin32 Issues (Windows)
```bash
python -m pip install --upgrade pywin32
```

### ZIP Extraction Issues
- Ensure ZIP files are not corrupted
- Check file permissions

### PDF Not Extracting
- Try updating pdfplumber: `pip install --upgrade pdfplumber`
- Some PDFs may have protection or unusual encoding

## 📝 Notes

- The application is designed for local use
- For production use, implement proper security measures
- Uploaded files are automatically deleted after extraction
- Emails are extracted using advanced pattern matching and validation
- The UI is fully responsive and works on mobile devices

## 📧 Output Files

### Simple Download
- Format: Plain text, one email per line
- Filename: `extracted_emails_YYYYMMDD_HHMMSS.txt`

### Detailed Download
- Format: Detailed report with statistics and file mapping
- Filename: `email_extraction_detailed_YYYYMMDD_HHMMSS.txt`
- Includes:
  - Extraction timestamp
  - Number of files processed
  - Files with emails
  - Unique emails found
  - Email to source file mapping

## 🎨 UI Features

- **Dark Mode Compatible**: Works with system preferences
- **Drag & Drop**: Intuitive file upload
- **Real-time Feedback**: Status messages and progress indicators
- **Responsive Design**: Works on desktop, tablet, and mobile
- **Email Copy**: Click any email to copy to clipboard
- **Statistics Dashboard**: Visual breakdown of extraction results

## 🔒 Security

- Files are temporarily stored and automatically deleted
- No files are permanently saved on the server
- Maximum file size limit prevents abuse
- Input validation on all uploads

## 📞 Support

For issues or questions, check the terminal output for detailed error messages.

---

**Version**: 1.0.0  
**Last Updated**: 2025
