# 💼 Universal PDF Toolkit
Universal PDF Toolkit is a Python-based command-line utility that converts Office documents, images, and source code into PDFs, merges multiple PDFs, and performs deep forensic scans on files to generate detailed reports.

---

## ✨ Features

- 📄 **Convert to PDF**
  - Office files (Excel only): `.xls`, `.xlsx` *(Windows-only)*
  - Code files with syntax highlighting: `.py`, `.js`, `.html`, `.css`, `.java`, etc.
  - Image files: `.jpg`, `.jpeg`, `.png`

- 📑 **Merge PDFs**
  - Combine multiple PDFs into one via menu

- 🧪 **Deep Scan**
  - Extracts file metadata:
    - File name
    - File size
    - Modified time
    - MIME type
    - First 256 bytes in hex
  - Outputs:
    - PDF Report
    - (Optional) JSON Metadata Export

---

## 🚀 Usage

Run the tool directly:

```
python universal_pdf_tool.py
```
You’ll see a menu like:

```
💼 Universal PDF Toolkit v2.0
1. Convert a File to PDF
2. Merge Multiple PDFs
3. Deep Scan & Generate Report
4. Exit
🛠 Requirements
```
📦 Install Python dependencies:
```
pip install pillow pdfkit python-magic PyPDF2 pygments reportlab comtypes
```
---

## 📦 External Dependency
💡 wkhtmltopdf is required for code-to-PDF conversion using pdfkit.

(Download it here)[https://wkhtmltopdf.org/downloads.html]

Ensure `wkhtmltopdf` is available in your system’s `PATH`.

📁 Folder Structure
The tool creates the following directory layout automatically:

```
outputs/
├── pdfs/       # Converted PDF files
├── logs/       # Log file: tool.log
├── reports/    # Deep scan PDF reports
├── json/       # (Optional) JSON metadata reports
```
---

##🔧 Technology Matrix
Task	Uses
Office to PDF	comtypes (Windows COM automation)
Image to PDF	Pillow
Code to PDF	Pygments + pdfkit + wkhtmltopdf
Scan report	python-magic + ReportLab

---

## 📝 Logging
All operations (success and errors) are logged to:
```
outputs/logs/tool.log
```

---

## 📦 Optional JSON Export
If you enable the following in the script: `ENABLE_JSON_EXPORT = True`
Then each deep scan will also save structured JSON `metadata` to:
```
outputs/json/
```

---

## ⚠️ Platform Support
**🪟 Windows:** Full support (including Office conversion)

**🐧 Linux / macOS:** Partial support (Office-to-PDF won't work without adaptation)

---

## 📃 License
This project is licensed under the [MIT License](LICENSE). You are free to use, modify, and share this project with proper attribution.

---

Built to automate your file conversions, PDF merges, and forensic file reporting.

---
