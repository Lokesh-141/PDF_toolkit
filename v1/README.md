# 🧰 Universal PDF Toolkit

A powerful Python-based utility that converts various file formats to PDF, merges PDFs, and performs deep file scans to generate forensic-style reports.

---

## ✨ Features

- 📄 **Convert to PDF**:
  - Office documents: `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`
  - Images: `.jpg`, `.jpeg`, `.png`
  - Source code with syntax highlighting: `.py`, `.js`, `.html`, `.java`, `.css`, `.json`, etc.

- 📑 **Merge PDFs**:
  - Combine multiple PDFs into a single document

- 🧪 **Deep Scan**:
  - Generate a detailed PDF report including metadata and hex preview of any file

---

## 📁 Output Structure

The script creates an `outputs/` directory automatically with the following subfolders:

outputs/
├── pdfs/ # Converted PDF files
├── logs/ # Logs (tool.log)
├── reports/ # Deep scan reports (PDF)
├── json/ # Optional JSON metadata (if enabled)

yaml
Copy
Edit

---

## 🚀 Usage

Run the script directly:

```bash
python your_script_name.py
Follow the interactive menu:

mathematica
Copy
Edit
📌 Universal PDF Toolkit
1. Convert a File to PDF
2. Merge Multiple PDFs
3. Deep Scan & Generate Report
4. Exit
🛠 Requirements
Install required Python packages:

bash
Copy
Edit
pip install pillow pdfkit python-magic PyPDF2 pygments reportlab comtypes
Also ensure:

✅ wkhtmltopdf is installed (required for code-to-PDF conversion)

✅ You're using Windows, as Office automation via comtypes only works there

🧠 How It Works
📄 Office to PDF: Uses COM automation (comtypes) to convert Office files (Windows only)

🖼️ Image to PDF: Uses Pillow (PIL.Image)

💻 Code to PDF: Uses Pygments + pdfkit + wkhtmltopdf to generate syntax-highlighted PDFs

🔍 Deep Scan: Uses python-magic to detect MIME type and ReportLab to generate a PDF report

📝 Logging
All activity and errors are recorded in:

bash
Copy
Edit
outputs/logs/tool.log
📦 JSON Export
If ENABLE_JSON_EXPORT = True in the script, metadata and scan results can be exported to:

bash
outputs/json/
⚠️ Platform Compatibility
⚠️ This tool is designed for Windows only, due to reliance on COM interfaces for Office automation.

To use it on macOS/Linux:

Replace Office automation with alternatives like LibreOffice via subprocess

Keep image, code, and scan features — those are cross-platform

---

📃 License
MIT License — Free to use, modify, and distribute.

---

Made with 🛠️ and ☕ by a Developer Who Hates Manual File Conversion
---
