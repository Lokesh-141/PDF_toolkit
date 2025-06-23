# 🧰 Universal PDF Toolkit

A powerful Python-based utility that converts various file formats to PDF, merges PDFs, and performs deep file scans to generate forensic-style reports.

## ✨ Features

- 📄 **Convert to PDF**:
  - Office files: `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`
  - Image files: `.jpg`, `.jpeg`, `.png`
  - Source code files: `.py`, `.js`, `.html`, `.java`, etc., with syntax highlighting

- 📑 **Merge PDFs**:
  - Combine multiple PDFs into a single file

- 🧪 **Deep Scan**:
  - Generates a report containing file metadata and a hex preview

## 📁 Output Structure

Upon running, the tool auto-generates an `outputs/` folder containing:

outputs/
├── pdfs/ # Converted PDF files
├── logs/ # Tool logs (tool.log)
├── reports/ # Scan reports in PDF format
├── json/ # (Optional) Exported metadata in JSON

## 🚀 Usage

Run the script directly:

```bash
python your_script_name.py
Then follow the menu prompts:

mathematica
Copy
Edit
📌 Universal PDF Toolkit
1. Convert a File to PDF
2. Merge Multiple PDFs
3. Deep Scan & Generate Report
4. Exit
🛠 Requirements
Install dependencies using pip:

bash
Copy
Edit
pip install pillow pdfkit python-magic PyPDF2 pygments reportlab comtypes
Also ensure you have:

wkhtmltopdf installed for pdfkit to work

Windows system (for Office automation via comtypes)

🧠 Behind the Scenes
Office file conversion uses COM automation (via comtypes) — Windows-only

Code-to-PDF uses Pygments for syntax highlighting and wkhtmltopdf for rendering

Deep scan uses magic for MIME type detection and ReportLab for PDF report generation

📝 Logging
All operations and errors are logged to:

bash
Copy
Edit
outputs/logs/tool.log
📦 JSON Export
If ENABLE_JSON_EXPORT = True is set, metadata and reports can optionally be exported to outputs/json/.

⚠️ Platform Notes
The tool is designed for Windows, due to reliance on Office COM interfaces.

On macOS/Linux, Office file conversion will not function unless alternatives are implemented (e.g., LibreOffice via subprocess).

📃 License
MIT License — use freely and modify as needed.
