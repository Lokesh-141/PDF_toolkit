# 🧰 Universal PDF Toolkit

**Universal PDF Toolkit** is a Python-based command-line utility that converts Office documents, images, and source code into PDFs, merges multiple PDFs, and performs deep forensic scans on files to generate detailed reports.

---

## ✨ Features

- 📄 **Convert to PDF**
  - Office documents: `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`
  - Image files: `.jpg`, `.jpeg`, `.png`
  - Source code: `.py`, `.js`, `.html`, `.java`, `.css`, `.json`, `.md`, etc., with syntax highlighting

- 📑 **Merge PDFs**
  - Combine multiple PDF files into a single document

- 🧪 **Deep Scan**
  - Analyze any file to extract metadata and a hex preview (first 256 bytes), then generate a report in PDF format

---

## 📂 Output Directory Structure

After running the tool, it creates an `outputs/` directory with organized folders:

outputs/
├── pdfs/ # All converted PDF files
├── logs/ # Logs (e.g., tool.log)
├── reports/ # Deep scan reports (PDF format)
├── json/ # Optional metadata and results (JSON format)

yaml
Copy
Edit

---

## 🚀 Getting Started

### ✅ Prerequisites

- **Python 3.7+**
- Works on **Windows** only (due to Office COM automation)
- `wkhtmltopdf` must be installed and added to your system PATH

### 📦 Install Required Python Libraries

```bash
pip install pillow pdfkit python-magic PyPDF2 pygments reportlab comtypes
🧭 How to Use
Run the script directly:

bash
Copy
Edit
python your_script_name.py
You'll see a menu:

mathematica
Copy
Edit
📌 Universal PDF Toolkit
1. Convert a File to PDF
2. Merge Multiple PDFs
3. Deep Scan & Generate Report
4. Exit
Example: Convert a Word document
vbnet
Copy
Edit
Choose an option (1–4): 1
Enter file path: C:\Users\Me\Documents\example.docx
✅ Saved: C:\Users\Me\Documents\example.pdf
🛠 Internals / How It Works
Feature	Implementation
Office to PDF	Uses comtypes for Windows COM automation
Images to PDF	Converts using Pillow
Code to PDF	Syntax highlighted via Pygments, rendered by pdfkit and wkhtmltopdf
Deep Scan Reporting	MIME detection via python-magic, PDF reports via ReportLab

📝 Logging
All actions are logged to:

bash
Copy
Edit
outputs/logs/tool.log
Useful for debugging failed conversions or deep scans.

🧾 Optional JSON Export
If you set ENABLE_JSON_EXPORT = True in the script config section, additional metadata and scan reports are exported as .json into:

bash
Copy
Edit
outputs/json/
⚠️ Platform Support
✅ Windows — Fully supported

❌ macOS / Linux — Limited support (Office conversion won't work unless replaced with LibreOffice or similar tools)

You can still use:

Image → PDF

Code → PDF

Deep scan
