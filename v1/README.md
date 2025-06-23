# 🧰 Universal PDF Toolkit
Universal PDF Toolkit is a Python-based command-line utility that converts Office documents, images, and source code into PDFs, merges multiple PDFs, and performs deep forensic scans on files to generate detailed reports.

---

## ✨ Features
📄 **Convert to PDF**

- **Office documents:** `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`

- **Images:** `.jpg`, `.jpeg`, `.png`

- **Source code (with syntax highlighting):** `.py`, `.js`, `.html`, `.java`, `.css`, `.json`, `.md`, etc.

**📑 Merge PDFs**

- Combine multiple PDF files into a single document

**🧪 Deep**

- Analyze any file to extract metadata and a hex preview (first 256 bytes)

- Generate a clean PDF report from the scan

---

## 📂 Output Directory Structure
After running, the tool creates the following folders inside `outputs/`:

```
outputs/
├── pdfs/       # Converted PDF files
├── logs/       # Execution logs (tool.log)
├── reports/    # Deep scan reports (PDFs)
├── json/       # Optional: Metadata and scan results in JSON
```

## 🚀 Getting Started
**✅ Prerequisites**

- 🐍 Python 3.7 or higher

- 🪟 Windows OS (required for Office-to-PDF conversion)

- [wkhtmltopdf](https://wkhtmltopdf.org/downloads.html) installed and available in system PATH

---

## 📦 Install Required Packages
```
pip install pillow pdfkit python-magic PyPDF2 pygments reportlab comtypes
```

---

## 🧭 How to Use
Run the script:

```
python universal_pdf_tool.py
```
You'll see this menu:
```
📌 Universal PDF Toolkit
1. Convert a File to PDF
2. Merge Multiple PDFs
3. Deep Scan & Generate Report
4. Exit
```
---

## 🛠 Internals

| **Task**      | **Uses**                              |
| ------------- | ------------------------------------- |
| Office to PDF | `comtypes` (Windows COM automation)   |
| Image to PDF  | `Pillow`                              |
| Code to PDF   | `Pygments` + `pdfkit` + `wkhtmltopdf` |
| Scan Report   | `python-magic` + `ReportLab`          |

---

## 📝 Logging
Execution logs are stored in:
```
outputs/logs/tool.log
```

---

## 📦 JSON Export
If `ENABLE_JSON_EXPORT = True` in the script, scan and file `metadata` are also saved to:
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
