# 🧰 Universal PDF Toolkit

A versatile command-line utility to convert files to PDF, merge PDFs, and perform deep scans on files to generate forensic-style reports. Supports Office documents, images, and code files with syntax highlighting.

---

## 🚀 Features

- **📄 Convert to PDF**
  - Word: `.doc`, `.docx`
  - Excel: `.xls`, `.xlsx`
  - PowerPoint: `.ppt`, `.pptx`
  - Images: `.jpg`, `.jpeg`, `.png`
  - Code/Text: `.py`, `.js`, `.html`, `.css`, `.md`, `.txt`, etc. with syntax highlighting

- **📎 Merge PDFs**
  - Combine multiple PDF files into a single document

- **🔍 Deep Scan Mode**
  - Generates a PDF report with:
    - File name, size, modified time, MIME type
    - Hex dump preview (first 256 bytes)

- **🗂 Structured Output**
  - Automatically creates folders:
    - `outputs/pdfs/`
    - `outputs/logs/`
    - `outputs/reports/`
    - `outputs/json/` (if enabled)

---

## 🛠 Requirements

- **Python 3.7+**
- **Windows OS only** (for Office conversion)
- **Microsoft Office** (Word, Excel, PowerPoint)
- **[wkhtmltopdf](https://wkhtmltopdf.org/downloads.html)** installed and available in PATH (for code → PDF)

### Install Python Dependencies

```bash
pip install -r requirements.txt
requirements.txt

matlab
Copy
Edit
reportlab
PyPDF2
Pillow
pygments
pdfkit
python-magic
comtypes
📂 Output Structure
bash
Copy
Edit
outputs/
├── pdfs/       # Converted PDFs
├── logs/       # Log file (tool.log)
├── reports/    # Scan reports
└── json/       # Optional metadata exports
🖥️ How to Use
Run the script:

bash
Copy
Edit
python main.py
You'll see a menu:

mathematica
Copy
Edit
📌 Universal PDF Toolkit
1. Convert a File to PDF
2. Merge Multiple PDFs
3. Deep Scan & Generate Report
4. Exit
💡 Example Use Cases
Convert slides.pptx to slides.pdf

Merge intro.pdf, chapter1.pdf into book.pdf

Scan suspicious.exe and generate a forensic-style PDF report

⚙️ Configuration
You can customize the following options in the script:

python
Copy
Edit
ENABLE_JSON_EXPORT = True
OUTPUT_DIR = "outputs"
⚠ Limitations
Office file conversion works only on Windows with Microsoft Office installed

Code-to-PDF conversion requires wkhtmltopdf binary installed

📄 License
This project is licensed under the MIT License.

🤝 Contributions
Pull requests and issues are welcome! Please open one if you have a feature suggestion or bug to report.

python
Copy
Edit

---
📄 License
This project is licensed under the MIT License.
---
