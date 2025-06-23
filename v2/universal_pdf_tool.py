import os, time, magic, pdfkit, logging, json, sys
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfMerger
from PIL import Image
from pygments import highlight
from pygments.lexers import get_lexer_for_filename
from pygments.formatters import HtmlFormatter
import comtypes.client
from datetime import datetime

# === CONFIG ===
ENABLE_JSON_EXPORT = True
OUTPUT_DIR = "outputs"
PDF_DIR = os.path.join(OUTPUT_DIR, "pdfs")
LOG_DIR = os.path.join(OUTPUT_DIR, "logs")
REPORT_DIR = os.path.join(OUTPUT_DIR, "reports")
JSON_DIR = os.path.join(OUTPUT_DIR, "json")
for d in [PDF_DIR, LOG_DIR, REPORT_DIR, JSON_DIR]: os.makedirs(d, exist_ok=True)

# === Logging Setup ===
logging.basicConfig(filename=os.path.join(LOG_DIR, "tool.log"), level=logging.INFO,
    format="[%(asctime)s] %(levelname)s: %(message)s")

# Supported extensions
word_exts  = ['.doc', '.docx']
excel_exts = ['.xls', '.xlsx']
ppt_exts   = ['.ppt', '.pptx']
img_exts   = ['.jpg', '.jpeg', '.png']
code_exts  = ['.py', '.js', '.css', '.java', '.html', '.xml', '.json', '.md', '.txt']

def convert_code_to_pdf(file_path):
    try:
        lexer = get_lexer_for_filename(file_path)
        formatter = HtmlFormatter(full=True, style='colorful')
        with open(file_path, 'r', encoding="utf-8", errors="ignore") as f:
            code = f.read()
        html_code = highlight(code, lexer, formatter)
        html_path = file_path + ".html"
        with open(html_path, 'w', encoding="utf-8") as f:
            f.write(html_code)
        pdf_name = os.path.splitext(os.path.basename(file_path))[0] + ".pdf"
        pdf_path = os.path.join(PDF_DIR, pdf_name)
        pdfkit.from_file(html_path, pdf_path)
        os.remove(html_path)
        logging.info(f"Code converted to PDF: {file_path} → {pdf_path}")
        print(f"✅ Saved to {pdf_path}")
    except Exception as e:
        logging.error(f"Failed code-to-PDF: {file_path}: {e}")
        print(f"❌ Failed to convert code to PDF: {e}")

def convert_image_to_pdf(file_path):
    try:
        image = Image.open(file_path)
        image = image.convert('RGB')
        pdf_name = os.path.splitext(os.path.basename(file_path))[0] + ".pdf"
        pdf_path = os.path.join(PDF_DIR, pdf_name)
        image.save(pdf_path)
        logging.info(f"Image converted: {file_path} → {pdf_path}")
        print(f"✅ Saved to {pdf_path}")
    except Exception as e:
        logging.error(f"Failed image-to-PDF: {file_path}: {e}")
        print(f"❌ Failed: {e}")

def convert_excel_to_pdf(file_path):
    try:
        excel = comtypes.client.CreateObject('Excel.Application')
        excel.Visible = False
        wb = excel.Workbooks.Open(file_path)
        pdf_name = os.path.splitext(os.path.basename(file_path))[0] + ".pdf"
        pdf_path = os.path.join(PDF_DIR, pdf_name)
        wb.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)
        excel.Quit()
        logging.info(f"Excel converted: {file_path} → {pdf_path}")
        print(f"✅ Saved to {pdf_path}")
    except Exception as e:
        logging.error(f"Failed Excel-to-PDF: {file_path}: {e}")
        print(f"❌ Failed: {e}")

def merge_pdfs(pdf_paths, output_path):
    try:
        merger = PdfMerger()
        for path in pdf_paths:
            merger.append(path.strip())
        output = os.path.join(PDF_DIR, output_path)
        merger.write(output)
        merger.close()
        logging.info(f"Merged PDFs: {output}")
        print(f"✅ Merged PDF saved to {output}")
    except Exception as e:
        logging.error(f"Failed PDF merge: {e}")
        print(f"❌ Merge failed: {e}")

def convert_to_pdf(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in code_exts:
        convert_code_to_pdf(file_path)
    elif ext in img_exts:
        convert_image_to_pdf(file_path)
    elif ext in excel_exts:
        convert_excel_to_pdf(file_path)
    else:
        print("⚠️ Format not supported for direct conversion. Try Deep Scan.")
        logging.warning(f"Unsupported format: {file_path}")

def scan_file(file_path):
    try:
        file_name = os.path.basename(file_path)
        size_kb = round(os.path.getsize(file_path) / 1024, 2)
        timestamp = os.path.getmtime(file_path)
        modified = datetime.fromtimestamp(timestamp).strftime("%A, %B %d, %Y at %H:%M:%S")
        mime_type = magic.Magic(mime=True).from_file(file_path)

        with open(file_path, "rb") as f:
            hex_preview = f.read(256).hex()

        report_name = f"{file_name}_report.pdf"
        report_path = os.path.join(REPORT_DIR, report_name)
        c = canvas.Canvas(report_path, pagesize=A4)

        y = 800
        label_font = "Helvetica-Bold"
        value_font = "Helvetica"
        size = 12

        def draw_label_value(label, value, offset=60):
            nonlocal y
            c.setFont(label_font, size)
            c.drawString(100, y, label)
            c.setFont(value_font, size)
            c.drawString(100 + offset, y, value)
            y -= 20

        draw_label_value("File:", file_name)
        draw_label_value("Size:", f"{size_kb} KB")
        draw_label_value("Modified:", modified)
        draw_label_value("Type:", mime_type)

        c.setFont(label_font, size)
        c.drawString(100, y, "Hex Preview (256 bytes):")
        y -= 20

        c.setFont(value_font, 10)
        hex_lines = [hex_preview[i:i + 64] for i in range(0, len(hex_preview), 64)]
        for line in hex_lines:
            c.drawString(100, y, line)
            y -= 12
            if y < 50:
                c.showPage()
                y = 800

        c.save()
        logging.info(f"Deep scan report saved: {report_path}")
        print(f"✅ Scan report saved to {report_path}")

        if ENABLE_JSON_EXPORT:
            json_data = {
                "File Name": file_name,
                "Size (KB)": size_kb,
                "Modified": modified,
                "Type": mime_type,
                "Hex Preview": hex_preview[:256]
            }
            json_path = os.path.join(JSON_DIR, f"{file_name}_report.json")
            with open(json_path, 'w') as f:
                json.dump(json_data, f, indent=4)
            logging.info(f"JSON exported: {json_path}")
            print(f"📄 Also exported JSON: {json_path}")

    except Exception as e:
        logging.error(f"Scan failed: {file_path}: {e}")
        print(f"❌ Scan failed: {e}")

def menu():
    while True:
        print("\n💼 Universal PDF Toolkit v2.0")
        print("1. Convert a File to PDF")
        print("2. Merge Multiple PDFs")
        print("3. Deep Scan & Generate Report")
        print("4. Exit")
        choice = input("Enter your choice: ").strip()
        if choice == '1':
            file_path = input("Enter file path to convert: ").strip()
            convert_to_pdf(file_path)
        elif choice == '2':
            pdfs = input("Enter PDF paths to merge (comma-separated): ").split(',')
            output = input("Enter output PDF name (e.g., merged.pdf): ").strip()
            merge_pdfs(pdfs, output)
        elif choice == '3':
            file_path = input("Enter file path to scan: ").strip()
            scan_file(file_path)
        elif choice == '4':
            print("👋 Exiting. Take care!")
            break
        else:
            print("❌ Invalid choice.")

if __name__ == "__main__":
    menu()
