import os, time, magic, pdfkit, logging, json
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfMerger
from PIL import Image
from pygments import highlight
from pygments.lexers import get_lexer_for_filename
from pygments.formatters import HtmlFormatter
import comtypes.client

# === CONFIG ===
ENABLE_JSON_EXPORT = True
OUTPUT_DIR = "outputs"
PDF_DIR = os.path.join(OUTPUT_DIR, "pdfs")
LOG_DIR = os.path.join(OUTPUT_DIR, "logs")
REPORT_DIR = os.path.join(OUTPUT_DIR, "reports")
JSON_DIR = os.path.join(OUTPUT_DIR, "json")

# Create folders if missing
for d in [PDF_DIR, LOG_DIR, REPORT_DIR, JSON_DIR]:
    os.makedirs(d, exist_ok=True)

# === Logging Setup ===
logging.basicConfig(
    filename=os.path.join(LOG_DIR, "tool.log"),
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s: %(message)s"
)

# Extension groups
word_exts  = ['.doc', '.docx']
excel_exts = ['.xls', '.xlsx']
ppt_exts   = ['.ppt', '.pptx']
img_exts   = ['.jpg', '.jpeg', '.png']
code_exts  = ['.py', '.js', '.css', '.java', '.html', '.xml', '.json', '.md', '.txt']

# Office → PDF
def office_to_pdf(file_path):
    ext_map = {
        '.doc': ('Word.Application', 17), '.docx': ('Word.Application', 17),
        '.xls': ('Excel.Application', 0), '.xlsx': ('Excel.Application', 0),
        '.ppt': ('PowerPoint.Application', 32), '.pptx': ('PowerPoint.Application', 32)
    }
    ext = os.path.splitext(file_path)[1].lower()
    app_name, fmt = ext_map[ext]
    app = comtypes.client.CreateObject(app_name)
    app.Visible = False
    doc = app.Documents.Open(file_path) if 'Word' in app_name else \
          app.Workbooks.Open(file_path) if 'Excel' in app_name else \
          app.Presentations.Open(file_path)
    pdf_path = file_path.replace(ext, '.pdf')
    if 'PowerPoint' in app_name:
        doc.SaveAs(pdf_path, fmt)
    elif 'Excel' in app_name:
        doc.ExportAsFixedFormat(fmt, pdf_path)
    else:
        doc.SaveAs(pdf_path, FileFormat=fmt)
    doc.Close()
    app.Quit()
    print(f"✅ Saved: {pdf_path}")

# Image → PDF
def image_to_pdf(file_path):
    img = Image.open(file_path).convert('RGB')
    pdf_path = os.path.splitext(file_path)[0] + '.pdf'
    img.save(pdf_path)
    print(f"✅ Saved: {pdf_path}")

# Code → Syntax-Highlighted PDF
def code_to_pdf(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        code = f.read()
    lexer = get_lexer_for_filename(file_path)
    formatter = HtmlFormatter(full=True, style='colorful')
    highlighted_code = highlight(code, lexer, formatter)

    html_file = file_path + '.html'
    pdf_file = file_path.rsplit('.', 1)[0] + '.pdf'

    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(highlighted_code)

    try:
        pdfkit.from_file(html_file, pdf_file)
        os.remove(html_file)
        print(f"✅ Saved: {pdf_file}")
    except Exception as e:
        print(f"❌ Failed to convert code to PDF: {e}")

# Merge PDFs
def merge_pdfs(pdf_list, output_name):
    merger = PdfMerger()
    for pdf in pdf_list:
        if os.path.exists(pdf.strip()):
            merger.append(pdf.strip())
    merger.write(output_name)
    merger.close()
    print(f"✅ Merged PDF saved: {output_name}")

# Deep Scan Mode
def scan_file(file_path):
    if not os.path.exists(file_path):
        print("❌ File not found.")
        return
    info = {
        "File Name": os.path.basename(file_path),
        "Size (KB)": round(os.path.getsize(file_path)/1024, 2),
        "Modified": time.ctime(os.path.getmtime(file_path)),
        "Type": magic.from_file(file_path, mime=True)
    }
    try:
        with open(file_path, "rb") as f:
            hex_preview = f.read(256).hex()
    except Exception as e:
        hex_preview = f"Read error: {e}"
    generate_scan_report(info, hex_preview, file_path)

def generate_scan_report(info, hex_data, original_path):
    report_path = original_path + "_report.pdf"
    c = canvas.Canvas(report_path, pagesize=A4)
    y = 800
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, y, "🧪 Deep Scan Report")
    c.setFont("Helvetica", 12)
    y -= 40
    for key, val in info.items():
        c.drawString(50, y, f"{key}: {val}")
        y -= 20
    y -= 10
    c.drawString(50, y, "🔍 Hex Preview (first 256 bytes):")
    c.setFont("Courier", 8)
    y -= 20
    for line in [hex_data[i:i+64] for i in range(0, len(hex_data), 64)]:
        c.drawString(50, y, line)
        y -= 10
        if y < 40:
            c.showPage()
            y = 800
    c.save()
    print(f"✅ Scan report created: {report_path}")

# Menu
def menu():
    while True:
        print("\n📌 Universal PDF Toolkit")
        print("1. Convert a File to PDF")
        print("2. Merge Multiple PDFs")
        print("3. Deep Scan & Generate Report")
        print("4. Exit")
        choice = input("Choose an option (1–4): ")

        if choice == '1':
            file_path = input("Enter file path: ").strip()
            ext = os.path.splitext(file_path)[1].lower()
            if ext in word_exts + excel_exts + ppt_exts:
                office_to_pdf(file_path)
            elif ext in img_exts:
                image_to_pdf(file_path)
            elif ext in code_exts:
                code_to_pdf(file_path)
            else:
                print("⏳ Format not directly supported. Try Deep Scan instead.")
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

# Run the tool
if __name__ == "__main__":
    menu()
