import pdfplumber
import csv
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

# === 弹出窗口选择 PDF 文件夹和 CSV 输出文件路径 ===
root = tk.Tk()
root.withdraw()

pdf_folder = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
if not pdf_folder:
    messagebox.showwarning("未选择文件夹", "你没有选择任何PDF文件夹，程序将退出。")
    exit()

output_csv = filedialog.asksaveasfilename(
    title="保存CSV文件",
    defaultextension=".csv",
    filetypes=[("CSV文件", "*.csv")]
)
if not output_csv:
    messagebox.showwarning("未选择CSV文件路径", "你没有设置CSV文件路径，程序将退出。")
    exit()

# === 日期格式化为 MM/DD/YYYY（符合 portal 要求） ===
def clean_date(raw_date):
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(raw_date.strip(), fmt).strftime("%m/%d/%Y")
        except:
            continue
    return raw_date.strip()

# === 提取 PDF 内容 ===
def extract_data_from_text(text):
    data = {}
    lines = text.splitlines()
    try:
        data["STORE NUMBER"] = lines[2].split()[0].strip()
        data["RTV PO#"] = str(lines[4].split()[0]).zfill(8)
        raw_date = lines[4].split()[1].strip()
        data["RTV PO DATE"] = clean_date(raw_date)
        data["RTV GRAND TOTAL"] = lines[-1].split("$")[-1].replace(",", "").strip()
        data["SKU"] = lines[22].split()[0].strip()
    except Exception:
        data["STORE NUMBER"] = "Not Found"
        data["RTV PO#"] = "Not Found"
        data["RTV PO DATE"] = "Not Found"
        data["RTV GRAND TOTAL"] = "Not Found"
        data["SKU"] = "Not Found"
    return data

# === 读取 PDF 文件并提取数据 ===
extracted_data = []
for pdf_file in os.listdir(pdf_folder):
    if pdf_file.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, pdf_file)
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    print(f"Processing {pdf_file} - Page {pdf.pages.index(page) + 1}")
                    pdf_data = extract_data_from_text(text)
                    row = {
                        "PACKAGE DATE": "",
                        "PACKAGE NUMBER": "",
                        "RTV GRAND TOTAL": str(pdf_data["RTV GRAND TOTAL"]),
                        "TYPE": "POR",
                        "COMMENTS": "Please provide us the POD of the return. We do not have the return record in our warehouse system.",
                        "P-VENDOR NUMBER": "877815",
                        "RTV PO#": str(pdf_data["RTV PO#"]),
                        "PO": "",
                        "STORE NUMBER": str(pdf_data["STORE NUMBER"]),
                        "RTV PO DATE": str(pdf_data["RTV PO DATE"]),
                        "SKU": str(pdf_data["SKU"]),
                        "PART": ""
                    }
                    extracted_data.append(row)

# === 写入 CSV 文件 ===
with open(output_csv, "w", newline="", encoding="utf-8") as csvfile:
    fieldnames = ["PACKAGE DATE", "PACKAGE NUMBER", "RTV GRAND TOTAL", "TYPE", "COMMENTS",
                  "P-VENDOR NUMBER", "RTV PO#", "PO", "STORE NUMBER", "RTV PO DATE", "SKU", "PART"]
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    for row in extracted_data:
        writer.writerow({k: str(v).strip() for k, v in row.items()})

messagebox.showinfo("完成", f"✅ 成功导出CSV文件：\n{output_csv}")
