import pdfplumber
import csv
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

# === 弹窗选择路径 ===
root = tk.Tk()
root.withdraw()

pdf_folder = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
if not pdf_folder:
    messagebox.showwarning("未选择PDF文件夹", "程序终止。")
    exit()

output_csv = filedialog.asksaveasfilename(
    title="保存CSV文件为...",
    defaultextension=".csv",
    filetypes=[("CSV Files", "*.csv")]
)
if not output_csv:
    messagebox.showwarning("未选择CSV文件路径", "程序终止。")
    exit()

# === 日期处理函数 ===
def clean_rtv_date(raw_date):
    """自动将各种日期格式转为 m/d/yyyy"""
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            dt = datetime.strptime(raw_date.strip(), fmt)
            return f"{dt.month}/{dt.day}/{dt.year}"
        except:
            continue
    return raw_date

# === 提取 PDF 内容 ===
def extract_data_from_text(text):
    data = {}
    lines = text.splitlines()
    try:
        data["STORE NUMBER"] = lines[2].split()[0].strip()
        data["RTV PO#"] = str(lines[4].split()[0]).strip()
        raw_date = lines[4].split()[1].strip()
        data["RTV PO DATE"] = clean_rtv_date(raw_date)
        rtv_value = next((line.split("$")[-1].strip() for line in lines if "RTV GRAND TOTAL" in line), "Not Found")
        data["RTV GRAND TOTAL"] = rtv_value.replace(",", "").strip()
        data["SKU"] = str(lines[19].split()[1].strip())
    except IndexError:
        data = {k: "Not Found" for k in ["STORE NUMBER", "RTV PO#", "RTV PO DATE", "RTV GRAND TOTAL", "SKU"]}
    return data

# === 扫描 PDF 并提取所有数据 ===
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
                        "RTV GRAND TOTAL": pdf_data["RTV GRAND TOTAL"],
                        "TYPE": "POR",
                        "COMMENTS": "Please provide us the POD of the return. We do not have the return record in our warehouse system.",
                        "P-VENDOR NUMBER": "27005187",  # 可自定义
                        "RTV PO#": pdf_data["RTV PO#"],
                        "PO": "",
                        "STORE NUMBER": pdf_data["STORE NUMBER"],
                        "RTV PO DATE": pdf_data["RTV PO DATE"],
                        "SKU": pdf_data["SKU"],
                        "PART": ""
                    }
                    extracted_data.append(row)

# === 写入 CSV 文件 ===
with open(output_csv, "w", newline="", encoding="utf-8") as csvfile:
    fieldnames = [
        "PACKAGE DATE", "PACKAGE NUMBER", "RTV GRAND TOTAL", "TYPE", "COMMENTS",
        "P-VENDOR NUMBER", "RTV PO#", "PO", "STORE NUMBER", "RTV PO DATE", "SKU", "PART"
    ]
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    for row in extracted_data:
        row["P-VENDOR NUMBER"] = row["P-VENDOR NUMBER"].strip()
        row["SKU"] = row["SKU"].strip()
        writer.writerow(row)

messagebox.showinfo("完成", f"✅ PDF 数据已导出为 CSV 文件：\n{output_csv}")
