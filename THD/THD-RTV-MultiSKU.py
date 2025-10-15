import pdfplumber
import csv
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

# === UI 选择路径 ===
root = tk.Tk()
root.withdraw()

pdf_folder = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
if not pdf_folder:
    messagebox.showwarning("未选择文件夹", "你没有选择任何PDF文件夹，程序将退出。")
    exit()

output_csv = filedialog.asksaveasfilename(
    title="保存CSV文件",
    defaultextension=".csv",
    filetypes=[("CSV 文件", "*.csv")]
)
if not output_csv:
    messagebox.showwarning("未选择CSV文件路径", "你没有设置CSV文件路径，程序将退出。")
    exit()

# === 提取 general 信息（门店、PO号、日期） ===
def extract_general_data(lines):
    try:
        store_number = lines[2].split()[0]
        rtv_po = str(lines[4].split()[0])
        raw_date = lines[4].split()[1]

        parsed_date = raw_date
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
            try:
                dt = datetime.strptime(raw_date, fmt)
                parsed_date = f"{dt.month}/{dt.day}/{dt.year}"  # Windows 兼容的 5/5/2025 格式
                break
            except:
                continue

        return {
            "STORE NUMBER": store_number,
            "RTV PO#": rtv_po,
            "RTV PO DATE": parsed_date
        }
    except Exception as e:
        print("❌ Error parsing general data:", e)
        return None


# === 提取 SKU 行 ===
def extract_sku_data(lines):
    sku_data = []
    start_index = next((i for i, line in enumerate(lines) if line.strip().startswith("PART")), None)
    end_index = next((i for i, line in enumerate(lines) if line.strip().startswith("MERCHANDISE")), None)

    if start_index is not None and end_index is not None:
        relevant_lines = lines[start_index + 1:end_index]
        for line in relevant_lines:
            parts = line.split()
            if len(parts) >= 9:
                sku_data.append({
                    "SKU": parts[0],
                    "RTV GRAND TOTAL": parts[-1]
                })
    return sku_data

# === 主逻辑：提取并汇总多SKU PDF数据 ===
extracted_data = {}

for pdf_file in os.listdir(pdf_folder):
    if pdf_file.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, pdf_file)
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines = text.splitlines()
                    general_data = extract_general_data(lines)
                    sku_data_list = extract_sku_data(lines)

                    if general_data and sku_data_list:
                        for sku_data in sku_data_list:
                            key = (general_data["RTV PO#"], sku_data["SKU"])
                            try:
                                rtv_total = round(float(sku_data["RTV GRAND TOTAL"].replace("$", "").replace(",", "")) * 1.1, 2)
                            except ValueError:
                                continue

                            if key in extracted_data:
                                extracted_data[key]["RTV GRAND TOTAL"] += rtv_total
                            else:
                                extracted_data[key] = {
                                    "PACKAGE DATE": "",
                                    "PACKAGE NUMBER": "",
                                    "RTV GRAND TOTAL": rtv_total,
                                    "TYPE": "POR",
                                    "COMMENTS": "Please provide us the POD of the return. We do not have the return record in our warehouse system.",
                                    "P-VENDOR NUMBER": "877815",
                                    "RTV PO#": general_data["RTV PO#"],
                                    "PO": "",
                                    "STORE NUMBER": general_data["STORE NUMBER"],
                                    "RTV PO DATE": general_data["RTV PO DATE"],
                                    "SKU": sku_data["SKU"],
                                    "PART": ""
                                }

# === 写入 CSV 文件 ===
with open(output_csv, "w", newline="", encoding="utf-8") as csvfile:
    fieldnames = [
        "PACKAGE DATE", "PACKAGE NUMBER", "RTV GRAND TOTAL", "TYPE", "COMMENTS",
        "P-VENDOR NUMBER", "RTV PO#", "PO", "STORE NUMBER", "RTV PO DATE", "SKU", "PART"
    ]
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    for row in extracted_data.values():
        row["RTV GRAND TOTAL"] = round(row["RTV GRAND TOTAL"], 2)
        writer.writerow(row)

messagebox.showinfo("完成", f"✅ CSV 已保存：\n{output_csv}")
