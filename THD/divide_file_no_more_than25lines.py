import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# 创建隐藏主窗口
root = tk.Tk()
root.withdraw()

# 选择输入CSV文件
input_csv = filedialog.askopenfilename(title="选择要拆分的CSV文件", filetypes=[("CSV 文件", "*.csv")])
if not input_csv:
    messagebox.showwarning("未选择文件", "你没有选择任何文件，程序将退出。")
    exit()

# 选择输出文件夹
output_folder = filedialog.askdirectory(title="选择保存拆分文件的文件夹")
if not output_folder:
    messagebox.showwarning("未选择文件夹", "你没有选择任何输出文件夹，程序将退出。")
    exit()

max_rows_per_file = 25  # 每个文件最大数据行数（不含表头）
original_filename = os.path.splitext(os.path.basename(input_csv))[0]
os.makedirs(output_folder, exist_ok=True)

with open(input_csv, newline='', encoding='utf-8') as csvfile:
    reader = csv.reader(csvfile)
    header = next(reader)
    file_count = 1
    rows = []

    for row in reader:
        row = [str(cell) if not isinstance(cell, str) else cell for cell in row]
        rows.append(row)

        if len(rows) >= max_rows_per_file:
            output_file = os.path.join(output_folder, f"{original_filename} ({file_count}).csv")
            with open(output_file, 'w', newline='', encoding='utf-8') as out_csv:
                writer = csv.writer(out_csv)
                writer.writerow(header)
                writer.writerows(rows)
            print(f"Saved: {output_file}")
            rows = []
            file_count += 1

    if rows:
        output_file = os.path.join(output_folder, f"{original_filename} ({file_count}).csv")
        with open(output_file, 'w', newline='', encoding='utf-8') as out_csv:
            writer = csv.writer(out_csv)
            writer.writerow(header)
            writer.writerows(rows)
        print(f"Saved: {output_file}")

messagebox.showinfo("完成", "✅ 所有文件已成功拆分完成。")
