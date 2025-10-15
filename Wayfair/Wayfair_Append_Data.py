import pandas as pd
import tkinter as tk
from tkinter import filedialog

# 初始化 tkinter，隐藏主窗口
root = tk.Tk()
root.withdraw()

# 🔹 弹窗选择要追加的新数据文件
print("📂 请选择要追加的新数据 Excel 文件：")
filtered_file = filedialog.askopenfilename(
    title="选择追加数据文件",
    filetypes=[("Excel 文件", "*.xlsx *.xls")]
)

if not filtered_file:
    print("❌ 未选择新数据文件，程序终止。")
    exit()

# 🔹 弹窗选择主表（要更新的 Excel 文件）
print("📂 请选择主表 Excel 文件（将被更新覆盖）：")
master_file = filedialog.askopenfilename(
    title="选择主表文件",
    filetypes=[("Excel 文件", "*.xlsx *.xls")]
)

if not master_file:
    print("❌ 未选择主表文件，程序终止。")
    exit()

# 🔹 读取文件
new_data = pd.read_excel(filtered_file, dtype=str)

try:
    master_data = pd.read_excel(master_file, dtype=str)
except FileNotFoundError:
    print("⚠️ 找不到主表，将新建主表。")
    master_data = pd.DataFrame()

# 🔹 合并数据（不去重）
combined = pd.concat([master_data, new_data], ignore_index=True)

# 🔹 覆盖保存到主表路径
combined.to_excel(master_file, index=False)

print(f"✅ 成功将 {filtered_file} 的内容追加到 {master_file} 中。")
