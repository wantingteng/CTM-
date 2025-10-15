import pandas as pd
from tkinter import Tk, filedialog
import os

# 指定要提取的目标 Store 列表
target_stores = [
    "14 Premium Home Source",
    "68 Wal-Mart Marketplace",
    "13 Home Outlet Direct",
    "9 Lowe's",
    "40:1 Home Best Price Products Inc. : Home Best Price dba Amazon",
    "11 Best Buy",
    "Home Depot / Forno"
]

# 打开文件选择窗口
root = Tk()
root.withdraw()  # 不显示主窗口

# 选择原始 Excel 文件
file_path = filedialog.askopenfilename(
    title="请选择原始Excel文件",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not file_path:
    print("未选择文件，程序退出。")
    exit()

# 读取 Excel 文件
df = pd.read_excel(file_path)

# 提取符合条件的行
df_extracted = df[df["Store"].isin(target_stores)]

# 删除这些行，得到剩余的数据
df_remaining = df[~df["Store"].isin(target_stores)]

# 选择保存提取数据的文件
save_extracted_path = filedialog.asksaveasfilename(
    title="保存提取出的数据为新文件",
    defaultextension=".xlsx",
    filetypes=[("Excel files", "*.xlsx")]
)

if save_extracted_path:
    df_extracted.to_excel(save_extracted_path, index=False)
    print(f"提取数据已保存至：{save_extracted_path}")
else:
    print("未保存提取文件，程序退出。")
    exit()

# 保存剩余数据（可选覆盖或另存为）
save_remaining_path = filedialog.asksaveasfilename(
    title="保存修改后的原始数据（原文件中移除了目标行）",
    defaultextension=".xlsx",
    filetypes=[("Excel files", "*.xlsx")]
)

if save_remaining_path:
    df_remaining.to_excel(save_remaining_path, index=False)
    print(f"剩余数据已保存至：{save_remaining_path}")
else:
    print("未保存剩余文件，程序退出。")
