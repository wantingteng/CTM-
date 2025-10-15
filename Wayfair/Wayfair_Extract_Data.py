import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

# 从原始 CSV 中提取某个月或某个范围内的 Settlement date 数据，并生成一个新的 Excel 文件。
#
# 主要功能：
#
# 弹出窗口让你选择一个原始 .csv 文件；
#
# 你可以输入：
#
# 1 → 提取某个 单月；
#
# 2 → 提取一个 月份范围；
#
# 自动添加 Month 列；
#
# 删除 Resources 列（如果有）；
#
# 结果按日期排序，保存为 Excel 到目录：C:\Reporting\Wayfair\CA\src_Data

# 初始化 tkinter，隐藏主窗口
root = tk.Tk()
root.withdraw()

# 🔹 弹窗选择原始 CSV 文件
print("📂 正在弹出窗口，请选择原始 CSV 文件...")
csv_file = filedialog.askopenfilename(
    title="请选择原始 CSV 文件",
    filetypes=[("CSV 文件", "*.csv")]
)

if not csv_file:
    print("❌ 未选择 CSV 文件，程序终止。")
    exit()

# 🔹 用户选择提取模式
mode = input("请选择提取模式：1（单月） 或 2（范围）: ").strip()

# 🔹 读取 CSV
df = pd.read_csv(csv_file)
df['Settlement date'] = pd.to_datetime(df['Settlement date'], errors='coerce')

# 🔹 插入 Month 列（在 Settlement date 后）
df.insert(df.columns.get_loc('Settlement date') + 1, 'Month', df['Settlement date'].dt.strftime('%m'))

# 🔹 删除 Resources 列（如存在）
if 'Resources' in df.columns:
    df.drop(columns=['Resources'], inplace=True)

# 🔹 数据筛选
if mode == '1':
    year = int(input("请输入年份（如 2025）: "))
    month = input("请输入月份（如 01, 02, ...）: ").zfill(2)
    mask = (df['Settlement date'].dt.year == year) & (df['Settlement date'].dt.strftime('%m') == month)
    filtered_df = df.loc[mask]
    output_filename = f'Filtered_Deductions_{year}_{month}.xlsx'

elif mode == '2':
    year = int(input("请输入年份（如 2025）: "))
    start_month = int(input("请输入起始月份（如 1）: "))
    end_month = int(input("请输入结束月份（如 4）: "))
    mask = (df['Settlement date'].dt.year == year) & \
           (df['Settlement date'].dt.month >= start_month) & \
           (df['Settlement date'].dt.month <= end_month)
    filtered_df = df.loc[mask]
    output_filename = f'Filtered_Deductions_{year}_{str(start_month).zfill(2)}_{str(end_month).zfill(2)}.xlsx'

else:
    print("❌ 模式输入错误，程序退出。")
    exit()

# 🔹 排序并格式化日期
filtered_df = filtered_df.sort_values(by='Settlement date')
filtered_df['Settlement date'] = filtered_df['Settlement date'].dt.strftime('%Y-%m-%d')

# 🔹 写死输出目录
output_dir = r'C:\Users\LZhu\Wanting_OneDrive\OneDrive - CTM Group\Report\Reporting\Wayfair'
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, output_filename)

# 🔹 保存为 Excel
filtered_df.to_excel(output_path, index=False)
print(f"✅ 提取成功，文件已保存至：{output_path}")
