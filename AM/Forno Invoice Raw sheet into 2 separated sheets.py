import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# 要保留的列
COLUMNS_TO_KEEP = [
    "Store",
    "Total Pieces",
    "Create Date",
    "Total Cost With Surcharge"
]

# 指定要提取的目标 Store 列表
TARGET_STORES = [
    "14 Premium Home Source",
    "68 Wal-Mart Marketplace",
    "13 Home Outlet Direct",
    "9 Lowe's",
    "40:1 Home Best Price Products Inc. : Home Best Price dba Amazon",
    "11 Best Buy",
    "Home Depot / Forno"
]

# 日期标准化函数
def normalize_dates(df):
    if "Create Date" in df.columns:
        df["Create Date"] = pd.to_datetime(df["Create Date"], errors="coerce", dayfirst=False).dt.date
    return df

# 打开文件选择对话框
def select_file(title="选择 Excel 文件"):
    return filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

# 保存文件对话框
def select_output_file(title="保存为"):
    return filedialog.asksaveasfilename(
        title=title,
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

# 加载 Excel（含工作表选择）
def select_sheet_and_load(file_path):
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheet_names = xls.sheet_names

    if len(sheet_names) == 1:
        return pd.read_excel(file_path, sheet_name=sheet_names[0], engine="openpyxl")

    df_result = None
    def confirm_selection():
        selected_sheet = sheet_var.get()
        top.destroy()
        nonlocal df_result
        df_result = pd.read_excel(file_path, sheet_name=selected_sheet, engine="openpyxl")

    top = tk.Toplevel()
    top.title("请选择工作表 Sheet")
    tk.Label(top, text="该文件有多个工作表，请选择一个：").pack(padx=10, pady=10)

    sheet_var = tk.StringVar(value=sheet_names[0])
    dropdown = tk.OptionMenu(top, sheet_var, *sheet_names)
    dropdown.pack(pady=5)

    confirm_btn = tk.Button(top, text="确认", command=confirm_selection)
    confirm_btn.pack(pady=10)

    top.wait_window()
    return df_result

# 主逻辑：清洗并拆分导出
def split_and_export():
    input_file = select_file("选择要处理的 Excel 文件")
    if not input_file:
        return

    try:
        df = select_sheet_and_load(input_file)

        # 保留列 + 清洗空值
        df = df[[col for col in df.columns if col in COLUMNS_TO_KEEP]]
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].fillna("").astype(str).str.strip()
            else:
                df[col] = df[col].fillna("")

        df = normalize_dates(df)

        # 拆分数据
        df_target = df[df["Store"].isin(TARGET_STORES)]
        df_remaining = df[~df["Store"].isin(TARGET_STORES)]

        # 保存两个文件
        target_path = select_output_file("保存【指定 Store】的文件")
        if target_path:
            df_target.to_excel(target_path, index=False)

        remaining_path = select_output_file("保存【剩余 Store】的文件")
        if remaining_path:
            df_remaining.to_excel(remaining_path, index=False)

        messagebox.showinfo("完成", f"已成功拆分并保存两个文件。")
    except Exception as e:
        messagebox.showerror("错误", str(e))

# UI 界面
def run_ui():
    root = tk.Tk()
    root.title("Excel 数据清洗拆分工具")
    canvas = tk.Canvas(root, width=400, height=180)
    canvas.pack()

    split_button = tk.Button(root, text="清洗并拆分导出两个文件", command=split_and_export, height=2, width=30)
    canvas.create_window(200, 80, window=split_button)

    root.mainloop()

if __name__ == "__main__":
    run_ui()
