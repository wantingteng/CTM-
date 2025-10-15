import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

COLUMNS_TO_KEEP = [
    "Store",
    "Total Pieces",
    "Create Date",
    "Total Cost With Surcharge"
]

# 日期标准化
def normalize_dates(df):
    if "Create Date" in df.columns:
        df["Create Date"] = pd.to_datetime(df["Create Date"], errors="coerce").dt.date
    return df

# UI选择文件
def select_file(title="请选择 Excel 文件"):
    return filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

# 支持多 sheet 选择
def select_sheet_and_load(file_path):
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheet_names = xls.sheet_names

    if len(sheet_names) == 1:
        return pd.read_excel(file_path, sheet_name=sheet_names[0], engine="openpyxl")

    df_result = None
    def confirm_selection():
        selected = sheet_var.get()
        top.destroy()
        nonlocal df_result
        df_result = pd.read_excel(file_path, sheet_name=selected, engine="openpyxl")

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

# 清洗函数
def clean_data(df):
    df = df[[col for col in df.columns if col in COLUMNS_TO_KEEP]].copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].fillna("").astype(str).str.strip()
        else:
            df[col] = df[col].fillna("")
    return normalize_dates(df)

# 主逻辑：清洗并追加到 big customer
def append_to_big_customer():
    new_file = select_file("选择要追加的新文件")
    if not new_file:
        return

    mother_file = select_file("选择 big customer 母表")
    if not mother_file:
        return

    try:
        new_df = select_sheet_and_load(new_file)
        new_df = clean_data(new_df)

        mother_df = select_sheet_and_load(mother_file)
        mother_df = clean_data(mother_df)

        subset_cols = [col for col in COLUMNS_TO_KEEP if col in new_df.columns and col in mother_df.columns]
        final_df = pd.concat([mother_df, new_df], ignore_index=True)
        final_df = final_df.drop_duplicates(subset=subset_cols)

        with pd.ExcelWriter(mother_file, engine="openpyxl", date_format="mm/dd/yyyy") as writer:
            final_df.to_excel(writer, index=False)

        messagebox.showinfo("完成", f"新数据已成功追加到 big customer 并保存：\n{mother_file}")
    except Exception as e:
        messagebox.showerror("错误", str(e))

# UI 入口
def run_ui():
    root = tk.Tk()
    root.title("追加数据到 big customer")
    canvas = tk.Canvas(root, width=400, height=180)
    canvas.pack()

    btn = tk.Button(root, text="追加新数据到 big customer", command=append_to_big_customer, height=2, width=30)
    canvas.create_window(200, 80, window=btn)

    root.mainloop()

if __name__ == "__main__":
    run_ui()
