import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

COLUMNS_TO_KEEP = [
    "Store",
    "Total Pieces",
    "Create Date",
    "Total Cost With Surcharge"
]

# 标准化日期
def normalize_dates(df):
    if "Create Date" in df.columns:
        df["Create Date"] = pd.to_datetime(df["Create Date"], errors="coerce", dayfirst=False).dt.date
    return df

# UI 文件选择
def select_file(title="请选择 Excel 文件"):
    return filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

# 将 new_df 追加进 master_df，去重
def append_data(master_df, new_df):
    master_df = normalize_dates(master_df)
    new_df = normalize_dates(new_df)

    subset_cols = [col for col in COLUMNS_TO_KEEP if col in master_df.columns and col in new_df.columns]
    combined = pd.concat([master_df, new_df], ignore_index=True)
    combined = combined.drop_duplicates(subset=subset_cols, keep="first")
    return combined

def append_one(master_title, new_title):
    new_file = select_file(f"请选择【{new_title}】新文件")
    if not new_file:
        return
    master_file = select_file(f"请选择【{master_title}】母表文件")
    if not master_file:
        return

    try:
        new_df = pd.read_excel(new_file, engine="openpyxl")
        master_df = pd.read_excel(master_file, engine="openpyxl")

        final_df = append_data(master_df, new_df)

        with pd.ExcelWriter(master_file, engine="openpyxl", date_format="mm/dd/yyyy") as writer:
            final_df.to_excel(writer, index=False)

        messagebox.showinfo("完成", f"已成功将【{new_title}】追加到母表中并保存。")
    except Exception as e:
        messagebox.showerror("错误", str(e))

def run_ui():
    root = tk.Tk()
    root.title("追加数据到历史母表工具")
    canvas = tk.Canvas(root, width=400, height=250)
    canvas.pack()

    btn1 = tk.Button(root, text="追加【目标 Store】数据", command=lambda: append_one("目标 Store 母表", "目标 Store 新文件"), height=2, width=30)
    btn2 = tk.Button(root, text="追加【其他 Store】数据", command=lambda: append_one("其他 Store 母表", "其他 Store 新文件"), height=2, width=30)

    canvas.create_window(200, 80, window=btn1)
    canvas.create_window(200, 150, window=btn2)

    root.mainloop()

if __name__ == "__main__":
    run_ui()
