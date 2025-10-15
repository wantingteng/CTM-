import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# 选择 Excel 文件
def select_file():
    return filedialog.askopenfilename(
        title="请选择要去重的 Excel 文件",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

def remove_duplicates():
    file_path = select_file()
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        original_rows = len(df)
        df_deduped = df.drop_duplicates()
        deduped_rows = len(df_deduped)

        # 保存去重后的文件
        new_file_path = os.path.splitext(file_path)[0] + "_去重后.xlsx"
        df_deduped.to_excel(new_file_path, index=False)

        messagebox.showinfo(
            "去重成功",
            f"原行数：{original_rows}\n去重后行数：{deduped_rows}\n已保存为：\n{new_file_path}"
        )
    except Exception as e:
        messagebox.showerror("错误", str(e))

# 简单 UI
def run_ui():
    root = tk.Tk()
    root.title("Excel 去重工具")
    canvas = tk.Canvas(root, width=400, height=200)
    canvas.pack()

    btn = tk.Button(root, text="选择文件并去重", command=remove_duplicates, height=2, width=30)
    canvas.create_window(200, 100, window=btn)

    root.mainloop()

if __name__ == "__main__":
    run_ui()
