#用于记录thd已经提交的disputed package 便于后续check
#此文件仅为初级清洗数据，清洗从thd网站下载的raw data
import pandas as pd
from tkinter import filedialog, Tk, messagebox
import os

# 选择 CSV 数据文件
def select_csv_file():
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="请选择CSV文件", filetypes=[("CSV Files", "*.csv")]
    )

# 选择输出目录
def select_output_folder():
    root = Tk()
    root.withdraw()
    return filedialog.askdirectory(title="请选择保存文件夹")

# 选择已有母表文件
def select_master_file():
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="请选择母表Excel文件", filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

# 清洗数据通用步骤
def clean_data(df):
    # 删除不需要的列
    for col in ["DISPUTED INVOICES", "RECEIVED DATE", "CLOSE DATE"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    # 去除所有字段中多余引号
    df = df.applymap(lambda x: x.replace("'", "") if isinstance(x, str) else x)

    # 添加 STATUS 列（如果不存在）
    if "STATUS" not in df.columns:
        df["STATUS"] = ""

    return df

# 模式 1：仅清洗
def clean_only():
    input_file = select_csv_file()
    if not input_file:
        print("❌ 未选择CSV文件")
        return

    output_folder = select_output_folder()
    if not output_folder:
        print("❌ 未选择输出目录")
        return

    df = pd.read_csv(input_file, dtype=str)
    df = clean_data(df)

    # 保存路径
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_path = os.path.join(output_folder, base_name + "_cleaned.xlsx")

    df.to_excel(output_path, index=False)
    print(f"✅ 清洗完成，保存为：{output_path}")

# 模式 2：清洗 + 追加到母表（保留所有原始列内容）
def clean_and_append():
    input_file = select_csv_file()
    if not input_file:
        print("❌ 未选择CSV文件")
        return

    master_file = select_master_file()
    if not master_file:
        print("❌ 未选择母表")
        return

    df_new = pd.read_csv(input_file, dtype=str)
    df_new = clean_data(df_new)

    # 读取母表
    df_master = pd.read_excel(master_file, dtype=str)

    # 确保 STATUS 列存在
    if "STATUS" not in df_master.columns:
        df_master["STATUS"] = ""

    # 对齐新数据的列顺序：与母表相同列名为准，缺失列补空
    for col in df_master.columns:
        if col not in df_new.columns:
            df_new[col] = ""
    df_new = df_new[df_master.columns]

    # 合并
    combined = pd.concat([df_master, df_new], ignore_index=True)

    # 去重：保留原母表版本（非空STATUS在前）
    if "DISPUTE PKG #" in combined.columns:
        combined = (
            combined.sort_values("STATUS", ascending=False)
            .drop_duplicates(subset=["DISPUTE PKG #"], keep="first")
            .reset_index(drop=True)
        )
    else:
        print("⚠️ 警告：未找到 DISPUTE PKG # 字段，无法去重")

    # 保存
    combined.to_excel(master_file, index=False)
    print(f"✅ 清洗并追加完成，母表已更新：{master_file}")

def main():
    root = Tk()
    root.withdraw()
    choice = messagebox.askyesno(
        "功能选择",
        "是否要将清洗结果追加到母表？\n\n是 = 清洗并追加\n否 = 只清洗保存"
    )
    if choice:
        clean_and_append()
    else:
        clean_only()

if __name__ == "__main__":
    main()
