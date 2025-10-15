import pandas as pd

# 读取 Excel 文件
file_path = r"\\SRV-AD01\Folder Redirection\LZhu\Documents\CTM Worksheet\THD\For_Report\Raw data\Terminal\1.xlsx"  # Excel 文件路径
df = pd.read_excel(file_path, header=None)  # 读取所有数据，无表头

# **删除空白行和空白列**
df = df.dropna(how="all")  # 删除完全空的行
df = df.dropna(axis=1, how="all")  # 删除完全空的列
df = df.reset_index(drop=True)  # 重新索引，防止行号错乱

# 存储提取的数据
records = []
current_rtv = None  # D列的RTV编号
current_amt = None  # A列的金额
current_status = None  # STATUS下方的状态信息

# 遍历 Excel 的每一行
for i in range(len(df)):
    row = df.iloc[i]

    # 识别 A 列是否是一个数字（表示新记录开始）
    if isinstance(row[0], (int, float)):  # 如果 A 列是数字 (1,2,3,4,...)
        current_rtv = row[3] if not pd.isna(row[3]) else None  # 提取 D 列的 RTV 号
        current_amt = df.iloc[i + 2, 0] if i + 2 < len(df) else None  # 提取 A 列的金额

        # 查找 STATUS 的行，并提取它下一行的状态信息
        current_status = None
        for j in range(i, len(df)):  # 从当前行开始向下查找
            if isinstance(df.iloc[j, 0], str) and df.iloc[j, 0].strip().upper() == "STATUS":
                current_status = df.iloc[j + 1, 0] if j + 1 < len(df) else None  # 取 STATUS 下面的内容
                break  # 找到 STATUS 后退出循环

        # 确保数据完整再记录
        if current_rtv and current_amt and current_status:
            records.append([current_rtv, current_amt, current_status])

# **创建 DataFrame 并去除可能的 NaN 值**
output_df = pd.DataFrame(records, columns=["RTV Number", "RTV Disputed Amount", "Status"])
output_df = output_df.dropna().reset_index(drop=True)  # 删除 NaN 值，重置索引

# **删除第一列包含 "Order" 的行**
output_df = output_df[~output_df["RTV Number"].astype(str).str.contains("Order", case=False, na=False)]

# **保存到 Excel**
output_file = r"\\SRV-AD01\Folder Redirection\LZhu\Documents\CTM Worksheet\THD\For_Report\Raw data\outputtest.xlsx"
output_df.to_excel(output_file, index=False, engine="openpyxl")

print(f"数据提取完成，已保存至 {output_file}")
