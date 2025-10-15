# -*- coding: utf-8 -*-
"""
从 THD ChargeBack PDF（第 2 页开始）提取三列：
- SCAC (数字)  -> 取紧随其后的 OMSID 数字
- Po Number    -> 行内的 PO（在 Pro-Tracking# 之后）
- Method Used  -> 'Assigned Method' 之后、金额 '$' 之前的承运人文本
并导出 Excel（UI 选择 PDF 和保存路径）
For Hayan
"""

import re
import pdfplumber
import pandas as pd
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename


# 版式：日期 + SCAC + ProTracking(9-12位) + Po(6-9位) + CustomerOrder + YES/NO + Assigned + MethodUsed + $金额
# 同时兼容 YYYY-MM-DD 与 MM/DD/YYYY 两种日期格式
MAIN_LINE = re.compile(
    r"""
    ^\s*
    (?:\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{4})     # 日期
    \s+[A-Z]{3,5}                                   # SCAC 缩写（如 EXLA/ABFS/CTII/PYLE/SAIA）
    \s+\d{9,12}                                     # Pro-Tracking #
    \s+(\d{6,9})                                    # <-- (1) Po Number
    \s+[A-Z0-9-]+                                   # Customer Order Num
    \s+(?:YES|NO)                                   # Shipped to Store?
    \s+\w+                                          # Assigned Method (e.g., Ground)
    \s+(.+?)                                        # <-- (2) Method Used（直到金额前）
    \s+\$\d+(?:\.\d{2})?                            # Penalty 金额
    \s*$
    """,
    re.VERBOSE,
)

# OMSID 出现在下一行（通常为紧随的一行），示例：OMSID 319614596: (...)
OMSID_RE = re.compile(r"OMSID\s+(\d+):")


def extract_rows(pdf_path: str):
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            if page_idx == 1:  # 跳过第一页封面/表头
                continue

            text = page.extract_text() or ""
            lines = text.splitlines()

            i = 0
            while i < len(lines):
                line = lines[i]

                m = MAIN_LINE.search(line)
                if m:
                    po = m.group(1).strip()           # 保留原始位数（包括可能的前导 0）
                    method = m.group(2).strip()

                    # 在后续 1~3 行里找 OMSID
                    omsid = ""
                    for j in range(1, 4):
                        if i + j < len(lines):
                            m2 = OMSID_RE.search(lines[i + j])
                            if m2:
                                omsid = m2.group(1)
                                break

                    if omsid:
                        rows.append((omsid, po, method))

                i += 1
    return rows


def main():
    root = Tk()
    root.withdraw()

    messagebox.showinfo("开始", "请选择需要处理的 PDF 文件")
    pdf_path = askopenfilename(filetypes=[("PDF 文件", "*.pdf")])
    if not pdf_path:
        return

    data = extract_rows(pdf_path)
    if not data:
        messagebox.showwarning(
            "未提取到数据",
            "可能是列顺序或版式不同。如果有样例页的抽取文本，我可再微调规则。"
        )
        return

    df = pd.DataFrame(data, columns=["SCAC (数字)", "Po Number", "Method Used"])
    df = df.drop_duplicates().reset_index(drop=True)

    save_path = asksaveasfilename(
        title="保存为 Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel 工作簿", "*.xlsx")],
        initialfile="scac_po_method.xlsx"
    )
    if not save_path:
        return

    df.to_excel(save_path, index=False)
    messagebox.showinfo("完成", f"已导出 {len(df)} 行到：\n{save_path}")


if __name__ == "__main__":
    main()
