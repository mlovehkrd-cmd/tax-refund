import os, sys, re, shutil, logging, traceback, warnings
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pdfplumber
from pdf2image import convert_from_path
import pytesseract
from openpyxl import Workbook

warnings.filterwarnings('ignore')

# ---------- 自动查找 Tesseract ----------
def find_tesseract():
    path = shutil.which('tesseract')
    if path:
        return path
    candidates = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        r'D:\Program Files\Tesseract-OCR\tesseract.exe',
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return 'tesseract'

pytesseract.pytesseract.tesseract_cmd = find_tesseract()
FORCE_OCR = False
OCR_LANG = 'chi_sim+eng'

# 列映射（不变，略，保持原脚本中的映射）
CONTRACT_COL_MAP = {
    '合同协议号': ['合同协议号','合同编号','合同号'],
    '商品名称': ['商品名称','货物名称','品名'],
    '不含税金额': ['不含税金额','金额','不含税价'],
    '税额': ['税额','税金']
}
RECEIPT_COLS = ['开票单位','合同号','商品名称','出口发票金额','出口报关单号',
                '增值税发票号','增值税发票金额','税额','报关日期','报关行']

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ---------- 所有工具函数与原脚本完全相同 ----------
# (为了节省篇幅，此处省略，请将你原脚本中 classify_pdf, extract_text_with_ocr,
#  extract_text, find_col, extract_customs, extract_invoice, extract_bill_text,
#  extract_contracts 等函数完整粘贴在此处)
# ---------- 工具函数结束 ----------

def main(folder_path):
    """主处理逻辑，与原 main 完全相同"""
    os.chdir(folder_path)
    # ... 将原 main 函数内容完整复制过来，注意缩进 ...
    print("处理完成！")

if __name__ == '__main__':
    try:
        # 初始化 Tk 根窗口（隐藏）
        root = tk.Tk()
        root.withdraw()
        # 弹出文件夹选择
        folder = filedialog.askdirectory(title='请选择包含所有文件的文件夹')
        if not folder:
            print("未选择文件夹，程序退出。")
            sys.exit(0)
        print(f"已选择文件夹：{folder}")
        main(folder)
        # 完成后弹出提示
        messagebox.showinfo("完成", "退税系统处理完毕！\n请查看已核验签收表和偏差日志。")
    except Exception as e:
        # 错误时弹出消息框
        error_msg = f"程序运行出错：\n{traceback.format_exc()}"
        print(error_msg)
        messagebox.showerror("错误", error_msg)
    finally:
        root.destroy()
