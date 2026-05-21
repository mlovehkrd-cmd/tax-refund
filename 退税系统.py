import os
import re
import sys
import shutil
import logging
import warnings
import tkinter as tk
from tkinter import filedialog

import pandas as pd
import pdfplumber
from pdf2image import convert_from_path
import pytesseract
from openpyxl import Workbook

warnings.filterwarnings('ignore')

# ================== 自动查找 Tesseract 路径 ==================
def find_tesseract():
    """自动查找 tesseract 可执行文件路径"""
    # 1. 尝试系统 PATH
    path = shutil.which('tesseract')
    if path:
        return path
    # 2. 尝试常见安装位置
    candidates = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        r'D:\Program Files\Tesseract-OCR\tesseract.exe',
        r'D:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    # 3. 没找到，使用默认（或许 pytesseract 能自己找到）
    return 'tesseract'

pytesseract.pytesseract.tesseract_cmd = find_tesseract()
# ============================================================

FORCE_OCR = False
OCR_LANG = 'chi_sim+eng'

CONTRACT_COL_MAP = {
    '合同协议号': ['合同协议号', '合同编号', '合同号'],
    '商品名称': ['商品名称', '货物名称', '品名'],
    '不含税金额': ['不含税金额', '金额', '不含税价'],
    '税额': ['税额', '税金']
}
RECEIPT_COLS = ['开票单位', '合同号', '商品名称', '出口发票金额', '出口报关单号',
                '增值税发票号', '增值税发票金额', '税额', '报关日期', '报关行']

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# -------------------- 工具函数 --------------------
def classify_pdf(filepath):
    """通过关键词识别 PDF 类型"""
    text = ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages[:2]:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except:
        pass
    if not text.strip() or len(text) < 20:
        return "unknown"
    if re.search(r'中华人民共和国海关出口货物报关单|出口货物报关单', text):
        return "customs"
    if re.search(r'增值税专用发票|增值税普通发票', text):
        return "invoice"
    if re.search(r'提单|BILL\s*OF\s*LADING|B/L', text):
        return "bill"
    return "unknown"

def extract_text_with_ocr(pdf_path):
    """OCR 提取 PDF 文本"""
    images = convert_from_path(pdf_path, dpi=300)
    full_text = ""
    for img in images:
        text = pytesseract.image_to_string(img, lang=OCR_LANG)
        full_text += text + "\n"
    return full_text

def extract_text(pdf_path):
    """统一文本提取"""
    if FORCE_OCR:
        return extract_text_with_ocr(pdf_path)
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except:
        pass
    if len(text.strip()) < 50:
        logging.info(f"{os.path.basename(pdf_path)} 疑似扫描件，启动 OCR...")
        text = extract_text_with_ocr(pdf_path)
    return text

def find_col(df, candidates):
    """在 DataFrame 列名中查找包含候选关键词的列"""
    for cand in candidates:
        for col in df.columns:
            if cand in str(col):
                return col
    return None

def extract_customs(text):
    """提取报关单信息"""
    info = {
        '报关单号': '',
        '合同协议号': '',
        '提单号': '',
        '境内发货人': '',
        '生产销售单位': '',
        '申报单位': '',
        '报关日期': '',
        '商品列表': []
    }
    m = re.search(r'报关单编号[：:\s]*(\d{18})', text)
    if m:
        info['报关单号'] = m.group(1)
    m = re.search(r'合同协议号[：:\s]*(\S+)', text)
    if m:
        info['合同协议号'] = m.group(1)
    m = re.search(r'提运单号[：:\s]*(\S+)', text)
    if m:
        info['提单号'] = m.group(1)
    m = re.search(r'境内发货人[：:\s]*(.*?)(生产销售单位|合同协议号)', text, re.DOTALL)
    if m:
        info['境内发货人'] = re.sub(r'\s+', '', m.group(1).strip())
    m = re.search(r'生产销售单位[：:\s]*(.*?)(申报单位|境内发货人)', text, re.DOTALL)
    if m:
        info['生产销售单位'] = re.sub(r'\s+', '', m.group(1).strip())
    m = re.search(r'申报单位[：:\s]*(.*?)(海关批注|放行日期|$)', text, re.DOTALL)
    if m:
        info['申报单位'] = re.sub(r'\s+', '', m.group(1).strip())
    m = re.search(r'申报日期[：:\s]*(\d{4}[-/]\d{2}[-/]\d{2})', text)
    if m:
        info['报关日期'] = m.group(1)
    pattern = r'项号\s+(\d+)\s+(\d{10})\s+([^\d]+?)\s+.*?总价[：:\s]*([\d,]+\.?\d*)'
    matches = re.findall(pattern, text, re.DOTALL)
    for m in matches:
        code = m[1]
        name = m[2].strip()
        name = re.sub(r'^\d{10}\s*', '', name)
        price = m[3].replace(',', '')
        info['商品列表'].append({'code': code, 'name': name, 'total_price': price})
    return info

def extract_invoice(text):
    """提取增值税发票信息"""
    info = {
        '发票号': '', '销售方名称': '', '购买方名称': '',
        '不含税金额': '', '税额': '', '货物名称列表': [], '合同号': ''
    }
    m = re.search(r'发票号码[：:\s]*(\S+)', text)
    if m:
        info['发票号'] = m.group(1)
    m = re.search(r'销售方[：:\s]*名称[：:\s]*(.*?)(纳税人识别号|地址|电话|开户行|$)', text, re.DOTALL)
    if m:
        info['销售方名称'] = m.group(1).strip().replace('\n', '')
    m = re.search(r'购买方[：:\s]*名称[：:\s]*(.*?)(纳税人识别号|地址|电话|$)', text, re.DOTALL)
    if m:
        info['购买方名称'] = m.group(1).strip().replace('\n', '')
    m = re.search(r'金额[：:\s]*([\d,]+\.\d{2})', text)
    if m:
        info['不含税金额'] = m.group(1).replace(',', '')
    m = re.search(r'税额[：:\s]*([\d,]+\.\d{2})', text)
    if m:
        info['税额'] = m.group(1).replace(',', '')
    m = re.search(r'货物或应税劳务[、]服务名称[：:\s]*(.*?)(合计|价税合计)', text, re.DOTALL)
    if m:
        block = m.group(1)
        items = re.findall(r'[\*]?([\u4e00-\u9fa5\w]+)\s', block)
        info['货物名称列表'] = [it.strip() for it in items if it.strip()]
    m = re.search(r'备注[：:\s]*(.*?)(收款人|复核|开票人|$)', text, re.DOTALL)
    if m:
        note = m.group(1)
        cm = re.search(r'(合同[：:\s]*\S+)', note)
        if cm:
            info['合同号'] = cm.group(1).replace('合同', '').replace('：', '').replace(':', '').strip()
    return info

def extract_bill_text(pdf_path):
    return extract_text(pdf_path)

def extract_contracts(files):
    all_rows = []
    for f in files:
        df = pd.read_excel(f, dtype=str)
        contract_col = find_col(df, CONTRACT_COL_MAP['合同协议号'])
        product_col = find_col(df, CONTRACT_COL_MAP['商品名称'])
        amount_col = find_col(df, CONTRACT_COL_MAP['不含税金额'])
        tax_col = find_col(df, CONTRACT_COL_MAP['税额'])
        if not contract_col or not product_col:
            logging.warning(f"内贸合同 {os.path.basename(f)} 列名不匹配，跳过")
            continue
        df = df[[contract_col, product_col, amount_col, tax_col]].copy()
        df.columns = ['合同协议号', '商品名称', '不含税金额', '税额']
        all_rows.append(df)
    if all_rows:
        combined = pd.concat(all_rows, ignore_index=True)
    else:
        combined = pd.DataFrame(columns=['合同协议号', '商品名称', '不含税金额', '税额'])
    return combined

# -------------------- 主流程 --------------------
def main(folder_path):
    os.chdir(folder_path)
    all_files = os.listdir(folder_path)
    pdf_files = [f for f in all_files if f.lower().endswith('.pdf')]
    xlsx_files = [f for f in all_files if f.lower().endswith('.xlsx')]

    customs_pdf = None
    invoice_pdfs = []
    bill_pdf = None
    contract_xlsxs = []
    receipt_xlsx = None

    for pdf_file in pdf_files:
        full_path = os.path.join(folder_path, pdf_file)
        ftype = classify_pdf(full_path)
        if ftype == 'customs':
            customs_pdf = full_path
        elif ftype == 'invoice':
            invoice_pdfs.append(full_path)
        elif ftype == 'bill':
            bill_pdf = full_path
        else:
            text = extract_text(full_path)
            if '出口货物报关单' in text:
                customs_pdf = full_path
            elif '增值税专用发票' in text:
                invoice_pdfs.append(full_path)
            elif '提单' in text:
                bill_pdf = full_path

    for xlsx_file in xlsx_files:
        full_path = os.path.join(folder_path, xlsx_file)
        df = pd.read_excel(full_path, nrows=0)
        if all(col in df.columns for col in RECEIPT_COLS):
            receipt_xlsx = full_path
        else:
            if find_col(df, CONTRACT_COL_MAP['合同协议号']):
                contract_xlsxs.append(full_path)

    if not customs_pdf:
        logging.error("未找到报关单 PDF！")
        return
    if not receipt_xlsx:
        logging.error("未找到签收表模板 Excel！")
        return

    customs_text = extract_text(customs_pdf)
    customs_data = extract_customs(customs_text)
    if not customs_data['商品列表']:
        logging.error("报关单商品提取失败，请检查格式。")
        return

    invoices = []
    for inv_path in invoice_pdfs:
        text = extract_text(inv_path)
        inv_data = extract_invoice(text)
        invoices.append(inv_data)

    bill_text = extract_bill_text(bill_pdf) if bill_pdf else ""
    contracts_df = extract_contracts(contract_xlsxs)

    deviation_log = []

    # 提单号校验
    bill_no = customs_data['提单号']
    if bill_no:
        if bill_no not in bill_text:
            deviation_log.append({
                '报关单号': customs_data['报关单号'],
                '商品行序号': '', '校验字段': '提单号',
                '期望值': bill_no, '实际值': '提单中未找到',
                '偏差类型': '缺失', '涉及文件名': os.path.basename(bill_pdf) if bill_pdf else ''
            })
    else:
        deviation_log.append({
            '报关单号': customs_data['报关单号'],
            '商品行序号': '', '校验字段': '提单号',
            '期望值': '', '实际值': '报关单未提取到',
            '偏差类型': '提取失败', '涉及文件名': '报关单'
        })

    # 合同协议号三方比对
    contract_customs = customs_data['合同协议号']
    contract_invoices = [inv['合同号'] for inv in invoices if inv['合同号']]
    contract_contracts = contracts_df['合同协议号'].unique().tolist() if not contracts_df.empty else []

    if contract_contracts and contract_customs not in contract_contracts:
        deviation_log.append({
            '报关单号': customs_data['报关单号'],
            '商品行序号': '', '校验字段': '合同协议号',
            '期望值': f"报关单: {contract_customs}",
            '实际值': f"内贸合同列表: {contract_contracts}",
            '偏差类型': '不一致', '涉及文件名': '内贸合同'
        })
    for inv_no in contract_invoices:
        if inv_no != contract_customs:
            deviation_log.append({
                '报关单号': customs_data['报关单号'],
                '商品行序号': '', '校验字段': '合同协议号',
                '期望值': f"报关单: {contract_customs}",
                '实际值': f"发票: {inv_no}",
                '偏差类型': '不一致', '涉及文件名': '增值税发票'
            })

    # 境内发货人/生产销售单位 vs 购买方名称
    buyer_name = invoices[0]['购买方名称'] if invoices else ''
    shipper = customs_data['境内发货人']
    producer = customs_data['生产销售单位']
    if shipper and buyer_name and shipper != buyer_name:
        deviation_log.append({
            '报关单号': customs_data['报关单号'],
            '商品行序号': '', '校验字段': '境内发货人',
            '期望值': shipper, '实际值': f"发票购买方: {buyer_name}",
            '偏差类型': '不一致', '涉及文件名': '增值税发票'
        })
    if producer and buyer_name and producer != buyer_name:
        deviation_log.append({
            '报关单号': customs_data['报关单号'],
            '商品行序号': '', '校验字段': '生产销售单位',
            '期望值': producer, '实际值': f"发票购买方: {buyer_name}",
            '偏差类型': '不一致', '涉及文件名': '增值税发票'
        })

    receipt_rows = []
    contract_products = contracts_df['商品名称'].dropna().unique().tolist() if not contracts_df.empty else []
    all_invoice_products = []
    for inv in invoices:
        all_invoice_products.extend(inv['货物名称列表'])

    for idx, item in enumerate(customs_data['商品列表']):
        item_name = item['name']
        if contract_products and item_name not in contract_products:
            deviation_log.append({
                '报关单号': customs_data['报关单号'],
                '商品行序号': idx+1, '校验字段': '商品名称',
                '期望值': item_name, '实际值': '内贸合同中未找到完全一致项',
                '偏差类型': '不一致/缺失', '涉及文件名': '内贸合同'
            })
        if all_invoice_products and item_name not in all_invoice_products:
            deviation_log.append({
                '报关单号': customs_data['报关单号'],
                '商品行序号': idx+1, '校验字段': '商品名称',
                '期望值': item_name, '实际值': '增值税发票中未找到完全一致项',
                '偏差类型': '不一致/缺失', '涉及文件名': '增值税发票'
            })

        matched_inv = None
        for inv in invoices:
            if item_name in inv['货物名称列表']:
                matched_inv = inv
                break
        if not matched_inv:
            deviation_log.append({
                '报关单号': customs_data['报关单号'],
                '商品行序号': idx+1, '校验字段': '增值税发票匹配',
                '期望值': item_name, '实际值': '无发票商品匹配',
                '偏差类型': '缺失', '涉及文件名': '增值税发票'
            })
            receipt_rows.append({
                '开票单位': '缺失',
                '合同号': customs_data['合同协议号'] if customs_data['合同协议号'] else '缺失',
                '商品名称': item_name,
                '出口发票金额': item['total_price'],
                '出口报关单号': customs_data['报关单号'] if customs_data['报关单号'] else '缺失',
                '增值税发票号': '缺失',
                '增值税发票金额': '缺失',
                '税额': '缺失',
                '报关日期': customs_data['报关日期'] if customs_data['报关日期'] else '缺失',
                '报关行': customs_data['申报单位'] if customs_data['申报单位'] else '缺失'
            })
        else:
            receipt_rows.append({
                '开票单位': matched_inv['销售方名称'] if matched_inv['销售方名称'] else '缺失',
                '合同号': customs_data['合同协议号'] if customs_data['合同协议号'] else '缺失',
                '商品名称': item_name,
                '出口发票金额': item['total_price'],
                '出口报关单号': customs_data['报关单号'] if customs_data['报关单号'] else '缺失',
                '增值税发票号': matched_inv['发票号'] if matched_inv['发票号'] else '缺失',
                '增值税发票金额': matched_inv['不含税金额'] if matched_inv['不含税金额'] else '缺失',
                '税额': matched_inv['税额'] if matched_inv['税额'] else '缺失',
                '报关日期': customs_data['报关日期'] if customs_data['报关日期'] else '缺失',
                '报关行': customs_data['申报单位'] if customs_data['申报单位'] else '缺失'
            })

    receipt_df = pd.DataFrame(receipt_rows)
    receipt_df = receipt_df[RECEIPT_COLS]
    output_receipt = os.path.join(folder_path, '已核验签收表.xlsx')
    receipt_df.to_excel(output_receipt, index=False)
    logging.info(f"已核验签收表已生成: {output_receipt}")

    if deviation_log:
        log_df = pd.DataFrame(deviation_log)
        log_output = os.path.join(folder_path, '偏差日志.xlsx')
        log_df.to_excel(log_output, index=False)
        logging.info(f"偏差日志已生成: {log_output}")
    else:
        logging.info("无偏差记录，未生成日志文件。")

    print("处理完成！")

# -------------------- 启动入口 --------------------
if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title='请选择包含所有文件（报关单、发票、提单、合同、签收表）的文件夹')
    if not folder:
        print("未选择文件夹，程序退出。")
    else:
        print(f"已选择文件夹：{folder}")
        main(folder)
    root.destroy()