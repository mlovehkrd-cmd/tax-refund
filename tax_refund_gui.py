import os, re, sys, shutil, logging, traceback, warnings, threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import pdfplumber
from pdf2image import convert_from_path
import pytesseract

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

CONTRACT_COL_MAP = {
    '合同协议号': ['合同协议号','合同编号','合同号'],
    '商品名称': ['商品名称','货物名称','品名'],
    '不含税金额': ['不含税金额','金额','不含税价'],
    '税额': ['税额','税金']
}
RECEIPT_COLS = ['开票单位','合同号','商品名称','出口发票金额','出口报关单号',
                '增值税发票号','增值税发票金额','税额','报关日期','报关行']

# -------------------- 工具函数 --------------------
def classify_pdf(filepath):
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
    images = convert_from_path(pdf_path, dpi=300)
    full_text = ""
    for img in images:
        text = pytesseract.image_to_string(img, lang=OCR_LANG)
        full_text += text + "\n"
    return full_text

def extract_text(pdf_path):
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
        log_msg(f"⚠️ {os.path.basename(pdf_path)} 疑似扫描件，启动 OCR...")
        text = extract_text_with_ocr(pdf_path)
    return text

def find_col(df, candidates):
    for cand in candidates:
        for col in df.columns:
            if cand in str(col):
                return col
    return None

def extract_customs(text):
    info = {
        '报关单号': '', '合同协议号': '', '提单号': '',
        '境内发货人': '', '生产销售单位': '', '申报单位': '',
        '报关日期': '', '商品列表': []
    }
    m = re.search(r'报关单编号[：:\s]*(\d{18})', text)
    if m: info['报关单号'] = m.group(1)
    m = re.search(r'合同协议号[：:\s]*(\S+)', text)
    if m: info['合同协议号'] = m.group(1)
    m = re.search(r'提运单号[：:\s]*(\S+)', text)
    if m: info['提单号'] = m.group(1)
    m = re.search(r'境内发货人[：:\s]*(.*?)(生产销售单位|合同协议号)', text, re.DOTALL)
    if m: info['境内发货人'] = re.sub(r'\s+', '', m.group(1).strip())
    m = re.search(r'生产销售单位[：:\s]*(.*?)(申报单位|境内发货人)', text, re.DOTALL)
    if m: info['生产销售单位'] = re.sub(r'\s+', '', m.group(1).strip())
    m = re.search(r'申报单位[：:\s]*(.*?)(海关批注|放行日期|$)', text, re.DOTALL)
    if m: info['申报单位'] = re.sub(r'\s+', '', m.group(1).strip())
    m = re.search(r'申报日期[：:\s]*(\d{4}[-/]\d{2}[-/]\d{2})', text)
    if m: info['报关日期'] = m.group(1)
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
    info = {
        '发票号': '', '销售方名称': '', '购买方名称': '',
        '不含税金额': '', '税额': '', '货物名称列表': [], '合同号': ''
    }
    m = re.search(r'发票号码[：:\s]*(\S+)', text)
    if m: info['发票号'] = m.group(1)
    m = re.search(r'销售方[：:\s]*名称[：:\s]*(.*?)(纳税人识别号|地址|电话|开户行|$)', text, re.DOTALL)
    if m: info['销售方名称'] = m.group(1).strip().replace('\n', '')
    m = re.search(r'购买方[：:\s]*名称[：:\s]*(.*?)(纳税人识别号|地址|电话|$)', text, re.DOTALL)
    if m: info['购买方名称'] = m.group(1).strip().replace('\n', '')
    m = re.search(r'金额[：:\s]*([\d,]+\.\d{2})', text)
    if m: info['不含税金额'] = m.group(1).replace(',', '')
    m = re.search(r'税额[：:\s]*([\d,]+\.\d{2})', text)
    if m: info['税额'] = m.group(1).replace(',', '')
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
            log_msg(f"⚠️ 内贸合同 {os.path.basename(f)} 列名不匹配，跳过")
            continue
        df = df[[contract_col, product_col, amount_col, tax_col]].copy()
        df.columns = ['合同协议号', '商品名称', '不含税金额', '税额']
        all_rows.append(df)
    if all_rows:
        return pd.concat(all_rows, ignore_index=True)
    else:
        return pd.DataFrame(columns=['合同协议号', '商品名称', '不含税金额', '税额'])

# -------------------- 日志与界面交互 --------------------
gui_log = None

def log_msg(msg):
    """在 GUI 日志框添加消息"""
    if gui_log:
        gui_log.insert(tk.END, msg + '\n')
        gui_log.see(tk.END)
        gui_log.update()
    print(msg)

# -------------------- 主流程 --------------------
def process_folder(folder_path):
    try:
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
            log_msg("❌ 未找到报关单 PDF！")
            return
        if not receipt_xlsx:
            log_msg("❌ 未找到签收表模板 Excel！")
            return

        log_msg("🔍 开始提取报关单信息...")
        customs_text = extract_text(customs_pdf)
        customs_data = extract_customs(customs_text)
        if not customs_data['商品列表']:
            log_msg("❌ 报关单商品提取失败，请检查格式。")
            return

        log_msg("🧾 读取增值税发票...")
        invoices = []
        for inv_path in invoice_pdfs:
            text = extract_text(inv_path)
            inv_data = extract_invoice(text)
            invoices.append(inv_data)

        log_msg("📄 读取提单...")
        bill_text = extract_bill_text(bill_pdf) if bill_pdf else ""
        log_msg("📑 读取内贸合同...")
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

        log_msg("🔁 开始逐项比对商品...")
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

        log_msg("💾 生成已核验签收表...")
        receipt_df = pd.DataFrame(receipt_rows)
        receipt_df = receipt_df[RECEIPT_COLS]
        output_receipt = os.path.join(folder_path, '已核验签收表.xlsx')
        receipt_df.to_excel(output_receipt, index=False)
        log_msg(f"✅ 已核验签收表已保存：{output_receipt}")

        if deviation_log:
            log_df = pd.DataFrame(deviation_log)
            log_output = os.path.join(folder_path, '偏差日志.xlsx')
            log_df.to_excel(log_output, index=False)
            log_msg(f"⚠️ 偏差日志已保存：{log_output}")
        else:
            log_msg("✅ 无偏差记录。")

        log_msg("🎉 处理完成！")
        return True

    except Exception as e:
        log_msg(f"💥 发生异常：{traceback.format_exc()}")
        return False

# -------------------- GUI 界面 --------------------
class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("十三部退税系统 v2.0")
        self.master.geometry("700x550")
        self.pack(fill=tk.BOTH, expand=True)
        self.create_widgets()

    def create_widgets(self):
        # 文件夹选择
        frame_top = tk.Frame(self)
        frame_top.pack(pady=10, padx=10, fill=tk.X)
        self.folder_path = tk.StringVar()
        tk.Label(frame_top, text="文件夹：").pack(side=tk.LEFT)
        tk.Entry(frame_top, textvariable=self.folder_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_top, text="选择文件夹", command=self.select_folder).pack(side=tk.LEFT)
        # 运行按钮
        self.btn_run = tk.Button(self, text="▶️ 开始处理", command=self.start_processing, bg="#4CAF50", fg="white", font=("Arial", 12))
        self.btn_run.pack(pady=10)
        # 日志区域
        frame_log = tk.Frame(self)
        frame_log.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        tk.Label(frame_log, text="处理日志：").pack(anchor=tk.W)
        global gui_log
        gui_log = scrolledtext.ScrolledText(frame_log, height=15, state='normal')
        gui_log.pack(fill=tk.BOTH, expand=True)
        # 底部按钮：打开结果
        frame_bottom = tk.Frame(self)
        frame_bottom.pack(pady=10)
        tk.Button(frame_bottom, text="打开已核验签收表", command=self.open_receipt).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_bottom, text="打开偏差日志", command=self.open_log).pack(side=tk.LEFT, padx=5)

    def select_folder(self):
        folder = filedialog.askdirectory(title='请选择包含所有文件的文件夹')
        if folder:
            self.folder_path.set(folder)

    def start_processing(self):
        folder = self.folder_path.get().strip()
        if not folder:
            messagebox.showwarning("提示", "请先选择文件夹！")
            return
        self.btn_run.config(state=tk.DISABLED, text="处理中...")
        gui_log.delete(1.0, tk.END)
        # 在线程中运行，避免界面卡死
        def task():
            success = process_folder(folder)
            self.btn_run.config(state=tk.NORMAL, text="▶️ 开始处理")
            if success:
                messagebox.showinfo("完成", "退税系统处理完毕！")
            else:
                messagebox.showerror("错误", "处理过程中发生错误，请查看日志。")
        threading.Thread(target=task, daemon=True).start()

    def open_receipt(self):
        folder = self.folder_path.get().strip()
        if folder:
            path = os.path.join(folder, '已核验签收表.xlsx')
            if os.path.exists(path):
                os.startfile(path)
            else:
                messagebox.showinfo("提示", "签收表尚未生成，请先处理。")
        else:
            messagebox.showwarning("提示", "请先选择文件夹。")

    def open_log(self):
        folder = self.folder_path.get().strip()
        if folder:
            path = os.path.join(folder, '偏差日志.xlsx')
            if os.path.exists(path):
                os.startfile(path)
            else:
                messagebox.showinfo("提示", "偏差日志尚未生成。")
        else:
            messagebox.showwarning("提示", "请先选择文件夹。")

if __name__ == '__main__':
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
