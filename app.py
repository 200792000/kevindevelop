from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl import load_workbook
from datetime import datetime, timedelta
import shutil, re, os, tempfile, zipfile, subprocess

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH     = os.path.join(os.path.dirname(__file__), 'template.xlsx')
TEMPLATE_BU1_PATH = os.path.join(os.path.dirname(__file__), 'template_bu1.xlsx')

WEEKDAY_MAP = ['一','二','三','四','五','六','日']

# ── 計算已發送截止日 ──────────────────────────────────────────
def get_sent_deadline(today=None):
    if today is None:
        today = datetime.today()
    # 找本週週一
    days_since_monday = today.weekday()  # 0=週一
    this_monday = today - timedelta(days=days_since_monday)
    this_monday = this_monday.replace(hour=0, minute=0, second=0, microsecond=0)
    # 本週週一發的是下下週（+14天起的那週週日 = +20天）
    deadline = this_monday + timedelta(days=20)
    return deadline

# ── 解析班表 ─────────────────────────────────────────────────
def load_schedule(file_obj):
    wb = openpyxl.load_workbook(file_obj)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hi = -1
    for i, row in enumerate(rows):
        vals = [str(v or '').strip() for v in row]
        if '店號' in vals and '日期' in vals:
            hi = i; break
    if hi < 0:
        raise ValueError('找不到欄位標題（店號/日期）')
    headers = [str(v or '').strip() for v in rows[hi]]
    idx_shop = headers.index('店號')
    idx_name = headers.index('店名')
    idx_date = headers.index('日期')
    idx_noon = headers.index('午別')
    map_ = {}
    first_date = None
    for row in rows[hi+1:]:
        shop = str(row[idx_shop] or '').strip()
        if not shop: continue
        d = str(row[idx_date] or '').strip()
        n = str(row[idx_noon] or '').strip()
        noon_full = '上午' if n == '上' else ('下午' if n == '下' else n)
        if d:
            if first_date is None: first_date = d
            map_.setdefault(shop, {'店名': str(row[idx_name] or '').strip(), 'dates': set(), 'noon': {}})
            map_[shop]['dates'].add(d)
            map_[shop]['noon'][d] = noon_full
    return map_, first_date

def parse_date(s):
    parts = s.split('/')
    return datetime(int(parts[0]), int(parts[1]), int(parts[2]))

def extract_version(filename):
    m = re.search(r'(\d{8})', filename)
    return m.group(1) if m else ''

# ── 比對差異並分類 ────────────────────────────────────────────
def compare_and_classify(m1, m2, deadline):
    異動通知 = []
    補1通知  = []
    for k in sorted(set(list(m1.keys()) + list(m2.keys()))):
        a, b = m1.get(k), m2.get(k)
        if not a or not b: continue
        s1 = '|'.join(d + '_' + a['noon'].get(d,'') for d in sorted(a['dates']))
        s2 = '|'.join(d + '_' + b['noon'].get(d,'') for d in sorted(b['dates']))
        if s1 != s2:
            orig = sorted(a['dates'])[0]
            chg  = sorted(b['dates'])[0]
            orig_dt = parse_date(orig)
            diff = {
                '店號':   k,
                '店名':   a['店名'],
                '原訂日': orig,
                '原訂午': a['noon'].get(orig, ''),
                '異動日': chg,
                '異動午': b['noon'].get(chg, ''),
            }
            # 版本一（原訂）日期 <= 截止日 → 已收到通知 → 異動通知書
            # 版本一（原訂）日期 >  截止日 → 尚未收到   → 補1通知書
            if orig_dt <= deadline:
                異動通知.append(diff)
            else:
                補1通知.append(diff)
    return 異動通知, 補1通知

# ── 產出異動通知書 Excel ───────────────────────────────────────
def generate_notice_excel(diffs, version_full):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    tmp.close()
    shutil.copy(TEMPLATE_PATH, tmp.name)
    wb = load_workbook(tmp.name)
    ws = wb['行程異動連絡單']
    DATE_FMT = 'M"月"D"日"'
    REASON = '因內部聯絡條，故行程異動'

    merges = [str(m) for m in list(ws.merged_cells.ranges)]
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))

    for i, d in enumerate(diffs):
        r_d1 = 6 + i * 5
        r_d2 = 8 + i * 5
        ws.cell(row=r_d1, column=1).value = d['店號']
        c = ws.cell(row=r_d1, column=4)
        c.value = parse_date(d['原訂日']); c.number_format = DATE_FMT
        ws.cell(row=r_d1, column=5).value = d['原訂午']
        ws.cell(row=r_d1, column=6).value = REASON
        ws.cell(row=r_d1, column=8).value = None
        ws.cell(row=r_d1, column=9).value = None
        c2 = ws.cell(row=r_d2, column=4)
        c2.value = parse_date(d['異動日']); c2.number_format = DATE_FMT
        ws.cell(row=r_d2, column=5).value = d['異動午']
        ws.cell(row=r_d2, column=8).value = version_full

    for i in range(len(diffs), 10):
        r_d1 = 6 + i * 5; r_d2 = 8 + i * 5
        for col in [1, 4, 5, 6, 8, 9]:
            ws.cell(row=r_d1, column=col).value = None
            ws.cell(row=r_d2, column=col).value = None

    for coord in merges:
        ws.merge_cells(coord)

    wb.save(tmp.name)
    return tmp.name

# ── 產出補1通知書 PDF（每間店一個）────────────────────────────
def generate_bu1_pdfs(diffs):
    pdf_files = []
    tmpdir = tempfile.mkdtemp()

    for d in diffs:
        wb = load_workbook(TEMPLATE_BU1_PATH)
        ws = wb['補實地盤點通知書']

        chg_dt = parse_date(d['異動日'])
        roc_year = chg_dt.year - 1911

        ws['B5'].value = d['店名']
        ws['A9'].value = str(roc_year)
        ws['D9'].value = str(chg_dt.month)
        ws['F9'].value = str(chg_dt.day)
        ws['I9'].value = WEEKDAY_MAP[chg_dt.weekday()]

        # 存成暫存 xlsx
        xlsx_path = os.path.join(tmpdir, f"{d['店號']}_{d['店名']}.xlsx")
        wb.save(xlsx_path)

        # 用 LibreOffice 轉 PDF
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf',
             xlsx_path, '--outdir', tmpdir],
            capture_output=True, timeout=30
        )
        pdf_path = xlsx_path.replace('.xlsx', '.pdf')
        if os.path.exists(pdf_path):
            # 改成正式檔名
            final_name = f"補1_{d['店號']}_{d['店名']}.pdf"
            final_path = os.path.join(tmpdir, final_name)
            os.rename(pdf_path, final_path)
            pdf_files.append((final_name, final_path))

    return pdf_files, tmpdir

# ── API 路由 ──────────────────────────────────────────────────
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/compare', methods=['POST'])
def api_compare():
    if 'file1' not in request.files or 'file2' not in request.files:
        return jsonify({'error': '請上傳兩個班表檔案'}), 400
    f1 = request.files['file1']
    f2 = request.files['file2']
    try:
        m1, first_date = load_schedule(f1)
        m2, _          = load_schedule(f2)
        ver8   = extract_version(f1.filename)
        month  = str(int(first_date.split('/')[1])).zfill(2)
        version_full = month + '-' + ver8
        deadline = get_sent_deadline()
        異動通知, 補1通知 = compare_and_classify(m1, m2, deadline)
        return jsonify({
            'version':        version_full,
            'month':          month,
            'v1_count':       len(m1),
            'v2_count':       len(m2),
            'deadline':       deadline.strftime('%Y/%m/%d'),
            'notice_count':   len(異動通知),
            'bu1_count':      len(補1通知),
            'notice_diffs':   異動通知,
            'bu1_diffs':      補1通知,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generate_notice', methods=['POST'])
def api_generate_notice():
    if 'file1' not in request.files or 'file2' not in request.files:
        return jsonify({'error': '請上傳兩個班表檔案'}), 400
    f1 = request.files['file1']
    f2 = request.files['file2']
    try:
        m1, first_date = load_schedule(f1)
        m2, _          = load_schedule(f2)
        ver8   = extract_version(f1.filename)
        month  = str(int(first_date.split('/')[1])).zfill(2)
        version_full = month + '-' + ver8
        deadline = get_sent_deadline()
        異動通知, _ = compare_and_classify(m1, m2, deadline)
        if not 異動通知:
            return jsonify({'error': '無異動通知書資料'}), 400
        out_path = generate_notice_excel(異動通知, version_full)
        today = datetime.today().strftime('%Y%m%d')
        filename = f'盤點行程異動聯絡單_{version_full}_{today}.xlsx'
        return send_file(out_path, as_attachment=True, download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generate_bu1', methods=['POST'])
def api_generate_bu1():
    if 'file1' not in request.files or 'file2' not in request.files:
        return jsonify({'error': '請上傳兩個班表檔案'}), 400
    f1 = request.files['file1']
    f2 = request.files['file2']
    try:
        m1, first_date = load_schedule(f1)
        m2, _          = load_schedule(f2)
        ver8   = extract_version(f1.filename)
        month  = str(int(first_date.split('/')[1])).zfill(2)
        version_full = month + '-' + ver8
        deadline = get_sent_deadline()
        _, 補1通知 = compare_and_classify(m1, m2, deadline)
        if not 補1通知:
            return jsonify({'error': '無補1通知書資料'}), 400
        pdf_files, tmpdir = generate_bu1_pdfs(補1通知)
        if not pdf_files:
            return jsonify({'error': 'PDF 產出失敗'}), 500

        # 打包成 ZIP
        zip_tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
        zip_tmp.close()
        with zipfile.ZipFile(zip_tmp.name, 'w') as zf:
            for name, path in pdf_files:
                zf.write(path, name)

        today = datetime.today().strftime('%Y%m%d')
        filename = f'補1通知書_{version_full}_{today}.zip'
        return send_file(zip_tmp.name, as_attachment=True, download_name=filename,
            mimetype='application/zip')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
