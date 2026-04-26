from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl import load_workbook
from datetime import datetime, timedelta
import shutil, re, os, tempfile, zipfile
import pymupdf
import requests

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH     = os.path.join(os.path.dirname(__file__), 'template.xlsx')
TEMPLATE_BU1_PATH = os.path.join(os.path.dirname(__file__), 'template_bu1.pdf')
WEEKDAY_MAP = ['一','二','三','四','五','六','日']
MAX_PER_SHEET = 10
RENDER_SERVICE_ID = 'srv-d7mecdb7uimc73crjh6g'

# ── 計算已發送截止日 ──────────────────────────────────────────
def get_sent_deadline(today=None):
    if today is None:
        today = datetime.today()
    days_since_monday = today.weekday()
    this_monday = today - timedelta(days=days_since_monday)
    this_monday = this_monday.replace(hour=0, minute=0, second=0, microsecond=0)
    return this_monday + timedelta(days=20)

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

# ── 比對並分類 ────────────────────────────────────────────────
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
            if orig_dt <= deadline:
                異動通知.append(diff)
            else:
                補1通知.append(diff)
    return 異動通知, 補1通知

# ── 產出單份異動通知書 Excel ──────────────────────────────────
def generate_single_notice(diffs, version_full):
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

    for i in range(len(diffs), MAX_PER_SHEET):
        r_d1 = 6 + i * 5; r_d2 = 8 + i * 5
        for col in [1, 4, 5, 6, 8, 9]:
            ws.cell(row=r_d1, column=col).value = None
            ws.cell(row=r_d2, column=col).value = None

    for coord in merges:
        ws.merge_cells(coord)

    wb.save(tmp.name)
    return tmp.name

# ── 產出異動通知書（超過10店自動分份）────────────────────────
def generate_notice_files(diffs, version_full):
    chunks = [diffs[i:i+MAX_PER_SHEET] for i in range(0, len(diffs), MAX_PER_SHEET)]
    files = []
    today = datetime.today().strftime('%Y%m%d')
    for idx, chunk in enumerate(chunks):
        path = generate_single_notice(chunk, version_full)
        if len(chunks) == 1:
            fname = f'盤點行程異動聯絡單_{version_full}_{today}.xlsx'
        else:
            fname = f'盤點行程異動聯絡單_{version_full}_{today}_第{idx+1}份.xlsx'
        files.append((fname, path))
    return files

# ── 產出補1通知書 PDF ────────────────────────────────────────
def make_bu1_pdf(store_name, chg_date, chg_noon, out_path):
    roc_year = str(chg_date.year - 1911)
    month    = str(chg_date.month)
    day      = str(chg_date.day)
    weekday  = WEEKDAY_MAP[chg_date.weekday()]

    doc = pymupdf.open(TEMPLATE_BU1_PATH)
    page = doc[0]
    blocks = page.get_text("dict")

    replacements = []
    for block in blocks['blocks']:
        if 'lines' not in block: continue
        for line in block['lines']:
            for span in line['spans']:
                if span['color'] == 255 and span['text'].strip():
                    replacements.append(span)

    for span in replacements:
        x, y = span['origin']
        old_text = span['text'].strip()
        size = span['size']

        if 150 < y < 170 and old_text not in ['114','5','一','上午','下午']:
            new_text = store_name
        elif 210 < y < 220:
            if x < 50:      new_text = roc_year
            elif x < 120:   new_text = month
            elif x < 175:   new_text = day
            elif x < 265:   new_text = weekday
            else:            new_text = chg_noon
        else:
            continue

        bbox = span['bbox']
        rect = pymupdf.Rect(bbox[0]-1, bbox[1]-1, bbox[2]+8, bbox[3]+1)
        page.draw_rect(rect, color=(1,1,1), fill=(1,1,1))
        page.insert_text((x, y), new_text, fontsize=size, color=(0,0,1), fontname='china-t')

    doc.save(out_path)

# ── API：健康檢查 ─────────────────────────────────────────────
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

# ── API：比對班表 ─────────────────────────────────────────────
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
        notice_files = (len(異動通知) + MAX_PER_SHEET - 1) // MAX_PER_SHEET if 異動通知 else 0
        return jsonify({
            'version':       version_full,
            'month':         month,
            'v1_count':      len(m1),
            'v2_count':      len(m2),
            'deadline':      deadline.strftime('%Y/%m/%d'),
            'notice_count':  len(異動通知),
            'bu1_count':     len(補1通知),
            'notice_diffs':  異動通知,
            'bu1_diffs':     補1通知,
            'notice_files':  notice_files,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── API：產出異動通知書 ───────────────────────────────────────
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

        files = generate_notice_files(異動通知, version_full)
        today = datetime.today().strftime('%Y%m%d')

        if len(files) == 1:
            return send_file(files[0][1], as_attachment=True,
                download_name=files[0][0],
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            zip_tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
            zip_tmp.close()
            with zipfile.ZipFile(zip_tmp.name, 'w') as zf:
                for fname, fpath in files:
                    zf.write(fpath, fname)
            return send_file(zip_tmp.name, as_attachment=True,
                download_name=f'盤點行程異動聯絡單_{version_full}_{today}.zip',
                mimetype='application/zip')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── API：產出補1通知書 ────────────────────────────────────────
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

        tmpdir = tempfile.mkdtemp()
        pdf_files = []
        for d in 補1通知:
            chg_dt = parse_date(d['異動日'])
            mmdd = f"{chg_dt.month:02d}{chg_dt.day:02d}"
            fname = f"(補1)實地盤點實施通知書-{d['店名']}{mmdd}.pdf"
            fpath = os.path.join(tmpdir, fname)
            make_bu1_pdf(d['店名'], chg_dt, d['異動午'], fpath)
            pdf_files.append((fname, fpath))

        today = datetime.today().strftime('%Y%m%d')
        zip_tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
        zip_tmp.close()
        with zipfile.ZipFile(zip_tmp.name, 'w') as zf:
            for fname, fpath in pdf_files:
                zf.write(fpath, fname)

        return send_file(zip_tmp.name, as_attachment=True,
            download_name=f'補1通知書_{version_full}_{today}.zip',
            mimetype='application/zip')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── 管理員：驗證密碼 ──────────────────────────────────────────
@app.route('/admin/verify', methods=['POST'])
def admin_verify():
    data = request.get_json()
    password = data.get('password', '')
    admin_pw = os.environ.get('ADMIN_PASSWORD', '')
    if password == admin_pw:
        return jsonify({'ok': True})
    return jsonify({'ok': False, 'error': '密碼錯誤'}), 401

# ── 管理員：更換密碼 ──────────────────────────────────────────
@app.route('/admin/change_password', methods=['POST'])
def admin_change_password():
    data = request.get_json()
    old_pw  = data.get('old_password', '')
    new_pw  = data.get('new_password', '')
    admin_pw = os.environ.get('ADMIN_PASSWORD', '')
    render_api_key = os.environ.get('RENDER_API_KEY', '')

    if old_pw != admin_pw:
        return jsonify({'error': '舊密碼錯誤'}), 401
    if not new_pw or len(new_pw) < 6:
        return jsonify({'error': '新密碼至少需要 6 個字元'}), 400

    headers = {
        'Authorization': f'Bearer {render_api_key}',
        'Content-Type': 'application/json',
    }
    url = f'https://api.render.com/v1/services/{RENDER_SERVICE_ID}/env-vars'
    payload = [
        {'key': 'ADMIN_PASSWORD', 'value': new_pw},
        {'key': 'RENDER_API_KEY', 'value': render_api_key},
    ]
    resp = requests.put(url, json=payload, headers=headers)
    if resp.status_code in [200, 201]:
        return jsonify({'ok': True, 'message': '密碼已更新！'})
    else:
        return jsonify({'error': f'更新失敗：{resp.text}'}), 500

# ── 管理員：上傳新範本 ────────────────────────────────────────
@app.route('/admin/upload_template', methods=['POST'])
def admin_upload_template():
    password = request.form.get('password', '')
    admin_pw = os.environ.get('ADMIN_PASSWORD', '')
    if password != admin_pw:
        return jsonify({'error': '密碼錯誤'}), 401
    if 'template' not in request.files:
        return jsonify({'error': '請上傳範本檔案'}), 400
    f = request.files['template']
    if not f.filename.endswith('.xlsx'):
        return jsonify({'error': '請上傳 .xlsx 格式'}), 400
    f.save(TEMPLATE_PATH)
    return jsonify({'ok': True, 'message': f'範本已更新！({f.filename})'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
