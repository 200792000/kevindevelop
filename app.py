from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
import shutil, re, os, io, tempfile

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.xlsx')

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

# ── 比對差異 ─────────────────────────────────────────────────
def compare(m1, m2):
    diffs = []
    for k in sorted(set(list(m1.keys()) + list(m2.keys()))):
        a, b = m1.get(k), m2.get(k)
        if not a or not b: continue
        s1 = '|'.join(d + '_' + a['noon'].get(d,'') for d in sorted(a['dates']))
        s2 = '|'.join(d + '_' + b['noon'].get(d,'') for d in sorted(b['dates']))
        if s1 != s2:
            orig = sorted(a['dates'])[0]
            chg  = sorted(b['dates'])[0]
            diffs.append({
                '店號':   k,
                '店名':   a['店名'],
                '原訂日': orig,
                '原訂午': a['noon'].get(orig, ''),
                '異動日': chg,
                '異動午': b['noon'].get(chg, ''),
            })
    return diffs

# ── 產出 Excel ────────────────────────────────────────────────
def generate_excel(diffs, version_full):
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
        c.value = parse_date(d['原訂日'])
        c.number_format = DATE_FMT

        ws.cell(row=r_d1, column=5).value = d['原訂午']
        ws.cell(row=r_d1, column=6).value = REASON
        ws.cell(row=r_d1, column=8).value = None
        ws.cell(row=r_d1, column=9).value = None

        c2 = ws.cell(row=r_d2, column=4)
        c2.value = parse_date(d['異動日'])
        c2.number_format = DATE_FMT

        ws.cell(row=r_d2, column=5).value = d['異動午']
        ws.cell(row=r_d2, column=8).value = version_full

    for i in range(len(diffs), 10):
        r_d1 = 6 + i * 5
        r_d2 = 8 + i * 5
        for col in [1, 4, 5, 6, 8, 9]:
            ws.cell(row=r_d1, column=col).value = None
            ws.cell(row=r_d2, column=col).value = None

    for coord in merges:
        ws.merge_cells(coord)

    wb.save(tmp.name)
    return tmp.name

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
        diffs  = compare(m1, m2)
        return jsonify({
            'version':   version_full,
            'month':     month,
            'v1_count':  len(m1),
            'v2_count':  len(m2),
            'diff_count': len(diffs),
            'diffs': diffs
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generate', methods=['POST'])
def api_generate():
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
        diffs  = compare(m1, m2)
        if not diffs:
            return jsonify({'error': '兩版本班表無差異'}), 400
        out_path = generate_excel(diffs, version_full)
        today = datetime.today().strftime('%Y%m%d')
        filename = f'盤點行程異動聯絡單_{version_full}_{today}.xlsx'
        return send_file(
            out_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
