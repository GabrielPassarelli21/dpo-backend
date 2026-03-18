from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os, json, openpyxl
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__, static_folder='static')
CORS(app)

UPLOAD_FOLDER = 'uploads'
DATA_FILE = 'data.json'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

PILAR_SHEETS = ['SEGURANÇA','GENTE','GESTÃO','FROTA','ENTREGA','ARMAZÉM','PLANEJAMENTO']
MESES = ['Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']

# ── PARSE PILAR SHEET ──
def parse_pilar(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows or len(rows) < 3:
        return None

    # Structure is always: row0=empty, row1=scores, row2=headers(Fev/Mar...)
    # Find header row (contains 'Fev')
    header_row = -1
    for i, row in enumerate(rows[:6]):
        row_str = [str(c or '').strip().lower() for c in row]
        if any(c.startswith('fev') or c.startswith('feb') for c in row_str):
            header_row = i
            break
    if header_row < 0:
        return None

    headers = [str(c or '').strip() for c in rows[header_row]]
    fev_idx = next((i for i, h in enumerate(headers) if h.lower().startswith('fev') or h.lower().startswith('feb')), -1)
    if fev_idx < 0:
        return None

    month_cols = list(range(fev_idx, min(fev_idx + 11, len(headers))))

    # Overall scores are ALWAYS on the row just before header (header_row - 1)
    overall = [0.0] * 11
    score_row_idx = header_row - 1
    if score_row_idx >= 0:
        score_row = rows[score_row_idx]
        for mi, col in enumerate(month_cols):
            if col < len(score_row):
                try:
                    v = float(score_row[col] or 0)
                    overall[mi] = round(v, 6)
                except:
                    pass

    itens = []
    current_bloco = ''
    for row in rows[header_row + 1:]:
        if not row or all(c is None or str(c).strip() == '' for c in row):
            continue
        col0 = str(row[0] or '').strip()
        col1 = str(row[1] or '').strip() if len(row) > 1 else ''
        col2 = str(row[2] or '').strip() if len(row) > 2 else ''
        col3 = str(row[3] or '').strip() if len(row) > 3 else ''

        import re
        if col0 and re.match(r'^\d+[\.\d]*\s+', col0):
            current_bloco = col0

        if not col2 and not col3:
            continue

        scores = []
        for col in month_cols:
            if col < len(row):
                try:
                    scores.append(int(float(row[col] or 0)))
                except:
                    scores.append(0)
            else:
                scores.append(0)

        itens.append({
            'bloco': current_bloco,
            'item': col2,
            'desc': col3 or col2,
            'mand': col1.lower() == 'sim',
            'scores': scores
        })

    return {'overallScores': overall, 'itens': itens}

# ── PARSE PLANO SHEET ──
def parse_plano(ws):
    rows = list(ws.iter_rows(values_only=True))
    acoes = []
    header_row = -1

    for i, row in enumerate(rows[:6]):
        row_str = [str(c or '').lower() for c in row]
        if any('ação' in c or 'acao' in c or 'responsável' in c for c in row_str):
            header_row = i
            break
    if header_row < 0:
        return []

    headers = [str(c or '').lower().strip() for c in rows[header_row]]

    def find_col(*keys):
        return next((i for i, h in enumerate(headers) if any(k in h for k in keys)), -1)

    i_acao  = find_col('ação','acao')
    i_resp  = find_col('responsável','responsavel')
    i_ini   = find_col('início','inicio')
    i_term  = find_col('término','termino')
    i_stat  = find_col('status')
    i_pilar = find_col('pilar')
    i_item  = find_col('item')

    last_pilar = ''
    for row in rows[header_row + 1:]:
        if not row:
            continue
        pilar = str(row[i_pilar] if i_pilar >= 0 and i_pilar < len(row) else row[4] if len(row) > 4 else '').strip()
        if pilar:
            last_pilar = pilar
        else:
            pilar = last_pilar

        acao = str(row[i_acao] if i_acao >= 0 and i_acao < len(row) else row[6] if len(row) > 6 else '').strip()
        if not acao or len(acao) < 5:
            continue

        def fmt_date(v):
            if not v:
                return '—'
            if isinstance(v, datetime):
                return v.strftime('%d/%m/%y')
            try:
                return str(v)[:10]
            except:
                return '—'

        acoes.append({
            'pilar':   pilar,
            'item':    str(row[i_item] if i_item >= 0 and i_item < len(row) else '').strip(),
            'acao':    acao,
            'resp':    str(row[i_resp] if i_resp >= 0 and i_resp < len(row) else '').strip(),
            'inicio':  fmt_date(row[i_ini] if i_ini >= 0 and i_ini < len(row) else None),
            'termino': fmt_date(row[i_term] if i_term >= 0 and i_term < len(row) else None),
            'status':  str(row[i_stat] if i_stat >= 0 and i_stat < len(row) else '').strip(),
        })

    return acoes

# ── ROUTES ──

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

@app.route('/api/data', methods=['GET'])
def get_data():
    if not os.path.exists(DATA_FILE):
        return jsonify({'status': 'empty', 'pilares': {}, 'acoes': [], 'meta': {}})
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        return jsonify(json.load(f))

@app.route('/api/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400

    file = request.files['file']
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Formato inválido. Use .xlsx ou .xls'}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet_names = [s.strip().upper() for s in wb.sheetnames]

        # Load existing data
        existing = {}
        existing_acoes = []
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                existing = saved.get('pilares', {})
                existing_acoes = saved.get('acoes', [])

        updated = []
        pilares = dict(existing)

        for ps in PILAR_SHEETS:
            idx = next((i for i, s in enumerate(sheet_names)
                        if s.strip() == ps or s.strip().startswith(ps[:4])), -1)
            if idx >= 0:
                parsed = parse_pilar(wb.worksheets[idx])
                if parsed:
                    pilares[ps] = {**parsed, 'nome': ps}
                    updated.append(ps)

        # Plano de ação
        plano_idx = next((i for i, s in enumerate(sheet_names)
                          if 'PLANO' in s or 'AÇÃO' in s or 'ACAO' in s), -1)
        acoes = existing_acoes
        if plano_idx >= 0:
            parsed_acoes = parse_plano(wb.worksheets[plano_idx])
            if parsed_acoes:
                acoes = parsed_acoes

        result = {
            'status': 'ok',
            'pilares': pilares,
            'acoes': acoes,
            'meta': {
                'fileName': filename,
                'updatedAt': datetime.now().strftime('%d/%m/%Y %H:%M'),
                'updatedPilares': updated,
            }
        }

        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, default=str)

        return jsonify({
            'status': 'ok',
            'updatedPilares': updated,
            'totalAcoes': len(acoes),
            'updatedAt': result['meta']['updatedAt'],
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
