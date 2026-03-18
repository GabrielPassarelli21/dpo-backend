from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os, json, openpyxl, re
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

    # Find the header row — it contains 'Fev' as a cell value
    header_row = -1
    fev_col = -1
    for i, row in enumerate(rows[:6]):
        for j, cell in enumerate(row):
            val = str(cell or '').strip().lower()
            if val.startswith('fev'):
                header_row = i
                fev_col = j
                break
        if header_row >= 0:
            break

    if header_row < 0 or fev_col < 0:
        return None

    # Month columns: 11 consecutive columns starting from fev_col
    month_cols = list(range(fev_col, min(fev_col + 11, len(rows[header_row]))))

    # Overall scores: always on the row immediately BEFORE the header row
    overall = [0.0] * 11
    if header_row >= 1:
        score_row = rows[header_row - 1]
        for mi, col in enumerate(month_cols):
            if col < len(score_row):
                try:
                    v = float(score_row[col] or 0)
                    if v > 0:
                        overall[mi] = round(v, 6)
                except:
                    pass

    # Parse itens
    itens = []
    current_bloco = ''
    for row in rows[header_row + 1:]:
        if not row:
            continue
        # Skip fully empty rows
        non_empty = [c for c in row if c is not None and str(c).strip() != '']
        if not non_empty:
            continue

        col0 = str(row[0] or '').strip()
        col1 = str(row[1] or '').strip() if len(row) > 1 else ''
        col2 = str(row[2] or '').strip() if len(row) > 2 else ''
        col3 = str(row[3] or '').strip() if len(row) > 3 else ''

        # Detect bloco header lines like "1.0 CONFORMIDADE"
        if col0 and re.match(r'^\d+[\.\d]*\s+\S', col0):
            current_bloco = col0

        # Skip rows without item or description
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
            'desc': col3 if col3 else col2,
            'mand': col1.lower() == 'sim',
            'scores': scores
        })

    return {'overallScores': overall, 'itens': itens}


# ── PARSE PLANO DE AÇÃO ──
def parse_plano(ws):
    rows = list(ws.iter_rows(values_only=True))
    acoes = []

    # Find header row — must contain BOTH a pilar/ação column AND responsável
    # This avoids matching the title "PLANO DE AÇÃO - DPO - ABS"
    header_row = -1
    for i, row in enumerate(rows[:8]):
        row_str = [str(c or '').lower().strip() for c in row]
        has_resp   = any('responsável' in c or 'responsavel' in c for c in row_str)
        has_pilar  = any(c == 'pilar' for c in row_str)
        has_status = any(c == 'status' for c in row_str)
        # Need at least responsável + (pilar or status) to confirm it's the real header
        if has_resp and (has_pilar or has_status):
            header_row = i
            break
    if header_row < 0:
        return []

    headers = [str(c or '').lower().strip() for c in rows[header_row]]

    def find_col(*keys):
        for i, h in enumerate(headers):
            if any(k in h for k in keys):
                return i
        return -1

    i_pilar = find_col('pilar')
    i_item  = find_col('item')
    i_acao  = find_col('ação', 'acao')
    i_resp  = find_col('responsável', 'responsavel')
    i_ini   = find_col('início', 'inicio')
    i_term  = find_col('término', 'termino', 'prazo')
    i_stat  = find_col('status')

    # Fallback column indices based on known structure
    if i_acao < 0:  i_acao  = 6
    if i_resp < 0:  i_resp  = 7
    if i_ini  < 0:  i_ini   = 9
    if i_term < 0:  i_term  = 10
    if i_stat < 0:  i_stat  = 13
    if i_pilar < 0: i_pilar = 4
    if i_item  < 0: i_item  = 5

    def fmt_date(v):
        if v is None or str(v).strip() in ('', 'None'):
            return '—'
        if isinstance(v, datetime):
            return v.strftime('%d/%m/%y')
        s = str(v).strip()
        # Excel serial date
        try:
            serial = float(s)
            if 40000 < serial < 60000:
                from datetime import date
                base = date(1899, 12, 30)
                from datetime import timedelta
                d = base + timedelta(days=int(serial))
                return d.strftime('%d/%m/%y')
        except:
            pass
        return s[:10] if len(s) > 10 else s

    last_pilar = ''
    for row in rows[header_row + 1:]:
        if not row:
            continue

        def safe(idx):
            if idx >= 0 and idx < len(row):
                return str(row[idx] or '').strip()
            return ''

        pilar = safe(i_pilar)
        if pilar:
            last_pilar = pilar
        else:
            pilar = last_pilar

        acao = safe(i_acao)
        if not acao or len(acao) < 4:
            continue

        acoes.append({
            'pilar':   pilar,
            'item':    safe(i_item),
            'acao':    acao,
            'resp':    safe(i_resp),
            'inicio':  fmt_date(row[i_ini] if i_ini < len(row) else None),
            'termino': fmt_date(row[i_term] if i_term < len(row) else None),
            'status':  safe(i_stat),
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
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Formato inválido. Use .xlsx ou .xls'}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)

        # Normalize sheet names (strip spaces, uppercase) but keep original index
        sheet_map = {s.strip().upper(): i for i, s in enumerate(wb.sheetnames)}

        # Load existing saved data
        existing_pilares = {}
        existing_acoes = []
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                existing_pilares = saved.get('pilares', {})
                existing_acoes   = saved.get('acoes', [])

        updated = []
        pilares = dict(existing_pilares)

        for ps in PILAR_SHEETS:
            # Try exact match first, then prefix match
            idx = sheet_map.get(ps)
            if idx is None:
                for key, i in sheet_map.items():
                    if key.startswith(ps[:4]):
                        idx = i
                        break
            if idx is not None:
                parsed = parse_pilar(wb.worksheets[idx])
                if parsed and parsed['itens']:
                    pilares[ps] = {**parsed, 'nome': ps}
                    updated.append(ps)

        # Parse plano de ação
        acoes = existing_acoes
        plano_idx = None
        for key, i in sheet_map.items():
            if 'PLANO' in key or 'AÇÃO' in key or 'ACAO' in key:
                plano_idx = i
                break
        if plano_idx is not None:
            parsed_acoes = parse_plano(wb.worksheets[plano_idx])
            if parsed_acoes:
                acoes = parsed_acoes

        if not updated and not acoes:
            return jsonify({'error': 'Nenhuma aba reconhecida. Verifique se os nomes das abas estão corretos (FROTA, GENTE, etc).'}), 400

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
        return jsonify({'error': f'Erro ao processar: {str(e)}'}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
