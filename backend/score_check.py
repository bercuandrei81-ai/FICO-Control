import io
import os
import re
import html
import uuid
from datetime import datetime, timedelta
from typing import Any

import openpyxl
from fastapi import APIRouter, UploadFile, File, Request
from fastapi.responses import HTMLResponse, FileResponse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

router = APIRouter()
_SCORE_EXPORT_DIR = os.path.join('uploads', 'score_check')
os.makedirs(_SCORE_EXPORT_DIR, exist_ok=True)


def _clean_header(value: Any) -> str:
    return re.sub(r'\s+', ' ', str(value or '')).strip()


def _num(value):
    if value is None or value == '':
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(',', '.')
    if not text or text.casefold() in {'n/a', 'na', '-', 'none'}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _read_xlsx(raw: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        rows = ws.iter_rows(values_only=True)
        headers = next(rows, None)
        if not headers:
            raise ValueError('Fișierul nu conține antet.')
        headers = [_clean_header(v) or f'Column {i+1}' for i, v in enumerate(headers)]
        records = []
        for row in rows:
            if not any(v is not None and str(v).strip() != '' for v in row):
                continue
            records.append({h: (row[i] if i < len(row) else None) for i, h in enumerate(headers)})
        return headers, records
    finally:
        wb.close()


def _encrypted_key(row):
    return f"{str(row.get('First Name') or '').strip()}\u241f{str(row.get('Last Name') or '').strip()}"


def _real_name(row):
    first = str(row.get('First Name') or '').strip()
    last = str(row.get('Last Name') or '').strip()
    return ' '.join(part for part in (first, last) if part).strip()


def _display_encrypted(row):
    first = str(row.get('First Name') or '').strip()
    last = str(row.get('Last Name') or '').strip()
    return ' | '.join(part for part in (first, last) if part)


def _fico_recent(row):
    last = _num(row.get('FICO® Safe Driving Score Last Period'))
    delta = _num(row.get('FICO® Safe Driving Score Δ'))
    if last is None:
        return None
    if delta is None:
        return int(round(last))
    recent = last * (1.0 + delta / 100.0)
    return int(round(max(300.0, min(850.0, recent))))


def _candidate_score(comp_row, shift_row):
    ckm = _num(comp_row.get('Total Driver km This Period'))
    skm = _num(shift_row.get('Total Driver km'))
    if ckm is None or skm is None:
        return 9999.0
    return abs(ckm - skm)


def _assign_real_names(comp_rows, shift_rows):
    all_pairs = []
    candidate_lists = {}
    for ci, comp in enumerate(comp_rows):
        candidates = []
        for si, shift in enumerate(shift_rows):
            diff = _candidate_score(comp, shift)
            candidates.append((diff, si))
            all_pairs.append((diff, ci, si))
        candidates.sort(key=lambda x: (x[0], _real_name(shift_rows[x[1]]).casefold()))
        candidate_lists[ci] = candidates

    used_c, used_s, chosen = set(), set(), {}
    for diff, ci, si in sorted(all_pairs, key=lambda x: x[0]):
        if ci in used_c or si in used_s:
            continue
        chosen[ci] = (diff, si)
        used_c.add(ci)
        used_s.add(si)

    result = []
    for ci, comp in enumerate(comp_rows):
        candidates = candidate_lists.get(ci, [])
        selected = chosen.get(ci) or (candidates[0] if candidates else None)
        if not selected:
            result.append({'name': '', 'confidence': 'Low', 'km_diff': None, 'candidate_2': '', 'candidate_3': '', 'shift_row': None})
            continue
        diff, si = selected
        alt = [(d, idx) for d, idx in candidates if idx != si]
        next_diff = alt[0][0] if alt else None
        if diff <= 0.05 and (next_diff is None or next_diff - diff >= 0.10):
            confidence = 'High'
        elif diff <= 0.50 and (next_diff is None or next_diff - diff >= 0.03):
            confidence = 'Medium'
        else:
            confidence = 'Low'

        def alt_name(n):
            if len(alt) <= n:
                return ''
            d, idx = alt[n]
            return f"{_real_name(shift_rows[idx])} (Δkm {d:.2f})"

        result.append({'name': _real_name(shift_rows[si]), 'confidence': confidence, 'km_diff': round(diff, 3), 'candidate_2': alt_name(0), 'candidate_3': alt_name(1), 'shift_row': shift_rows[si]})
    return result


def _build_rows(comp_headers, comp_rows, driver_headers, driver_rows, shift_headers, shift_rows):
    driver_by_key = {_encrypted_key(r): r for r in driver_rows}
    matches = _assign_real_names(comp_rows, shift_rows)
    front_headers = ['Nume real estimat', 'Confidence', 'Diferență km', 'Candidat alternativ 2', 'Candidat alternativ 3', 'Encrypted ID', 'Cod șofer / Vehicle Identifier', 'FICO Recent', 'FICO Last Period', 'FICO Δ', 'Ore lucrate', 'Km Shift Report']
    all_headers = front_headers + [f'Comparison · {h}' for h in comp_headers] + [f'Driver Report · {h}' for h in driver_headers] + [f'Shift Report · {h}' for h in shift_headers]
    output_rows = []
    for i, comp in enumerate(comp_rows):
        match = matches[i]
        shift = match['shift_row'] or {}
        driver = driver_by_key.get(_encrypted_key(comp), {})
        row = {
            'Nume real estimat': match['name'], 'Confidence': match['confidence'], 'Diferență km': match['km_diff'],
            'Candidat alternativ 2': match['candidate_2'], 'Candidat alternativ 3': match['candidate_3'],
            'Encrypted ID': _display_encrypted(comp), 'Cod șofer / Vehicle Identifier': shift.get('Vehicle Identifier'),
            'FICO Recent': _fico_recent(comp), 'FICO Last Period': comp.get('FICO® Safe Driving Score Last Period'),
            'FICO Δ': comp.get('FICO® Safe Driving Score Δ'), 'Ore lucrate': shift.get('Total Driver Hours'), 'Km Shift Report': shift.get('Total Driver km'),
        }
        for h in comp_headers:
            row[f'Comparison · {h}'] = comp.get(h)
        for h in driver_headers:
            row[f'Driver Report · {h}'] = driver.get(h)
        for h in shift_headers:
            row[f'Shift Report · {h}'] = shift.get(h)
        output_rows.append(row)
    return all_headers, output_rows


def _write_excel(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Verificare Scor'
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h) for h in headers])
    dark = PatternFill('solid', fgColor='17212B')
    white_bold = Font(color='FFFFFF', bold=True)
    for cell in ws[1]:
        cell.fill = dark
        cell.font = white_bold
        cell.alignment = Alignment(vertical='center', wrap_text=True)
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    high_fill = PatternFill('solid', fgColor='E9F8EF')
    med_fill = PatternFill('solid', fgColor='FFF4D6')
    low_fill = PatternFill('solid', fgColor='FDEEEE')
    incident_fill = PatternFill('solid', fgColor='FFE0E0')
    header_index = {h: idx + 1 for idx, h in enumerate(headers)}
    conf_col = header_index.get('Confidence')
    if conf_col:
        for r in range(2, ws.max_row + 1):
            value = ws.cell(r, conf_col).value
            ws.cell(r, conf_col).fill = high_fill if value == 'High' else med_fill if value == 'Medium' else low_fill
            ws.cell(r, conf_col).font = Font(bold=True)
    incident_terms = ('speeding', 'distraction', 'hard braking', 'hard cornering', 'acceleration')
    for h, col_idx in header_index.items():
        hl = h.casefold()
        if any(term in hl for term in incident_terms) and 'rating' not in hl:
            for r in range(2, ws.max_row + 1):
                val = _num(ws.cell(r, col_idx).value)
                if val is not None and val > 0:
                    ws.cell(r, col_idx).fill = incident_fill
                    ws.cell(r, col_idx).font = Font(bold=True)
    for idx, h in enumerate(headers, start=1):
        width = min(max(14, len(str(h)) + 2), 38)
        if h in {'Nume real estimat', 'Encrypted ID', 'Candidat alternativ 2', 'Candidat alternativ 3'}:
            width = 34
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = width
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def _cleanup_exports():
    now = datetime.now()
    try:
        for name in os.listdir(_SCORE_EXPORT_DIR):
            path = os.path.join(_SCORE_EXPORT_DIR, name)
            try:
                if now - datetime.fromtimestamp(os.path.getmtime(path)) > timedelta(hours=24):
                    os.remove(path)
            except Exception:
                pass
    except Exception:
        pass


def _page(body: str, title='Verificare Scor'):
    return HTMLResponse(f'''<!doctype html><html lang="ro"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(title)}</title><style>
*{{box-sizing:border-box}}body{{margin:0;font-family:Arial,Helvetica,sans-serif;background:#f4f6f8;color:#17212b}}.wrap{{width:min(96%,1380px);margin:28px auto 60px}}.top{{display:flex;justify-content:space-between;align-items:center;gap:15px;margin-bottom:20px}}h1{{margin:0;font-size:34px}}.back{{text-decoration:none;color:#17212b;border:1px solid #d8dde3;background:#fff;padding:11px 14px;border-radius:10px;font-weight:800}}.panel{{background:#fff;border-radius:16px;padding:22px;box-shadow:0 5px 18px rgba(0,0,0,.05);margin-bottom:16px}}.grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}}label{{display:block;font-weight:800;margin-bottom:8px}}input[type=file]{{width:100%;padding:13px;border:1px solid #d8dde3;border-radius:10px;background:#fff}}button,.download{{display:inline-flex;align-items:center;justify-content:center;text-decoration:none;padding:13px 16px;border:0;border-radius:10px;background:#17212b;color:#fff;font-weight:900;cursor:pointer;margin-top:16px}}.help{{color:#667085;line-height:1.45}}.notice{{padding:12px 14px;border-radius:10px;background:#fff4d6;color:#8a5a00;margin:12px 0}}.table-wrap{{overflow:auto;max-height:68vh;border:1px solid #eceff2;border-radius:12px}}table{{border-collapse:collapse;width:max-content;min-width:100%;font-size:13px}}th,td{{padding:10px 11px;border-bottom:1px solid #eceff2;white-space:nowrap;text-align:left}}th{{position:sticky;top:0;background:#17212b;color:#fff;z-index:1}}.high{{background:#e9f8ef}}.medium{{background:#fff4d6}}.low{{background:#fdeeee}}@media(max-width:850px){{.grid{{grid-template-columns:1fr}}h1{{font-size:28px}}}}</style></head><body><main class="wrap">{body}</main></body></html>''')


@router.get('/admin/score-check', response_class=HTMLResponse)
def score_check_page(request: Request):
    body = '''<div class="top"><div><div style="font-size:12px;color:#667085;font-weight:800;letter-spacing:1px">INSTRUMENTE FICO</div><h1>Verificare Scor</h1></div><a class="back" href="/admin">← Admin Dashboard</a></div><section class="panel"><p class="help">Încarcă cele trei rapoarte Amazon. Sistemul păstrează toate datele, corelează ID-ul criptat cu Driver Report și estimează numele real folosind kilometrul din Comparison și Shift Report. Vei vedea FICO Recent, FICO Last Period, codul vehiculului/șoferului și toate incidentele.</p><form action="/admin/score-check/generate" method="post" enctype="multipart/form-data"><div class="grid"><div><label>1. Comparison</label><input type="file" name="comparison_file" accept=".xlsx" required></div><div><label>2. Driver Report</label><input type="file" name="driver_file" accept=".xlsx" required></div><div><label>3. Shift Report</label><input type="file" name="shift_file" accept=".xlsx" required></div></div><button type="submit">Generează raportul complet</button></form></section><section class="panel"><strong>Ce păstrăm:</strong><p class="help">100% din coloanele originale ale celor 3 fișiere. Speeding, Distraction, Hard Braking, Hard Cornering, Acceleration și orice alt indicator rămân în raport.</p></section>'''
    return _page(body)


@router.post('/admin/score-check/generate', response_class=HTMLResponse)
async def score_check_generate(comparison_file: UploadFile = File(...), driver_file: UploadFile = File(...), shift_file: UploadFile = File(...)):
    for f in (comparison_file, driver_file, shift_file):
        if not (f.filename or '').lower().endswith('.xlsx'):
            return _page('<div class="notice">Toate cele trei fișiere trebuie să fie XLSX.</div><a class="back" href="/admin/score-check">Înapoi</a>')
    try:
        comp_h, comp_r = _read_xlsx(await comparison_file.read())
        drv_h, drv_r = _read_xlsx(await driver_file.read())
        shift_h, shift_r = _read_xlsx(await shift_file.read())
        required_comp = {'First Name','Last Name','Total Driver km This Period','FICO® Safe Driving Score Last Period'}
        required_shift = {'First Name','Last Name','Total Driver Hours','Total Driver km'}
        if not required_comp.issubset(set(comp_h)):
            raise ValueError('Comparison nu are coloanele Amazon așteptate.')
        if not {'First Name','Last Name'}.issubset(set(drv_h)):
            raise ValueError('Driver Report nu are First Name / Last Name.')
        if not required_shift.issubset(set(shift_h)):
            raise ValueError('Shift Report nu are coloanele necesare pentru nume, ore și km.')
        headers, rows = _build_rows(comp_h, comp_r, drv_h, drv_r, shift_h, shift_r)
        data = _write_excel(headers, rows)
    except Exception as exc:
        return _page(f'<div class="notice">Nu am putut procesa fișierele: {html.escape(str(exc))}</div><a class="back" href="/admin/score-check">Înapoi</a>')
    _cleanup_exports()
    token = uuid.uuid4().hex
    filename = f'Verificare_Scor_{datetime.now().strftime("%Y-%m-%d_%H%M%S")}_{token[:8]}.xlsx'
    path = os.path.join(_SCORE_EXPORT_DIR, filename)
    with open(path, 'wb') as handle:
        handle.write(data)
    preview_headers = ['Nume real estimat','Confidence','Encrypted ID','Cod șofer / Vehicle Identifier','FICO Recent','FICO Last Period','Ore lucrate','Km Shift Report']
    preview = ''
    for row in rows[:250]:
        cls = str(row.get('Confidence') or '').lower()
        preview += '<tr class="%s">%s</tr>' % (cls, ''.join(f'<td>{html.escape(str(row.get(h) if row.get(h) is not None else "—"))}</td>' for h in preview_headers))
    table = '<div class="table-wrap"><table><thead><tr>' + ''.join(f'<th>{html.escape(h)}</th>' for h in preview_headers) + '</tr></thead><tbody>' + preview + '</tbody></table></div>'
    body = f'''<div class="top"><div><div style="font-size:12px;color:#667085;font-weight:800;letter-spacing:1px">INSTRUMENTE FICO</div><h1>Verificare Scor</h1></div><a class="back" href="/admin">← Admin Dashboard</a></div><section class="panel"><strong>Raport generat: {len(rows)} șoferi</strong><p class="help">Preview-ul arată coloanele principale. Excelul descărcat păstrează absolut toate coloanele din Comparison, Driver Report și Shift Report.</p><a class="download" href="/admin/score-check/download/{html.escape(filename)}">Descarcă Excel complet</a></section><section class="panel">{table}</section>'''
    return _page(body)


@router.get('/admin/score-check/download/{filename}')
def score_check_download(filename: str):
    safe = os.path.basename(filename)
    path = os.path.join(_SCORE_EXPORT_DIR, safe)
    if not os.path.isfile(path):
        return HTMLResponse('Fișierul nu mai este disponibil.', status_code=404)
    return FileResponse(path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=safe)


def register_score_check(app):
    if getattr(app.state, '_score_check_registered', False):
        return
    app.state._score_check_registered = True
    app.include_router(router)

    @app.middleware('http')
    async def _score_check_nav_middleware(request, call_next):
        response = await call_next(request)
        try:
            ctype = response.headers.get('content-type', '')
            if request.url.path.startswith('/admin') and 'text/html' in ctype:
                body = b''
                async for chunk in response.body_iterator:
                    body += chunk
                text = body.decode('utf-8', errors='replace')
                if 'href="/admin/score-check"' not in text and 'side-nav-links' in text:
                    marker = '<nav class="side-nav-links">'
                    link = '<a class="side-link" href="/admin/score-check"><i></i>Verificare Scor</a>'
                    text = text.replace(marker, marker + link, 1)
                headers = dict(response.headers)
                headers.pop('content-length', None)
                return HTMLResponse(content=text, status_code=response.status_code, headers=headers)
        except Exception:
            return response
        return response
