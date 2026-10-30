"""Arma Candidatos_Busqueda_Octubre_2026.xlsx a partir de registros/*.json"""
import json, glob, os, re, unicodedata, datetime as dt
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

BASE = os.environ.get('CVS_BASE', '/home/claude/cvs')
OUT = os.path.join(BASE, 'Candidatos_Busqueda_Octubre_2026.xlsx')
CORTE = os.environ.get('CVS_CORTE', '28/10/2026 01:10')  # fecha y hora del último mail leído
HOY = dt.date(2026, 10, 1)  # "actual" = octubre 2026
ZONAS_OK = {'vicente lopez': 'Vicente López', 'san isidro': 'San Isidro', 'san martin': 'San Martín',
            'general san martin': 'San Martín', 'gral. san martin': 'San Martín', 'gral san martin': 'San Martín'}
# Localidades -> partido (por si el partido vino vacío o mal)
LOC = {'florida': 'Vicente López', 'florida oeste': 'Vicente López', 'olivos': 'Vicente López', 'munro': 'Vicente López',
       'carapachay': 'Vicente López', 'la lucila': 'Vicente López', 'villa martelli': 'Vicente López',
       'villa adelina': 'San Isidro', 'boulogne': 'San Isidro', 'beccar': 'San Isidro', 'martinez': 'San Isidro',
       'acassuso': 'San Isidro', 'san isidro': 'San Isidro', 'vicente lopez': 'Vicente López',
       'villa ballester': 'San Martín', 'san andres': 'San Martín', 'jose leon suarez': 'San Martín',
       'villa maipu': 'San Martín', 'billinghurst': 'San Martín', 'chilavert': 'San Martín', 'san martin': 'San Martín',
       'villa lynch': 'San Martín', 'loma hermosa': 'Tres de Febrero', 'caseros': 'Tres de Febrero',
       'tigre': 'Tigre', 'don torcuato': 'Tigre', 'pacheco': 'Tigre', 'general pacheco': 'Tigre',
       'el talar': 'Tigre', 'benavidez': 'Tigre', 'san fernando': 'San Fernando', 'victoria': 'San Fernando',
       'virreyes': 'San Fernando', 'jose c. paz': 'José C. Paz', 'jose c paz': 'José C. Paz',
       'grand bourg': 'Malvinas Argentinas', 'los polvorines': 'Malvinas Argentinas', 'pilar': 'Pilar',
       'del viso': 'Pilar', 'escobar': 'Escobar', 'garin': 'Escobar', 'san miguel': 'San Miguel',
       'bella vista': 'San Miguel', 'munoz': 'San Miguel'}


def norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode().lower().strip()
    return re.sub(r'\s+', ' ', s)


def partido_de(r):
    p = norm(r.get('partido'))
    if p in ZONAS_OK: return ZONAS_OK[p]
    if p.startswith('caba') or 'capital' in p or 'ciudad autonoma' in p: return 'CABA'
    loc = norm(r.get('localidad'))
    for k, v in sorted(LOC.items(), key=lambda x: -len(x[0])):
        if loc == k or loc.startswith(k + ' ') or loc.endswith(' ' + k):
            return v
    if loc.startswith('caba') or 'capital federal' in loc: return 'CABA'
    return (r.get('partido') or '').strip()


def a_fecha(v, es_hasta):
    """Devuelve (date, aproximada:bool). 'YYYY-MM' exacto; 'YYYY' -> 1-jul (aprox)."""
    if v is None: return None, False
    v = str(v).strip().lower()
    if v in ('actual', 'actualidad', 'presente', 'hoy'): return HOY, False
    m = re.match(r'^(\d{4})-(\d{1,2})', v)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12: return dt.date(y, mo, 1), False
    m = re.match(r'^(\d{4})$', v)
    if m: return dt.date(int(m.group(1)), 7, 1), True
    return None, False


def trabajo(t):
    if isinstance(t, list):
        keys = ['empresa', 'cargo', 'rubro', 'gastronomia', 'atencion_publico', 'cocina', 'desde', 'hasta', 'fechas_texto']
        t = dict(zip(keys, t + [None] * (9 - len(t))))
    d, ad = a_fecha(t.get('desde'), False)
    h, ah = a_fecha(t.get('hasta'), True)
    if d and h and ad and ah and d == h:  # mismo año sin meses -> 6 meses
        d, h = dt.date(d.year, 1, 1), dt.date(d.year, 6, 1)
    if d and h and h < d: d, h = None, None
    if d and not h and str(t.get('hasta')) not in ('None', ''):
        pass
    if (d and not h) or (h and not d): d, h = None, None  # una sola punta: no se puede medir
    return {'empresa': t.get('empresa') or '', 'cargo': t.get('cargo') or '', 'rubro': t.get('rubro') or '',
            'gastro': bool(t.get('gastronomia')), 'atencion': bool(t.get('atencion_publico')),
            'cocina': bool(t.get('cocina')), 'desde': d, 'hasta': h, 'aprox': bool(ad or ah),
            'fechas_texto': t.get('fechas_texto') or ''}


def meses(t):
    if not (t['desde'] and t['hasta']): return None
    return (t['hasta'].year - t['desde'].year) * 12 + t['hasta'].month - t['desde'].month + 1


# ---------- cargar ----------
regs = []
for f in sorted(glob.glob(os.path.join(BASE, 'registros', '*.json'))):
    try:
        r = json.load(open(f))
    except Exception as e:
        print('JSON roto', f, e); continue
    mid = r.get('message_id') or os.path.basename(f)[:-5]
    mid_base = mid.split('_')[0]
    meta_p = os.path.join(BASE, 'textos', mid_base + '.meta.json')
    meta = json.load(open(meta_p)) if os.path.exists(meta_p) else {}
    r['link'] = r.get('link') or meta.get('viewUrl') or f'https://mail.google.com/mail/u/lapresumidapastas@gmail.com/#all/{mid_base}'
    r['trabajos_n'] = [trabajo(t) for t in (r.get('trabajos') or [])]
    r['partido_n'] = partido_de(r)
    regs.append(r)

# ---------- duplicados ----------
def clave(r):
    m = norm(r.get('mail'))
    if '@' in m and 'lapresumida' not in m: return 'm:' + m
    tel = re.sub(r'\D', '', str(r.get('telefono') or ''))
    if len(tel) >= 8: return 't:' + tel[-8:]
    return 'n:' + norm(r.get('nombre')) if r.get('nombre') else 'id:' + r['message_id']

grupos = defaultdict(list)
for r in regs: grupos[clave(r)].append(r)
cands, dups = [], []
for k, g in grupos.items():
    g.sort(key=lambda r: (len(r['trabajos_n']), r.get('estado_lectura') == 'ok', r.get('fecha_mail') or ''), reverse=True)
    best = g[0]
    best['veces'] = len(g)
    best['otros_mails'] = [x['message_id'] for x in g[1:]]
    cands.append(best)
    for x in g[1:]: dups.append((x, best))

# ---------- evaluación ----------
ZONA_PTS = {'Vicente López': 20, 'San Isidro': 20, 'San Martín': 12}


def evaluar(r):
    ts = r['trabajos_n']
    puesto = r.get('puesto') or 'No dice'
    deducido = False
    if puesto not in ('Cocina', 'Atención/Venta'):
        puesto = 'Cocina' if any(t['cocina'] for t in ts) else 'Atención/Venta'
        deducido = True
    r['puesto_eval'] = puesto; r['puesto_deducido'] = deducido
    con_fecha = [meses(t) for t in ts if meses(t) is not None]
    r['n_trabajos'] = len(ts); r['n_con_fecha'] = len(con_fecha)
    r['prom'] = sum(con_fecha) / len(con_fecha) if con_fecha else None
    r['cortos'] = sum(1 for m in con_fecha if m < 12)
    r['largo'] = max(con_fecha) if con_fecha else None

    motivos, alertas = [], []
    part = r['partido_n']
    if part in ZONA_PTS: pass
    elif part: motivos.append(f'Vive fuera de zona ({r.get("localidad") or part}, partido {part})')
    else: alertas.append('Zona sin dato')
    ft = norm(r.get('full_time'))
    if ft.startswith('no dice') or not ft: alertas.append('Full time: no lo dice')
    elif ft.startswith('no'): motivos.append(f'No tiene disponibilidad full time ({r.get("disponibilidad")})')
    sin_leer = (not ts) and (str(r.get('estado_lectura','')).startswith(('sin adjunto','ilegible')))
    r['sin_leer'] = sin_leer
    if sin_leer:
        alertas.append('No se pudo leer el CV (' + (r.get('estado_lectura') or '') + '): abrirlo a mano')
    elif not ts:
        motivos.append('Sin experiencia laboral en el CV' + (' (el aviso pide experiencia comprobable)' if puesto == 'Cocina' else ' (el aviso pide referencias comprobables)'))
    elif puesto == 'Cocina' and not any(t['cocina'] or t['gastro'] for t in ts):
        motivos.append('Se postula a cocina pero no tiene experiencia en cocina ni gastronomía')
    if r.get('estado_lectura', '').startswith('ilegible'): alertas.append('CV ilegible: abrirlo a mano')

    if sin_leer: grupo = 'CV sin leer'
    elif not ts: grupo = 'Sin experiencia'
    elif not con_fecha: grupo = 'Sin fechas, revisar'
    elif r['prom'] >= 24: grupo = '1. Estable (2 años o más)'
    elif r['prom'] < 12 and r['cortos'] >= 2: grupo = 'Cambia seguido'
    elif con_fecha and len(con_fecha) < len(ts): grupo = '2. Intermedio'
    else: grupo = '2. Intermedio'
    if con_fecha and len(con_fecha) < len(ts): alertas.append(f'{len(ts)-len(con_fecha)} trabajo(s) sin fechas')
    r['grupo'] = grupo; r['motivos'] = motivos; r['alertas'] = alertas

    # experiencia relevante (meses)
    def m_de(cond): return sum(meses(t) or 0 for t in ts if cond(t))
    def hay(cond): return any(cond(t) for t in ts)
    if puesto == 'Cocina':
        mc = m_de(lambda t: t['cocina'])
        if mc >= 24: pe, txt = 30, f'{mc} meses en cocina'
        elif mc >= 12: pe, txt = 22, f'{mc} meses en cocina'
        elif hay(lambda t: t['cocina']): pe, txt = 15, 'algo de cocina (menos de 1 año o sin fechas)'
        elif hay(lambda t: t['gastro']): pe, txt = 8, 'gastronomía pero no cocinando'
        else: pe, txt = 0, 'sin experiencia en cocina'
    else:
        mg = m_de(lambda t: t['gastro'] and t['atencion'])
        ma = m_de(lambda t: t['atencion'])
        if mg >= 24: pe, txt = 30, f'{mg} meses atendiendo en gastronomía'
        elif hay(lambda t: t['gastro']): pe, txt = 22, 'atención/trabajo en gastronomía (menos de 2 años o sin fechas)'
        elif ma >= 24: pe, txt = 20, f'{ma} meses de atención al público (no gastronomía)'
        elif hay(lambda t: t['atencion']): pe, txt = 14, 'algo de atención al público'
        else: pe, txt = 0, 'sin atención al público'
    r['pts_exp'] = pe; r['exp_txt'] = txt
    r['pts_zona'] = ZONA_PTS.get(part, 5 if not part else 0)
    pa = 10 if ft.startswith('si') else (5 if (ft.startswith('no dice') or not ft) else 0)
    ref = norm(r.get('referencias'))
    if ref.startswith('si') or r.get('estudios_relevantes'): pa += 5
    elif ref.startswith('a pedido'): pa += 3
    r['pts_aviso'] = pa


for r in cands: evaluar(r)

# ---------- Excel ----------
wb = Workbook()
F = 'Arial'
H_FILL = PatternFill('solid', start_color='1F1F1F'); H_FONT = Font(name=F, bold=True, color='FFFFFF', size=10)
AMARILLO = PatternFill('solid', start_color='FFE066'); VERDE = PatternFill('solid', start_color='E2F0D9')
GRIS = PatternFill('solid', start_color='F2F2F2'); ROJO = PatternFill('solid', start_color='FCE4E4')
BODY = Font(name=F, size=10); LINK = Font(name=F, size=10, color='0563C1', underline='single')
thin = Side(style='thin', color='D0D0D0'); BORDE = Border(bottom=thin)
WRAP = Alignment(wrap_text=True, vertical='top')


def encabezado(ws, cols, widths, fila=1):
    for i, (c, w) in enumerate(zip(cols, widths), 1):
        cell = ws.cell(row=fila, column=i, value=c)
        cell.font = H_FONT; cell.fill = H_FILL; cell.alignment = Alignment(wrap_text=True, vertical='center')
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[fila].height = 32
    ws.freeze_panes = ws.cell(row=fila + 1, column=5)
    ws.auto_filter.ref = f'A{fila}:{get_column_letter(len(cols))}{fila}'


def fmt_fecha(d, aprox):
    if not d: return ''
    return d.strftime('%m/%Y') + (' (aprox)' if aprox else '')


def exp_texto(r):
    partes = []
    for t in r['trabajos_n']:
        m = meses(t)
        dur = (f'{m // 12} a {m % 12} m' if m >= 12 else f'{m} m') if m is not None else 'sin fechas'
        partes.append(f"{t['empresa'] or '?'} ({t['cargo'] or t['rubro']}) – {dur}")
    return '; '.join(partes) if partes else 'Sin experiencia en el CV'


# Hoja Detalle de trabajos (fuente de los cálculos)
wsd = wb.active; wsd.title = 'Detalle trabajos'
colsd = ['ID candidato', 'Nombre', 'Empresa', 'Puesto/cargo', 'Rubro', 'Gastronomía', 'Atención al público', 'Cocina',
         'Fechas como dice el CV', 'Desde', 'Hasta', 'Meses que duró']
encabezado(wsd, colsd, [19, 24, 26, 26, 22, 11, 11, 9, 24, 12, 12, 11])
wsd.freeze_panes = 'C2'
fila = 2
for i, r in enumerate(cands, 1):
    r['id'] = 'm' + min([r['message_id']] + r['otros_mails'])  # código fijo: el primer mail de esa persona
    for t in r['trabajos_n']:
        vals = [r['id'], r.get('nombre'), t['empresa'], t['cargo'], t['rubro'], 'Sí' if t['gastro'] else 'No',
                'Sí' if t['atencion'] else 'No', 'Sí' if t['cocina'] else 'No', t['fechas_texto'] or 'sin fechas',
                t['desde'], t['hasta']]
        for c, v in enumerate(vals, 1):
            cell = wsd.cell(row=fila, column=c, value=v); cell.font = BODY
        for c in (10, 11):
            wsd.cell(row=fila, column=c).number_format = 'mm/yyyy'
        if t['aprox']:
            wsd.cell(row=fila, column=10).comment = Comment('El CV da solo el año: se tomó julio de ese año (mitad de año). Si empezó y terminó el mismo año, se tomaron 6 meses.', 'Salva')
        if t['desde'] and t['hasta']:
            wsd.cell(row=fila, column=12, value=f'=(YEAR(K{fila})-YEAR(J{fila}))*12+MONTH(K{fila})-MONTH(J{fila})+1').font = BODY
        else:
            wsd.cell(row=fila, column=12, value='sin fechas').font = BODY
        fila += 1
DET_LAST = max(fila - 1, 2)
RNG = lambda col: f"'Detalle trabajos'!${col}$2:${col}${DET_LAST}"

COLS = ['Pos.', 'Puntaje (0-100)', 'Permanencia', 'Nombre', 'Teléfono', 'Mail', 'Abrir CV', 'Localidad', 'Partido',
        'Puesto al que se postula', 'Experiencia (trabajo – cuánto duró)', 'Trabajos con fecha',
        'Promedio meses por trabajo', 'Trabajos de menos de 1 año', 'Trabajo más largo (meses)',
        'Disponibilidad horaria', 'Full time', 'Resumen', 'Alertas / revisar', 'Pts permanencia (35)',
        'Pts experiencia (30)', 'Pts cercanía (20)', 'Pts aviso (15)', 'Por qué esos puntos de experiencia',
        'Referencias', 'Nombre del archivo', 'Fecha del mail', 'Mandó el CV (veces)', 'ID candidato']
WID = [6, 10, 20, 24, 15, 28, 10, 18, 15, 16, 60, 9, 11, 11, 11, 22, 10, 45, 30, 11, 11, 11, 11, 34, 14, 34, 11, 10, 19]


GRUPO_TXT_X = lambda g: re.sub(r'^\d\.\s*', '', g)


def hoja_candidatos(nombre, lista, con_puntaje=True, extra_motivo=False):
    ws = wb.create_sheet(nombre)
    cols = (['Motivo'] if extra_motivo else []) + COLS
    wid = ([45] if extra_motivo else []) + WID
    encabezado(ws, cols, wid)
    off = 1 if extra_motivo else 0
    ci = {c: i + 1 + off for i, c in enumerate(COLS)}
    L = lambda c: get_column_letter(ci[c])
    for n, r in enumerate(lista, 1):
        f = n + 1
        if extra_motivo: ws.cell(row=f, column=1, value='; '.join(r['motivos']) or r['grupo'])
        puesto = r['puesto_eval'] + (' (deducido)' if r['puesto_deducido'] else '')
        idc = r['id']; idref = f'${L("ID candidato")}{f}'
        vals = {
            'Pos.': n, 'Permanencia': r['grupo'], 'Nombre': r.get('nombre'), 'Teléfono': r.get('telefono'),
            'Mail': r.get('mail'), 'Localidad': r.get('localidad') or r.get('zona_texto'), 'Partido': r['partido_n'] or 'sin dato',
            'Puesto al que se postula': puesto, 'Experiencia (trabajo – cuánto duró)': exp_texto(r),
            'Disponibilidad horaria': r.get('disponibilidad'), 'Full time': r.get('full_time'),
            'Resumen': r.get('resumen'), 'Alertas / revisar': '; '.join(r['alertas']),
            'Pts experiencia (30)': r['pts_exp'], 'Pts cercanía (20)': r['pts_zona'], 'Pts aviso (15)': r['pts_aviso'],
            'Por qué esos puntos de experiencia': r['exp_txt'], 'Referencias': r.get('referencias'),
            'Nombre del archivo': r.get('archivo'), 'Fecha del mail': r.get('fecha_mail'),
            'Mandó el CV (veces)': r['veces'], 'ID candidato': idc,
        }
        for c, v in vals.items():
            ws.cell(row=f, column=ci[c], value=v)
        # fórmulas que salen de la hoja Detalle trabajos
        ws.cell(row=f, column=ci['Trabajos con fecha'], value=f'=COUNTIFS({RNG("A")},{idref},{RNG("L")},">0")')
        tf = f'{L("Trabajos con fecha")}{f}'
        ws.cell(row=f, column=ci['Promedio meses por trabajo'], value=f'=IF({tf}=0,"sin fechas",ROUND(AVERAGEIFS({RNG("L")},{RNG("A")},{idref}),1))')
        ws.cell(row=f, column=ci['Trabajos de menos de 1 año'], value=f'=COUNTIFS({RNG("A")},{idref},{RNG("L")},"<12")')
        ws.cell(row=f, column=ci['Trabajo más largo (meses)'], value=f'=IF({tf}=0,"sin fechas",_xlfn.MAXIFS({RNG("L")},{RNG("A")},{idref}))')
        pr = f'{L("Promedio meses por trabajo")}{f}'; co = f'{L("Trabajos de menos de 1 año")}{f}'
        nt = len(r['trabajos_n'])
        base_sin = 0 if nt == 0 else 10
        ws.cell(row=f, column=ci['Pts permanencia (35)'],
                value=f'=IF({tf}=0,{base_sin},MAX(0,IF({pr}>=36,35,IF({pr}>=24,30,IF({pr}>=12,20,IF({pr}>=6,10,0))))-MIN(9,3*{co})))')
        ws.cell(row=f, column=ci['Puntaje (0-100)'],
                value=f'=SUM({L("Pts permanencia (35)")}{f}:{L("Pts aviso (15)")}{f})')
        link = ws.cell(row=f, column=ci['Abrir CV'], value='Abrir CV'); link.hyperlink = r['link']
        for c in range(1, len(cols) + 1):
            cell = ws.cell(row=f, column=c)
            cell.font = LINK if c == ci['Abrir CV'] else BODY
            cell.alignment = WRAP; cell.border = BORDE
        ws.cell(row=f, column=ci['Puntaje (0-100)']).font = Font(name=F, size=11, bold=True)
        ws.cell(row=f, column=ci['Puntaje (0-100)']).fill = AMARILLO
        if r['grupo'].startswith('1.'): ws.cell(row=f, column=ci['Permanencia']).fill = VERDE
        elif r['grupo'].startswith('Sin fechas'): ws.cell(row=f, column=ci['Permanencia']).fill = GRIS
        elif r['grupo'] == 'Cambia seguido': ws.cell(row=f, column=ci['Permanencia']).fill = ROJO
    return ws


def py_pts_perm(r):
    if r['n_con_fecha'] == 0: return 0 if not r['trabajos_n'] else 10
    p = round(r['prom'], 1)
    b = 35 if p >= 36 else 30 if p >= 24 else 20 if p >= 12 else 10 if p >= 6 else 0
    return max(0, b - min(9, 3 * r['cortos']))


for r in cands:
    r['pts_perm'] = py_pts_perm(r)
    r['total'] = r['pts_perm'] + r['pts_exp'] + r['pts_zona'] + r['pts_aviso']


# ---------- experiencia como encargado/a ----------
_ENC_SI = re.compile(r"\b(sub ?)?encargad[oa]s?\b(?!\s+de\s+(limpieza|edificio|cartel|repuesto|traslado|e-?commerce|publicidad|maestranza|mantenimiento|deposito|stock|redes))"
                     r"|\bjef[ea] de (cocina|salon|barra|produccion|camareros|plaza|turno|local|sucursal|partida)"
                     r"|\bgerente( de turno| de local| de sucursal)?\b(?! de (contabilidad|sistemas|marketing))"
                     r"|\bsupervis(or|ora)\b(?! de contabilidad)"
                     r"|\bresponsable (de|del) (turno|local|sucursal|salon|tienda|cocina|equipo)"
                     r"|\blider de (experiencia|sector|equipo|turno)|\bsegundo a cargo")
_ENC_NO = re.compile(r"project manager|social media|community manager|contabilidad|remiser|jefa? de catedra|jefa? de compras")
def cargos_encargado(ts):
    out = []
    for t in ts:
        txt = norm(t['cargo'])
        if _ENC_NO.search(txt) and not re.search(r"cocina|local|turno|salon", txt): continue
        if _ENC_SI.search(txt):
            m = meses(t)
            out.append(f"{t['cargo']} en {t['empresa'] or '?'}" + (f" ({m} meses)" if m else ""))
    return out
for r in cands: r['enc'] = cargos_encargado(r['trabajos_n'])

ORDEN = {'1. Estable (2 años o más)': 0, '2. Intermedio': 1, 'Sin fechas, revisar': 2}
cumplen = [r for r in cands if not r['motivos'] and r['grupo'] in ORDEN]
sinleer = sorted([r for r in cands if not r['motivos'] and r['grupo'] == 'CV sin leer'], key=lambda r: r.get('nombre') or '')
cambian = [r for r in cands if not r['motivos'] and r['grupo'] == 'Cambia seguido']
nocumplen = [r for r in cands if r['motivos']]
key = lambda r: (ORDEN.get(r['grupo'], 9), -r['total'], r.get('nombre') or '')
rank_at = sorted([r for r in cumplen if r['puesto_eval'] == 'Atención/Venta'], key=key)
rank_co = sorted([r for r in cumplen if r['puesto_eval'] == 'Cocina'], key=key)
cambian.sort(key=lambda r: -r['total'])
nocumplen.sort(key=lambda r: (r['motivos'][0][:20], -r['total']))

hoja_candidatos('Ranking Atención', rank_at)
hoja_candidatos('Ranking Cocina', rank_co)
if cambian: hoja_candidatos('Cambian seguido', cambian)
if nocumplen: hoja_candidatos('No cumplen', nocumplen, extra_motivo=True)
ws_sl = hoja_candidatos('Revisar a mano', sinleer)
for row in range(2, ws_sl.max_row + 1):
    r = sinleer[row - 2]
    ws_sl.cell(row=row, column=COLS.index('Resumen') + 1, value=((r.get('resumen') or '') + ' | ' + (r.get('notas') or '')).strip(' |'))

# Duplicados
wsdup = wb.create_sheet('Duplicados')
encabezado(wsdup, ['Nombre', 'Mail', 'Teléfono', 'Fecha del mail', 'Archivo', 'Abrir mail', 'Quedó en la planilla como'], [26, 30, 16, 12, 34, 10, 14])
for n, (x, best) in enumerate(dups, 2):
    for c, v in enumerate([x.get('nombre'), x.get('mail'), x.get('telefono'), x.get('fecha_mail'), x.get('archivo'), 'Abrir', best['id']], 1):
        cell = wsdup.cell(row=n, column=c, value=v); cell.font = BODY
    wsdup.cell(row=n, column=6).hyperlink = x['link']; wsdup.cell(row=n, column=6).font = LINK

# Hoja resumen + explicación (primera)
wsr = wb.create_sheet('Cómo leer esto', 0)
wsr.column_dimensions['A'].width = 34; wsr.column_dimensions['B'].width = 14; wsr.column_dimensions['C'].width = 90
T = lambda r, c, v, **k: wsr.cell(row=r, column=c, value=v)
rows = [
    ('Búsqueda Cocinera/o y Vendedor/a – CVs recibidos desde el 23/10 hasta el ' + CORTE, None, None),
    ('Datos hasta', CORTE, 'Fecha y hora del último mail con CV leído. Los CVs que llegaron después se suman en la próxima actualización.'),
    ('', None, None),
    ('NÚMEROS', None, None),
    ('Mails con CV leídos', len(regs), 'Respuestas a lapresumidapastas@gmail.com desde el 23/10 hasta el ' + CORTE + '.'),
    ('Personas distintas', len(cands), 'Si alguien mandó el CV más de una vez, cuenta una sola vez (ver pestaña Duplicados).'),
    ('Ranking Atención', len(rank_at), 'Cumplen el aviso y no cambian de trabajo cada pocos meses.'),
    ('Ranking Cocina', len(rank_co), 'Ídem, para cocina.'),
    ('   de esos, estables (2 años o más)', sum(1 for r in rank_at + rank_co if r['grupo'].startswith('1.')), 'Aparecen primero en cada ranking, pintados de verde.'),
    ('   de esos, sin fechas, revisar', sum(1 for r in rank_at + rank_co if r['grupo'].startswith('Sin fechas')), 'El CV no tiene fechas: no se puede medir permanencia. Van al final, en gris.'),
    ('Descartados (no cumplen, cambian seguido o menos de 79 puntos)', int(os.environ.get('CVS_DESCARTADOS', '0') or 0) + len(nocumplen), 'No cumplen el aviso (zona, full time o experiencia), cambian de trabajo seguido o sacaron menos de 79 puntos. Se borraron del tablero y del Excel; solo queda anotado que el mail ya se leyó.'),
    ('Encargados', sum(1 for r in cands if r['enc']), 'Personas que ya trabajaron como encargado/a, jefe/a o supervisor/a. Pestaña Encargados (agrupados por lista).'),
    ('Encargados estables', sum(1 for r in cands if r['enc'] and r['grupo'].startswith('1.') and not r['motivos']), 'Fueron encargado/a y duran 2 años o más en promedio, y cumplen el aviso. Pestaña Encargados estables.'),
    ('Revisar a mano', len(sinleer), 'No se pudo leer el CV: mandaron un link de Drive, el mail no traía adjunto o el archivo es ilegible. Cumplen la zona o no se sabe; abrilos con el link.'),
    ('', None, None),
    ('DE DÓNDE SALE CADA DATO', None, None),
    ('Nombre, teléfono, mail, zona', None, 'Copiados del CV. Si el CV no trae mail, se usa el mail desde el que lo mandó.'),
    ('Puesto', None, 'Del asunto del mail ("Búsqueda Cocinera/o" o "Vendedor/a"). Si el asunto no lo dice, del texto del mail o del CV. Si nadie lo dice, se asume Atención, salvo que tenga experiencia cocinando: ahí va a Cocina. Se marca "(deducido)".'),
    ('Experiencia', None, 'Cada trabajo del CV con sus fechas. El detalle completo (empresa, rubro, desde, hasta, meses) está en la pestaña Detalle trabajos.'),
    ('Partido', None, 'Sale del barrio que pone el CV (ej. Munro → Vicente López, Boulogne → San Isidro, Villa Ballester → San Martín).'),
    ('Abrir CV', None, 'Link al mail en Gmail con el CV adjunto. Hay que estar logueado en lapresumidapastas@gmail.com.'),
    ('', None, None),
    ('CÓMO SE CALCULÓ', None, None),
    ('Meses en cada trabajo', None, 'Desde el mes que empezó hasta el que terminó, contando los dos. Ej: 03/2021 a 02/2023 = 24 meses. "Actualidad" = octubre 2026. Si el CV pone solo el año, se toma julio de ese año (en la celda hay un comentario).'),
    ('Promedio meses por trabajo', None, 'Suma de los meses de los trabajos con fecha ÷ cantidad de esos trabajos. Fórmula sobre la pestaña Detalle trabajos. Los trabajos sin fechas no entran.'),
    ('Trabajos de menos de 1 año', None, 'Cuántos trabajos con fecha duraron menos de 12 meses.'),
    ('Permanencia', None, '1. Estable: promedio de 24 meses o más.  2. Intermedio: el resto.  Cambia seguido: promedio menor a 12 meses Y 2 o más trabajos de menos de 1 año (va a su pestaña).  Sin fechas, revisar: ningún trabajo tiene fechas.'),
    ('', None, None),
    ('REQUISITOS DEL AVISO (filtro)', None, None),
    ('Zona', None, 'Tiene que vivir en Vicente López, San Isidro o San Martín. Si vive en otro partido o en CABA → No cumple. Si el CV no dice dónde vive → queda, con alerta "Zona sin dato".'),
    ('Full time', None, 'Si el CV dice solo medio día, un turno o fines de semana → No cumple. Si no dice nada → queda, con alerta para preguntar en la entrevista.'),
    ('Experiencia / referencias', None, 'Cocina: tiene que tener algún trabajo en cocina o gastronomía. Atención: tiene que tener al menos un trabajo (si nunca trabajó, no puede dar referencias).'),
    ('Lo que NO se usó', None, 'Edad, sexo, estado civil, hijos, nacionalidad ni foto. No se miraron ni se cargaron.'),
    ('', None, None),
    ('PUNTAJE (0 a 100)', None, 'Columna amarilla = suma de las cuatro partes. Cada parte tiene su columna, así se ve de dónde sale.'),
    ('Permanencia', 35, 'Promedio 36 meses o más: 35 · 24 a 36: 30 · 12 a 24: 20 · 6 a 12: 10 · menos de 6: 0. Se resta 3 por cada trabajo de menos de 1 año (máximo 9). Sin fechas: 10. Sin experiencia: 0.'),
    ('Experiencia', 30, 'Atención: 2 años o más atendiendo en gastronomía 30 · algo de gastronomía 22 · 2 años o más de atención al público en otro rubro 20 · algo de atención 14 · nada 0.  Cocina: 2 años o más cocinando 30 · 1 a 2 años 22 · menos 15 · gastronomía sin cocinar 8 · nada 0.'),
    ('Cercanía', 20, 'Vicente López o San Isidro (donde están los locales) 20 · San Martín 12 · sin dato 5.'),
    ('Aviso', 15, 'Full time: dice que sí 10 · no dice 5.  Más 5 si da referencias o tiene formación en cocina/manipulación de alimentos (3 si dice "referencias a pedido").'),
    ('', None, None),
    ('ORDEN', None, 'En cada ranking: primero los estables, después los intermedios, al final los sin fechas. Dentro de cada grupo, de mayor a menor puntaje.'),
    ('OJO', None, 'Los datos los leyó Claude de cada CV; puede haber errores en CVs con mal formato o escaneados. Antes de llamar, abrí el CV con el link.'),
]
for i, (a, b, c) in enumerate(rows, 1):
    ca = wsr.cell(row=i, column=1, value=a); cb = wsr.cell(row=i, column=2, value=b); cc = wsr.cell(row=i, column=3, value=c)
    for x in (ca, cb, cc): x.font = BODY; x.alignment = WRAP
    if a and a.isupper() or i == 1: ca.font = Font(name=F, bold=True, size=12 if i == 1 else 11)
for i in range(1, len(rows) + 1):
    if wsr.cell(row=i, column=1).value in ('NÚMEROS', 'DE DÓNDE SALE CADA DATO', 'CÓMO SE CALCULÓ', 'REQUISITOS DEL AVISO (filtro)', 'PUNTAJE (0 a 100)'):
        for c in (1, 2, 3): wsr.cell(row=i, column=c).fill = AMARILLO

wse = wb.create_sheet('Encargados')
encabezado(wse, ['Lista', 'Pos.', 'Puntaje', 'Nombre', 'Teléfono', 'Partido', 'Puesto', 'Permanencia', 'Experiencia como encargado/a', 'Abrir CV', 'ID candidato'], [16, 6, 9, 26, 15, 15, 16, 20, 70, 10, 19])
wse.freeze_panes = 'E2'
fe = 2
for nombre_l, rs in (('Atención', rank_at), ('Cocina', rank_co), ('Cambian seguido', cambian), ('Revisar a mano', sinleer), ('No cumplen', nocumplen)):
    for pos, r in enumerate(rs, 1):
        if not r['enc']: continue
        vals = [nombre_l, pos, r['total'], r.get('nombre'), r.get('telefono'), r['partido_n'] or 'sin dato', r['puesto_eval'], GRUPO_TXT_X(r['grupo']), ' ; '.join(r['enc']), 'Abrir CV', r['id']]
        for c, v in enumerate(vals, 1):
            cell = wse.cell(row=fe, column=c, value=v); cell.font = BODY; cell.alignment = WRAP; cell.border = BORDE
        wse.cell(row=fe, column=10).hyperlink = r['link']; wse.cell(row=fe, column=10).font = LINK
        fe += 1
wse.cell(row=fe + 1, column=1, value='Se marca "encargado/a" cuando algún puesto del CV dice encargado/a (de local, turno, salón, caja, cocina…), subencargado/a, jefe/a de cocina o salón, gerente de turno, supervisor/a, responsable de turno o local, o líder de equipo. No cuenta "encargada de limpieza", "de edificio" y similares, que no son a cargo de un local o equipo.').font = Font(name=F, size=9, italic=True)
wsx = wb.create_sheet('Encargados estables')
encabezado(wsx, ['Lista', 'Pos.', 'Puntaje', 'Nombre', 'Teléfono', 'Partido', 'Puesto', 'Permanencia', 'Experiencia como encargado/a', 'Abrir CV', 'ID candidato'], [16, 6, 9, 26, 15, 15, 16, 20, 70, 10, 19])
wsx.freeze_panes = 'E2'
fe = 2
for nombre_l, rs in (('Atención', rank_at), ('Cocina', rank_co), ('No cumplen', nocumplen)):
    for pos, r in enumerate(rs, 1):
        if not r['enc'] or not r['grupo'].startswith('1.'): continue
        vals = [nombre_l, pos, r['total'], r.get('nombre'), r.get('telefono'), r['partido_n'] or 'sin dato', r['puesto_eval'], GRUPO_TXT_X(r['grupo']), ' ; '.join(r['enc']), 'Abrir CV', r['id']]
        for c, v in enumerate(vals, 1):
            cell = wsx.cell(row=fe, column=c, value=v); cell.font = BODY; cell.alignment = WRAP; cell.border = BORDE
        wsx.cell(row=fe, column=10).hyperlink = r['link']; wsx.cell(row=fe, column=10).font = LINK
        fe += 1
wsx.cell(row=fe + 1, column=1, value='Se marca "encargado/a" cuando algún puesto del CV dice encargado/a (de local, turno, salón, caja, cocina…), subencargado/a, jefe/a de cocina o salón, gerente de turno, supervisor/a, responsable de turno o local, o líder de equipo. Estables = en promedio duraron 2 años o más en sus trabajos (columna Permanencia).').font = Font(name=F, size=9, italic=True)
wb.move_sheet('Detalle trabajos', offset=len(wb.sheetnames))
wb.save(OUT)
json.dump({'regs': len(regs), 'cands': len(cands), 'at': len(rank_at), 'co': len(rank_co), 'cambian': len(cambian),
           'nocumplen': len(nocumplen), 'sinleer': len(sinleer), 'dups': len(dups),
           'motivos': sorted(set(m.split('(')[0].strip() for r in nocumplen for m in r['motivos'])),
           'top_at': [(r['nombre'], r['total'], r['grupo'], r['partido_n']) for r in rank_at[:8]],
           'top_co': [(r['nombre'], r['total'], r['grupo'], r['partido_n']) for r in rank_co[:8]],
           'pyscores': {r['id']: r['total'] for r in cands}},
          open(os.path.join(BASE, 'resumen.json'), 'w'), ensure_ascii=False, indent=1)
print('OK', OUT)

# ---------- datos para el tablero ----------
def wa(tel):
    d = re.sub(r'\D', '', str(tel or ''))
    if d.startswith('549') and len(d) == 13: return d
    if d.startswith('54') and len(d) == 12: return '549' + d[2:]
    if d.startswith('0'): d = d[1:]
    if len(d) == 12 and d[2:4] == '15' and d.startswith('11'): d = '11' + d[4:]
    if len(d) == 10: return '549' + d
    return ''

def bloque(r, lista):
    ts = []
    for t in r['trabajos_n']:
        ts.append({'e': t['empresa'], 'c': t['cargo'], 'r': t['rubro'], 'g': t['gastro'], 'a': t['atencion'], 'k': t['cocina'],
                   'f': t['fechas_texto'] or 'sin fechas', 'm': meses(t), 'ap': t['aprox']})
    return {'id': r['id'], 'lista': lista, 'n': r.get('nombre') or '(sin nombre)', 'tel': r.get('telefono') or '', 'wa': wa(r.get('telefono')),
            'mail': r.get('mail') or '', 'loc': r.get('localidad') or r.get('zona_texto') or '', 'part': r['partido_n'] or '',
            'puesto': r['puesto_eval'], 'ded': r['puesto_deducido'], 'grupo': r['grupo'], 'prom': round(r['prom'], 1) if r['prom'] is not None else None,
            'cortos': r['cortos'], 'largo': r['largo'], 'disp': r.get('disponibilidad') or '', 'ft': r.get('full_time') or '',
            'res': r.get('resumen') or '', 'al': r['alertas'], 'mot': r['motivos'], 'pp': r['pts_perm'], 'pe': r['pts_exp'],
            'pz': r['pts_zona'], 'pa': r['pts_aviso'], 'tot': r['total'], 'ext': r['exp_txt'], 'ref': r.get('referencias') or '',
            'arch': r.get('archivo') or '', 'link': r['link'], 'fecha': r.get('fecha_mail') or '', 'veces': r['veces'], 'ts': ts, 'enc': r['enc'], 'mids': [r['message_id']] + r['otros_mails']}

tab = []
for lista, rs in (('atencion', rank_at), ('cocina', rank_co), ('cambian', cambian), ('nocumplen', nocumplen), ('revisar', sinleer)):
    for pos, r in enumerate(rs, 1):
        b = bloque(r, lista); b['pos'] = pos; tab.append(b)
json.dump({'generado': CORTE, 'corte': CORTE, 'descartados': int(os.environ.get('CVS_DESCARTADOS', '0') or 0) + len(nocumplen), 'mails': len(regs), 'personas': len(cands), 'cands': tab}, open(os.path.join(BASE, 'datos_tablero.json'), 'w'), ensure_ascii=False)
print('tablero', len(tab))
