"""Actualiza el tablero de selección de La Presumida.

Uso:
  python3 -I construir.py abrir  REPO WORK CLAVE
      Desencripta REPO/estado.bin y deja en WORK: registros/*.json y estado.json
  python3 -I construir.py armar  REPO WORK CLAVE "DD/MM/AAAA HH:MM"
      Arma el Excel y los dos tableros con todo lo que hay en WORK/registros,
      encripta los CVs nuevos (los busca en WORK/adjuntos/<message_id>/),
      y guarda en REPO: index.html, datos.bin, cv/*.bin y estado.bin (encriptado).
      Deja en WORK: el Excel, tablero_seleccion.html (versión Claude) y
      nuevos_artifact.json (CVs nuevos para subir al tablero de Claude).

REPO = carpeta del repositorio de GitHub; WORK = carpeta de trabajo.
"""
import sys, os, json, glob, base64, secrets, shutil, subprocess

from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

AQUI = os.path.dirname(os.path.abspath(__file__))
EXT_OK = ('pdf', 'jpg', 'jpeg', 'png', 'webp', 'docx', 'doc', 'odt', 'rtf')


def clave(repo, pw):
    k = json.load(open(os.path.join(repo, 'kdf.json')))
    return AESGCM(PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=base64.b64decode(k['salt']),
                             iterations=k['it']).derive(pw.encode()))


def enc(g, b):
    iv = secrets.token_bytes(12)
    return iv + g.encrypt(iv, b, None)


def dec(g, b):
    return g.decrypt(b[:12], b[12:], None)


def abrir(repo, work, pw):
    g = clave(repo, pw)
    st = json.loads(dec(g, open(os.path.join(repo, 'estado.bin'), 'rb').read()))
    os.makedirs(os.path.join(work, 'registros'), exist_ok=True)
    for r in st['registros']:
        json.dump(r, open(os.path.join(work, 'registros', r['_archivo']), 'w'), ensure_ascii=False, indent=1)
    json.dump(st['estado'], open(os.path.join(work, 'estado.json'), 'w'), ensure_ascii=False, indent=1)
    print('registros:', len(st['registros']), '| procesados:', len(st['estado']['procesados']),
          '| datos hasta:', st['estado']['corte'])


def tipo_archivo(p):
    b = os.path.basename(p)
    e = b.lower().rsplit('.', 1)[-1] if '.' in b else ''
    if e in EXT_OK: return e
    kind = subprocess.run(['file', '-b', '--mime-type', p], capture_output=True, text=True).stdout.strip()
    return {'application/pdf': 'pdf', 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp'}.get(kind, '')


def buscar_cv(work, mids, destino_sin_ext):
    """Elige el mejor adjunto de los mails de la persona y lo deja como PDF o imagen."""
    def prio(p):
        b = os.path.basename(p).lower(); e = tipo_archivo(p)
        return (0 if ('cv' in b or 'curr' in b) else 1, 0 if e == 'pdf' else 1 if e in ('docx', 'doc', 'odt', 'rtf') else 2)
    cands = []
    for mid in mids:
        d = os.path.join(work, 'adjuntos', mid.split('_')[0])
        if os.path.isdir(d):
            cands += [os.path.join(d, x) for x in os.listdir(d) if not x.startswith('pag') and not x.endswith('.txt')]
    cands = [p for p in cands if tipo_archivo(p)]
    if not cands: return None
    p = sorted(cands, key=prio)[0]; e = tipo_archivo(p)
    if e in ('docx', 'doc', 'odt', 'rtf'):
        tmp = destino_sin_ext + '_conv'; os.makedirs(tmp, exist_ok=True)
        src = os.path.join(tmp, 'cv.' + e); shutil.copy(p, src)
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, src], capture_output=True, timeout=180)
        out = os.path.join(tmp, 'cv.pdf')
        if not os.path.exists(out): return None
        shutil.move(out, destino_sin_ext + '.pdf'); shutil.rmtree(tmp, ignore_errors=True)
        return destino_sin_ext + '.pdf'
    if e in ('png', 'jpg', 'jpeg', 'webp'):
        from PIL import Image, ImageOps
        im = ImageOps.exif_transpose(Image.open(p)).convert('RGB'); im.thumbnail((1700, 1700))
        im.save(destino_sin_ext + '.jpg', 'JPEG', quality=82, optimize=True)
        return destino_sin_ext + '.jpg'
    shutil.copy(p, destino_sin_ext + '.pdf')
    return destino_sin_ext + '.pdf'


M = 'https://macproia88.github.io/marcas/'
PAGINAS = [(M + 'tablero-presumida-v2.html', 'Tablero La Presumida'), (M + 'mostrador.html', 'Tablero de mostrador'),
           (M + 'apps.html', 'Tablero de apps'), (M + 'tablero-club.html', 'Tablero del club'),
           (M + 'reconocimiento.html', 'Agenda de reconocimiento'), (M + 'ranking-anual.html', 'Ranking del año'), (M + 'ranking-mostrador.html', 'Ranking de mostrador'), (None, 'Selección de personal (CVs)')]
TABLEROS = {}  # clave -> (url, nombre); se completa en armar()


def sello_y_links(s, paginas=None):
    """Arriba: fecha y hora de la última actualización. Abajo: links a las otras páginas del negocio (igual que en el repo marcas)."""
    import re, datetime as dt
    ahora = dt.datetime.now(dt.timezone(dt.timedelta(hours=-3)))
    s = re.sub(r'<div id="lp-upd"[^>]*>.*?</div>\s*', '', s, flags=re.S)
    s = re.sub(r'<nav id="lp-nav".*?</nav>\s*', '', s, flags=re.S)
    upd = (f'<div id="lp-upd" style="font:12px/1.4 system-ui,-apple-system,Segoe UI,sans-serif;color:#6b6b6b;text-align:right;'
           f'padding:6px 16px;position:relative;z-index:31">Última actualización: {ahora:%d/%m/%Y} a las {ahora:%H:%M} (hora Argentina)</div>\n')
    links = ' · '.join(f'<b>{n}</b>' if not h else f'<a href="{h}" style="color:#b85042;text-decoration:none;white-space:nowrap">{n}</a>'
                       for h, n in (paginas or PAGINAS))
    nav = (f'<nav id="lp-nav" style="font:13px/1.6 system-ui,-apple-system,Segoe UI,sans-serif;max-width:1180px;margin:24px auto 16px;'
           f'padding:12px 16px;border-top:1px solid rgba(128,128,128,.35);color:#6b6b6b">'
           f'<div style="font-weight:600;margin-bottom:4px">Páginas del negocio</div>{links}</nav>\n')
    i = s.find('<body>') + len('<body>')
    s = s[:i] + '\n' + upd + s[i:]
    j = s.rfind('</body>')
    return s[:j] + nav + s[j:]


def github_index(t, datos_file="datos.bin", store_key="sel_estados", tablero="equipo"):
    """Convierte la plantilla del tablero de Claude en la versión con contraseña para GitHub Pages."""
    s = '<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><meta name="robots" content="noindex,nofollow">' + t.replace('<meta name="claude:viewport-fit-sides" content="cover">', '')
    s = s.replace('</style>', """.lock{position:fixed;inset:0;z-index:30;background:var(--paper);display:flex;align-items:center;justify-content:center;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
.lock form{width:100%;max-width:360px;display:grid;gap:12px}
.lock h2{font-family:var(--display);font-size:30px;text-transform:uppercase;margin:0}
.lock input[type=password]{font:inherit;font-size:17px;padding:12px;border:1.5px solid var(--ink);border-radius:8px;background:var(--card);color:var(--ink)}
.lock .err{color:var(--bad);font-size:14px;min-height:1.2em}
.lock label.rec{display:flex;gap:8px;align-items:center;font-size:14px;color:var(--muted)}
.lock button{font:inherit;font-weight:700;font-size:16px;background:var(--yolk);color:var(--ink-on-yolk);border:0;border-radius:8px;padding:12px;cursor:pointer}
</style></head><body>
<div class="lock" id="lock"><form id="lockForm">
  <h2>Selección <span style="background:var(--yolk);padding:0 6px">Octubre</span></h2>
  <p style="margin:0;color:var(--muted)">Lista de candidatos de La Presumida. Ingresá la contraseña para ver los datos y los CVs.</p>
  <input type="password" id="pw" autocomplete="current-password" placeholder="Contraseña" required>
  <label class="rec"><input type="checkbox" id="rec"> Recordar en este dispositivo</label>
  <button type="submit" id="pwBtn">Entrar</button>
  <div class="err" id="pwErr"></div>
</form></div>""", 1)
    viejo = '<script id="datos" type="application/json">/*DATA*/</script>\n<script>\n(function(){\n  const D = JSON.parse(document.getElementById(\'datos\').textContent);'
    assert viejo in s
    s = s.replace(viejo, """<script>
(async function(){
  const b64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
  let KEY = null;
  async function derivar(pw){
    const k = await (await fetch('kdf.json', {cache:'no-store'})).json();
    const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({name:'PBKDF2', salt:b64(k.salt), iterations:k.it, hash:'SHA-256'}, base, {name:'AES-GCM', length:256}, true, ['decrypt']);
  }
  async function descifrar(buf){
    const u = new Uint8Array(buf);
    return crypto.subtle.decrypt({name:'AES-GCM', iv:u.slice(0,12)}, KEY, u.slice(12));
  }
  async function cargarDatos(){
    const buf = await (await fetch('datos.bin', {cache:'no-store'})).arrayBuffer();
    return JSON.parse(new TextDecoder().decode(await descifrar(buf)));
  }
  let D = null;
  try { const g = localStorage.getItem('sel_key'); if (g){ KEY = await crypto.subtle.importKey('raw', b64(g), 'AES-GCM', true, ['decrypt']); D = await cargarDatos(); } } catch(e){ KEY = null; D = null; }
  if (!D){
    D = await new Promise(resolve => {
      const f = document.getElementById('lockForm');
      f.addEventListener('submit', async ev => {
        ev.preventDefault();
        const btn = document.getElementById('pwBtn'); btn.disabled = true; btn.textContent = 'Abriendo…';
        document.getElementById('pwErr').textContent = '';
        try {
          KEY = await derivar(document.getElementById('pw').value);
          const datos = await cargarDatos();
          if (document.getElementById('rec').checked){
            try { const raw = new Uint8Array(await crypto.subtle.exportKey('raw', KEY)); localStorage.setItem('sel_key', btoa(String.fromCharCode(...raw))); } catch(e){}
          }
          resolve(datos);
        } catch(e){
          document.getElementById('pwErr').textContent = 'Contraseña incorrecta. Probá de nuevo.';
          btn.disabled = false; btn.textContent = 'Entrar';
        }
      });
      document.getElementById('pw').focus();
    });
  }
  document.getElementById('lock').remove();
""")
    viejo = """    if (/\\.(jpg|png|webp)$/i.test(c.cv)){ const im = el('img'); im.alt = 'CV de '+c.n; im.src = c.cv; body.append(im); return; }
    const msg = el('p','vmsg','Cargando el CV…'); body.append(msg);
    try {
      const lib = await pdfjs();
      const buf = await (await fetch(c.cv)).arrayBuffer();"""
    assert viejo in s
    s = s.replace(viejo, """    const msg = el('p','vmsg','Cargando el CV…'); body.append(msg);
    try {
      const buf = await descifrar(await (await fetch(c.cv)).arrayBuffer());
      if (c.cvt === 'img'){ msg.remove(); const im = el('img'); im.alt = 'CV de '+c.n; im.src = URL.createObjectURL(new Blob([buf])); body.append(im); return; }
      const lib = await pdfjs();""")
    viejo = """  function saveEstado(id, patch){
    estados[id] = Object.assign({estado:'', nota:''}, estados[id], patch);
    if (!db) return;"""
    assert viejo in s
    s = s.replace(viejo, """  try { Object.assign(estados, JSON.parse(localStorage.getItem('sel_estados') || '{}')); } catch(e){}
  function saveEstado(id, patch){
    estados[id] = Object.assign({estado:'', nota:''}, estados[id], patch);
    try { localStorage.setItem('sel_estados', JSON.stringify(estados)); } catch(e){}
    if (!db) return;""")
    s = s.replace("Los estados se guardan para todos los que abren este tablero.", "Los estados y notas se guardan solo en este dispositivo.")
    s = s.replace("fetch('datos.bin', {cache:'no-store'})", "fetch('" + datos_file + "', {cache:'no-store'})")
    s = s.replace("localStorage.getItem('sel_estados')", "localStorage.getItem('" + store_key + "')").replace("localStorage.setItem('sel_estados',", "localStorage.setItem('" + store_key + "',")
    s = s.replace("'Descartado: se borra de este tablero en la próxima actualización (todos los días 7:40). Si fue un error, cambiá el estado antes.'",
                  "'Descartado: tocá «Mandar descartados» (abajo) y mandá el mail; al otro día a las 7:40 se borra de este tablero.'")
    # barra para mandar los descartados por mail (la actualización diaria lee esos mails)
    s = s.replace('</style>', '''.envbar{position:fixed;left:0;right:0;bottom:0;z-index:15;background:var(--ink);color:var(--paper);display:flex;flex-wrap:wrap;gap:8px;align-items:center;justify-content:space-between;padding:10px max(16px,env(safe-area-inset-right,0px)) calc(10px + env(safe-area-inset-bottom,0px)) max(16px,env(safe-area-inset-left,0px));font-size:14px}
.envbar[hidden]{display:none!important}
.envbar button{font:inherit;font-weight:700;background:var(--yolk);color:var(--ink-on-yolk);border:0;border-radius:8px;padding:8px 12px;cursor:pointer}
.reenviar{font-size:13px;color:var(--link);background:none;border:0;padding:0;cursor:pointer;text-decoration:underline}
</style>''', 1)
    s = s.replace('<div class="list" id="list"></div>', '<div class="list" id="list"></div>\n  <p><button type="button" class="reenviar" id="reenviar">Volver a mandar todos mis descartados</button></p>\n  <div style="height:70px"></div>', 1)
    s = s.replace('</main>', '</main>\n<div class="envbar" id="envbar" hidden><span id="envtxt"></span><button type="button" id="envbtn">Mandar descartados</button></div>', 1)
    viejo = "    try { localStorage.setItem('" + store_key + "', JSON.stringify(estados)); } catch(e){}"
    assert viejo in s
    s = s.replace(viejo, viejo + "\n    if (patch && 'estado' in patch) barraEnvio();")
    viejo = "  fillSelects(); renderTabs(); render();"
    assert viejo in s
    s = s.replace(viejo, viejo + '''
  const ENV_KEY = '__SK___enviados';
  let enviados = {}; try { enviados = JSON.parse(localStorage.getItem(ENV_KEY) || '{}'); } catch(e){}
  async function codigo(){ try { const raw = new Uint8Array(await crypto.subtle.exportKey('raw', KEY)); const h = new Uint8Array(await crypto.subtle.digest('SHA-256', raw)); return [...h].map(x=>x.toString(16).padStart(2,'0')).join('').slice(0,10); } catch(e){ return ''; } }
  function descartadosDe(todos){ return D.cands.filter(c => (estados[c.id]||{}).estado === 'descartado' && (todos || !enviados[c.id])); }
  function barraEnvio(){
    const p = descartadosDe(false), bar = document.getElementById('envbar');
    if (!p.length){ bar.hidden = true; return; }
    document.getElementById('envtxt').textContent = 'Tenés ' + p.length + ' descartado' + (p.length===1?'':'s') + ' sin mandar. Mandalos para que mañana se saquen de la lista.';
    bar.hidden = false;
  }
  async function mandar(todos){
    const p = descartadosDe(todos); if (!p.length){ alertaSuave('No hay descartados para mandar.'); return; }
    const cuerpo = 'TABLERO: __TAB__\\nCODIGO: ' + await codigo() + '\\n\\nDescartados (no tocar este texto):\\n' + p.map(c => c.id + ' | ' + c.n).join('\\n') + '\\n';
    location.href = 'mailto:lapresumidapastas@gmail.com?subject=' + encodeURIComponent('DESCARTES SELECCION __TAB__') + '&body=' + encodeURIComponent(cuerpo);
    p.forEach(c => enviados[c.id] = 1); try { localStorage.setItem(ENV_KEY, JSON.stringify(enviados)); } catch(e){}
    barraEnvio();
  }
  function alertaSuave(t){ document.getElementById('envtxt').textContent = t; document.getElementById('envbar').hidden = false; setTimeout(barraEnvio, 2500); }
  document.getElementById('envbtn').onclick = () => mandar(false);
  document.getElementById('reenviar').onclick = () => mandar(true);
  barraEnvio();''')
    s = s.replace("      if (!window.claude || !window.claude.use) { $('#sync').textContent = 'Los estados se ven solo en esta pantalla.'; return; }",
                  "      if (!window.claude || !window.claude.use) { return; }")
    s = s.replace("})();\n</script>", "})();\n</script></body></html>")
    s = s.replace('__SK__', store_key).replace('__TAB__', tablero)
    assert '/*DATA*/' not in s
    return s


def es_encargado(c):
    return bool(c.get('enc')) and c['lista'] in ('atencion', 'cocina')


TABS_EQUIPO = None  # usa las pestañas de la plantilla (Atención, Cocina, Revisar a mano)
TABS_ENC = [
    ['enc', 'Encargados', 'Fueron encargados/as (de local, turno, salón, caja o cocina, supervisores, gerentes de turno) y cumplen el aviso. De mayor a menor puntaje. Con el filtro "Encargados estables" ves solo los que duraron 2 años o más en promedio.'],
]


def datos_tablero(da, clave_t, excl_ids, corte):
    t = json.loads(json.dumps(da))
    if clave_t == 'equipo':
        t['cands'] = [c for c in t['cands'] if c['id'] not in excl_ids]
        t['tabs'] = [
            ['atencion', 'Atención', 'Cumplen el aviso y no cambian de trabajo seguido. Primero los estables (2 años o más), después los intermedios y al final los que no tienen fechas; dentro de cada grupo, de mayor a menor puntaje.'],
            ['cocina', 'Cocina', 'Igual que Atención, para la búsqueda de cocinera/o.'],
            ['enc', 'Encargados', 'Los de Atención y Cocina que fueron encargados/as (de local, turno, salón, caja o cocina, supervisores, gerentes de turno). Son los mismos que ve Sebastián en su tablero; acá siguen apareciendo también en Atención o Cocina.'],
            ['revisar', 'Revisar a mano', 'No se pudo leer el CV (mandaron un link de Drive, el mail no traía adjunto o el archivo no se abrió). Abrilos con el botón.']]
        t['titulo'] = ['Selección', 'Equipo']; t['quien'] = 'Tablero de Ale (vendedores/as, cocina y encargados)'
    else:
        cs = [c for c in t['cands'] if es_encargado(c) and c['id'] not in excl_ids]
        for c in cs: c['lista'] = 'enc'
        cs.sort(key=lambda c: (not c['grupo'].startswith('1.'), -c['tot']))
        cont = {}
        for c in cs: cont[c['lista']] = cont.get(c['lista'], 0) + 1; c['pos'] = cont[c['lista']]
        t['cands'] = cs; t['tabs'] = TABS_ENC
        t['titulo'] = ['Selección', 'Encargados']; t['quien'] = 'Tablero de Sebastián (búsqueda de encargados/as)'
    t['personas'] = len(t['cands'])
    t['desc_tablero'] = len(excl_ids)
    return t


def descartes(work, clave_t, carpeta):
    """Lee los estados guardados en la base de un tablero (bajados con ArtifactData list + out_dir)
    y suma a estado.json los IDs marcados 'descartado' en ese tablero."""
    est_p = os.path.join(work, 'estado.json'); est = json.load(open(est_p))
    excl = est.setdefault('excl', {'equipo': [], 'encargados': []})
    nuevos = []
    for f in glob.glob(os.path.join(carpeta, '**', '*.json'), recursive=True):
        doc = json.load(open(f))
        v = doc.get('data', doc) if isinstance(doc, dict) else {}
        did = doc.get('id') or doc.get('doc_id') or os.path.basename(f)[:-5]
        if isinstance(v, dict) and v.get('estado') == 'descartado' and did not in excl[clave_t]:
            excl[clave_t].append(did); nuevos.append(did)
    json.dump(est, open(est_p, 'w'), ensure_ascii=False, indent=1)
    print(f'{clave_t}: {len(nuevos)} descartados nuevos | total en ese tablero: {len(excl[clave_t])}')


def descartes_mail(repo, work, pw, archivo):
    """Lee el texto de un mail 'DESCARTES SELECCION <tablero>' que mandan las páginas de GitHub
    y suma esos IDs a estado.json -> excl[tablero]. Controla el CODIGO (sale de la contraseña)."""
    import hashlib, re
    k = json.load(open(os.path.join(repo, 'kdf.json')))
    raw = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=base64.b64decode(k['salt']), iterations=k['it']).derive(pw.encode())
    ok_cod = hashlib.sha256(raw).hexdigest()[:10]
    txt = open(archivo, encoding='utf-8', errors='replace').read()
    mt = re.search(r'TABLERO:\s*(equipo|encargados)', txt); mc = re.search(r'CODIGO:\s*([0-9a-f]{10})', txt)
    if not mt or not mc or mc.group(1) != ok_cod:
        print('IGNORADO: el mail no trae TABLERO/CODIGO válidos'); return
    est_p = os.path.join(work, 'estado.json'); est = json.load(open(est_p))
    excl = est.setdefault('excl', {'equipo': [], 'encargados': []}); t = mt.group(1)
    ids = sorted(set(re.findall(r'\b(m[0-9a-f]{12,20}(?:_\d+)?)\s*\|', txt)))
    nuevos = [i for i in ids if i not in excl[t]]
    excl[t].extend(nuevos)
    json.dump(est, open(est_p, 'w'), ensure_ascii=False, indent=1)
    print(f'{t}: {len(nuevos)} descartados nuevos (de {len(ids)} en el mail) | total en ese tablero: {len(excl[t])}')


def armar(repo, work, pw, corte):
    g = clave(repo, pw)
    est_p = os.path.join(work, 'estado.json')
    est = json.load(open(est_p)) if os.path.exists(est_p) else {'procesados': [], 'cv': {}, 'artifact_url': ''}
    # 1) Excel + datos (armar_excel.py)
    def correr_excel():
        env = dict(os.environ, CVS_BASE=work, CVS_CORTE=corte, CVS_DESCARTADOS=str(est.get('descartados', 0)))
        r = subprocess.run([sys.executable, '-I', os.path.join(AQUI, 'armar_excel.py')], env=env, capture_output=True, text=True)
        print(r.stdout[-800:])
        if r.returncode: print(r.stderr[-3000:]); sys.exit('armar_excel.py falló')
        return json.load(open(os.path.join(work, 'datos_tablero.json')))
    d = correr_excel()
    # 1b) DESCARTAR a los que no cumplen el aviso: se borran sus datos y su CV; solo queda el ID del mail como "ya leído"
    borrar_art = []
    PUNTAJE_MIN = int(est.get('puntaje_min', 79))  # Walter pidió descartar también los de menos de 79 puntos
    fuera = [c for c in d['cands'] if c['lista'] in ('nocumplen', 'cambian') or (c['lista'] in ('atencion', 'cocina') and c['tot'] < PUNTAJE_MIN)]
    # Descartados a mano en cada tablero (los lee la actualización de la base de cada tablero, ver 'descartes').
    # Se borra del todo solo si ya no queda en ningún tablero: equipo (todos) y encargados (los que fueron encargados).
    excl = est.setdefault('excl', {'equipo': [], 'encargados': []})
    ex_eq, ex_en = set(excl.get('equipo', [])), set(excl.get('encargados', []))
    ids_fuera = {c['id'] for c in fuera}
    for c in d['cands']:
        if c['id'] in ids_fuera: continue
        en_enc = es_encargado(c)
        if c['id'] in ex_eq and (not en_enc or c['id'] in ex_en):
            fuera.append(c); ids_fuera.add(c['id'])
            est['descartados_a_mano'] = est.get('descartados_a_mano', 0) + 1
    for c in fuera:
        for mid in c['mids']:
            base_mid = mid.split('_')[0]
            for f in glob.glob(os.path.join(work, 'registros', base_mid + '*.json')):
                os.remove(f)
            est['procesados'] = sorted(set(est['procesados']) | {base_mid})
            shutil.rmtree(os.path.join(work, 'adjuntos', base_mid), ignore_errors=True)
            for ext in ('.txt', '.meta.json'):
                pth = os.path.join(work, 'textos', base_mid + ext)
                if os.path.exists(pth): os.remove(pth)
        info = est['cv'].pop(c['id'], None)
        if info:
            pth = os.path.join(repo, info['gh'])
            if os.path.exists(pth): os.remove(pth)
            borrar_art.append(info['art'])
        est['descartados'] = est.get('descartados', 0) + 1
    if fuera:
        print('descartados ahora:', len(fuera), '| total descartados:', est['descartados'])
        d = correr_excel()
    json.dump(borrar_art, open(os.path.join(work, 'borrar_artifact.json'), 'w'))
    # 2) CVs: los que ya estaban se reutilizan; los nuevos se buscan en WORK/adjuntos
    os.makedirs(os.path.join(repo, 'cv'), exist_ok=True)
    art_dir = os.path.join(work, 'artifact', 'cv'); os.makedirs(art_dir, exist_ok=True)
    nuevos_art, nuevos_cv, sin_cv = [], 0, 0
    for c in d['cands']:
        info = est['cv'].get(c['id'])
        if not info or not os.path.exists(os.path.join(repo, info['gh'])):
            p = buscar_cv(work, c['mids'], os.path.join(art_dir, c['id']))
            if p:
                t = 'pdf' if p.endswith('.pdf') else 'img'
                gh = 'cv/' + c['id'] + '.bin'
                open(os.path.join(repo, gh), 'wb').write(enc(g, open(p, 'rb').read()))
                info = {'gh': gh, 'art': 'cv/' + os.path.basename(p), 't': t}
                est['cv'][c['id']] = info; nuevos_art.append(info['art']); nuevos_cv += 1
            else:
                info = None
        if not info: sin_cv += 1
        c['cv_gh'] = info['gh'] if info else ''; c['cv_art'] = info['art'] if info else ''; c['cvt'] = info['t'] if info else ''
        c.pop('mids', None)
    # 3) tablero GitHub (encriptado)
    dg = json.loads(json.dumps(d))
    for c in dg['cands']: c['cv'] = c.pop('cv_gh'); c.pop('cv_art')
    plantilla = open(os.path.join(AQUI, 'tablero_template.html')).read()
    G = 'https://macproia88.github.io/sel-oct26-presumida/'
    pags_g = PAGINAS[:-1] + [(G, 'Selección: equipo (Ale)'), (G + 'encargados.html', 'Selección: encargados (Sebastián)')]
    for clave_t, dat, pag, sk in (('equipo', 'datos.bin', 'index.html', 'sel_estados'), ('encargados', 'datos_enc.bin', 'encargados.html', 'sel_estados_enc')):
        tg = datos_tablero(dg, clave_t, set(est.get('excl', {}).get(clave_t, [])), corte)
        open(os.path.join(repo, dat), 'wb').write(enc(g, json.dumps(tg, ensure_ascii=False).encode()))
        yo = 'Selección: equipo (Ale)' if clave_t == 'equipo' else 'Selección: encargados (Sebastián)'
        open(os.path.join(repo, pag), 'w').write(sello_y_links(github_index(plantilla, dat, sk, clave_t), [(None if n == yo else h, n) for h, n in pags_g]))
    open(os.path.join(repo, '.nojekyll'), 'w').write('')
    # 4) tablero Claude (datos adentro de la página)
    da = json.loads(json.dumps(d))
    for c in da['cands']: c['cv'] = c.pop('cv_art'); c.pop('cv_gh')
    urls = {'equipo': est.get('artifact_url', ''), 'encargados': est.get('artifact_url_encargados', '')}
    pags = PAGINAS[:-1] + [(urls['equipo'] or None, 'Selección: equipo (Ale)'), (urls['encargados'] or None, 'Selección: encargados (Sebastián)')]
    for clave_t, nombre in (('equipo', 'tablero_seleccion.html'), ('encargados', 'tablero_encargados.html')):
        t = datos_tablero(da, clave_t, set(est['excl'].get(clave_t, [])), corte)
        pg = [(h if n != pags[-2 if clave_t == 'equipo' else -1][1] else None, n) for h, n in pags]
        html = plantilla.replace('/*DATA*/', json.dumps(t, ensure_ascii=False).replace('</', '<\\/'))
        open(os.path.join(work, nombre), 'w').write(sello_y_links(html, pg))
        print(f'tablero {clave_t}: {t["personas"]} candidatos | descartados a mano ahí: {t["desc_tablero"]}')
        if clave_t == 'encargados':
            ya = set(est.get('enc_publicados', []))
            hacen_falta = sorted({c['cv'] for c in t['cands'] if c.get('cv')} - ya)
            json.dump(hacen_falta, open(os.path.join(work, 'nuevos_encargados.json'), 'w'))
            est['enc_publicados'] = sorted(ya | set(hacen_falta))
    json.dump(nuevos_art, open(os.path.join(work, 'nuevos_artifact.json'), 'w'))
    # 5) estado encriptado (registros + mails procesados + mapa de CVs)
    regs = []
    for f in sorted(glob.glob(os.path.join(work, 'registros', '*.json'))):
        rr = json.load(open(f)); rr['_archivo'] = os.path.basename(f)
        mid = rr.get('message_id') or os.path.basename(f)[:-5]
        meta_p = os.path.join(work, 'textos', mid.split('_')[0] + '.meta.json')
        if not rr.get('link') and os.path.exists(meta_p):
            rr['link'] = json.load(open(meta_p)).get('viewUrl', '')
        regs.append(rr)
    est['procesados'] = sorted(set(est['procesados']) | {r_['message_id'].split('_')[0] for r_ in regs if r_.get('message_id')})
    est['corte'] = corte
    json.dump(est, open(est_p, 'w'), ensure_ascii=False, indent=1)
    open(os.path.join(repo, 'estado.bin'), 'wb').write(enc(g, json.dumps({'registros': regs, 'estado': est}, ensure_ascii=False).encode()))
    print(f'OK | descartados total: {est.get("descartados", 0)} | personas: {len(d["cands"])} | CVs nuevos encriptados: {nuevos_cv} | sin archivo de CV: {sin_cv} | datos hasta: {corte}')


if __name__ == '__main__':
    acc = sys.argv[1]
    if acc == 'abrir': abrir(sys.argv[2], sys.argv[3], sys.argv[4])
    elif acc == 'descartes': descartes(sys.argv[2], sys.argv[3], sys.argv[4])
    elif acc == 'descartes_mail': descartes_mail(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
    elif acc == 'armar': armar(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
    else: sys.exit(__doc__)
