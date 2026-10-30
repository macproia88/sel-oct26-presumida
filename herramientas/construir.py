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


def github_index(t):
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
    s = s.replace("      if (!window.claude || !window.claude.use) { $('#sync').textContent = 'Los estados se ven solo en esta pantalla.'; return; }",
                  "      if (!window.claude || !window.claude.use) { return; }")
    s = s.replace("})();\n</script>", "})();\n</script></body></html>")
    assert '/*DATA*/' not in s
    return s


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
    open(os.path.join(repo, 'datos.bin'), 'wb').write(enc(g, json.dumps(dg, ensure_ascii=False).encode()))
    plantilla = open(os.path.join(AQUI, 'tablero_template.html')).read()
    open(os.path.join(repo, 'index.html'), 'w').write(github_index(plantilla))
    open(os.path.join(repo, '.nojekyll'), 'w').write('')
    # 4) tablero Claude (datos adentro de la página)
    da = json.loads(json.dumps(d))
    for c in da['cands']: c['cv'] = c.pop('cv_art'); c.pop('cv_gh')
    open(os.path.join(work, 'tablero_seleccion.html'), 'w').write(
        plantilla.replace('/*DATA*/', json.dumps(da, ensure_ascii=False).replace('</', '<\\/')))
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
    elif acc == 'armar': armar(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
    else: sys.exit(__doc__)
