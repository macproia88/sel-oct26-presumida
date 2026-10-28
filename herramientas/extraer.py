"""Uso: python3 -I extraer.py <json_guardado_por_get_message_RAW> <message_id>
Saca los adjuntos del mail (pdf/docx/doc/imagenes/odt), los guarda en adjuntos/<message_id>/
y deja el texto en textos/<message_id>.txt. Imprime un resumen."""
import sys, json, base64, email, os, subprocess, re
from email import policy
src, mid = sys.argv[1], sys.argv[2]
base = '/home/claude/cvs'
outdir = os.path.join(base, 'adjuntos', mid); os.makedirs(outdir, exist_ok=True)
txt_out = os.path.join(base, 'textos', mid + '.txt')
d = json.load(open(src))
raw = d['raw']; raw += '=' * (-len(raw) % 4)
msg = email.message_from_bytes(base64.urlsafe_b64decode(raw), policy=policy.default)
print('SUBJECT:', msg['Subject']); print('FROM:', msg['From']); print('DATE:', msg['Date'])
body = ''
try:
    b = msg.get_body(preferencelist=('plain','html'))
    if b: body = b.get_content()
    body = re.sub(r'<[^>]+>', ' ', body)
except Exception as e:
    body = f'[error cuerpo: {e}]'
partes = []
for part in msg.walk():
    fn = part.get_filename()
    if not fn: continue
    fn = os.path.basename(fn).replace('/', '_')
    data = part.get_payload(decode=True)
    if not data: continue
    p = os.path.join(outdir, fn); open(p, 'wb').write(data)
    ext = fn.lower().rsplit('.', 1)[-1] if '.' in fn else ''
    texto, metodo = '', ''
    try:
        if ext == 'pdf':
            texto = subprocess.run(['pdftotext', '-layout', p, '-'], capture_output=True, text=True, timeout=60).stdout
            metodo = 'pdftotext'
            if len(texto.strip()) < 150:
                pref = os.path.join(outdir, 'pag')
                subprocess.run(['pdftoppm', '-r', '200', '-png', '-l', '3', p, pref], timeout=120)
                ocr = ''
                for img in sorted(f for f in os.listdir(outdir) if f.startswith('pag') and f.endswith('.png')):
                    ocr += subprocess.run(['tesseract', os.path.join(outdir, img), '-', '-l', 'eng'], capture_output=True, text=True, timeout=120).stdout
                texto, metodo = ocr, 'OCR (PDF escaneado)'
        elif ext in ('docx', 'doc', 'odt', 'rtf'):
            subprocess.run(['libreoffice', '--headless', '--convert-to', 'txt:Text', '--outdir', outdir, p], capture_output=True, timeout=120)
            tp = os.path.splitext(p)[0] + '.txt'
            texto = open(tp, errors='ignore').read() if os.path.exists(tp) else ''
            metodo = 'libreoffice'
        elif ext in ('jpg', 'jpeg', 'png', 'webp', 'heic'):
            texto = subprocess.run(['tesseract', p, '-', '-l', 'eng'], capture_output=True, text=True, timeout=120).stdout
            metodo = 'OCR (imagen)'
        else:
            metodo = 'no soportado'
    except Exception as e:
        metodo += f' ERROR {e}'
    partes.append((fn, metodo, texto))
with open(txt_out, 'w') as f:
    f.write(f"ASUNTO: {msg['Subject']}\nDE: {msg['From']}\nFECHA: {msg['Date']}\n\n=== CUERPO DEL MAIL ===\n{body.strip()[:3000]}\n")
    for fn, metodo, texto in partes:
        f.write(f"\n=== ADJUNTO: {fn} | método: {metodo} | {len(texto.strip())} caracteres ===\n{texto}\n")
for fn, metodo, texto in partes:
    print(f'ADJUNTO: {fn} | {metodo} | {len(texto.strip())} chars | ruta: {os.path.join(outdir, fn)}')
if not partes: print('SIN ADJUNTOS')
json.dump({'message_id': mid, 'viewUrl': d.get('viewUrl',''), 'archivos': [fn for fn,_,_ in partes]}, open(os.path.join(base,'textos',mid+'.meta.json'),'w'))
print('TEXTO EN:', txt_out)
