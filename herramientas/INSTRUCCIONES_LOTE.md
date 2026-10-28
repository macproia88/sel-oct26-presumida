# Instructions for processing a batch of CV emails

You process a list of Gmail message IDs. Each one is a job application to "La Presumida" (fresh pasta shops in Florida/Vicente López, Beccar/San Isidro and Olivos, northern Greater Buenos Aires). Two openings: COOK ("Cocinera/o") and COUNTER SALES ("Vendedor/a", "Empleada para atención").

## Per message
1. Load once: ToolSearch "select:mcp__Gmail__get_message".
2. Call mcp__Gmail__get_message(messageId=ID, messageFormat="RAW"). The result is saved to a file and you get a path back. (If it ever comes back inline instead, it was a tiny mail — call it again with PLAIN_TEXT and treat the body as the CV text; note "sin adjunto".)
3. Run: `python3 -I /home/claude/cvs/scripts/extraer.py "<saved path>" <ID>`
4. Read /home/claude/cvs/textos/<ID>.txt (use the Read tool).
5. If the text is empty/garbage (scanned PDF, photo of a CV; English-only OCR mangles Spanish) open the image/PDF directly with the Read tool to look at it (PDF: use pages "1-2"; images under /home/claude/cvs/adjuntos/<ID>/, including pag-*.png renders). Read it yourself.
6. If the mail only has a Drive/other link and no attachment, use the mail body and set estado_lectura accordingly.
7. Write ONE JSON file per message: /home/claude/cvs/registros/<ID>.json (UTF-8, ensure_ascii=False) with this exact schema:

```json
{
  "message_id": "...",
  "fecha_mail": "YYYY-MM-DD",
  "asunto": "...",
  "archivo": "file name of the CV attachment (or 'sin adjunto')",
  "nombre": "Full name as written",
  "telefono": "as written, or ''",
  "mail": "email from the CV, else the sender address",
  "zona_texto": "where they live, exactly as written (address/neighbourhood/city), or ''",
  "localidad": "best city/neighbourhood, e.g. 'Florida', 'Munro', 'Olivos', 'Boulogne', 'José C. Paz', 'CABA - Belgrano', or ''",
  "partido": "partido/municipio, e.g. 'Vicente López', 'San Isidro', 'San Martín', 'Tigre', 'San Fernando', 'Malvinas Argentinas', 'CABA', or '' if unknown",
  "puesto": "Cocina | Atención/Venta | No dice",
  "puesto_fuente": "asunto | cuerpo del mail | CV | deducido",
  "trabajos": [
    {
      "empresa": "...",
      "cargo": "...",
      "rubro": "short, e.g. 'gastronomía - restaurante', 'panadería', 'comercio - indumentaria', 'supermercado', 'casa de familia', 'call center', 'administrativo'",
      "gastronomia": true,
      "atencion_publico": true,
      "cocina": false,
      "desde": "YYYY-MM or YYYY or null",
      "hasta": "YYYY-MM or YYYY or 'actual' or null",
      "fechas_texto": "dates exactly as written in the CV, or 'sin fechas'"
    }
  ],
  "referencias": "si (names/phones given) | 'a pedido' | no menciona",
  "disponibilidad": "what it says about schedule availability (e.g. 'full time', 'turno tarde', 'inmediata', 'fines de semana'), or 'No dice'",
  "full_time": "si (says full time / amplia / total / any shift) | no (only part-time, one shift, few hours, only weekends, or studies that limit hours) | no dice",
  "estudios_relevantes": "only food-related training (cocina, manipulación de alimentos, pastelería), else ''",
  "resumen": "ONE line in Rioplatense Spanish, max ~20 words, about the work profile only",
  "estado_lectura": "ok | OCR/imagen leído a mano | sin adjunto (solo cuerpo) | ilegible",
  "notas": "anything odd: duplicate candidate, two CVs, non-CV attachment, etc. Else ''"
}
```

## Rules
- Today is 2026-10-28. "Actualidad/presente/hoy/actualmente" → hasta = "actual".
- Dates: month names → number ("Mar 2021" → "2021-03"). Only a year → "2021". Do NOT invent dates. If a job has no dates → desde=null, hasta=null, fechas_texto="sin fechas". If only a duration is given (e.g. "2 años") put that in fechas_texto and leave desde/hasta null.
- gastronomia = restaurants, bars, cafés, bakeries/panaderías, pastas, rotisería, heladería, catering, fast food, kitchen, confitería, delivery of food shops. Supermarket/almacén/fiambrería counts as gastronomia only if it's a food counter (fiambrería, panadería section, carnicería).
- atencion_publico = job facing customers (sales, cashier, waiter, counter, receptionist, call center).
- cocina = actually cooking/prep (cocinero, ayudante de cocina, pastero, pizzero, pastelero, bachero no).
- List ALL jobs in the CV, most recent first. Include informal jobs (e.g. "emprendimiento propio") only as they're written.
- DO NOT record age, date of birth, sex/gender, marital status, children, nationality, DNI, CUIL, health, photo description or religion. Not in any field, not in notas, not in resumen.
- Never guess phone or email; copy exactly.
- puesto: subject containing "cociner" → Cocina; "vended", "atención", "empleada" → Atención/Venta; else look in the mail body / CV objective; else "No dice".
- If one message has several CVs of different people, write one file per person with ID suffix (<ID>_2.json) and note it.
- Never send, reply, label or move any email. Read only.

## At the end
Report back only: how many processed OK, how many with problems (list IDs + problem). Do not paste the records.
