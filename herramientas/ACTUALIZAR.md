# Actualización diaria del tablero de selección (La Presumida)

Instrucciones para la sesión programada de Claude. Repositorio: `macproia88/sel-oct26-presumida`.
Clave de los datos: la que figura en el mensaje de la tarea programada.
Carpeta de trabajo: `/home/claude/cvs` (las rutas de `INSTRUCCIONES_LOTE.md` y `extraer.py` asumen esa carpeta).

Hablale en español rioplatense y llamalo "Juan" (desde el 10/11/2026 pidió que le digan Juan, no Walter ni Salva). En Gmail no se manda, responde ni borra nada. La ÚNICA modificación permitida (la pidió Walter el 04/11/2026) es archivar los CVs ya leídos en la carpeta "CV" (paso 5.3).

**REGLA DE JUAN (13/11/2026):** en las automatizaciones no le consultes nada ni le hagas preguntas; resolvé todo directamente y solo informá lo que hiciste.

## 1. Preparar
1. `add_repo` (owner `macproia88`, repo `sel-oct26-presumida`, access `push`) y clonar con `git clone --depth 1 https://github.com/macproia88/sel-oct26-presumida /home/claude/sel-oct26-presumida` (timeout largo, ~10 min).
2. `mkdir -p /home/claude/cvs/scripts /home/claude/cvs/registros /home/claude/cvs/textos /home/claude/cvs/adjuntos && cp /home/claude/sel-oct26-presumida/herramientas/* /home/claude/cvs/scripts/`
3. `cd /home/claude/cvs && python3 -I scripts/construir.py abrir /home/claude/sel-oct26-presumida /home/claude/cvs CLAVE`
   Imprime cuántos registros hay y "datos hasta" (fecha de la última revisión). `estado.json` tiene la lista `procesados` (IDs de mails ya leídos).
4. Anotá la hora de ahora en Argentina (herramienta de hora actual): esa va a ser la nueva fecha de corte, formato `DD/MM/AAAA HH:MM`.

## 1b. Leer los descartados a mano de los dos tableros (pedido del 10/11/2026)
Hay DOS tableros en Claude, cada uno con su base de estados (colección `estados`, un documento por candidato con `estado` y `nota`):
- Equipo (Ale): `estado.json` → `artifact_url` (https://claude.ai/artifact/GMFa4HpdDihqj6XSQb71aC)
- Encargados (Sebastián): `estado.json` → `artifact_url_encargados` (https://claude.ai/artifact/UWARpAyUbD5pQWnRtr5avz)
Para cada uno: `ArtifactData` action `list`, collection `estados`, `query: {"limit": 1000}` (seguí `next_cursor` si hay), `out_dir: /home/claude/cvs/db_<equipo|encargados>` (borrá esa carpeta antes). Después:
`cd /home/claude/cvs && python3 -I scripts/construir.py descartes /home/claude/cvs equipo /home/claude/cvs/db_equipo` y lo mismo con `encargados /home/claude/cvs/db_encargados`.
Eso anota en `estado.json` → `excl` los IDs marcados "Descartado" en cada tablero. Al armar (paso 4) salen de ESE tablero; si alguien quedó descartado en los dos (o en Equipo y no fue encargado), se borran del todo sus datos y su CV. No borres ni cambies nada en las bases de los tableros.

### 1c. Descartados que mandan por mail desde las páginas de GitHub (Ale y Sebastián usan estas)
Las páginas con contraseña (`https://macproia88.github.io/sel-oct26-presumida/` = Equipo/Ale y `.../encargados.html` = Encargados/Sebastián) guardan los estados en el celular de cada uno; cuando marcan "Descartado" aparece la barra "Mandar descartados", que arma un mail a lapresumidapastas@gmail.com con asunto `DESCARTES SELECCION equipo` o `DESCARTES SELECCION encargados`, un `CODIGO` (sale de la contraseña) y una línea `<id> | <nombre>` por candidato.
- Buscá con `mcp__Gmail__search_threads`: `subject:"DESCARTES SELECCION" newer_than:4d` (busca también en Enviados). Saltá los message IDs que ya están en `estado.json` → `descartes_mails`.
- Para cada mensaje: `mcp__Gmail__get_message` con `PLAIN_TEXT`, guardá el cuerpo en `/home/claude/cvs/desc_<ID>.txt` y corré
  `cd /home/claude/cvs && python3 -I scripts/construir.py descartes_mail /home/claude/sel-oct26-presumida /home/claude/cvs CLAVE /home/claude/cvs/desc_<ID>.txt`
  (imprime cuántos sumó; si dice IGNORADO, el código no coincide: no lo cuentes y avisalo). Después agregá ese message ID a `descartes_mails` en `estado.json`.
- No respondas ni muevas esos mails.

## 2. Buscar CVs nuevos en Gmail
- Herramienta `mcp__Gmail__search_threads` (cargar con ToolSearch). Consultas (cambiar FECHA por el día anterior al último corte, formato AAAA/MM/DD):
  - `to:lapresumidapastas@gmail.com has:attachment after:FECHA`
  - `to:lapresumidapastas@gmail.com after:FECHA (cv OR curriculum OR currículum OR búsqueda OR busqueda OR postulación OR postulacion OR vendedora OR vendedor OR cocinera OR cocinero)`
  - Si alguna consulta devuelve 50 resultados, partila por día (`after:`/`before:`) para no perder mails.
- Descartá: IDs que ya están en `procesados`; mails que no son postulaciones (proveedores, facturas, contadoratantucci@gmail.com, deliverylapresumida@gmail.com, cuentas del negocio, apps, bancos, newsletters, mensajes enviados por lapresumidapastas). Ante la duda, procesalo.

## 3. Leer cada CV nuevo
- Seguí `/home/claude/cvs/scripts/INSTRUCCIONES_LOTE.md` para cada mail nuevo (RAW → `extraer.py` → leer → `registros/<ID>.json`).
- Si son más de 30, repartilos en lotes de ~30 entre subagentes (Agent tool), cada uno con su lista de IDs en un archivo y diciéndoles que usen su propia subcarpeta para archivos auxiliares.
- Verificá al final que haya un `registros/<ID>.json` por cada mail nuevo.

## 4. Reconstruir
`cd /home/claude/cvs && python3 -I scripts/construir.py armar /home/claude/sel-oct26-presumida /home/claude/cvs CLAVE "DD/MM/AAAA HH:MM"`
- Tiene que terminar con `OK | descartados total: ... | personas: ...`. Los que no cumplen el aviso (zona, full time, experiencia) o cambian de trabajo seguido o sacan menos de 79 puntos se DESCARTAN solos: se borran sus datos y su CV y solo queda el ID del mail en `procesados` (para no volver a leerlo). Descartados + personas no puede bajar respecto de la vez anterior.
- Los CVs de descartados que estaban publicados en el tablero de Claude están en `/home/claude/cvs/borrar_artifact.json`: al publicar, pasalos en `files` con valor null.
- Imprime también `tablero equipo: N candidatos | descartados a mano ahí: X` y lo mismo para `encargados`.
- Deja: `/home/claude/cvs/Candidatos_Busqueda_Octubre_2026.xlsx`, `/home/claude/cvs/tablero_seleccion.html` (Equipo), `/home/claude/cvs/tablero_encargados.html` (Encargados), `/home/claude/cvs/nuevos_encargados.json` (CVs que le faltan al tablero de Encargados), `/home/claude/cvs/nuevos_artifact.json` (CVs nuevos para el tablero de Claude, rutas `cv/<archivo>` que están en `/home/claude/cvs/artifact/cv/`).
- Recalculá el Excel con el script `recalc.py` de la skill xlsx (cargar la skill) y confirmá 0 errores.

## 5. Publicar
1. GitHub: `cd /home/claude/sel-oct26-presumida && git fetch origin main && git add -A && git commit -m "Actualización DD/MM" && git push`
   (si el push falla por acceso, avisale a Walter que hay que reconectar GitHub en claude.ai → Personalización → Conectores → Integración con GitHub).
2. Tablero de Claude (URL en `estado.json` → `artifact_url`): primero `Artifact` action `read` con esa url; después `Artifact` publish con `url`, `file_path: /home/claude/cvs/tablero_seleccion.html` y `files` = los CVs nuevos de `nuevos_artifact.json` mapeados como `{"cv/<archivo>": "artifact/cv/<archivo>"}` (rutas relativas a `/home/claude/cvs`; hasta 250 por publicación). Sin `capabilities` (se mantienen). Si el publish se rechaza por una versión más nueva, hacé `Artifact` action `read` de la url y volvé a publicar con `url`.
   Tablero de Encargados (`artifact_url_encargados`): igual, con `file_path: /home/claude/cvs/tablero_encargados.html` y `files` = cada ruta de `nuevos_encargados.json` copiada del tablero de Equipo: `{"cv/<archivo>": {"artifact": "<artifact_url>", "path": "cv/<archivo>"}}` (publicá primero el de Equipo para que los CVs nuevos ya estén ahí). Los CVs de `borrar_artifact.json` también van con null en este tablero si estaban publicados ahí.

3. Archivar en Gmail los CVs leídos (pedido de Walter): para cada mail nuevo de esta corrida que sea una postulación (tenga o no `registros/<ID>.json` después de armar, o haya quedado descartado), llamá `mcp__Gmail__update_message_labels` con `messageId=<ID>`, `addLabelIds=["Label_1028625076126494661"]` (carpeta "CV") y `removeLabelIds=["INBOX"]`. NO toques los mails que no son CVs (facturas, Mercado Pago, avisos), ni los que quedaron sin leer porque el adjunto no se pudo abrir (esos quedan en Recibidos para que Walter los vea). Si son muchos, repartilos entre subagentes. Si una llamada falla, reintentá una vez y mencioná en el aviso cuántos no se pudieron archivar.

## 6. Avisar
- `SendUserFile` con el Excel actualizado.
- `SendUserMessage` corto: cuántos mails nuevos se leyeron, cuántos se archivaron en la carpeta CV, cuántos entraron a cada ranking (Atención / Cocina) y los mejores nuevos con su puntaje, cuántos encargados nuevos entraron al tablero de Sebastián, cuántos se sacaron por estar marcados "Descartado" en cada tablero, cuántos quedaron en "Revisar a mano", y la nueva fecha de corte. Si no hubo CVs nuevos, decilo igual y actualizá la fecha de corte (pasos 4 y 5 igual).
