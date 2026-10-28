# Actualización diaria del tablero de selección (La Presumida)

Instrucciones para la sesión programada de Claude. Repositorio: `macproia88/sel-oct26-presumida`.
Clave de los datos: la que figura en el mensaje de la tarea programada.
Carpeta de trabajo: `/home/claude/cvs` (las rutas de `INSTRUCCIONES_LOTE.md` y `extraer.py` asumen esa carpeta).

Hablale a Walter en español rioplatense y llamalo "Salva". Solo lectura en Gmail: nunca mandar, responder, etiquetar ni mover mails.

## 1. Preparar
1. `add_repo` (owner `macproia88`, repo `sel-oct26-presumida`, access `push`) y clonar con `git clone --depth 1 https://github.com/macproia88/sel-oct26-presumida /home/claude/sel-oct26-presumida` (timeout largo, ~10 min).
2. `mkdir -p /home/claude/cvs/scripts /home/claude/cvs/registros /home/claude/cvs/textos /home/claude/cvs/adjuntos && cp /home/claude/sel-oct26-presumida/herramientas/* /home/claude/cvs/scripts/`
3. `cd /home/claude/cvs && python3 -I scripts/construir.py abrir /home/claude/sel-oct26-presumida /home/claude/cvs CLAVE`
   Imprime cuántos registros hay y "datos hasta" (fecha de la última revisión). `estado.json` tiene la lista `procesados` (IDs de mails ya leídos).
4. Anotá la hora de ahora en Argentina (herramienta de hora actual): esa va a ser la nueva fecha de corte, formato `DD/MM/AAAA HH:MM`.

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
- Tiene que terminar con `OK | personas: ...`. La cantidad de personas no puede bajar respecto de la vez anterior.
- Deja: `/home/claude/cvs/Candidatos_Busqueda_Octubre_2026.xlsx`, `/home/claude/cvs/tablero_seleccion.html`, `/home/claude/cvs/nuevos_artifact.json` (CVs nuevos para el tablero de Claude, rutas `cv/<archivo>` que están en `/home/claude/cvs/artifact/cv/`).
- Recalculá el Excel con el script `recalc.py` de la skill xlsx (cargar la skill) y confirmá 0 errores.

## 5. Publicar
1. GitHub: `cd /home/claude/sel-oct26-presumida && git fetch origin main && git add -A && git commit -m "Actualización DD/MM" && git push`
   (si el push falla por acceso, avisale a Walter que hay que reconectar GitHub en claude.ai → Personalización → Conectores → Integración con GitHub).
2. Tablero de Claude (URL en `estado.json` → `artifact_url`): primero `Artifact` action `read` con esa url; después `Artifact` publish con `url`, `file_path: /home/claude/cvs/tablero_seleccion.html` y `files` = los CVs nuevos de `nuevos_artifact.json` mapeados como `{"cv/<archivo>": "artifact/cv/<archivo>"}` (rutas relativas a `/home/claude/cvs`; hasta 250 por publicación). Sin `capabilities` (se mantienen).

## 6. Avisar
- `SendUserFile` con el Excel actualizado.
- `SendUserMessage` corto: cuántos mails nuevos se leyeron, cuántos entraron a cada ranking (Atención / Cocina) y los mejores nuevos con su puntaje, cuántos quedaron en "Revisar a mano", y la nueva fecha de corte. Si no hubo CVs nuevos, decilo igual y actualizá la fecha de corte (pasos 4 y 5 igual).
