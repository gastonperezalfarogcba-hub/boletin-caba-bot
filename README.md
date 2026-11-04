# Bot Boletín Oficial CABA

Revisa el Boletín Oficial de la Ciudad 4 veces por día hábil. Cuando encuentra
un boletín nuevo:

1. Manda al canal de Telegram un **resumen**: N° y fecha, total de
   publicaciones, cantidades por poder/sección y por tipo de norma, los 10
   organismos con más publicaciones, y las leyes y decretos con link al PDF.
2. Suma todas las normas a la **base histórica** (`data/normas_AAAA.csv`, un
   archivo por año, queda guardada en el repo).
3. Regenera y manda el **Excel histórico** con estas pestañas:
   - `Normas`: todas las normas, con filtros y link al PDF
   - `Tabla_Resumen`: mes × tipo de norma (como la tabla de Cultura, pero de todo)
   - `Por_Organismo`: organismo × tipo de norma
   - `Por_Poder`: poder/sección × tipo de norma
   - `Por_Dia`: cantidad de normas por boletín

Si no hay boletín nuevo (ya se procesó, todavía no salió, fin de semana o
feriado), la corrida termina sin enviar nada.

Si algún día fallaran las 4 corridas, la siguiente corrida que funcione
recupera también los boletines que faltaron.

## Puesta en marcha (una sola vez)

### 1. Bot y canal de Telegram
1. En Telegram, hablá con **@BotFather** → `/newbot` → guardá el **token**.
2. Creá un canal (puede ser privado) y agregá el bot como **administrador**
   con permiso para publicar.
3. ID del canal:
   - canal público: `@nombre_del_canal`
   - canal privado: reenviá un mensaje del canal a **@userinfobot** (o
     @RawDataBot) y copiá el id que empieza con `-100…`

### 2. Repositorio en GitHub
1. Creá un repo nuevo (privado si querés) y subí estos archivos tal cual,
   incluida la carpeta `.github/workflows/`.
2. **Settings → Secrets and variables → Actions → New repository secret**:
   - `TELEGRAM_TOKEN` = token del bot
   - `TELEGRAM_CHAT_ID` = id del canal
3. **Settings → Actions → General → Workflow permissions** → marcar
   *Read and write permissions* (para que pueda guardar la base histórica).
4. Pestaña **Actions → Boletín Oficial CABA → Run workflow** para probarlo.

### 3. (Opcional) Cargar el histórico
**Actions → Boletín Oficial CABA → Run workflow** y en el campo *desde*
poné el N° de boletín inicial:
- `7275` = primer boletín de 2026 (02/01/2026)

Procesa todos los boletines desde ese número hasta hoy (aprox. 1 segundo por
boletín) y manda un solo mensaje de "Carga histórica" con el Excel completo.
Después sigue solo con la corrida diaria.

## Horarios
Lunes a viernes 08:30, 10:30, 13:00 y 17:00 (hora Argentina). Se cambian en
`.github/workflows/boletin.yml` (los cron van en UTC = hora argentina + 3).
GitHub puede demorar los cron programados entre 5 y 30 minutos.

## Correrlo en la PC
```
pip install -r requirements.txt
set TELEGRAM_TOKEN=...        (en Mac/Linux: export TELEGRAM_TOKEN=...)
set TELEGRAM_CHAT_ID=...
python boletin_bot.py                 # corrida normal
python boletin_bot.py --desde 7275    # carga histórica desde ese boletín
python boletin_bot.py --no-telegram   # solo actualiza la base y el Excel
python boletin_bot.py --forzar        # reprocesa el último boletín
```

## Notas técnicas
- Fuente: `https://api-restboletinoficial.buenosaires.gob.ar/obtenerBoletin/{N}/true`
  (`0` = último publicado). Es la misma API que usa la web del Boletín.
- Si se pide un N° que no existe, la API devuelve el último boletín en vez de
  un error; el bot lo detecta y no duplica datos.
- Los nombres de organismo se normalizan (la API trae algunos con un guion al
  final, como "Ministerio de Salud-").
- En "Leyes y decretos" del resumen solo entran normas del Poder Legislativo y
  Ejecutivo (en Comunicados y Avisos hay tipos como "Ley 6009 - …" que no son
  leyes nuevas).
- El Excel se genera en cada corrida y se manda por Telegram (también queda
  como artefacto en la pestaña Actions). Un año completo da ~6 MB.
