"""
Bot Boletín Oficial CABA
------------------------
1. Consulta la API del Boletín Oficial de la Ciudad (último boletín publicado).
2. Procesa todos los boletines nuevos desde el último que ya tenía
   (normalmente uno; si algún día falló, recupera los que faltaron):
   - agrega sus normas a la base histórica  data/normas_AAAA.csv
   - manda al canal un resumen + un Excel liviano con las normas del día
   - el Excel histórico completo se manda solo a pedido (--historico)
3. Si no hay nada nuevo (ya procesado o todavía no salió) termina sin enviar nada.

Pensado para correr 4 veces por día: la primera corrida que encuentra el
boletín nuevo lo procesa y las siguientes no hacen nada.

Variables de entorno:
  TELEGRAM_TOKEN    token del bot (BotFather)
  TELEGRAM_CHAT_ID  id del canal (@nombre_canal o -100xxxxxxxxxx)
  DESDE             (opcional) igual que --desde

Uso:
  python boletin_bot.py                 # corrida normal
  python boletin_bot.py --desde 7275    # carga histórica desde ese N° de boletín
  python boletin_bot.py --no-telegram   # no enviar nada, solo actualizar datos
  python boletin_bot.py --forzar        # reprocesar el último boletín
  python boletin_bot.py --historico     # mandar también el Excel histórico
  python boletin_bot.py --json x.json   # procesar un JSON guardado (pruebas)
"""

import argparse
import html
import json
import os
import re
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

API = "https://api-restboletinoficial.buenosaires.gob.ar/obtenerBoletin/{n}/true"
# n = 0 -> último boletín publicado; n = 7488 -> boletín N° 7488

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
ESTADO = DATA_DIR / "estado.json"
SALIDA = BASE / "salida"
EXCEL = SALIDA / "boletin_caba_historico.xlsx"

MAX_RESUMENES = 5   # si hay más boletines pendientes, es carga histórica:
                    # se manda un solo mensaje en vez de uno por boletín

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

COLUMNAS = ["Fecha", "Año", "Mes", "N° Boletín", "Poder", "Tipo de Norma",
            "Organismo", "Nombre", "Sumario", "ID Norma", "URL", "Tiene Anexos"]

# Paleta GCBA
NAVY = "1B2A4A"
AMARILLO = "FFD200"
GRIS_FILA = "F2F5FA"


# ───────────────────────────── Descarga ──────────────────────────────

def descargar(n: int = 0) -> dict:
    ultimo_error = None
    for intento in range(3):
        try:
            r = requests.get(API.format(n=n), timeout=60,
                             headers={"User-Agent": "boletin-caba-bot/1.0"})
            r.raise_for_status()
            return r.json()
        except Exception as ex:
            ultimo_error = ex
            time.sleep(5 * (intento + 1))
    raise RuntimeError(f"No se pudo descargar el boletín {n}: {ultimo_error}")


def parsear_fecha(txt: str) -> datetime:
    txt = str(txt).strip()[:10]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(txt, fmt)
        except ValueError:
            pass
    raise ValueError(f"Fecha no reconocida: {txt!r}")


def limpiar(txt: str) -> str:
    """Saca espacios dobles y guiones sueltos al final
    (la API trae, por ejemplo, 'Ministerio de Salud-')."""
    txt = re.sub(r"\s+", " ", str(txt or "")).strip()
    return re.sub(r"[\s\-–]+$", "", txt)


def aplanar(data: dict) -> tuple[dict, list[dict]]:
    """Convierte normas.normas[poder][tipo][organismo] -> lista de filas."""
    boletin = data.get("boletin", {}) or {}
    fecha = parsear_fecha(boletin.get("fecha_publicacion", ""))
    numero = int(boletin.get("numero"))
    mes = f"{fecha.month:02d} - {MESES[fecha.month - 1]}"

    filas = []
    arbol = (data.get("normas") or {}).get("normas") or {}
    for poder, tipos in arbol.items():
        if not isinstance(tipos, dict):
            continue
        for tipo, organismos in tipos.items():
            if not isinstance(organismos, dict):
                continue
            for organismo, normas in organismos.items():
                if not isinstance(normas, list):
                    continue
                for n in normas:
                    filas.append({
                        "Fecha": fecha.strftime("%Y-%m-%d"),
                        "Año": fecha.year,
                        "Mes": mes,
                        "N° Boletín": numero,
                        "Poder": limpiar(poder),
                        "Tipo de Norma": limpiar(tipo),
                        "Organismo": limpiar(organismo),
                        "Nombre": limpiar(n.get("nombre")),
                        "Sumario": limpiar(n.get("sumario")),
                        "ID Norma": n.get("id_norma", ""),
                        "URL": n.get("url_norma", "") or "",
                        "Tiene Anexos": "Sí" if n.get("anexos") else "No",
                    })
    info = {"numero": numero, "fecha": fecha, "total": len(filas)}
    return info, filas


# ───────────────────────────── Estado / base ─────────────────────────

def leer_estado() -> dict:
    if ESTADO.exists():
        return json.loads(ESTADO.read_text(encoding="utf-8"))
    return {}


def guardar_estado(estado: dict) -> None:
    ESTADO.parent.mkdir(parents=True, exist_ok=True)
    ESTADO.write_text(json.dumps(estado, ensure_ascii=False, indent=2),
                      encoding="utf-8")


def actualizar_base(filas: list[dict], numeros: set[int]) -> pd.DataFrame:
    """Base histórica en data/normas_AAAA.csv (un archivo por año, para no
    superar el límite de tamaño de archivo de GitHub)."""
    nuevo = pd.DataFrame(filas, columns=COLUMNAS).astype(str)
    previos = [pd.read_csv(p, dtype=str, keep_default_na=False)
               for p in sorted(DATA_DIR.glob("normas_*.csv"))]
    if previos:
        base = pd.concat(previos, ignore_index=True)
        # si algún boletín se reprocesa, se reemplaza (no se duplica)
        base = base[~base["N° Boletín"].isin({str(n) for n in numeros})]
        base = pd.concat([base, nuevo], ignore_index=True)
    else:
        base = nuevo
    base = base.sort_values(["Fecha", "Poder", "Tipo de Norma", "Organismo"],
                            kind="stable").reset_index(drop=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    for anio, parte in base.groupby("Año"):
        parte.to_csv(DATA_DIR / f"normas_{anio}.csv", index=False,
                     encoding="utf-8-sig")
    return base


# ───────────────────────────── Excel ─────────────────────────────────

def _borde():
    s = Side(style="thin", color="C9D1E0")
    return Border(left=s, right=s, top=s, bottom=s)


def _encabezado(ws, fila, ncols):
    for c in range(1, ncols + 1):
        cel = ws.cell(row=fila, column=c)
        cel.font = Font(bold=True, color="FFFFFF", name="Montserrat")
        cel.fill = PatternFill("solid", fgColor=NAVY)
        cel.alignment = Alignment(horizontal="center", vertical="center",
                                  wrap_text=True)
        cel.border = _borde()


def _titulo(ws, texto):
    cel = ws.cell(row=1, column=1, value=texto)
    cel.font = Font(bold=True, size=14, color=NAVY, name="Montserrat")
    ws.cell(row=2, column=1,
            value=f"Actualizado: {datetime.now():%d/%m/%Y %H:%M}").font = Font(
        italic=True, size=9, color="666666")


def _tabla_cruzada(ws, df, filas, columnas, titulo, etiqueta_filas):
    piv = pd.crosstab(df[filas], df[columnas], margins=True,
                      margins_name="Total general")
    cols = list(piv.columns)
    ws.cell(row=4, column=1, value=f"Cuenta de {columnas}")
    ws.cell(row=4, column=2, value=columnas)
    hdr = [etiqueta_filas] + cols
    for j, h in enumerate(hdr, 1):
        ws.cell(row=5, column=j, value=h)
    _encabezado(ws, 5, len(hdr))
    for i, (idx, fila) in enumerate(piv.iterrows(), start=6):
        total = idx == "Total general"
        ws.cell(row=i, column=1, value=idx)
        for j, v in enumerate(fila, start=2):
            ws.cell(row=i, column=j, value=int(v) if v else None)
        for j in range(1, len(hdr) + 1):
            cel = ws.cell(row=i, column=j)
            cel.border = _borde()
            if j > 1:
                cel.alignment = Alignment(horizontal="center")
            if total:
                cel.font = Font(bold=True, color=NAVY)
                cel.fill = PatternFill("solid", fgColor=AMARILLO)
            elif i % 2 == 0:
                cel.fill = PatternFill("solid", fgColor=GRIS_FILA)
    for i in range(6, 5 + len(piv)):
        ws.cell(row=i, column=len(hdr)).font = Font(bold=True, color=NAVY)
    ws.column_dimensions["A"].width = max(
        16, min(60, max(len(str(x)) for x in piv.index) + 2))
    for j in range(2, len(hdr) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14
    ws.row_dimensions[5].height = 75
    ws.freeze_panes = "B6"
    _titulo(ws, titulo)


def generar_excel(base: pd.DataFrame, destino: Path = EXCEL,
                  diario: bool = False) -> Path:
    """Arma el Excel. diario=True: versión liviana de un solo boletín
    (sin las pestañas por mes y por día, que no aportan para un día)."""
    SALIDA.mkdir(parents=True, exist_ok=True)
    df = base.copy()
    df["Período"] = df["Año"].astype(str) + " · " + df["Mes"]
    wb = Workbook()

    # 1) Base completa
    ws = wb.active
    ws.title = "Normas"
    ws.append(COLUMNAS)
    _encabezado(ws, 1, len(COLUMNAS))
    enteros = {"Año", "N° Boletín", "ID Norma"}
    idx_ent = [COLUMNAS.index(c) for c in enteros]
    for fila in df[COLUMNAS].itertuples(index=False):
        fila = list(fila)
        for k in idx_ent:
            if str(fila[k]).isdigit():
                fila[k] = int(fila[k])
        ws.append(fila)
    url_col = COLUMNAS.index("URL") + 1
    for r in range(2, ws.max_row + 1):
        cel = ws.cell(row=r, column=url_col)
        if cel.value:
            cel.hyperlink = cel.value
            cel.font = Font(color="1F5FBF", underline="single")
    anchos = [11, 7, 14, 10, 18, 22, 38, 30, 70, 11, 40, 10]
    for j, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # 2) Mes x Tipo de Norma (como la tabla de Cultura, pero de todo)
    if not diario:
        _tabla_cruzada(wb.create_sheet("Tabla_Resumen"), df, "Período",
                       "Tipo de Norma", "Normas por mes y tipo de norma",
                       "Mes")
    # 3) Organismo x Tipo de Norma
    _tabla_cruzada(wb.create_sheet("Por_Organismo"), df, "Organismo",
                   "Tipo de Norma", "Normas por organismo y tipo de norma",
                   "Organismo")
    # 4) Poder x Tipo
    _tabla_cruzada(wb.create_sheet("Por_Poder"), df, "Poder",
                   "Tipo de Norma", "Normas por poder y tipo de norma",
                   "Poder / Sección")

    if diario:
        wb.save(destino)
        return destino

    # 5) Serie diaria
    ws = wb.create_sheet("Por_Dia")
    dia = (df.groupby(["Fecha", "N° Boletín"]).size()
             .reset_index(name="Normas publicadas")
             .sort_values("Fecha", ascending=False))
    hdr = ["Fecha", "N° Boletín", "Normas publicadas"]
    for j, h in enumerate(hdr, 1):
        ws.cell(row=5, column=j, value=h)
    _encabezado(ws, 5, 3)
    for i, f in enumerate(dia.itertuples(index=False), start=6):
        ws.cell(row=i, column=1, value=f[0])
        ws.cell(row=i, column=2, value=int(f[1]))
        ws.cell(row=i, column=3, value=int(f[2]))
        for j in range(1, 4):
            ws.cell(row=i, column=j).border = _borde()
            ws.cell(row=i, column=j).alignment = Alignment(horizontal="center")
    for c, w in zip("ABC", (14, 12, 18)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A6"
    _titulo(ws, "Normas publicadas por boletín")

    wb.save(destino)
    return destino


def cargar_base() -> pd.DataFrame:
    partes = [pd.read_csv(p, dtype=str, keep_default_na=False)
              for p in sorted(DATA_DIR.glob("normas_*.csv"))]
    if not partes:
        return pd.DataFrame(columns=COLUMNAS)
    return pd.concat(partes, ignore_index=True)


# ───────────────────────────── Telegram ──────────────────────────────

def es_destacada(f: dict) -> bool:
    """Leyes y decretos del Poder Legislativo / Ejecutivo.
    (Ojo: en 'Comunicados y Avisos' hay tipos como 'Ley 6009 - ...'
    que no son leyes nuevas, por eso se filtra también por poder.)"""
    tipo = f["Tipo de Norma"].lower()
    return (f["Poder"] in ("Poder Legislativo", "Poder Ejecutivo")
            and (tipo == "ley" or tipo.startswith("decreto")))


def _recortar(lineas: list[str], limite: int = 3950) -> str:
    # Límite de Telegram: 4096. Se cortan líneas enteras (nunca HTML a medias).
    final, largo = [], 0
    for ln in lineas:
        if largo + len(ln) + 1 > limite:
            final.append("<i>…mensaje recortado, el detalle está en el Excel</i>")
            break
        final.append(ln)
        largo += len(ln) + 1
    return "\n".join(final)


def armar_resumen(info: dict, filas: list[dict]) -> str:
    e = html.escape
    fecha = info["fecha"]
    dias = ["lunes", "martes", "miércoles", "jueves", "viernes",
            "sábado", "domingo"]
    lineas = [
        f"📰 <b>Boletín Oficial CABA N° {info['numero']}</b>",
        f"🗓 {dias[fecha.weekday()].capitalize()} {fecha:%d/%m/%Y}",
        f"📄 <b>{info['total']}</b> publicaciones",
        "",
        "🏛 <b>Por poder / sección</b>",
    ]
    for p, n in Counter(f["Poder"] for f in filas).most_common():
        lineas.append(f"• {e(p)}: {n}")
    lineas += ["", "📑 <b>Por tipo de norma</b>"]
    for t, n in Counter(f["Tipo de Norma"] for f in filas).most_common():
        lineas.append(f"• {e(t)}: {n}")
    por_org = Counter(f["Organismo"] for f in filas)
    lineas += ["", "🏢 <b>Organismos con más publicaciones</b>"]
    for o, n in por_org.most_common(10):
        lineas.append(f"• {e(o)}: {n}")
    if len(por_org) > 10:
        lineas.append(f"<i>…y {len(por_org) - 10} organismos más</i>")

    destacadas = [f for f in filas if es_destacada(f)]
    if destacadas:
        lineas += ["", "⭐ <b>Leyes y decretos</b>"]
        for f in destacadas[:12]:
            txt = f"{f['Nombre']} — {f['Sumario']}"
            txt = txt if len(txt) <= 160 else txt[:157] + "…"
            if f["URL"]:
                lineas.append(f'• <a href="{e(f["URL"], quote=True)}">'
                              f"{e(txt)}</a>")
            else:
                lineas.append(f"• {e(txt)}")
        if len(destacadas) > 12:
            lineas.append(f"<i>…y {len(destacadas) - 12} más en el Excel</i>")
    return _recortar(lineas)


def armar_carga_historica(nuevos: list, base: pd.DataFrame) -> str:
    f0, f1 = nuevos[0][0]["fecha"], nuevos[-1][0]["fecha"]
    n = f"{sum(i['total'] for i, _ in nuevos):,}".replace(",", ".")
    return (f"📚 <b>Carga histórica del Boletín Oficial CABA</b>\n"
            f"Se incorporaron <b>{len(nuevos)}</b> boletines "
            f"(N° {nuevos[0][0]['numero']} a {nuevos[-1][0]['numero']}, "
            f"{f0:%d/%m/%Y} al {f1:%d/%m/%Y}) con <b>{n}</b> normas.")


def pie_historico(base: pd.DataFrame) -> str:
    return (f"📊 Histórico acumulado: {len(base):,} normas en "
            f"{base['N° Boletín'].nunique()} boletines").replace(",", ".")


def telegram(metodo: str, **kwargs):
    token = os.environ["TELEGRAM_TOKEN"]
    r = requests.post(f"https://api.telegram.org/bot{token}/{metodo}",
                      timeout=180, **kwargs)
    if not r.ok:
        raise RuntimeError(f"Telegram {metodo}: {r.status_code} {r.text}")
    return r.json()


def enviar_mensaje(texto: str) -> None:
    telegram("sendMessage", data={
        "chat_id": os.environ["TELEGRAM_CHAT_ID"], "text": texto,
        "parse_mode": "HTML", "disable_web_page_preview": "true"})


def enviar_excel(excel: Path, nombre: str, pie: str) -> None:
    with open(excel, "rb") as fh:
        telegram("sendDocument",
                 data={"chat_id": os.environ["TELEGRAM_CHAT_ID"],
                       "caption": pie},
                 files={"document": (nombre, fh)})


# ───────────────────────────── Main ──────────────────────────────────

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--desde", type=int,
                    default=int(os.environ["DESDE"]) if os.environ.get("DESDE")
                    else None,
                    help="N° de boletín desde el cual cargar (carga histórica)")
    ap.add_argument("--historico", action="store_true",
                    default=os.environ.get("HISTORICO", "").lower() == "true",
                    help="mandar también el Excel histórico completo")
    ap.add_argument("--json", nargs="+",
                    help="procesar JSON guardados en vez de la API (pruebas)")
    ap.add_argument("--no-telegram", action="store_true")
    ap.add_argument("--forzar", action="store_true")
    args = ap.parse_args()

    estado = leer_estado()
    ultimo_proc = estado.get("ultimo_numero")

    # 1) ¿Qué boletines hay que procesar?
    if args.json:
        disponibles = [json.loads(Path(p).read_text(encoding="utf-8"))
                       for p in args.json]
        datos = {int(d["boletin"]["numero"]): d for d in disponibles}
        ultimo_pub = max(datos)
    else:
        try:
            ultimo = descargar(0)
        except Exception as ex:  # API caída: se reintenta en la próxima corrida
            print(f"No se pudo consultar la API: {ex}")
            ultimo, ultimo_pub, datos = None, 0, {}
        else:
            ultimo_pub = int(ultimo["boletin"]["numero"])
            datos = {ultimo_pub: ultimo}

    if args.desde:
        inicio = args.desde
    elif args.forzar:
        inicio = ultimo_pub
    elif ultimo_proc:
        inicio = int(ultimo_proc) + 1
    else:
        inicio = ultimo_pub        # primera vez: arranca por el de hoy

    pendientes = [n for n in range(inicio, ultimo_pub + 1)
                  if ultimo_pub and (not args.json or n in datos)]
    if pendientes:
        print(f"Boletines a procesar: {pendientes[0]} a {pendientes[-1]} "
              f"({len(pendientes)})")
    elif ultimo_pub:
        print(f"Sin novedades: el último publicado es el N° {ultimo_pub} "
              "y ya está procesado.")

    # 2) Descargar y aplanar
    nuevos = []
    for n in pendientes:
        try:
            data = datos.get(n) or descargar(n)
            time.sleep(0 if n in datos else 1)   # no saturar la API
            info, filas = aplanar(data)
        except Exception as ex:
            # Se corta acá: lo que falte se reintenta en la próxima corrida
            print(f"Error con el boletín {n}: {ex}. Se reintenta luego.")
            break
        if info["numero"] != n:
            # Si el número no existe, la API no da error: devuelve el último
            # boletín publicado. Se descarta para no duplicar datos.
            print(f"Boletín {n} no existe (la API devolvió el "
                  f"{info['numero']}); se saltea.")
            continue
        if not filas:
            if n == ultimo_pub:
                print(f"Boletín {n} todavía sin normas; se reintenta luego.")
                break
            print(f"Boletín {n} sin normas (se saltea).")
            continue
        nuevos.append((info, filas))
        print(f"  N° {n} ({info['fecha']:%d/%m/%Y}): {len(filas)} normas")

    carga_historica = len(nuevos) > MAX_RESUMENES
    mandar_historico = args.historico or carga_historica
    if not nuevos and not mandar_historico:
        return 0

    # 3) Base histórica (+ Excel histórico completo, queda como artefacto)
    if nuevos:
        todas = [f for _, filas in nuevos for f in filas]
        base = actualizar_base(todas, {i["numero"] for i, _ in nuevos})
    else:
        base = cargar_base()
    if base.empty:
        print("Todavía no hay datos acumulados.")
        return 0
    excel_hist = generar_excel(base)
    print(f"Base histórica: {len(base)} normas. Excel: {excel_hist}")

    # 4) Telegram
    #    Diario: resumen + Excel liviano con las normas de ese boletín.
    #    Histórico completo: solo a pedido o en una carga histórica.
    if not args.no_telegram:
        if carga_historica:
            enviar_mensaje(armar_carga_historica(nuevos, base))
        else:
            for info, filas in nuevos:
                enviar_mensaje(armar_resumen(info, filas))
                dia = pd.DataFrame(filas, columns=COLUMNAS).astype(str)
                nombre = (f"boletin_caba_{info['fecha']:%Y-%m-%d}"
                          f"_N{info['numero']}.xlsx")
                xl = generar_excel(dia, SALIDA / nombre, diario=True)
                enviar_excel(xl, nombre,
                             f"Normas del Boletín N° {info['numero']} "
                             f"({info['fecha']:%d/%m/%Y})")
        if mandar_historico:
            hoy = datetime.now()
            enviar_excel(excel_hist,
                         f"boletin_caba_historico_{hoy:%Y-%m-%d}.xlsx",
                         pie_historico(base))
        print("Enviado a Telegram.")

    # 5) Estado (solo avanza hasta lo procesado con éxito)
    if nuevos:
        ult_info = nuevos[-1][0]
        if not ultimo_proc or ult_info["numero"] > int(ultimo_proc):
            estado["ultimo_numero"] = ult_info["numero"]
        estado["ultima_corrida"] = {
            "numero": ult_info["numero"],
            "fecha": ult_info["fecha"].strftime("%Y-%m-%d"),
            "boletines": len(nuevos),
            "normas": len(todas),
            "procesado": datetime.now().isoformat(timespec="seconds"),
        }
        guardar_estado(estado)
    return 0


if __name__ == "__main__":
    sys.exit(main())
