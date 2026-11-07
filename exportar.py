# exportar.py — escribe el xlsx final con el layout EXACTO del export de API_2_Grill
# (26 columnas, en este orden, nada más). Sin IA, $0.
#
# Lee el archivo de trabajo (entregas/<stem>_limpio.xlsx, que puede traer "Fila" y
# otras columnas internas) y copia cada columna final desde la columna equivalente
# (sin importar tildes ni mayúsculas). Lo que no exista sale vacío pero presente.
# Formatos portados de pipeline._write_xlsx_rows de Grill (xlsxwriter).

import datetime
import re
import unicodedata
from pathlib import Path
from typing import Any

import openpyxl
import xlsxwriter

COLUMNAS_FINALES = [
    "ID Noticia", "Fecha", "Medio", "Región", "Tipo de Medio", "Sección - Programa", "Título",
    "Duración - Nro. Caracteres", "CPE", "Audiencia", "Tier", "Tono_IA", "Subtema_IA", "TIPO",
    "GÉNERO", "Link Nota", "texto", "Link (Streaming - Imagen)", "Link Ext", "FECHA INFORME",
    "HORA B", "ID duplicada", "Contexto analizado", "Resumen - Aclaracion", "Audio", "Web Ext",
]

# Nombres alternos de entrada para cada columna final (además del propio nombre).
ALTERNOS = {
    "ID Noticia": ["NoticiaId", "Id Noticia"],
    "Título": ["Titular", "Title"],
    "Tono_IA": ["Tono"],
    "Subtema_IA": ["Subtema"],
    "GÉNERO": ["Genero"],
}
COLS_IA = ("Tono_IA", "Subtema_IA")
COLS_LINK = {"Link Nota", "Link (Streaming - Imagen)", "Link Ext", "Web Ext", "Audio"}
COLS_FECHA = {"Fecha", "FECHA INFORME"}
COLS_ID = {"ID Noticia", "ID duplicada"}
COLS_MILES = {"Duración - Nro. Caracteres", "Audiencia", "Tier"}
COLS_MONEDA = {"CPE"}
COLS_ANCHAS = {"Título", "Resumen - Aclaracion", "Contexto analizado", "Audio", "texto",
               "Link Ext", "Web Ext"}


def _n(texto: Any) -> str:
    texto = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return " ".join(re.sub(r"[^a-z0-9]+", " ", texto.lower()).split())


def mapa_columnas(encabezado: list[Any]) -> dict[str, int | None]:
    """columna final → índice en el archivo de trabajo (o None si no existe)."""
    indice = {}
    for i, h in enumerate(encabezado):
        if h is not None:
            indice.setdefault(_n(h), i)
    return {col: next((indice[_n(c)] for c in [col, *ALTERNOS.get(col, [])] if _n(c) in indice), None)
            for col in COLUMNAS_FINALES}


def _como_id(v: Any) -> Any:
    limpio = re.sub(r"[^\d.]", "", str(v))
    try:
        return int(float(limpio)) if limpio else str(v)
    except ValueError:
        return str(v)


def exportar(trabajo: Path, destino: Path, con_ia: bool = True) -> dict:
    """Escribe `destino` con las 26 columnas finales. con_ia=False deja Tono_IA/Subtema_IA vacías."""
    hoja = openpyxl.load_workbook(trabajo).active       # no read_only: necesitamos los hipervínculos
    filas = list(hoja.iter_rows())
    encabezado = [c.value for c in filas[0]]
    mapa = mapa_columnas(encabezado)
    if not con_ia:
        mapa.update({c: None for c in COLS_IA})

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(destino), {"strings_to_urls": False, "nan_inf_to_errors": False})
    ws = wb.add_worksheet("Resultado")
    f_head = wb.add_format({"bold": True})
    f_link = wb.add_format({"font_color": "#0563C1", "underline": 1, "align": "left"})
    f_fecha = wb.add_format({"num_format": "dd/mm/yyyy"})
    f_moneda = wb.add_format({"num_format": "$#,##0"})
    f_miles = wb.add_format({"num_format": "#,##0"})
    f_id = wb.add_format({"num_format": "0"})
    for i, col in enumerate(COLUMNAS_FINALES):
        ancho = 55 if col in COLS_ANCHAS else 15 if col in ("Link Nota", "Link (Streaming - Imagen)") \
            else 28 if col in COLS_IA else 20
        ws.set_column(i, i, ancho)
        ws.write(0, i, col, f_head)

    escritas = enlaces = 0
    for fila in filas[1:]:
        if all(c.value in (None, "") for c in fila):
            continue
        escritas += 1
        r = escritas
        for ci, col in enumerate(COLUMNAS_FINALES):
            i = mapa[col]
            celda = fila[i] if i is not None and i < len(fila) else None
            v = celda.value if celda is not None else None
            url = celda.hyperlink.target if celda is not None and celda.hyperlink else None
            if col in COLS_LINK and not url and isinstance(v, str) and re.match(r"https?://", v.strip()):
                url = v.strip()
            if url:
                ws.write_url(r, ci, url, f_link, string=str(v) if v not in (None, "") else url)
                enlaces += 1
            elif v in (None, ""):
                ws.write_blank(r, ci, None)
            elif col in COLS_FECHA and isinstance(v, (datetime.datetime, datetime.date)):
                ws.write_datetime(r, ci, v, f_fecha)
            elif col in COLS_ID:
                ws.write(r, ci, _como_id(v), f_id)
            elif col in COLS_MONEDA and isinstance(v, (int, float)):
                ws.write_number(r, ci, v, f_moneda)
            elif col in COLS_MILES and isinstance(v, (int, float)):
                ws.write_number(r, ci, v, f_miles)
            elif isinstance(v, (int, float)):
                ws.write_number(r, ci, v)
            elif isinstance(v, (datetime.datetime, datetime.date)):
                ws.write_datetime(r, ci, v, f_fecha)
            else:
                ws.write_string(r, ci, str(v))
    wb.close()
    return {"filas": escritas, "enlaces": enlaces,
            "vacias": [c for c, i in mapa.items() if i is None]}
