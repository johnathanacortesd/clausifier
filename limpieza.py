# limpieza.py — limpieza y orden DETERMINISTAS del dossier, antes del agente.
#
# Portado de API_2_Grill (pipeline.py), solo la parte sin IA: carga con
# hipervínculos, normalización de columnas, expansión de menciones, detección
# de duplicados y orden por título. No llama a ningún modelo: cuesta $0.
#
# Punto de entrada:  limpiar_dossier(ruta_entrada, ruta_salida) -> dict

import datetime
import html
import io
import logging
import os
import re
import unicodedata
from pathlib import Path
from typing import Dict, List, Tuple
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string, range_boundaries

logger = logging.getLogger("limpieza")

TIPOS_AV = frozenset({"Radio", "Televisión"})
TIPO_MEDIO_MAP = {
    "online": "Internet", "internet": "Internet",
    "diario": "Prensa",
    "am": "Radio", "fm": "Radio", "radio": "Radio",
    "aire": "Televisión", "cable": "Televisión", "tv": "Televisión",
    "television": "Televisión", "televisión": "Televisión",
    "revista": "Revistas", "revistas": "Revistas",
}

KEY_MAP = {
    "idnoticia": "ID Noticia", "fecha": "Fecha", "hora": "Hora", "medio": "Medio",
    "tipodemedio": "Tipo de Medio", "titulo": "Título", "link_streaming": "Link (Streaming - Imagen)",
    "menciones": "Menciones - Empresa", "idduplicada": "ID duplicada",
}

THOUSANDS_COLS = {"Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", "Tier", "Audiencia"}
NUMERIC_COLS = {"ID Noticia", "NoticiaId", "ID duplicada", "CPE"} | THOUSANDS_COLS

REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
HYPERLINK_TAIL_BYTES = 4 * 1024 * 1024
MAX_FILAS_POR_DOSSIER = 1000

COL_FILA = "Fila"                       # clave única y estable de cada fila de salida
COL_DUP = "ID duplicada"


class LimiteFilasError(ValueError):
    """El dossier supera MAX_FILAS_POR_DOSSIER filas."""


def _fmt_miles(n: int) -> str:
    return f"{int(n):,}".replace(",", ".")


# ─── Hipervínculos (lectura directa del XML del xlsx) ────────────────────
def _local_tag(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _workbook_rel_path(target: str) -> str:
    t = (target or "").lstrip("/")
    if t.startswith("xl/"):
        return t
    if t.startswith("worksheets/") or t.startswith("theme/") or t == "styles.xml":
        return "xl/" + t
    return t


def _read_xml_tail(sheet_fh, tail_bytes: int) -> bytes:
    try:
        sheet_fh.seek(0, os.SEEK_END)
        size = sheet_fh.tell()
        sheet_fh.seek(max(0, size - tail_bytes))
        return sheet_fh.read()
    except (OSError, AttributeError):
        return sheet_fh.read()


def _assign_hyperlink_ref(result: dict, ref: str, target: str) -> None:
    if ":" in ref:
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                result[(r, c)] = target
        return
    col_letter, row = coordinate_from_string(ref)
    result[(row, column_index_from_string(col_letter))] = target


def extract_hyperlinks_from_xlsx(xlsx_bytes: bytes, sheet_title: str) -> Dict[Tuple[int, int], str]:
    """{(fila, columna): url} de la hoja pedida. Si algo falla, devuelve {} y sigue."""
    result: Dict[Tuple[int, int], str] = {}
    try:
        with ZipFile(io.BytesIO(xlsx_bytes)) as zf:
            wb_tree = ET.fromstring(zf.read("xl/workbook.xml"))
            rels_tree = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            rid_to_target = {r.attrib.get("Id"): r.attrib.get("Target")
                             for r in rels_tree if r.attrib.get("Id") and r.attrib.get("Target")}
            sheet_path = None
            for el in wb_tree.iter():
                if _local_tag(el.tag) == "sheet" and el.attrib.get("name") == sheet_title:
                    rid = el.attrib.get(REL_NS + "id") or el.attrib.get("r:id")
                    sheet_path = _workbook_rel_path(rid_to_target.get(rid, ""))
                    break
            if not sheet_path:
                return result
            if sheet_path not in zf.namelist():
                alt = sheet_path[3:] if sheet_path.startswith("xl/") else "xl/" + sheet_path
                if alt not in zf.namelist():
                    return result
                sheet_path = alt
            dirname, fname = sheet_path.rsplit("/", 1)
            rels_path = f"{dirname}/_rels/{fname}.rels"
            url_by_rid = {}
            if rels_path in zf.namelist():
                for rel in ET.fromstring(zf.read(rels_path)):
                    if rel.attrib.get("TargetMode") == "External" or "hyperlink" in rel.attrib.get("Type", ""):
                        if rel.attrib.get("Id") and rel.attrib.get("Target"):
                            url_by_rid[rel.attrib["Id"]] = rel.attrib["Target"]
            with zf.open(sheet_path) as fh:
                tail = _read_xml_tail(fh, HYPERLINK_TAIL_BYTES)
            idx = tail.rfind(b"<hyperlinks")
            if idx < 0:
                return result
            end = tail.find(b"</hyperlinks>", idx)
            chunk = tail[idx:] if end < 0 else tail[idx:end + len(b"</hyperlinks>")]
            # el bloque suelto pierde la declaración del prefijo r:, la agregamos
            if b"xmlns:r=" not in chunk:
                chunk = chunk.replace(
                    b"<hyperlinks",
                    b'<hyperlinks xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"', 1)
            try:
                hyper_tree = ET.fromstring(chunk)
            except ET.ParseError:
                logger.warning("No se pudo parsear el bloque de hipervínculos; se sigue sin ellos.")
                return result
            for el in hyper_tree.iter():
                if _local_tag(el.tag) != "hyperlink":
                    continue
                ref = el.attrib.get("ref")
                rid = el.attrib.get(REL_NS + "id") or el.attrib.get("r:id")
                target = (url_by_rid.get(rid) if rid else None) or el.attrib.get("location")
                if ref and target:
                    _assign_hyperlink_ref(result, ref, target)
    except Exception:  # noqa: BLE001
        logger.exception("Fallo al extraer hipervínculos; se sigue con los valores de celda.")
    return result


# ─── Utilidades de texto ─────────────────────────────────────────────────
def norm_key(text) -> str:
    if text is None:
        return ""
    t = unicodedata.normalize("NFKD", str(text).strip().lower()).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "", t)


def get_column_robust(df, name):
    name_norm = norm_key(name)
    for col in df.columns:
        if norm_key(col) == name_norm:
            return df[col]
    return pd.Series([np.nan] * len(df), index=df.index)


def clean_text(text):
    if not isinstance(text, str):
        return text
    return re.sub(r"\s+", " ", text).strip()


def decode_html_entities(text):
    """'&#xF3;' -> 'ó', '&amp;' -> '&'. Hasta 3 pasadas (doble codificación)."""
    if not isinstance(text, str):
        return text
    out = text
    for _ in range(3):
        decoded = html.unescape(out)
        if decoded == out:
            break
        out = decoded
    return out


def clean_cuerpo(text):
    if not isinstance(text, str) or text.strip() in ("", "nan", "None"):
        return ""
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    return decode_html_entities(text).strip()


def clean_title_for_output(title):
    if not isinstance(title, str) or title in ("nan", "None"):
        return ""
    return re.sub(r"\s+", " ", decode_html_entities(title)).strip()


def normalizar_tipo_medio(tipo_raw) -> str:
    if not isinstance(tipo_raw, str):
        return str(tipo_raw)
    t = norm_key(tipo_raw) if tipo_raw.strip() else ""
    for k, v in TIPO_MEDIO_MAP.items():
        if norm_key(k) == t:
            return v
    return tipo_raw.strip().title() or "Otro"


def parse_numeric(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        if isinstance(val, float) and (math_isnan(val)):
            return None
        if isinstance(val, float) and val.is_integer():
            return int(val)
        return val
    s = re.sub(r"[^\d.,\-eE]", "", str(val).strip())
    if not s:
        return None
    dots, commas = s.count("."), s.count(",")
    if dots > 1 and commas == 0:
        s = s.replace(".", "")
    elif commas > 1 and dots == 0:
        s = s.replace(",", "")
    elif dots and commas:
        s = s.replace(",", "") if s.rfind(".") > s.rfind(",") else s.replace(".", "").replace(",", ".")
    elif dots == 1 and len(s.split(".")[1]) == 3:
        s = s.replace(".", "")
    elif commas == 1:
        s = s.replace(",", "") if len(s.split(",")[1]) == 3 else s.replace(",", ".")
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return None


def math_isnan(v) -> bool:
    return isinstance(v, float) and v != v


# ─── Duplicados ──────────────────────────────────────────────────────────
def _normalizar_url(url: str) -> str:
    if not url:
        return ""
    url = re.sub(r"^https?://", "", url.strip().lower())
    return re.sub(r"^www\.", "", url).rstrip("/")


def _extract_url(val) -> str:
    if val is None:
        return ""
    if isinstance(val, dict):
        return str(val.get("url") or "").strip()
    if math_isnan(val):
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("", "nan", "none", "link") else s


def _normalizar_hora(val) -> str:
    if val is None or math_isnan(val):
        return ""
    if isinstance(val, (pd.Timestamp, datetime.datetime, datetime.time)):
        return "" if (isinstance(val, pd.Timestamp) and pd.isna(val)) else val.strftime("%H:%M:%S")
    s = str(val).strip()
    if s.lower() in ("", "nan", "nat", "none"):
        return ""
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?", s)
    if m:
        h, mi, se = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        if 0 <= h < 24 and 0 <= mi < 60 and 0 <= se < 60:
            return f"{h:02d}:{mi:02d}:{se:02d}"
    return s


def detectar_duplicados_avanzado(rows, km=KEY_MAP):
    """Marca is_duplicate y llena 'ID duplicada' con el ID de la primera aparición.

    - Radio/TV: misma mención + mismo medio + misma hora.
    - Resto: misma URL normalizada (URL Nota o Link Streaming) + misma mención.
    """
    seen_url, seen_bcast = {}, {}
    for i, row in enumerate(rows):
        if row.get("is_duplicate"):
            continue
        tipo = normalizar_tipo_medio(str(row.get(km["tipodemedio"], "")))
        mencion = norm_key(row.get(km["menciones"], ""))
        if not mencion:
            continue
        if tipo in TIPOS_AV:
            medio = norm_key(row.get(km["medio"], ""))
            hora = _normalizar_hora(row.get(km["hora"]))
            if not medio or not hora:
                continue
            k, seen = (mencion, medio, hora), seen_bcast
        else:
            url = _normalizar_url(_extract_url(row.get("URL Nota")) or _extract_url(row.get(km["link_streaming"])))
            if not url:
                continue
            k, seen = (url, mencion), seen_url
        if k in seen:
            row["is_duplicate"] = True
            row[km["idduplicada"]] = rows[seen[k]].get(km["idnoticia"], "")
        else:
            seen[k] = i
    return rows


# ─── Lectura ─────────────────────────────────────────────────────────────
def _rows_to_dataframe(raw_headers, data_rows, hyperlinks) -> pd.DataFrame:
    rows, empty_streak = [], 0
    for ridx, values in enumerate(data_rows, start=2):
        if all(v is None or v == "" for v in values):
            empty_streak += 1
            if empty_streak >= 50:
                break
            continue
        empty_streak = 0
        row_data = {}
        for i, h in enumerate(raw_headers):
            if not h:
                continue
            val = values[i] if i < len(values) else None
            url = hyperlinks.get((ridx, i + 1))
            row_data[h] = {"value": val or "Link", "url": url} if url else val
        rows.append(row_data)
    return pd.DataFrame(rows, columns=[h for h in raw_headers if h])


def load_dossier_dataframe(file_bytes: bytes) -> pd.DataFrame:
    """Primera hoja con openpyxl + hipervínculos de las celdas."""
    bio = io.BytesIO(file_bytes)
    wb = openpyxl.load_workbook(bio, read_only=True, data_only=True)
    try:
        sheet = wb.active
        hyperlinks = extract_hyperlinks_from_xlsx(file_bytes, sheet.title or "Sheet")
        it = sheet.iter_rows(values_only=True)
        header = next(it, None)
        if header is None:
            return pd.DataFrame()
        raw_headers = [str(h).strip() if h is not None else None for h in header]
        return _rows_to_dataframe(raw_headers, (list(r) for r in it), hyperlinks)
    finally:
        wb.close()


def cargar_mapa_csv(url: str | None) -> dict:
    """CSV de 2 columnas (clave, valor) sin encabezado. Sin URL o con error: {}."""
    if not url:
        return {}
    try:
        df = pd.read_csv(url, header=None, dtype=str).dropna(how="all")
        m = pd.Series(df.iloc[:, 1].values, index=df.iloc[:, 0].astype(str).str.lower().str.strip()).to_dict()
        return {k: v for k, v in m.items() if k not in ("nan", "")}
    except Exception as e:  # noqa: BLE001 — sin mapa se sigue con "N/A"
        logger.warning("No se pudo cargar el CSV %s: %s", url, e)
        return {}


# ─── Normalización ───────────────────────────────────────────────────────
def normalize_dossier_dataframe(df, region_map, internet_map):
    if df is None or df.empty:
        return pd.DataFrame()
    n = len(df)
    vacia = pd.Series([None] * n, index=df.index, dtype=object)

    if "Tipo de Medio" in df.columns:
        df["Tipo de Medio"] = df["Tipo de Medio"].apply(
            lambda v: normalizar_tipo_medio(v) if isinstance(v, str) and v.strip() else v)
    tipo = df.get("Tipo de Medio", pd.Series(["Otro"] * n, index=df.index)).astype(str)
    is_av = tipo.isin(list(TIPOS_AV))
    is_grafica = tipo.isin(["Prensa", "Internet", "Revistas"])
    is_internet = tipo == "Internet"

    # Cuerpo: primera columna con datos de la cadena de respaldo
    cuerpo_col = None
    for nombre in ("CuerpoEs", "Resumen - Aclaracion", "Resumen", "Cuerpo", "resumen corto"):
        s = get_column_robust(df, nombre)
        if not s.dropna().empty:
            cuerpo_col = next(c for c in df.columns if norm_key(c) == norm_key(nombre))
            break
    if cuerpo_col is not None:
        df[cuerpo_col] = df[cuerpo_col].apply(lambda v: clean_cuerpo(v) if isinstance(v, str) else v)
    if "Resumen - Aclaracion" in df.columns and cuerpo_col != "Resumen - Aclaracion":
        # Como en Grill: el Resumen - Aclaracion lleva el cuerpo (CuerpoEs) cuando viene vacío.
        resumen = df["Resumen - Aclaracion"].apply(lambda v: clean_cuerpo(v) if isinstance(v, str) else v)
        df["Resumen - Aclaracion"] = resumen.where(
            resumen.notna() & (resumen != ""), df[cuerpo_col] if cuerpo_col else resumen)

    if "Medio" in df.columns:
        clave_medio = df["Medio"].astype(str).str.lower().str.strip()
        previa = df["Región"] if "Región" in df.columns else vacia    # conserva la región que ya traiga
        df["Región"] = clave_medio.map(region_map).fillna(previa).fillna("N/A").replace("", "N/A")
        if internet_map:
            df.loc[is_internet, "Medio"] = clave_medio[is_internet].map(internet_map).fillna(df.loc[is_internet, "Medio"])
    else:
        df["Región"] = "N/A"

    df["ID Noticia"] = df["NoticiaId"] if "NoticiaId" in df.columns else df.get("ID Noticia", vacia)
    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"].apply(lambda v: v if not isinstance(v, dict) else v.get("value")),
                                     dayfirst=True, errors="coerce").dt.normalize()
    for c in ("Sección - Programa", "Autor - Conductor"):
        if c in df.columns:
            df[c] = df[c].apply(clean_text)
    for c in ("Título", "Titulo"):
        if c in df.columns:
            df[c] = df[c].apply(lambda v: clean_title_for_output(v) if isinstance(v, str) else v)

    # Radio/TV: la duración viene en "Duración - Nro. Caracteres"; va a Dimensión y Duración queda en 0
    if "Dimensioncm2" in df.columns and "Dimensión" in df.columns:
        df["Dimensión"] = df["Dimensión"].where(df["Dimensión"].notna(), df["Dimensioncm2"])
    if "Dimensión" in df.columns and "Duración - Nro. Caracteres" in df.columns:
        df["Dimensión"] = df["Dimensión"].astype(object)
        df["Duración - Nro. Caracteres"] = df["Duración - Nro. Caracteres"].astype(object)
        df.loc[is_av, "Dimensión"] = df.loc[is_av, "Duración - Nro. Caracteres"]
        df.loc[is_av, "Duración - Nro. Caracteres"] = 0

    if "Tipo de Medio" in df.columns:
        if "Valor de Nota" in df.columns:
            df["CPE"] = np.where(is_av, get_column_robust(df, "CPE"),
                                 np.where(is_grafica, get_column_robust(df, "Valor de Nota"), np.nan))
        # Enlaces: Link Nota = AV -> URL Nota AV; resto -> URL (Streaming - Imagen)
        url_av = df.get("URL Nota AV", df.get("Link Nota AV", vacia))
        url_str = df.get("URL (Streaming - Imagen)", vacia)
        hay_av, hay_str = ("URL Nota AV" in df.columns or "Link Nota AV" in df.columns), "URL (Streaming - Imagen)" in df.columns
        if "Link Nota" in df.columns and (hay_av or hay_str):
            ln = []
            for v_av, v_str, av in zip(url_av, url_str, is_av):
                u = _extract_url(v_av).replace(".com.ar", ".com.co") if av else _extract_url(v_str)
                ln.append({"value": "Link", "url": u} if u else None)
            df["Link Nota"] = ln
        if "Link (Streaming - Imagen)" in df.columns and "URL Nota" in df.columns:
            df["Link (Streaming - Imagen)"] = [
                ({"value": "Link", "url": _extract_url(u)} if _extract_url(u) else None) if inet else None
                for u, inet in zip(df.get("URL Nota", vacia), is_internet)]
        if hay_str or "Audio" not in df.columns:
            df["Audio"] = [(_extract_url(u) if av else "") for u, av in zip(url_str, is_av)]
        if "Link Ext" not in df.columns or hay_av or hay_str:
            df["Link Ext"] = [_url_link_ext(r) for r in df.to_dict("records")]
        if "Web Ext" not in df.columns or "URL Nota" in df.columns:
          df["Web Ext"] = [(_extract_url(r.get("URL Nota")) if r.get("Tipo de Medio") == "Internet" else "")
                         for r in df.to_dict("records")]
        # Menciones: Radio/TV usa "Menciones - Empresa"; gráfica usa "Empresa rel."
        men_av = df.get("Menciones - Empresa", vacia).fillna("").astype(str).apply(clean_text)
        men_gr = df.get("Empresa rel.", vacia).fillna("").astype(str).apply(clean_text)
        if "Empresa rel." not in df.columns:
            men_gr = men_av
        df["Menciones - Empresa"] = np.where(is_av, men_av, np.where(is_grafica, men_gr, men_av))

    for c in NUMERIC_COLS & set(df.columns):
        df[c] = df[c].apply(lambda v: v if isinstance(v, dict) else parse_numeric(v))
    return df


def _url_link_ext(row) -> str:
    tipo = row.get("Tipo de Medio")
    if tipo in TIPOS_AV:
        return _extract_url(row.get("Link Nota"))
    if tipo == "Internet":
        return _extract_url(row.get("URL (Streaming - Imagen)")) or _extract_url(row.get("Link Nota"))
    if tipo in ("Prensa", "Revistas"):
        return (_extract_url(row.get("Link (Streaming - Imagen)")) or _extract_url(row.get("Link Nota"))
                or _extract_url(row.get("URL (Streaming - Imagen)")))
    return ""


def expand_menciones(df) -> List[dict]:
    """Una fila por cada mención separada por ';' en 'Menciones - Empresa'."""
    out = []
    for idx, rec in enumerate(df.to_dict("records")):
        crudo = rec.get("Menciones - Empresa")
        crudo = "" if crudo is None or math_isnan(crudo) else str(crudo)
        menciones = [m.strip() for m in crudo.split(";") if m.strip()] or [None]
        for m in menciones:
            r = dict(rec)
            if "Menciones - Empresa" in df.columns:
                r["Menciones - Empresa"] = m or ""
            r["original_index"] = idx
            r["is_duplicate"] = False
            out.append(r)
    return out


def _columnas_preservadas(columnas_entrada, agregadas) -> List[str]:
    """Todas las columnas originales en su orden + las nuevas al final."""
    cols = [c for c in columnas_entrada if c]
    for c in agregadas:
        if c not in cols:
            cols.append(c)
    return cols


# ─── Escritura ───────────────────────────────────────────────────────────
def _valor_celda(v):
    if v is None or math_isnan(v):
        return None
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.to_pydatetime()
    if isinstance(v, np.generic):
        return v.item()
    return v


def _escribir_xlsx(rows, cols, ruta_salida: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dossier"
    ws.append(cols)
    for r_i, row in enumerate(rows, start=2):
        for c_i, col in enumerate(cols, start=1):
            v = row.get(col)
            cell = ws.cell(row=r_i, column=c_i)
            if isinstance(v, dict):
                url = v.get("url")
                cell.value = _valor_celda(v.get("value")) or ("Link" if url else None)
                if url:
                    cell.hyperlink = url
            else:
                v = _valor_celda(v)
                cell.value = v
                if isinstance(v, datetime.datetime):
                    cell.number_format = "DD/MM/YYYY"
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ruta_salida)


# ─── Punto de entrada ────────────────────────────────────────────────────
def limpiar_dossier(ruta_entrada, ruta_salida) -> dict:
    """Limpia, expande menciones, marca duplicados y ordena por título. Sin IA.

    Escribe ruta_salida y devuelve un resumen:
    total, unicos, duplicados, filas_expandidas, medios_sin_region, columnas.
    """
    ruta_entrada, ruta_salida = Path(ruta_entrada), Path(ruta_salida)
    file_bytes = ruta_entrada.read_bytes()
    df = load_dossier_dataframe(file_bytes)
    n_filas = len(df)
    if n_filas > MAX_FILAS_POR_DOSSIER:
        raise LimiteFilasError(
            f"El archivo tiene {_fmt_miles(n_filas)} filas y el máximo por corrida es "
            f"{_fmt_miles(MAX_FILAS_POR_DOSSIER)}. Divide el xlsx en partes de hasta "
            f"{_fmt_miles(MAX_FILAS_POR_DOSSIER)} filas y procesa cada una por separado.")
    if df.empty:
        raise ValueError("El archivo no tiene filas con datos.")

    columnas_entrada = [c for c in df.columns if c and c != COL_FILA]
    formato_grill = "Tipo de Medio" in columnas_entrada
    df = normalize_dossier_dataframe(df, cargar_mapa_csv(os.environ.get("REGIONES_CSV_URL")),
                                     cargar_mapa_csv(os.environ.get("INTERNET_CSV_URL")))

    medios_sin_region = []
    if "Medio" in columnas_entrada:
        medios_sin_region = sorted(set(df.loc[df["Región"] == "N/A", "Medio"].astype(str).str.strip())
                                   - {"", "nan", "None"})

    rows = detectar_duplicados_avanzado(expand_menciones(df))
    col_titulo = "Título" if "Título" in df.columns else ("Titulo" if "Titulo" in df.columns else None)
    rows.sort(key=lambda r: (norm_key(r.get(col_titulo, "")) if col_titulo else "", str(r.get("ID Noticia", ""))))
    for i, r in enumerate(rows, start=1):
        r[COL_FILA] = i
        if not r.get("is_duplicate"):
            r[COL_DUP] = None

    agregadas = (["Región"] if "Medio" in columnas_entrada else [])
    if formato_grill:
        agregadas += ["Audio", "Link Ext", "Web Ext"]
    agregadas += [COL_DUP, COL_FILA]
    cols = _columnas_preservadas(columnas_entrada, agregadas)
    _escribir_xlsx(rows, cols, ruta_salida)

    duplicados = sum(1 for r in rows if r.get("is_duplicate"))
    return {
        "archivo": str(ruta_salida),
        "filas_entrada": n_filas,
        "total": len(rows),
        "unicos": len(rows) - duplicados,
        "duplicados": duplicados,
        "filas_expandidas": len(rows) - n_filas,
        "medios_sin_region": medios_sin_region,
        "columnas": cols,
    }
