# contexto.py — "Contexto analizado" y "texto", portados de API_2_Grill (analyzer_tono_tema.py
# v4.74 y pipeline.py). Sin IA, $0. El agente clasifica TONO y TEMA leyendo esta columna.
#
#   Contexto analizado = [Título si menciona marca/alias/vocero] + oraciones literales con
#                        mención (deduplicadas, tope 2000 / 1200 en radio-TV), cortadas al
#                        3er punto; sin mención → lead del cuerpo (3 oraciones); sin cuerpo →
#                        título. Si el texto resultante es corto (≤60 palabras o ≤370 car.)
#                        se agrega una oración más.
#   texto              = Contexto analizado sin el Título.

from __future__ import annotations

import html
import re
import unicodedata
from functools import lru_cache
from pathlib import Path
from typing import List, Optional, Sequence

import openpyxl

COL_CONTEXTO = "Contexto analizado"
COL_TEXTO = "texto"


# ─── utilidades de texto (analyzer_tono_tema.py) ──────────────────────────
_CTRL_RE = re.compile('[\x00-\x08\x0b-\x1f￾￿]')


def ctrl(s) -> str:
    return _CTRL_RE.sub('', str(s or ''))


@lru_cache(maxsize=50000)
def _nz_str(s: str) -> str:
    s = unicodedata.normalize('NFKD', ctrl(s))
    s = ''.join(c for c in s if not unicodedata.combining(c)).lower()
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9ñ ]+', ' ', s)).strip()


def nz(s) -> str:
    return _nz_str(s if type(s) is str else str(s or ''))


def sq(s) -> str:
    return re.sub(r'\s+', ' ', ctrl(s)).strip()


def decode_html_entities(text):
    """&#xF3; → ó, &amp; → &; hasta 3 pasadas por doble codificación (pipeline.py)."""
    if not isinstance(text, str):
        return text
    out = text
    for _ in range(3):
        decoded = html.unescape(out)
        if decoded == out:
            break
        out = decoded
    return out


def expandir_voceros(voceros: Sequence[str]) -> List[str]:
    """Nombre completo + variantes nombre + primer apellido.

    «Miguel Silva Moyano» → «Miguel Silva»; «Juan Carlos Pérez» → «Juan Pérez».
    Con 3 palabras no se sabe si la 2ª es nombre o apellido: se agregan ambas.
    (Reconstruido: coherencia_texto.py de Grill no estaba disponible.)
    """
    out: List[str] = []
    for v in voceros or []:
        v = sq(v)
        if not v:
            continue
        out.append(v)
        p = v.split()
        if len(p) == 3:
            out += [f"{p[0]} {p[1]}", f"{p[0]} {p[2]}"]
        elif len(p) >= 4:
            out.append(f"{p[0]} {p[2]}")
    vistos, res = set(), []
    for x in out:
        if nz(x) and nz(x) not in vistos:
            vistos.add(nz(x))
            res.append(x)
    return res


def texto_sin_titulo(contexto: str, titulo: str) -> str:
    """Contexto sin el Título, aunque difiera en forma (comillas, espacios, entidades);
    nunca corta dentro de una palabra. (Reconstruido desde README de Grill.)"""
    c = sq(decode_html_entities(contexto or ''))
    t = sq(decode_html_entities(titulo or ''))
    if not c or not t:
        return c
    palabras = re.findall(r'\w+', t)
    if not palabras:
        return c
    patron = r'(?<!\w)' + r'\W*'.join(re.escape(w) for w in palabras) + r'(?!\w)\W*'
    m = re.match(r'\W*' + patron, c, flags=re.I)
    if not m:
        m = re.search(patron, c, flags=re.I)
    if not m:
        return c
    return sq(c[:m.start()] + ' ' + c[m.end():])


# ─── coincidencia de marca ────────────────────────────────────────────────
def _regex_coincidencia(brand: str, aliases: Sequence[str], voceros: Sequence[str]) -> List[str]:
    rx = []
    alineados = [x for x in aliases if x and len(str(x).strip()) >= 2]
    vocs = [v for v in voceros if v and len(str(v).strip()) >= 2]
    for nombre in [brand] + list(alineados) + list(vocs):
        n = nz(nombre)
        if not n or len(n) < 2:
            continue
        rx.append(r'(?<![a-z0-9])' + re.escape(n) + r'(?![a-z0-9])')
    return rx


def _mención_normalizada_en(p_norm: str, rx) -> bool:
    return bool(rx) and any(re.search(r, p_norm) for r in rx)


def _titulo_menciona_marca(titulo, brand, aliases, voceros) -> bool:
    if not (titulo or '').strip():
        return False
    rx = _regex_coincidencia(brand, aliases or [], voceros or [])
    return bool(rx) and _mención_normalizada_en(nz(titulo), rx)


# ─── recortes ─────────────────────────────────────────────────────────────
_ABREV_PUNTO = {'a', 'm', 'p', 'dr', 'dra', 'sr', 'sra', 'srta', 'ing',
                'lic', 'no', 'art', 'aprox', 'ej', 'etc', 'vs', 'ud',
                'uds', 'sres', 'd', 'c', 'e', 'g', 'ph', 'jr'}


def _cortar_tercer_punto(texto: str, n: int = 3) -> str:
    """Recorta al n-ésimo punto de oración (no cuentan abreviaturas, iniciales ni decimales)."""
    s = re.sub(r'\s+', ' ', (texto or '')).strip()
    if not s:
        return ''
    contados = 0
    for m in re.finditer(r'[.!?…]+', s):
        if (m.start() > 0 and m.end() < len(s)
                and s[m.start() - 1].isdigit() and s[m.end()].isdigit()):
            continue
        antes = s[:m.start()].strip()
        prev = re.split(r'\s+', antes)[-1] if antes else ''
        prev_limpia = re.sub(r'[^a-záéíóúüñ]', '', prev.lower())
        if prev_limpia in _ABREV_PUNTO:
            continue
        if len(prev_limpia) == 1 and prev_limpia.isalpha():
            continue
        sig = s[m.end():].lstrip()
        if sig and not (sig[0].isupper() or sig[0].isdigit() or sig[0] in '"“('):
            continue
        contados += 1
        if contados == n:
            return s[:m.end()].strip()
    return s


VENTANA_CTX_MINIMO = 1500


def _ventana_lead(cuerpo: str) -> str:
    c = re.sub(r'\s+', ' ', (cuerpo or '')).strip()
    if len(c) <= VENTANA_CTX_MINIMO:
        return c
    corte = c[:VENTANA_CTX_MINIMO]
    fines = [m.end() for m in re.finditer(r'[.!?…]\s', corte) if m.end() >= 600]
    return corte[:fines[-1]].strip() if fines else corte.strip()


UMBRAL_TEXTO_PALABRAS = 60
UMBRAL_TEXTO_CARACTERES = 370


def _oracion_siguiente_en_cuerpo(ctx_exact: str, cuerpo: str) -> str:
    if not (ctx_exact or '').strip() or not (cuerpo or '').strip():
        return ''

    def _norm(x):
        return re.sub(r'\s+', ' ', x or '').strip()
    cuerpo_or = [_norm(o) for o in re.split(r'(?<=[.!?…])\s+|\r?\n+', cuerpo) if _norm(o)]
    extracto_or = [_norm(o) for o in re.split(r'(?<=[.!?…])\s+|\r?\n+', ctx_exact) if _norm(o)]
    if not cuerpo_or or not extracto_or:
        return ''
    idxs = [i for i, o in enumerate(cuerpo_or) if o == extracto_or[-1]]
    if not idxs or idxs[-1] + 1 >= len(cuerpo_or):
        return ''
    siguiente = cuerpo_or[idxs[-1] + 1]
    return '' if siguiente in set(extracto_or) else siguiente


def _contexto_minimo_util(ctx: str, titulo: str, texto: str) -> str:
    c = (ctx or '').strip()
    if c:
        return _cortar_tercer_punto(c)
    lead = _cortar_tercer_punto(_ventana_lead(texto or ''))
    if lead:
        return lead
    return _cortar_tercer_punto(titulo or '')


def _contexto_con_titulo_y_extension(ctx_exact, titulo, cuerpo, brand, aliases, voceros) -> str:
    base = _contexto_minimo_util(ctx_exact, titulo, cuerpo)
    titulo_limpio = (titulo or '').strip()
    prefijo = ''
    if (titulo_limpio and titulo_limpio not in base
            and _titulo_menciona_marca(titulo_limpio, brand, aliases, voceros)):
        prefijo = titulo_limpio + ' '
    texto = texto_sin_titulo(prefijo + base, titulo_limpio)
    if (len(texto.split()) <= UMBRAL_TEXTO_PALABRAS
            or len(texto) <= UMBRAL_TEXTO_CARACTERES):
        fuente = (ctx_exact or '').strip() or _ventana_lead(cuerpo or '')
        extendido = _cortar_tercer_punto(fuente, 4)
        if extendido and extendido != base:
            base = extendido
        elif (ctx_exact or '').strip():
            siguiente = _oracion_siguiente_en_cuerpo(ctx_exact, cuerpo)
            if siguiente:
                base = (base + ' ' + siguiente).strip() if base else siguiente
    return (prefijo + base).strip()


MEDIOS_RADIODIFUSION = {'radio', 'television', 'tv', 'aire', 'cable', 'am', 'fm'}
TOPE_CTX_GENERAL = 2000
TOPE_CTX_RADIODIFUSION = 1200


def _contexto_exacto_marca(texto, titulo, brand, aliases, voceros=(), tipo_medio='') -> str:
    """Oraciones literales con mención, deduplicadas, hasta el tope (1200 radio-TV / 2000)."""
    fuente = str(texto or '')
    nombres = [str(x).strip() for x in [brand, *(aliases or []), *(voceros or [])]
               if str(x or '').strip() and len(nz(x)) >= 2]
    if not fuente or not nombres:
        return str(titulo or '').strip()[:TOPE_CTX_GENERAL]
    patrones = [re.compile(r'(?<![a-z0-9])' + re.escape(nz(x)) + r'(?![a-z0-9])') for x in nombres]

    def menciona(fragmento: str) -> bool:
        return any(p.search(nz(fragmento)) for p in patrones)

    oraciones = [o.strip() for o in re.split(r'(?<=[.!?…])\s+|\r?\n+', fuente) if o.strip()]
    hits = [o for o in oraciones if menciona(o)]
    if not hits:
        return ''
    tope = TOPE_CTX_RADIODIFUSION if nz(tipo_medio) in MEDIOS_RADIODIFUSION else TOPE_CTX_GENERAL
    vistas: List[str] = []
    for o in hits:
        o = re.sub(r'\s+', ' ', o).strip()
        if o and nz(o) not in {nz(v) for v in vistas}:
            vistas.append(o)
    salida, total = [], 0
    for o in vistas:
        if salida and total + len(o) + 1 > tope:
            break
        salida.append(o)
        total += len(o) + 1
    texto_out = ' '.join(salida).strip()
    return texto_out[:tope] if len(texto_out) > tope else texto_out


CANDIDATOS_FUENTE = ['CuerpoEs', 'Cuerpo', 'Texto', 'Texto completo', 'Resumen - Aclaracion',
                     'resumen corto', 'Resumen']


def _limpio(v) -> Optional[str]:
    if v is None or (isinstance(v, float) and v != v):
        return None
    s = decode_html_entities(ctrl(v))
    return s if s.strip() and s.strip().lower() != 'nan' else None


def _texto_contexto_fila(row: dict, brand='', aliases=(), voceros=()) -> str:
    """CuerpoEs si menciona marca/alias/vocero; si no, el primer texto que la mencione;
    si ninguno, el más largo."""
    rx = _regex_coincidencia(brand, aliases or [], voceros or [])
    textos = [t for t in (_limpio(row.get(c)) for c in CANDIDATOS_FUENTE) if t]
    if not textos:
        return ''
    if rx:
        for t in textos:
            if _mención_normalizada_en(nz(t), rx):
                return t
    return max(textos, key=len)


# ─── API pública ──────────────────────────────────────────────────────────
def contexto_fila(row: dict, cliente: str, alias: Sequence[str], voceros: Sequence[str]) -> tuple[str, str, str]:
    """(Contexto analizado, texto, origen) con origen ∈ {'mencion', 'lead', 'titulo'}."""
    titulo = sq(decode_html_entities(ctrl(row.get('Título') or '')))
    fuente = _texto_contexto_fila(row, cliente, alias, voceros)
    ctx = _contexto_exacto_marca(fuente, titulo, cliente, alias, voceros,
                                 tipo_medio=str(row.get('Tipo de Medio') or ''))
    final = _contexto_con_titulo_y_extension(ctx, titulo, fuente, cliente, alias, voceros)
    origen = 'mencion' if ctx else ('lead' if _ventana_lead(fuente) else 'titulo')
    return final, texto_sin_titulo(final, titulo), origen


def agregar_contexto(ruta_xlsx, cliente: str, alias: list[str], voceros: list[str]) -> dict:
    """Escribe (o recalcula) "Contexto analizado" y "texto" en el xlsx limpio, in situ."""
    if not (cliente or '').strip():
        raise ValueError("Falta el cliente: el Contexto analizado se arma con sus menciones.")
    ruta = Path(ruta_xlsx)
    alias = [a.strip() for a in alias or [] if a and a.strip()]
    vocs = expandir_voceros([v for v in voceros or [] if v and v.strip()])
    libro = openpyxl.load_workbook(ruta)
    hoja = libro.active
    encabezado = [c.value for c in hoja[1]]
    while encabezado and encabezado[-1] is None:
        encabezado.pop()
    recalculadas = [n for n in (COL_CONTEXTO, COL_TEXTO) if n in encabezado]
    for nombre in (COL_CONTEXTO, COL_TEXTO):
        if nombre not in encabezado:
            hoja.cell(row=1, column=len(encabezado) + 1, value=nombre)
            encabezado.append(nombre)
    c_ctx, c_txt = encabezado.index(COL_CONTEXTO) + 1, encabezado.index(COL_TEXTO) + 1
    conteo = {'mencion': 0, 'lead': 0, 'titulo': 0}
    filas = 0
    for r in range(2, hoja.max_row + 1):
        valores = [hoja.cell(row=r, column=i + 1).value for i in range(len(encabezado))]
        if all(v in (None, '') for v in valores):
            continue
        row = {str(h): v for h, v in zip(encabezado, valores) if h}
        ctx, txt, origen = contexto_fila(row, cliente, alias, vocs)
        hoja.cell(row=r, column=c_ctx, value=ctx)
        hoja.cell(row=r, column=c_txt, value=txt)
        conteo[origen] += 1
        filas += 1
    libro.save(ruta)
    return {'filas': filas, 'con_mencion': conteo['mencion'], 'desde_lead': conteo['lead'],
            'solo_titulo': conteo['titulo'], 'recalculadas': recalculadas, 'voceros_buscados': vocs}
