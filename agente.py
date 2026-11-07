# clausifier — clasifica noticias (TONO y TEMA) de un xlsx para cualquier cliente
# y compara cómo lo cubren cinco portales colombianos.
#
# Generado por la skill agent-sdk-wizard. Corre con:
#   uv run agente.py --archivo "datos/noticias.xlsx" --cliente "Mi marca" --alias "MM" --voceros "Ana Pérez"
#   uv run agente.py --modo portales --cliente "Mi marca" --alias "MM" --voceros "Ana Pérez"
#   uv run agente.py "pedido libre"                 (o sin nada: te pregunta lo que falte)
# La credencial sale de .env (ver credenciales.py): ant, API key o suscripción.

import asyncio
import json
import re
import shutil
import unicodedata
from pathlib import Path
from typing import Any

from credenciales import preparar_credenciales

AQUI = Path(__file__).resolve().parent         # funciona desde cualquier carpeta (y tus funciones la usan)
WORKSPACE = AQUI / "workspace"
ENTREGAS = WORKSPACE / "entregas"

from claude_agent_sdk import (
    query,
    ClaudeSDKError,
    SystemMessage,
    ClaudeAgentOptions,
    AssistantMessage,
    ResultMessage,
    TextBlock,
    ToolUseBlock,
    ToolResultBlock,
)
from claude_agent_sdk import tool, create_sdk_mcp_server
from claude_agent_sdk import HookMatcher
from claude_agent_sdk import AgentDefinition

from hooks import HOOK  # la función vigilante vive en hooks.py, al lado

import openpyxl

import limpieza  # limpieza y orden deterministas (sin IA) antes de clasificar
import contexto  # Contexto analizado y texto, como API_2_Grill (sin IA)
import exportar  # el xlsx final con el layout exacto del export de Grill


# ─── Paso 5 · tus funciones ───────────────────────────────────────────────
# Claude NO ejecuta esto. Pide la herramienta por su nombre y este programa
# la corre. Los nombres completos que ve el modelo son:
#     mcp__medios__leer_noticias   y   mcp__medios__guardar_clasificacion
# (el prefijo sale de la CLAVE del dict mcp_servers, no del name= de abajo)

def _normal(texto: Any) -> str:
    """minúsculas, sin tildes ni signos: para comparar nombres y títulos."""
    texto = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9 ]+", " ", texto.lower()).split().__str__()


def _ruta(archivo: str) -> Path:
    """El archivo puede venir relativo a workspace/ o como ruta completa."""
    p = Path(archivo)
    if not p.is_absolute():
        p = WORKSPACE / p
        if not p.exists():
            p = WORKSPACE / "datos" / Path(archivo).name
    return p


def _columna(encabezado: list[Any], nombre: str) -> int:
    objetivo = _normal(nombre)
    for i, h in enumerate(encabezado):
        if _normal(h) == objetivo:
            return i
    raise KeyError(f"No existe la columna '{nombre}'. Columnas: {[h for h in encabezado if h]}")


# Nombres conocidos en los distintos formatos de monitoreo (orden = preferencia).
CANDIDATOS_ID = ["ID Noticia", "NoticiaId", "Id Noticia", "ID"]
CANDIDATOS_TITULO = ["Título", "Titular", "Title"]
CANDIDATOS_CUERPO = ["Resumen - Aclaracion", "CuerpoEs", "Cuerpo", "Resumen", "texto", "Contenido"]


def _buscar(encabezado: list[Any], candidatos: list[str]) -> int | None:
    for c in candidatos:
        try:
            return _columna(encabezado, c)
        except KeyError:
            continue
    return None


def detectar_columnas(encabezado: list[Any]) -> list[str]:
    """Título + todas las columnas de cuerpo que existan (p. ej. Resumen - Aclaracion y CuerpoEs)."""
    elegidas = []
    i = _buscar(encabezado, CANDIDATOS_TITULO)
    if i is not None:
        elegidas.append(str(encabezado[i]))
    for c in CANDIDATOS_CUERPO:
        i = _buscar(encabezado, [c])
        if i is not None and str(encabezado[i]) not in elegidas:
            elegidas.append(str(encabezado[i]))
        if len(elegidas) >= 3:                 # título + 2 cuerpos: suficiente señal, pocos tokens
            break
    return elegidas


COL_TONO, COL_SUBTEMA = "Tono_IA", "Subtema_IA"   # columnas de salida (como el export de API_2_Grill)

# Progreso para la app (las tools corren en este mismo proceso y lo actualizan).
PROGRESO: dict[str, Any] = {}


def reiniciar_progreso() -> None:
    PROGRESO.clear()
    PROGRESO.update(paso="Clasificando", total=0, leidas=0, pagina=0, paginas=0, faltan=None,
                    guardadas=0, cambios_revision=0, revisado=False, subtemas_antes=None,
                    subtemas_despues=None)


reiniciar_progreso()
POR_PAGINA =25          # noticias por llamada: con contextos de hasta 1200 car. la página queda < ~15k tokens
TOPE_TEXTO = 1200        # caracteres de Contexto analizado por noticia (Grill lo arma con tope 2000)
TOPE_TITULO = 160        # el título va solo como etiqueta corta, para agrupar e igual_a
COLUMNAS_CONTEXTO = [contexto.COL_CONTEXTO, contexto.COL_TEXTO]   # orden de preferencia


@tool(
    "leer_noticias",
    "Lee un archivo .xlsx de monitoreo de medios y devuelve una fila por noticia con id, medio, "
    "fecha, 'titulo' (etiqueta corta) y 'contexto' (el Contexto analizado: las oraciones literales "
    "donde aparece el cliente). TONO y TEMA se juzgan desde 'contexto'. "
    "'archivo': nombre o ruta del xlsx (p. ej. 'entregas/noticias sec_limpio.xlsx'). "
    f"Devuelve PÁGINAS de {POR_PAGINA} noticias: 'desde' = 0 la primera vez y luego el número "
    "que diga 'siguiente', hasta que 'siguiente' sea null. Las filas con 'igual_a' repiten el "
    "título de otra fila: dales su mismo tema y tono.",
    {"archivo": str, "desde": int},
)
async def leer_noticias(args: dict[str, Any]) -> dict[str, Any]:
    try:
        ruta = _ruta(args["archivo"])
        hoja = openpyxl.load_workbook(ruta, read_only=True, data_only=True).active
        filas = list(hoja.iter_rows(values_only=True))
        encabezado = list(filas[0])
        i_tit = _buscar(encabezado, CANDIDATOS_TITULO)
        # Fuente del contexto: Contexto analizado → texto → columnas de cuerpo detectadas (respaldo).
        fuentes = [i for i in (_buscar(encabezado, [c]) for c in COLUMNAS_CONTEXTO) if i is not None]
        respaldo = [_columna(encabezado, c) for c in detectar_columnas(encabezado)
                    if _buscar(encabezado, [c]) != i_tit and c not in COLUMNAS_CONTEXTO]
        if i_tit is None and not fuentes and not respaldo:
            raise KeyError(f"No encuentro título ni contexto. Columnas: {[h for h in encabezado if h]}")
        i_id = _buscar(encabezado, CANDIDATOS_ID)
        i_medio = next((i for i, h in enumerate(encabezado) if _normal(h) == _normal("Medio")), None)
        i_fecha = next((i for i, h in enumerate(encabezado) if _normal(h) == _normal("Fecha")), None)

        # Archivo limpio (limpieza.py): "Fila" es la clave única y los duplicados
        # no se mandan al modelo (heredan la clasificación de su original al guardar).
        i_fila = _buscar(encabezado, [limpieza.COL_FILA])
        i_dup = _buscar(encabezado, [limpieza.COL_DUP])
        if i_fila is not None:
            i_id = i_fila
        grupos: dict[str, str] = {}
        salida = []
        omitidos = 0
        de_respaldo = 0
        for n, fila in enumerate(filas[1:], start=2):
            if not any(fila):
                continue
            if i_dup is not None and fila[i_dup] not in (None, ""):
                omitidos += 1
                continue
            titulo = str(fila[i_tit] or "") if i_tit is not None else ""
            clave = _normal(titulo)
            fila_id = str(fila[i_id]) if i_id is not None else str(n)
            registro = {
                "id": fila_id,
                "medio": fila[i_medio] if i_medio is not None else None,
                "fecha": str(fila[i_fecha])[:10] if i_fecha is not None else None,
                "titulo": titulo[:TOPE_TITULO],
            }
            if titulo and clave in grupos:
                registro["igual_a"] = grupos[clave]      # mismo título: no repetir el contexto
            else:
                grupos.setdefault(clave, fila_id)
                texto = next((str(fila[i]).strip() for i in fuentes if str(fila[i] or "").strip()), "")
                if not texto:
                    texto = " | ".join(str(fila[i]).strip() for i in respaldo if str(fila[i] or "").strip())
                    de_respaldo += bool(texto)
                registro["contexto"] = texto[:TOPE_TEXTO]
            salida.append(registro)
    except Exception as e:  # noqa: BLE001 — el error vuelve como contexto, no tumba al agente
        return {"content": [{"type": "text", "text": f"Error leyendo el archivo: {e}"}], "is_error": True}

    desde = max(0, int(args.get("desde") or 0))
    pagina = salida[desde:desde + POR_PAGINA]
    siguiente = desde + POR_PAGINA if desde + POR_PAGINA < len(salida) else None
    texto = json.dumps({"archivo": ruta.name, "total": len(salida),
                        "desde": desde, "siguiente": siguiente,
                        "duplicados_omitidos": omitidos, "noticias": pagina},
                       ensure_ascii=False, separators=(",", ":"))
    TAMANOS_PAGINA.append(len(texto))
    PROGRESO.update(total=len(salida), leidas=max(PROGRESO.get("leidas", 0), desde + len(pagina)),
                    pagina=desde // POR_PAGINA + 1, paginas=-(-len(salida) // POR_PAGINA) or 1)
    print(f"  >>> leer_noticias: filas {desde + 1}-{desde + len(pagina)} de {len(salida)}, "
          f"{len(grupos)} títulos distintos{f', {omitidos} duplicados omitidos' if omitidos else ''}"
          f"{f', {de_respaldo} sin contexto (columnas de respaldo)' if de_respaldo else ''}"
          f" · {len(texto):,} car.")
    return {"content": [{"type": "text", "text": texto}]}


TAMANOS_PAGINA: list[int] = []   # caracteres del JSON de cada página (para medir)


def _destino(origen: Path) -> Path:
    return ENTREGAS / f"{origen.stem.removesuffix('_limpio').removesuffix('_trabajo')}_clasificado.xlsx"


def _trabajo(origen: Path) -> Path:
    """Archivo de trabajo con la clave interna "Fila": el _limpio de entregas/ (o una copia
    _trabajo del original si se clasifica sin limpiar). El original nunca se toca."""
    if origen.parent == ENTREGAS and origen.stem.endswith(("_limpio", "_trabajo")):
        return origen
    trabajo = ENTREGAS / f"{origen.stem}_trabajo.xlsx"
    if not trabajo.exists() or not _destino(origen).exists():
        ENTREGAS.mkdir(parents=True, exist_ok=True)
        shutil.copy2(origen, trabajo)
    return trabajo


def _cols_salida(hoja, encabezado: list[Any]) -> tuple[int, int]:
    """Columnas Tono_IA / Subtema_IA (base 1): se reutilizan si existen, si no se agregan."""
    for nombre in (COL_TONO, COL_SUBTEMA):
        if _buscar(encabezado, [nombre]) is None:
            hoja.cell(row=1, column=len(encabezado) + 1, value=nombre)
            encabezado.append(nombre)
    return _buscar(encabezado, [COL_TONO]) + 1, _buscar(encabezado, [COL_SUBTEMA]) + 1


@tool(
    "guardar_clasificacion",
    "Guarda TONO y TEMA (columnas 'Tono_IA' y 'Subtema_IA') y escribe en entregas/ el xlsx final "
    "con el layout de entrega (el original no se toca). "
    "'archivo': el mismo xlsx que leíste. 'filas_json': lista JSON "
    '[{"id":"123","tema":"...","tono":"Positivo|Neutro|Negativo"}, ...]. Puedes llamarla '
    "varias veces por tandas: las filas ya guardadas se conservan y las que repitas se corrigen. "
    "Devuelve cuántas filas faltan por clasificar.",
    {"archivo": str, "filas_json": str},
)
async def guardar_clasificacion(args: dict[str, Any]) -> dict[str, Any]:
    try:
        origen = _ruta(args["archivo"])
        nuevas = {str(f["id"]): f for f in json.loads(args["filas_json"])}
        malos = [i for i, f in nuevas.items() if f.get("tono") not in ("Positivo", "Neutro", "Negativo")]
        if malos:
            raise ValueError(f"tono inválido en ids {malos[:10]}: usa Positivo, Neutro o Negativo")

        destino = _destino(origen)
        nuevo = not destino.exists()
        trabajo = _trabajo(origen)
        libro = openpyxl.load_workbook(trabajo)
        hoja = libro.active
        encabezado = [c.value for c in hoja[1]]
        while encabezado and encabezado[-1] is None:
            encabezado.pop()
        c_tono, c_tema = _cols_salida(hoja, encabezado)
        if nuevo:                              # corrida nueva: se vacían las columnas de salida
            for r in range(2, hoja.max_row + 1):   # OJO: cell(value=None) NO borra; hay que asignar .value
                hoja.cell(row=r, column=c_tono).value = None
                hoja.cell(row=r, column=c_tema).value = None
        c_fila = _buscar(encabezado, [limpieza.COL_FILA])
        c_dup = _buscar(encabezado, [limpieza.COL_DUP])
        c_noticia = _buscar(encabezado, CANDIDATOS_ID)
        c_id = c_fila if c_fila is not None else c_noticia      # archivo limpio: "Fila" es la clave
        c_id = c_id + 1 if c_id is not None else None
        c_men = _buscar(encabezado, ["Menciones - Empresa"])

        def valor(r: int, c: int | None) -> Any:
            return hoja.cell(row=r, column=c + 1).value if c is not None else None

        faltantes = []
        cambios = 0                            # filas ya clasificadas cuyo tono o tema cambió
        duplicadas = []                        # filas que heredan de su original
        originales: dict[tuple[str, str], int] = {}
        for r in range(2, hoja.max_row + 1):
            if all(c.value is None for c in hoja[r]):
                continue
            if c_dup is not None and valor(r, c_dup) not in (None, ""):
                duplicadas.append(r)
                continue
            originales.setdefault((str(valor(r, c_noticia)), _normal(valor(r, c_men))), r)
            originales.setdefault((str(valor(r, c_noticia)), ""), r)
            clave = str(hoja.cell(row=r, column=c_id).value) if c_id else str(r)
            if clave in nuevas:
                antes = (hoja.cell(row=r, column=c_tono).value, hoja.cell(row=r, column=c_tema).value)
                ahora = (nuevas[clave]["tono"], nuevas[clave]["tema"])
                cambios += bool(antes[1]) and antes != ahora
                hoja.cell(row=r, column=c_tono, value=ahora[0])
                hoja.cell(row=r, column=c_tema, value=ahora[1])
            elif not hoja.cell(row=r, column=c_tema).value:
                faltantes.append(clave)
        # Cada duplicado copia Tono / Subtema de su original (mismo ID y misma mención).
        for r in duplicadas:
            dup_id = str(valor(r, c_dup))
            o = originales.get((dup_id, _normal(valor(r, c_men)))) or originales.get((dup_id, ""))
            if o:
                hoja.cell(row=r, column=c_tono, value=hoja.cell(row=o, column=c_tono).value)
                hoja.cell(row=r, column=c_tema, value=hoja.cell(row=o, column=c_tema).value)
        subtemas = len({_normal(hoja.cell(row=r, column=c_tema).value) for r in range(2, hoja.max_row + 1)
                        if hoja.cell(row=r, column=c_tema).value})
        libro.save(trabajo)
        exportar.exportar(trabajo, destino)    # el entregable: 26 columnas exactas, sin "Fila"
    except Exception as e:  # noqa: BLE001
        return {"content": [{"type": "text", "text": f"Error guardando: {e}"}], "is_error": True}

    PROGRESO.update(guardadas=PROGRESO.get("guardadas", 0) + len(nuevas), faltan=len(faltantes),
                    subtemas_despues=subtemas)
    if PROGRESO.get("revisado"):
        PROGRESO["cambios_revision"] = PROGRESO.get("cambios_revision", 0) + cambios
        PROGRESO["paso"] = "Guardando"
    print(f"  >>> guardar_clasificacion: {len(nuevas)} filas ({cambios} corregidas) → "
          f"{destino.relative_to(AQUI)} (faltan {len(faltantes)}, {subtemas} subtemas distintos)")
    if faltantes:
        # Normal mientras se avanza por páginas; al final debe llegar a 0.
        ids = ", ".join(faltantes) if len(faltantes) <= 40 else ", ".join(faltantes[:40]) + ", ..."
        return {"content": [{"type": "text", "text":
                f"Guardado parcial en entregas/{destino.name}. Faltan por clasificar: {len(faltantes)} "
                f"(ids: {ids}). Si quedan páginas, sigue leyendo; si ya leíste todas, clasifica esos "
                "ids y vuelve a llamar."}]}
    return {"content": [{"type": "text", "text":
            f"Guardado en entregas/{destino.name}. Filas recibidas: {len(nuevas)}. Faltan por clasificar: 0."}]}


# ─── Revisión de consistencia (determinista, sin IA) ──────────────────────
UMBRAL_TITULO = 0.92      # como Grill: títulos casi iguales
UMBRAL_CONTEXTO = 0.85    # como Grill: resúmenes casi iguales
UMBRAL_SUBTEMA = 0.90     # subtemas con redacción casi idéntica
UMBRAL_SOBRE = 0.35       # mismo subtema pero contextos muy distintos: posible sobreagrupación
MIN_CTX = 120             # contextos más cortos ("Fuente: Alcaldía...") no sirven para comparar
GRUPOS_POR_PAGINA = 30
TOPE_CTX_GRUPO = 300

try:
    from rapidfuzz.fuzz import ratio as _rf_ratio

    def _similitud(a: str, b: str) -> float:
        return _rf_ratio(a, b) / 100.0
except ImportError:  # respaldo sin dependencia
    from difflib import SequenceMatcher

    def _similitud(a: str, b: str) -> float:
        return SequenceMatcher(None, a, b).ratio()


def _simple(texto: Any) -> str:
    t = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return " ".join(re.sub(r"[^a-z0-9]+", " ", t.lower()).split())


def construir_candidatos(ruta: Path) -> list[dict]:
    """Grupos de filas que PODRÍAN ser el mismo hecho y no tienen el mismo tono/subtema,
    más pares con el mismo subtema y contextos muy distintos (posible sobreagrupación)."""
    hoja = openpyxl.load_workbook(ruta, read_only=True, data_only=True).active
    filas = list(hoja.iter_rows(values_only=True))
    enc = list(filas[0])
    i_id = _buscar(enc, [limpieza.COL_FILA]) if _buscar(enc, [limpieza.COL_FILA]) is not None \
        else _buscar(enc, CANDIDATOS_ID)
    i_dup, i_tit = _buscar(enc, [limpieza.COL_DUP]), _buscar(enc, CANDIDATOS_TITULO)
    i_ctx = next((i for i in (_buscar(enc, [c]) for c in COLUMNAS_CONTEXTO) if i is not None), None)
    i_tono, i_tema = _buscar(enc, [COL_TONO]), _buscar(enc, [COL_SUBTEMA])
    regs = []
    for n, f in enumerate(filas[1:], start=2):
        if not any(f) or (i_dup is not None and f[i_dup] not in (None, "")):
            continue
        g = lambda i: f[i] if i is not None and i < len(f) else None  # noqa: E731
        regs.append({"id": str(g(i_id) if i_id is not None else n), "titulo": str(g(i_tit) or ""),
                     "contexto": str(g(i_ctx) or ""), "tono": g(i_tono), "subtema": g(i_tema) or ""})
    for r in regs:
        r["_t"], r["_c"], r["_s"] = _simple(r["titulo"]), _simple(r["contexto"])[:600], _simple(r["subtema"])

    padre = list(range(len(regs)))

    def raiz(x: int) -> int:
        while padre[x] != x:
            padre[x] = padre[padre[x]]
            x = padre[x]
        return x

    motivos: dict[tuple[int, int], str] = {}
    sobre: list[tuple[int, int, float]] = []
    for a in range(len(regs)):
        for b in range(a + 1, len(regs)):
            ra, rb = regs[a], regs[b]
            motivo = None
            if ra["_t"] and ra["_t"] == rb["_t"]:
                motivo = "mismo título"
            elif ra["_t"] and rb["_t"] and _similitud(ra["_t"], rb["_t"]) >= UMBRAL_TITULO:
                motivo = "título casi igual"
            elif min(len(ra["_c"]), len(rb["_c"])) >= MIN_CTX and _similitud(ra["_c"], rb["_c"]) >= UMBRAL_CONTEXTO:
                motivo = "contexto casi igual"
            elif ra["_s"] and rb["_s"] and ra["_s"] != rb["_s"] and _similitud(ra["_s"], rb["_s"]) >= UMBRAL_SUBTEMA:
                motivo = "subtema con redacción casi idéntica"
            if motivo:
                motivos[(a, b)] = motivo
                padre[raiz(a)] = raiz(b)
            elif ra["_s"] and ra["_s"] == rb["_s"] and min(len(ra["_c"]), len(rb["_c"])) >= MIN_CTX:
                sim = _similitud(ra["_c"], rb["_c"])
                if sim < UMBRAL_SOBRE:
                    sobre.append((a, b, sim))

    def fila(i: int) -> dict:
        r = regs[i]
        return {"id": r["id"], "tono": r["tono"], "subtema": r["subtema"], "titulo": r["titulo"][:TOPE_TITULO],
                "contexto": r["contexto"][:TOPE_CTX_GRUPO]}

    grupos: dict[int, list[int]] = {}
    for i in range(len(regs)):
        grupos.setdefault(raiz(i), []).append(i)
    salida = []
    for miembros in grupos.values():
        if len(miembros) < 2:
            continue
        if len({(regs[i]["tono"], regs[i]["_s"]) for i in miembros}) == 1:
            continue                           # ya consistentes: no hace falta revisarlos
        razones = sorted({m for (a, b), m in motivos.items() if a in miembros})
        salida.append({"tipo": "posible_mismo_hecho", "motivo": ", ".join(razones),
                       "filas": [fila(i) for i in miembros]})
    # Sobreagrupación: un par por subtema (el de contextos más distintos).
    por_subtema: dict[str, tuple[int, int, float]] = {}
    for a, b, sim in sobre:
        s = regs[a]["_s"]
        if s not in por_subtema or sim < por_subtema[s][2]:
            por_subtema[s] = (a, b, sim)
    for a, b, sim in por_subtema.values():
        salida.append({"tipo": "posible_sobreagrupacion",
                       "motivo": f"mismo subtema, contextos poco parecidos ({sim:.2f})",
                       "filas": [fila(a), fila(b)]})
    for n, g in enumerate(salida):
        g["grupo"] = n
    return salida


@tool(
    "candidatos_consistencia",
    "Revisión de consistencia (sin IA). Después de guardar TODAS las páginas, lee la clasificación "
    "y devuelve grupos de filas que PODRÍAN ser el mismo hecho pero tienen distinto tono o subtema "
    "(tipo 'posible_mismo_hecho'), y pares con el mismo subtema pero contextos muy distintos "
    "(tipo 'posible_sobreagrupacion'). Cada fila trae id, tono, subtema, titulo y contexto (300 car.). "
    f"'archivo': el mismo xlsx que leíste. 'desde': 0 y luego 'siguiente' ({GRUPOS_POR_PAGINA} grupos "
    "por llamada) hasta null. Devuelve 'total' de grupos.",
    {"archivo": str, "desde": int},
)
async def candidatos_consistencia(args: dict[str, Any]) -> dict[str, Any]:
    try:
        origen = _ruta(args["archivo"])
        ruta = _trabajo(origen) if _destino(origen).exists() else origen
        grupos = construir_candidatos(ruta)
    except Exception as e:  # noqa: BLE001
        return {"content": [{"type": "text", "text": f"Error buscando candidatos: {e}"}], "is_error": True}
    desde = max(0, int(args.get("desde") or 0))
    pagina = grupos[desde:desde + GRUPOS_POR_PAGINA]
    siguiente = desde + GRUPOS_POR_PAGINA if desde + GRUPOS_POR_PAGINA < len(grupos) else None
    if not PROGRESO.get("revisado"):
        PROGRESO["subtemas_antes"] = PROGRESO.get("subtemas_despues")
    PROGRESO.update(paso="Revisando consistencia", revisado=True, grupos=len(grupos))
    print(f"  >>> candidatos_consistencia: grupos {desde + 1}-{desde + len(pagina)} de {len(grupos)}")
    return {"content": [{"type": "text", "text": json.dumps(
        {"total": len(grupos), "desde": desde, "siguiente": siguiente, "grupos": pagina},
        ensure_ascii=False, separators=(",", ":"))}]}


medios = create_sdk_mcp_server(
    name="medios", version="1.0.0", tools=[leer_noticias, guardar_clasificacion, candidatos_consistencia]
)


# ─── Paso 4 · el system prompt ────────────────────────────────────────────

SYSTEM_PROMPT = """<rol>
Eres un analista de monitoreo de medios que clasifica noticias en español de Colombia para una agencia de monitoreo. Tu trabajo lo usa un equipo que entrega informes a marcas e instituciones, así que la consistencia entre filas importa tanto como acertar cada una.
</rol>

<contexto>
Trabajas dentro de la carpeta workspace/. Los archivos de entrada están en datos/ (por ejemplo datos/noticias sec.xlsx). Lo que entregas va en entregas/. Nunca modificas el archivo original.
</contexto>

<objetivo>
Tienes dos modos. Elige según lo que pida el usuario:
1. Clasificar un xlsx: el usuario te da el archivo, el nombre del CLIENTE (marca) y, opcionalmente, ALIAS y VOCEROS. Antes de llegarte, el archivo ya viene limpio y con la columna "Contexto analizado" (las oraciones literales donde aparece el cliente). A cada fila le asignas TEMA y TONO y los guardas en una copia del archivo.
2. Comparar portales: el usuario te da un CLIENTE (y quizá ALIAS y VOCEROS) y quiere ver cómo lo cubren cinco portales colombianos con las noticias más recientes. Para eso usa la skill comparar-portales.
Si falta el cliente o el archivo (en el modo 1), pídelos en una frase y termina sin llamar herramientas: adivinar el cliente cambia todo el tono.
Sirves para cualquier cliente: no supongas ninguno que el usuario no haya nombrado.
</objetivo>

<cliente_alias_voceros>
- ALIAS son siglas o nombres alternos del cliente (por ejemplo "SED" por "Secretaría de Educación"). Una mención de cualquier alias cuenta como mención del cliente.
- VOCEROS son las personas que hablan por la marca (su director, secretario, gerente). Una mención de un vocero, o lo que él dice o se dice de él en su rol, cuenta como mención del cliente y pesa en el TONO.
- Si una fila no menciona al cliente, a sus alias ni a sus voceros, clasifícala igual: TONO Neutro, salvo que la noticia afecte claramente al cliente (por ejemplo, un hecho en el sector que el cliente gestiona).
</cliente_alias_voceros>

<como_trabajas>
TONO y TEMA se juzgan desde el "contexto" de cada noticia (la columna Contexto analizado), no desde el título: el título es solo una etiqueta para agrupar.

TEMA: una frase específica de 4 a 7 palabras que defina realmente la noticia y la distinga de otras. Evita temas genéricos como "Seguridad en Bogotá" o "Gestión de la alcaldía": un lector debe saber de qué hecho se trata solo con el tema.

TONO: Positivo, Neutro o Negativo, evaluado solo desde la perspectiva del cliente, no del tono general de la noticia. Una tragedia puede ser Positiva para el cliente si muestra su gestión a favor.
- Positivo: gestiones a favor, logros, reconocimientos, inversiones u obras del cliente.
- Negativo: críticas, denuncias, sanciones o afectación directa a la marca.
- Neutro: menciones generales o informativas sin efecto claro para el cliente, o noticias donde el cliente no aparece.
Si el pedido trae un bloque <criterio_tono>, ese criterio del cliente MANDA: aplícalo a cada fila por encima de estas definiciones generales, y pásaselo completo y literal al ayudante revisor-subtemas cuando lo lances.

Consistencia, en este orden:
1. Lee TODAS las filas antes de asignar nada.
2. Agrupa las noticias con títulos o contextos iguales o similares ("igual_a" ya junta los títulos idénticos; tú juntas además las que cuentan el mismo hecho con otra redacción).
3. A cada grupo dale exactamente el mismo TEMA y el mismo TONO. Reutiliza un tema existente cuando el hecho sea el mismo.
4. Antes de guardar, repasa la lista de temas y fusiona los casi duplicados con redacción distinta.
5. Revisión de consistencia, al final (solo modo xlsx): cuando guardar_clasificacion diga "Faltan por clasificar: 0", llama candidatos_consistencia con desde=0 para ver cuántos grupos hay ("total"). Si total es 0, terminas. Si no, lanza el ayudante revisor-subtemas con Agent pasándole el archivo y el rango de grupos (desde y cuántos); si hay más de 30 grupos, lanza dos en paralelo, cada uno con su mitad. No le copies los grupos: él los lee. Con sus correcciones llama UNA vez guardar_clasificacion con la lista [{id, tema, tono}] corregida (solo las filas que cambian). No repitas la revisión.
</como_trabajas>

<herramientas>
- mcp__medios__leer_noticias: pasa siempre archivo y desde. Cada noticia trae id, medio, fecha, "titulo" (etiqueta corta) y "contexto" (Contexto analizado): juzga TONO y TEMA con "contexto". Devuelve páginas de noticias: empieza con desde=0, clasifica esa página, guárdala con guardar_clasificacion y pide la siguiente con el número de "siguiente", hasta que sea null. Lleva tu lista de temas ya usados de una página a otra y reutilízalos cuando el hecho sea el mismo. Las filas con "igual_a" tienen el mismo título que la fila indicada: mismo tema y tono. Si el archivo viene limpio (entregas/..._limpio.xlsx), el id es la columna "Fila" y los duplicados ya vienen omitidos: no los pidas ni los clasifiques, heredan la clasificación de su original al guardar.
- mcp__medios__guardar_clasificacion: pásale el archivo y la lista JSON de {id, tema, tono} de la página que acabas de clasificar. Mientras queden páginas es normal que diga que faltan filas. Al terminar todas las páginas debe decir "Faltan por clasificar: 0"; si no, clasifica los ids que nombra y vuelve a llamarla. El resultado queda en las columnas "Tono_IA" y "Subtema_IA" del xlsx final.
- mcp__medios__candidatos_consistencia: solo después de guardar todas las páginas (paso 5). Devuelve grupos candidatos, sin decidir nada.
- Agent (ayudante revisor-subtemas): solo en el modo xlsx, en el paso 5. Devuelve JSON con correcciones; tú las guardas.
- Agent (ayudante analista-portal): solo en el modo comparar portales, siguiendo la skill; lanza los cinco en paralelo, uno por portal. Nunca en el modo xlsx.
- WebSearch y WebFetch: solo en el modo comparar portales, si hace falta completar lo que trajo un ayudante.
- Write: solo para escribir dentro de entregas/ (por ejemplo entregas/comparativo.md).
No tienes terminal ni puedes editar archivos existentes.
</herramientas>

<limites>
- Nunca inventes una noticia, una fecha, un medio ni un enlace. Si algo no está en el archivo o en la web, dilo.
- Nunca cambies el xlsx original: la copia clasificada la escribe guardar_clasificacion en entregas/.
- No clasifiques el tono por lo que opinas del cliente, solo por cómo lo deja la noticia.
</limites>

<formato_de_salida>
Al terminar, responde en español, en máximo 6 viñetas: archivo generado, número de filas clasificadas, cuántas por tono, los 5 temas más frecuentes y cualquier fila dudosa. Cierra con UNA frase de conclusión. En el modo portales, resume el comparativo y da la ruta de entregas/comparativo.md.
Paras cuando la herramienta diga que no faltan filas (modo 1) o cuando el comparativo esté escrito (modo 2).
</formato_de_salida>"""


# ─── Las nueve decisiones, en un solo bloque ──────────────────────────────
# Cambiar el agente = cambiar una línea de aquí. La tabla paso → línea está
# en README.md, y el porqué de cada una en decisiones.md.

opciones = ClaudeAgentOptions(
    model="claude-haiku-5-5",                 # paso 3: rápido y barato (sube a "claude-sonnet-5-5" si el tono sale flojo)
    system_prompt=SYSTEM_PROMPT,              # paso 4: quién es, qué logra, reglas y cuándo parar
    allowed_tools=[                           # paso 5: qué puede hacer en el mundo
        "mcp__medios__leer_noticias",
        "mcp__medios__guardar_clasificacion",
        "mcp__medios__candidatos_consistencia",
        "WebSearch",
        "WebFetch",
        "Write",
        "Skill",
        "Agent",                              # paso 7: auto-aprueba lanzar a los ayudantes
    ],
    mcp_servers={"medios": medios},           # paso 5: tus funciones (sin MCP externo, paso 5-bis)
    setting_sources=["project"],              # paso 6: descubre workspace/.claude/skills/
    skills=["comparar-portales"],             # paso 6: el procedimiento que le enseñaste
    agents={                                  # paso 7: varios en paralelo (solo modo portales)
        "analista-portal": AgentDefinition(
            # description = CUÁNDO usarlo. Es lo que el modelo lee para decidir.
            description=(
                "Busca las noticias más recientes de un cliente en UN portal colombiano y devuelve "
                "título, fecha, URL, tema y tono de cada nota. Úsalo en el modo comparar portales: "
                "lanza uno por portal, los cinco en paralelo."
            ),
            # prompt = el system prompt del ayudante: su rol y sus límites.
            prompt=(
                "Eres analista de un solo portal de noticias colombiano. Recibes CLIENTE, ALIAS, "
                "VOCEROS y UN dominio de portal. Busca con WebSearch (cliente, alias y voceros, con "
                "site:<dominio>) las 2-3 notas más recientes que hablen del cliente; abre con WebFetch "
                "como máximo una, la más reciente. Una mención de un alias o un vocero cuenta como "
                "mención del cliente. Devuelve máximo 10 líneas: por nota, título | fecha | URL | "
                "TEMA (4-7 palabras, específico) | TONO para el cliente (Positivo, Neutro o "
                "Negativo); al final una línea con el enfoque del portal. Nunca inventes notas, "
                "fechas ni URL: si no encuentras nada, responde 'sin notas recientes'. No escribes "
                "archivos."
            ),
            # tools = lista blanca. Lo que no está aquí, no existe para él.
            tools=["WebSearch", "WebFetch"],
            model="haiku",                    # trabajo de búsqueda: modelo barato
            maxTurns=8,                       # OJO: camelCase también en Python
        ),
        "revisor-subtemas": AgentDefinition(
            description=(
                "Revisa la consistencia de Tono_IA y Subtema_IA en el modo xlsx: decide si cada grupo "
                "candidato es el MISMO hecho y devuelve correcciones en JSON. Úsalo una vez al final, "
                "después de guardar todas las páginas; pásale el archivo y el rango de grupos."
            ),
            prompt=(
                "Eres revisor de consistencia de una clasificación de noticias en español. Recibes un "
                "ARCHIVO y un rango de grupos (desde y cuántos). Lee los grupos con "
                "mcp__medios__candidatos_consistencia (archivo, desde), solo los de tu rango.\n"
                "Para cada grupo decide si sus filas cuentan el MISMO hecho: el mismo evento, la misma "
                "entidad y el mismo momento. Mismo tema general o mismo medio NO es el mismo hecho.\n"
                "- Si es el mismo hecho: todas las filas llevan un solo subtema (reutiliza el más "
                "específico de los que ya tienen) y un solo tono (el que mejor sostiene el contexto).\n"
                "- Si no lo es: déjalas separadas y haz cada subtema lo bastante específico (4 a 7 "
                "palabras) para distinguirlas. Los grupos 'posible_sobreagrupacion' casi siempre son "
                "hechos distintos que comparten un subtema demasiado genérico.\n"
                "Nunca fusiones solo porque los subtemas comparten palabras. No agrupes por agrupar.\n"
                "Tono solo Positivo, Neutro o Negativo, para el cliente. Si te pasan un CRITERIO DEL "
                "TONO, decide el tono con ese criterio. No escribes archivos.\n"
                "Responde SOLO con JSON compacto, sin texto alrededor: "
                '{"correcciones":[{"id":"...","tema":"...","tono":"..."}],'
                '"grupos":[{"grupo":0,"mismo_hecho":true,"razon":"una línea"}]}. '
                "En correcciones incluye solo las filas cuyo tema o tono cambia."
            ),
            tools=["mcp__medios__candidatos_consistencia"],
            model="haiku",
            maxTurns=6,
        ),
    },
    forward_subagent_text=True,               # paso 7: ver también el TEXTO del ayudante
    permission_mode="default",                # paso 8: aprueba solo lo de allowed_tools
    disallowed_tools=["Bash", "Edit", "Read", "Grep", "Glob"],   # paso 8: nada de leer volcados por fuera de las tools
    hooks={"PreToolUse": [HookMatcher(matcher="Write|Edit", hooks=[HOOK])]},   # paso 9: escribe solo en entregas/
    cwd=WORKSPACE,                            # la carpeta que el agente ve (junto a este archivo)
    max_turns=70,                             # freno: páginas × (leer + guardar) + revisión de consistencia
    max_budget_usd=0.60,                      # freno: clasificar + revisión (modo portales: TOPE_PORTALES)
    stderr=lambda linea: None,                # cambia a print(linea, end="") para depurar
)


# ─── El lector del stream ─────────────────────────────────────────────────

def imprimir(mensaje: Any) -> None:
    sangria = "   ↳ " if getattr(mensaje, "parent_tool_use_id", None) else ""

    if isinstance(mensaje, SystemMessage) and mensaje.subtype == "init":
        print(f"[init] skill «comparar-portales» cargada: {'comparar-portales' in mensaje.data.get('skills', [])}")
        print(f"[init] servidores: {[(s.get('name'), s.get('status')) for s in mensaje.data.get('mcp_servers', [])]}")
        return

    if isinstance(mensaje, AssistantMessage):
        for bloque in mensaje.content:
            if isinstance(bloque, ToolUseBlock):
                ruta = bloque.input.get("file_path") or bloque.input.get("path")
                if ruta:
                    ruta = str(ruta).removeprefix(str(opciones.cwd)).lstrip("/\\")   # relativa (Mac y Windows)
                arg = (
                    bloque.input.get("skill")
                    or bloque.input.get("command")
                    or ruta
                    or bloque.input.get("url")
                    or bloque.input.get("query")
                    or bloque.input.get("pattern")
                    or bloque.input.get("prompt")
                    or ", ".join(f"{k}={v}" for k, v in bloque.input.items())
                )
                print(f"{sangria}🔧 {bloque.name}({str(arg)[:80]})")
            elif isinstance(bloque, TextBlock) and bloque.text.strip():
                print(f"{sangria}{bloque.text.strip()}")

    elif isinstance(mensaje, ResultMessage):
        if mensaje.subtype == "success" and not mensaje.is_error:
            print(f"\n{'─' * 60}")
        else:
            FALLO[0] = True
            print(f"\n⚠️  terminó por: {mensaje.subtype}")
        costo = mensaje.total_cost_usd or 0.0
        segundos = (mensaje.duration_ms or 0) / 1000
        print(f"{mensaje.num_turns} turnos · ${costo:.4f} · {segundos:.0f}s")


def imprimir_error_de_tool(mensaje: Any) -> None:
    for bloque in getattr(mensaje, "content", []) or []:
        if isinstance(bloque, ToolResultBlock) and bloque.is_error:
            print(f"   ❌ la herramienta falló: {str(bloque.content)[:200]}")


PROMPT = ""
FALLO = [False]
TOPE_PORTALES = 0.60                          # 5 ayudantes en paralelo: cada uno paga sus tokens
TOPE_CLASIFICAR = 0.60                        # clasificar + revisor-subtemas


def exportar_limpio(limpio: Path) -> Path:
    """Solo limpiar: el mismo layout de 26 columnas, con Tono_IA/Subtema_IA vacías."""
    destino = limpio.with_name(f"{limpio.stem}_export.xlsx")
    exportar.exportar(limpio, destino, con_ia=False)
    return destino


def armar_prompt(a: Any) -> str:
    """Convierte los flags en un pedido explícito para el agente."""
    extra = []
    if a.alias:
        extra.append(f"ALIAS: {a.alias}")
    if a.voceros:
        extra.append(f"VOCEROS: {a.voceros}")
    extra_txt = ("\n" + "\n".join(extra)) if extra else ""
    if a.modo == "portales":
        return (f"Modo comparar portales. CLIENTE: {a.cliente}{extra_txt}\n"
                "Busca las noticias más recientes del cliente (usa cliente, alias y voceros en las "
                "búsquedas) en cinco portales colombianos y escribe entregas/comparativo.md.")
    return (f"Modo clasificar xlsx.\nARCHIVO: {a.archivo}\nCLIENTE: {a.cliente}{extra_txt}\n"
            f"{bloque_criterio(a)}"
            "Clasifica todas las filas con TONO y TEMA y guárdalas con guardar_clasificacion.")


def bloque_criterio(a: Any) -> str:
    """Criterio del tono elegido en la app (radio de Grill) o texto propio del cliente (tiene prioridad)."""
    from criterios import CRITERIOS_TONO
    propio = (getattr(a, "criterio_texto", "") or "").strip()
    nombre = getattr(a, "criterio", None)
    texto = propio or CRITERIOS_TONO.get(nombre or "", "")
    if not texto:
        return ""
    etiqueta = "personalizado del cliente" if propio else nombre
    return f"<criterio_tono nombre=\"{etiqueta}\">\n{texto}\n</criterio_tono>\n"


def separar(texto: str | None) -> list[str]:
    return [x.strip() for x in (texto or "").split(",") if x.strip()]


def limpiar_antes(archivo: str, cliente: str, alias: str = "", voceros: str = "") -> str:
    """Limpia, ordena y agrega Contexto analizado (sin IA, $0) y devuelve la ruta del limpio para el agente."""
    origen = _ruta(archivo)
    destino = ENTREGAS / f"{origen.stem}_limpio.xlsx"
    (ENTREGAS / f"{origen.stem}_clasificado.xlsx").unlink(missing_ok=True)   # no mezclar con corridas viejas
    r = limpieza.limpiar_dossier(origen, destino)
    print(f"🧹 limpieza: {r['total']} filas ({r['filas_expandidas']} por menciones), "
          f"{r['unicos']} únicas, {r['duplicados']} duplicadas, "
          f"{len(r['medios_sin_region'])} medios sin región → {destino.relative_to(AQUI)}")
    c = contexto.agregar_contexto(destino, cliente, separar(alias), separar(voceros))
    print(f"🔎 contexto: {c['con_mencion']} con mención, {c['desde_lead']} desde el lead, "
          f"{c['solo_titulo']} solo título")
    return f"entregas/{destino.name}"


def leer_argumentos() -> str:
    import argparse
    p = argparse.ArgumentParser(description="clausifier: clasifica noticias y compara portales")
    p.add_argument("pedido", nargs="*", help="pedido libre (opcional)")
    p.add_argument("--modo", choices=["xlsx", "portales"], default="xlsx")
    p.add_argument("--cliente", help="nombre del cliente (marca), obligatorio")
    p.add_argument("--alias", default="", help="siglas o nombres alternos, separados por coma")
    p.add_argument("--voceros", default="", help="personas que hablan por la marca, separadas por coma")
    p.add_argument("--archivo", help="xlsx relativo a workspace/, p. ej. datos/noticias.xlsx")
    p.add_argument("--criterio", default=None,
                   help="criterio del tono: 'aspectual' o 'sector' (como el radio de la app)")
    p.add_argument("--criterio-texto", dest="criterio_texto", default="",
                   help="criterio del tono propio del cliente (tiene prioridad sobre --criterio)")
    p.add_argument("--limpiar", dest="limpiar", action="store_true", default=True,
                   help="limpia y ordena el xlsx antes de clasificar (por defecto)")
    p.add_argument("--sin-limpiar", dest="limpiar", action="store_false",
                   help="clasifica el xlsx tal cual, sin limpieza previa")
    a = p.parse_args()
    if a.criterio:
        from criterios import CRITERIOS_TONO
        a.criterio = next((k for k in CRITERIOS_TONO if a.criterio.lower() in k.lower()), a.criterio)

    usa_flags = any([a.cliente, a.archivo, a.alias, a.voceros, a.modo == "portales"])
    if a.modo == "portales" or (a.pedido and "portal" in " ".join(a.pedido).lower()):
        opciones.max_budget_usd = TOPE_PORTALES
    if a.pedido and not usa_flags:
        return " ".join(a.pedido)             # pedido libre: el agente pide lo que falte
    if not a.cliente:
        a.cliente = input("Cliente (marca): ").strip()
        if not a.alias:
            a.alias = input("Alias (opcional, separados por coma): ").strip()
        if not a.voceros:
            a.voceros = input("Voceros (opcional, separados por coma): ").strip()
    if a.modo == "xlsx" and not a.archivo:
        a.archivo = input("Archivo xlsx (relativo a workspace/, p. ej. datos/noticias.xlsx): ").strip()
    if a.modo == "xlsx" and a.limpiar:
        a.archivo = limpiar_antes(a.archivo, a.cliente, a.alias, a.voceros)
    prompt = armar_prompt(a)
    if a.pedido:
        prompt += "\nIndicaciones adicionales: " + " ".join(a.pedido)
    return prompt


async def main() -> None:
    print("▶ clausifier\n")
    try:
        async for mensaje in query(prompt=PROMPT, options=opciones):
            imprimir(mensaje)
            imprimir_error_de_tool(mensaje)
    except ClaudeSDKError as error:
        print(f"\n⚠️  el agente paró: {error}")
        FALLO[0] = True


if __name__ == "__main__":
    import sys
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")   # emojis y tildes en la consola de Windows
    PROMPT = leer_argumentos()                # flags, pedido libre o preguntas en consola
    preparar_credenciales()                   # ANTES de crear la sesión
    asyncio.run(main())
    sys.exit(1 if FALLO[0] else 0)
