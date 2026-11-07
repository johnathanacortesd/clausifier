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


# ─── Paso 5 · tus funciones ───────────────────────────────────────────────
# Claude NO ejecuta esto. Pide la herramienta por su nombre y este programa
# la corre. Los nombres completos que ve el modelo son:
#     mcp__medios__leer_noticias   y   mcp__medios__guardar_clasificacion
# (el prefijo sale de la CLAVE del dict mcp_servers, no del name= de abajo)

COLUMNAS_POR_DEFECTO = ["Título", "Resumen - Aclaracion"]


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


@tool(
    "leer_noticias",
    "Lee un archivo .xlsx de monitoreo de medios y devuelve una fila por noticia con id, "
    "medio, fecha, el texto de las columnas pedidas y 'grupo' (noticias con el mismo título "
    "comparten grupo). 'archivo': nombre o ruta del xlsx (p. ej. 'datos/noticias sec.xlsx'). "
    "'columnas': nombres separados por ',' o '|'; vacío o inexistentes = detecta solas título y "
    "cuerpo (Título, Resumen - Aclaracion, CuerpoEs...). La respuesta dice qué columnas usó.",
    {"archivo": str, "columnas": str},
)
async def leer_noticias(args: dict[str, Any]) -> dict[str, Any]:
    try:
        ruta = _ruta(args["archivo"])
        columnas = [c.strip() for c in re.split(r"[|,]", args.get("columnas") or "") if c.strip()]
        hoja = openpyxl.load_workbook(ruta, read_only=True, data_only=True).active
        filas = list(hoja.iter_rows(values_only=True))
        encabezado = list(filas[0])
        # Usa las columnas pedidas que existan; si ninguna existe, detecta título y cuerpo solas.
        idx = {}
        for c in columnas:
            i = _buscar(encabezado, [c])
            if i is not None:
                idx[str(encabezado[i])] = i
        if not idx:
            idx = {c: _columna(encabezado, c) for c in detectar_columnas(encabezado)}
        if not idx:
            raise KeyError(f"No encuentro columnas de título ni de cuerpo. Columnas: {[h for h in encabezado if h]}")
        columnas = list(idx)
        i_id = _buscar(encabezado, CANDIDATOS_ID)
        i_medio = next((i for i, h in enumerate(encabezado) if _normal(h) == _normal("Medio")), None)
        i_fecha = next((i for i, h in enumerate(encabezado) if _normal(h) == _normal("Fecha")), None)

        # Archivo limpio (limpieza.py): "Fila" es la clave única y los duplicados
        # no se mandan al modelo (heredan la clasificación de su original al guardar).
        i_fila = _buscar(encabezado, [limpieza.COL_FILA])
        i_dup = _buscar(encabezado, [limpieza.COL_DUP])
        if i_fila is not None:
            i_id = i_fila
        grupos: dict[str, int] = {}
        salida = []
        omitidos = 0
        for n, fila in enumerate(filas[1:], start=2):
            if not any(fila):
                continue
            if i_dup is not None and fila[i_dup] not in (None, ""):
                omitidos += 1
                continue
            clave = _normal(fila[idx[columnas[0]]])
            grupos.setdefault(clave, len(grupos) + 1)
            registro = {
                "id": str(fila[i_id]) if i_id is not None else str(n),
                "grupo": grupos[clave],
                "medio": fila[i_medio] if i_medio is not None else None,
                "fecha": str(fila[i_fecha])[:10] if i_fecha is not None else None,
            }
            for c, i in idx.items():
                registro[c] = str(fila[i] or "")[:350]   # señal, no volcados
            salida.append(registro)
    except Exception as e:  # noqa: BLE001 — el error vuelve como contexto, no tumba al agente
        return {"content": [{"type": "text", "text": f"Error leyendo el archivo: {e}"}], "is_error": True}

    print(f"  >>> leer_noticias: {len(salida)} noticias, {len(grupos)} títulos distintos"
          f"{f', {omitidos} duplicados omitidos' if omitidos else ''}")
    texto = json.dumps({"archivo": ruta.name, "columnas_usadas": columnas, "total": len(salida),
                        "duplicados_omitidos": omitidos, "noticias": salida},
                       ensure_ascii=False, separators=(",", ":"))
    return {"content": [{"type": "text", "text": texto}]}


@tool(
    "guardar_clasificacion",
    "Guarda TONO y TEMA (columnas 'Tono C' y 'Tema C', al final) en una COPIA exacta del xlsx "
    "dentro de entregas/ (el original no se toca). "
    "'archivo': el mismo xlsx que leíste. 'filas_json': lista JSON "
    '[{"id":"123","tema":"...","tono":"Positivo|Neutro|Negativo"}, ...]. Puedes llamarla '
    "varias veces por tandas: las filas ya guardadas se conservan. Devuelve cuántas filas "
    "faltan por clasificar.",
    {"archivo": str, "filas_json": str},
)
async def guardar_clasificacion(args: dict[str, Any]) -> dict[str, Any]:
    try:
        origen = _ruta(args["archivo"])
        nuevas = {str(f["id"]): f for f in json.loads(args["filas_json"])}
        malos = [i for i, f in nuevas.items() if f.get("tono") not in ("Positivo", "Neutro", "Negativo")]
        if malos:
            raise ValueError(f"tono inválido en ids {malos[:10]}: usa Positivo, Neutro o Negativo")

        ENTREGAS.mkdir(parents=True, exist_ok=True)
        destino = ENTREGAS / f"{origen.stem.removesuffix('_limpio')}_clasificado.xlsx"
        if not destino.exists():
            shutil.copy2(origen, destino)     # copia exacta: hojas, formatos y orden intactos
        libro = openpyxl.load_workbook(destino)
        hoja = libro.active
        encabezado = [c.value for c in hoja[1]]
        while encabezado and encabezado[-1] is None:
            encabezado.pop()
        for nombre in ("Tono C", "Tema C"):   # dos columnas nuevas al FINAL, en este orden
            if nombre not in encabezado:
                hoja.cell(row=1, column=len(encabezado) + 1, value=nombre)
                encabezado.append(nombre)
        c_tono, c_tema = encabezado.index("Tono C") + 1, encabezado.index("Tema C") + 1
        c_fila = _buscar(encabezado, [limpieza.COL_FILA])
        c_dup = _buscar(encabezado, [limpieza.COL_DUP])
        c_noticia = _buscar(encabezado, CANDIDATOS_ID)
        c_id = c_fila if c_fila is not None else c_noticia      # archivo limpio: "Fila" es la clave
        c_id = c_id + 1 if c_id is not None else None
        c_men = _buscar(encabezado, ["Menciones - Empresa"])

        def valor(r: int, c: int | None) -> Any:
            return hoja.cell(row=r, column=c + 1).value if c is not None else None

        faltantes = []
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
                hoja.cell(row=r, column=c_tono, value=nuevas[clave]["tono"])
                hoja.cell(row=r, column=c_tema, value=nuevas[clave]["tema"])
            elif not hoja.cell(row=r, column=c_tema).value:
                faltantes.append(clave)
        # Cada duplicado copia Tono C / Tema C de su original (mismo ID y misma mención).
        for r in duplicadas:
            dup_id = str(valor(r, c_dup))
            o = originales.get((dup_id, _normal(valor(r, c_men)))) or originales.get((dup_id, ""))
            if o:
                hoja.cell(row=r, column=c_tono, value=hoja.cell(row=o, column=c_tono).value)
                hoja.cell(row=r, column=c_tema, value=hoja.cell(row=o, column=c_tema).value)
        libro.save(destino)
    except Exception as e:  # noqa: BLE001
        return {"content": [{"type": "text", "text": f"Error guardando: {e}"}], "is_error": True}

    print(f"  >>> guardar_clasificacion: {len(nuevas)} filas → {destino.relative_to(AQUI)} (faltan {len(faltantes)})")
    if faltantes:
        # Una fila sin clasificar es un ERROR, no un hueco silencioso: el modelo debe completarla.
        return {"content": [{"type": "text", "text":
                f"Guardado parcial en entregas/{destino.name}. ERROR: faltan {len(faltantes)} filas sin "
                f"clasificar, ids: {', '.join(faltantes)}. Clasifícalas y vuelve a llamar."}], "is_error": True}
    return {"content": [{"type": "text", "text":
            f"Guardado en entregas/{destino.name}. Filas recibidas: {len(nuevas)}. Faltan por clasificar: 0."}]}


medios = create_sdk_mcp_server(
    name="medios", version="1.0.0", tools=[leer_noticias, guardar_clasificacion]
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
1. Clasificar un xlsx: el usuario te da el archivo, el nombre del CLIENTE (marca) y, opcionalmente, ALIAS, VOCEROS y las columnas con el texto de la noticia (por defecto "Título" y "Resumen - Aclaracion"). A cada fila le asignas TEMA y TONO y los guardas en una copia del archivo.
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
TEMA: una frase específica de 4 a 7 palabras que defina realmente la noticia y la distinga de otras. Evita temas genéricos como "Seguridad en Bogotá" o "Gestión de la alcaldía": un lector debe saber de qué hecho se trata solo con el tema.

TONO: Positivo, Neutro o Negativo, evaluado solo desde la perspectiva del cliente, no del tono general de la noticia. Una tragedia puede ser Positiva para el cliente si muestra su gestión a favor.
- Positivo: gestiones a favor, logros, reconocimientos, inversiones u obras del cliente.
- Negativo: críticas, denuncias, sanciones o afectación directa a la marca.
- Neutro: menciones generales o informativas sin efecto claro para el cliente, o noticias donde el cliente no aparece.

Consistencia, en este orden:
1. Lee TODAS las filas antes de asignar nada.
2. Agrupa las noticias con títulos o resúmenes iguales o similares (el campo "grupo" ya junta los títulos idénticos; tú juntas además los que cuentan el mismo hecho con otra redacción).
3. A cada grupo dale exactamente el mismo TEMA y el mismo TONO. Reutiliza un tema existente cuando el hecho sea el mismo.
4. Antes de guardar, repasa la lista de temas y fusiona los casi duplicados con redacción distinta.
</como_trabajas>

<herramientas>
- mcp__medios__leer_noticias: llámala una vez con el archivo y las columnas (vacío = por defecto). Te devuelve todas las filas con su id. Si el archivo viene limpio (entregas/..._limpio.xlsx), el id es la columna "Fila" y los duplicados ya vienen omitidos: no los pidas ni los clasifiques, heredan la clasificación de su original al guardar.
- mcp__medios__guardar_clasificacion: pásale el archivo y la lista JSON de {id, tema, tono}. Manda todas las filas en una sola llamada si caben; si no, en tandas de unas 60. Si responde con un ERROR de filas faltantes, clasifica esos ids y vuelve a llamarla; sigue hasta que diga "Faltan por clasificar: 0". El resultado queda en las columnas "Tono C" y "Tema C" al final de la copia.
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
    },
    forward_subagent_text=True,               # paso 7: ver también el TEXTO del ayudante
    permission_mode="default",                # paso 8: aprueba solo lo de allowed_tools
    disallowed_tools=["Bash", "Edit"],        # paso 8: allowed_tools aprueba; esto prohíbe
    hooks={"PreToolUse": [HookMatcher(matcher="Write|Edit", hooks=[HOOK])]},   # paso 9: escribe solo en entregas/
    cwd=WORKSPACE,                            # la carpeta que el agente ve (junto a este archivo)
    max_turns=25,                             # freno: máximo de vueltas del ciclo
    max_budget_usd=0.40,                      # freno: tope de gasto (modo portales: TOPE_PORTALES)
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
COLUMNAS_TEXTO = ",".join(COLUMNAS_POR_DEFECTO)


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
            f"COLUMNAS: {a.columnas}\n"
            "Clasifica todas las filas con TONO y TEMA y guárdalas con guardar_clasificacion.")


def limpiar_antes(archivo: str) -> str:
    """Limpia y ordena el xlsx (sin IA, $0) y devuelve la ruta del limpio para el agente."""
    origen = _ruta(archivo)
    destino = ENTREGAS / f"{origen.stem}_limpio.xlsx"
    (ENTREGAS / f"{origen.stem}_clasificado.xlsx").unlink(missing_ok=True)   # no mezclar con corridas viejas
    r = limpieza.limpiar_dossier(origen, destino)
    print(f"🧹 limpieza: {r['total']} filas ({r['filas_expandidas']} por menciones), "
          f"{r['unicos']} únicas, {r['duplicados']} duplicadas, "
          f"{len(r['medios_sin_region'])} medios sin región → {destino.relative_to(AQUI)}")
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
    p.add_argument("--columnas", default=COLUMNAS_TEXTO, help=f"por defecto: {COLUMNAS_TEXTO}")
    p.add_argument("--limpiar", dest="limpiar", action="store_true", default=True,
                   help="limpia y ordena el xlsx antes de clasificar (por defecto)")
    p.add_argument("--sin-limpiar", dest="limpiar", action="store_false",
                   help="clasifica el xlsx tal cual, sin limpieza previa")
    a = p.parse_args()

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
        a.archivo = limpiar_antes(a.archivo)
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
