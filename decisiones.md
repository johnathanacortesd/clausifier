# Decisiones de clausifier

La credencial y las nueve respuestas del wizard, la opción exacta del SDK que produjo cada una, y
dónde leer más. Generado el 2026-11-07.

Versiones contra las que se probó: **Python SDK 0.2.176** (empaca Claude Code 2.1.316).

| # | Pregunta | Tu respuesta | Código exacto | Dónde leer más |
|---|---|---|---|---|
| 0 | ¿Con qué pagas? | API key de la Console | `AGENTE_AUTH=api_key` en `.env` (lo aplica `credenciales.py`) | [authentication](https://code.claude.com/docs/en/authentication) |
| 1 | ¿En qué lenguaje? | Python | `agente.py` + `pyproject.toml`, `query()` | [quickstart](https://code.claude.com/docs/en/agent-sdk/quickstart.md) |
| 2 | ¿Qué tiene que lograr? | Llamar a mis funciones (xlsx) + investigar en la web | una corrida `query()` por pedido; `cwd=workspace` | [agent-loop](https://code.claude.com/docs/en/agent-sdk/agent-loop.md) |
| 3 | ¿Qué modelo? | Haiku 5.5 | `model="claude-haiku-5-5"` | [modelos](https://platform.claude.com/docs/en/about-claude/models/overview.md) |
| 4 | ¿Qué reglas fijas? | Mínimo + las reglas de tu idea (cliente, alias, voceros, TEMA, TONO, consistencia) | `system_prompt=SYSTEM_PROMPT` | [system prompt](https://code.claude.com/docs/en/agent-sdk/modifying-system-prompts.md) |
| 5 | ¿Qué puede hacer? | Mis funciones + web; sin MCP externo | `mcp_servers={"medios": medios}`, `allowed_tools=[mcp__medios__leer_noticias, mcp__medios__guardar_clasificacion, WebSearch, WebFetch, Write, Skill]` | [custom-tools](https://code.claude.com/docs/en/agent-sdk/custom-tools.md) |
| 6 | ¿Un procedimiento tuyo? | Skill nueva `comparar-portales` | `setting_sources=["project"]`, `skills=["comparar-portales"]` | [skills](https://code.claude.com/docs/en/skills.md) |
| 7 | ¿Repartir la tarea? | Sí, varios en paralelo (solo modo portales; clasificar sigue con un solo agente) | `agents={"analista-portal": AgentDefinition(tools=["WebSearch", "WebFetch"], model="haiku", maxTurns=8, …)}`, `"Agent"` en `allowed_tools`, `forward_subagent_text=True`; tope del modo `TOPE_PORTALES = 0.60` | [subagents](https://code.claude.com/docs/en/agent-sdk/subagents.md) |
| 8 | ¿Quién autoriza? | Solo lo que listé | `permission_mode="default"`, `disallowed_tools=["Bash", "Edit"]` | [permissions](https://code.claude.com/docs/en/agent-sdk/permissions.md) |
| 9 | ¿Un vigilante? | El guardián de tu idea: escribir solo en `entregas/` | `hooks={"PreToolUse": [HookMatcher(matcher="Write\|Edit", hooks=[HOOK])]}` | [hooks](https://code.claude.com/docs/en/agent-sdk/hooks.md) |

## Frenos que se pusieron sin preguntar

| Opción | Valor | Por qué |
|---|---|---|
| `max_turns` | `25` | Corta el ciclo tras esas vueltas. Al cortar, `result.subtype = "error_max_turns"`. |
| `max_budget_usd` | `0.40` | Corta cuando el costo estimado llega a ese valor. Al cortar, `result.subtype = "error_max_budget_usd"`. |
| `setting_sources` | `["project"]` | Carga solo `workspace/.claude/` (la skill), no tu `~/.claude`. |

## Las tres cosas que más confunden de este SDK

1. **`allowed_tools` aprueba, no restringe.** Para quitar una herramienta de verdad, va en
   `disallowed_tools` (por eso Bash y Edit están ahí).
2. **Cada llamada abre una sesión nueva.** El agente no recuerda la corrida anterior.
3. **Un hook corre antes que todo lo demás**: aunque cambies el modo de permisos, `Write` sigue
   sin poder salir de `entregas/`.

## Limpieza antes de clasificar (agregado después)

- **Qué**: `limpieza.py` porta de API_2_Grill solo la parte sin IA (carga con hipervínculos,
  normalización, expansión de menciones, duplicados, orden por título). Va antes del agente porque
  es determinista: no tiene sentido pagar tokens por algo que el código hace igual siempre.
- **Duplicados**: el modelo no los ve; heredan Tono / Subtema del original al guardar (por
  "ID duplicada" + mención). Menos tokens y la misma clasificación garantizada para la misma nota.
- **Fila**: tras expandir menciones el ID se repite, así que la clave de cada fila es "Fila".
- **Región**: los mapas solo vienen de `REGIONES_CSV_URL` / `INTERNET_CSV_URL`; sin ellos, "N/A".
  No portado: el analizador OpenAI, Jev/Typesafe y los PKL.
- **Desactivar**: `--sin-limpiar` en la terminal (sin limpieza ni Contexto: el agente usa las
  columnas de cuerpo). En la app la limpieza ya no es opcional.

## Contexto analizado (agregado después)

- **Qué**: `contexto.py` porta de API_2_Grill el extracto de menciones (`_texto_contexto_fila`,
  `_contexto_exacto_marca`, `_contexto_con_titulo_y_extension`) y la columna "texto". Corre tras la
  limpieza, en terminal y app (también en "Solo limpiar", que por eso exige cliente).
- **Por qué**: el modelo juzga el tono por lo que se dice del cliente, no por el tono general de la
  nota; y enviar solo el extracto (más el título como etiqueta) ahorra tokens frente al cuerpo.
- **Página**: 25 noticias, contexto recortado a 1200 car.; en `noticias sec.xlsx` cada página mide
  5,7-13,8k caracteres (≈2-5k tokens). Corrida real: 18 turnos, $0.036, 0 filas faltantes.
- **Reconstruido**: `texto_sin_titulo` y `expandir_voceros` vivían en `coherencia_texto.py`, que no
  estaba en la copia de Grill; se rehicieron según su README. Los duplicados también reciben
  Contexto (Grill les pone "-").

## Qué probar después

- Subir a `claude-sonnet-5-5` y comparar los tonos dudosos y el recibo.
- Ajustar el párrafo de TONO del system prompt con ejemplos reales de tu equipo.
- Poner `stderr` en modo depuración si algo falla sin mensaje.

## Revisión de consistencia y layout de entrega (agregado después)

- **Layout**: el xlsx final lo escribe `exportar.py` con las 26 columnas exactas de `noticias sec.xlsx` (export de Grill). El estado de la clasificación vive en el `_limpio` (con `Fila`); cada `guardar_clasificacion` re-exporta. Solo limpiar usa el mismo layout con Tono_IA/Subtema_IA vacías.
- **Candidatos sin IA, decisión con IA**: agrupar es determinista (rapidfuzz, umbrales de Grill); decidir si es el mismo hecho lo hace `revisor-subtemas` (Haiku, solo la tool de candidatos). Solo se mandan grupos inconsistentes + pares sospechosos de sobreagrupación, así cuesta poco (~0.02 USD en 189 filas).
- Contextos de menos de 120 car. (p. ej. 'Fuente: Alcaldía de Bogotá') no cuentan para similitud: encadenaban hechos distintos.
- Bug corregido: `hoja.cell(..., value=None)` no borra en openpyxl; ahora se asigna `.value = None` al empezar una corrida (antes una entrada ya clasificada dejaba el 'faltan 0' sin sentido).
- App: el log crudo va a 'Ver detalle técnico'; la vista muestra paso, barra de páginas y resumen. El ayudante corre en segundo plano y llegan dos ResultMessage: se suman turnos y se toma el costo acumulado.
