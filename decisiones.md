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

## Qué probar después

- Subir a `claude-sonnet-5-5` y comparar los tonos dudosos y el recibo.
- Ajustar el párrafo de TONO del system prompt con ejemplos reales de tu equipo.
- Poner `stderr` en modo depuración si algo falla sin mensaje.
