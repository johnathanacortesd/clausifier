# clausifier

Clasifica las noticias de un xlsx de monitoreo para cualquier cliente: agrega a una copia del
archivo dos columnas al final, **Tono C** y **Tema C**. También compara cómo cubren al cliente
cinco portales colombianos (El Tiempo, El Espectador, Semana, Infobae Colombia y Caracol Radio).

Hecho con el [Claude Agent SDK](https://code.claude.com/docs/en/agent-sdk).

## Cómo correrlo

En esta máquina `uv` necesita los certificados del sistema para llegar a PyPI (si no, sale
`invalid peer certificate: UnknownIssuer`). En PowerShell:

```powershell
$env:UV_SYSTEM_CERTS = "1"     # una vez por terminal
uv sync                        # la primera vez: instala el SDK y openpyxl
```

### Cómo probar con tu xlsx

1. Copia tu archivo a `workspace\datos\`.
2. Corre (CLIENTE es obligatorio; ALIAS y VOCEROS son opcionales, separados por coma):

```powershell
uv run agente.py --archivo "datos/noticias sec.xlsx" --cliente "Secretaría de Educación" --alias "SED" --voceros "Nombre Apellido"
```

Otras columnas de texto: `--columnas "Título,Texto"` (por defecto `Título,Resumen - Aclaracion`).
Si no pasas `--cliente` o `--archivo`, el programa te los pregunta en la consola.

El resultado queda en `workspace\entregas\<nombre original>_clasificado.xlsx`: copia exacta del
original (hojas, formatos y orden de filas intactos) con **Tono C** y **Tema C** al final. El
original no se toca. Si queda alguna fila sin clasificar, la herramienta lo marca como error y
el agente la completa.

### Comparar 5 portales

```powershell
uv run agente.py --modo portales --cliente "Secretaría de Educación" --alias "SED" --voceros "Nombre Apellido"
```

El agente lanza cinco ayudantes `analista-portal` **en paralelo**, uno por portal (solo buscan y
leen; no escriben). Después une sus respuestas en `workspace\entregas\comparativo.md`: una tabla
por nota con portal, fecha, titular, TEMA, TONO y enlace, un resumen por portal y en qué se
diferencia la información para la marca. En pantalla las líneas de los ayudantes salen con `↳`.

Este modo cuesta **~5× más** que clasificar, porque cada ayudante paga sus propios tokens. Por eso
su tope sube a `TOPE_PORTALES = 0.60` dólares.

### Pedido libre

```powershell
uv run agente.py "Clasifica datos/noticias sec.xlsx. Cliente: <X>"
```

### La credencial

`credenciales.py` lee `.env` y usa **una sola** forma de pagar (la primera línea que imprime
el agente dice cuál):

| `AGENTE_AUTH` en `.env` | Qué hace | Quién paga |
|---|---|---|
| `ant` | usa tu login de `ant auth login` (perfil en `ANTHROPIC_PROFILE`) | la organización de ese login |
| `api_key` (configurado) | usa `ANTHROPIC_API_KEY` | la organización dueña de la key (tus créditos de la Console) |
| `suscripcion` | usa `CLAUDE_CODE_OAUTH_TOKEN` de `claude setup-token` | tu plan Pro/Max |

Para cambiar, edita `.env` (hay un ejemplo en `.env.example`). `.env` está en `.gitignore`:
nunca lo subas a git.

## Qué vas a ver en pantalla

```
[auth] API key de .env (paga la organización dueña de la key)
▶ clausifier

[init] skill «comparar-portales» cargada: True
[init] servidores: [('medios', 'connected')]
🔧 mcp__medios__leer_noticias(archivo=datos/noticias sec.xlsx, columnas=Título,Resumen - Aclaracion)
  >>> leer_noticias: 189 noticias, 158 títulos distintos
🔧 mcp__medios__guardar_clasificacion(archivo=datos/noticias sec.xlsx, filas_json=[…])
  >>> guardar_clasificacion: 68 filas → workspace\entregas\noticias sec_clasificado.xlsx (faltan 0)
- Archivo generado: entregas/noticias sec_clasificado.xlsx …

────────────────────────────────────────────────────────────
9 turnos · $0.0893 · 152s
```

- `🔧` es el agente **pidiendo** una herramienta; la ejecuta este programa.
- `>>>` es tu función corriendo de verdad.
- La última línea es el recibo: vueltas del ciclo y costo de esa corrida.
- `⚠️ terminó por: error_max_turns` o `error_max_budget_usd`: chocó con un freno.

## Cuánto cuesta

Clasificar ~190 noticias con Haiku 5.5: **~$0.09**. Tope duro: `max_budget_usd=0.40` por corrida.
Comparar portales (5 ayudantes): cuesta ~5× más; tope `TOPE_PORTALES = 0.60`.

## Cómo funciona

```
 tú: --archivo --cliente --alias --voceros
        │
        ▼
 ┌────────────────────────────┐  leer_noticias ─────► tu xlsx (openpyxl)
 │ Claude · claude-haiku-5-5  │  guardar_clasificacion► entregas/*_clasificado.xlsx
 │ reglas: TEMA 4-7 palabras, │  Agent ×5 en paralelo► analista-portal (WebSearch/WebFetch)
 │ TONO desde el cliente      │  Write ─[hook: solo entregas/]─► comparativo.md
 │ skill: comparar-portales   │  ✗ Bash ✗ Edit
 └────────────────────────────┘
        │  máx 25 vueltas · máx $0.40
        ▼
 resumen: filas, tonos, temas top, dudosas
```

## Cómo cambiar una decisión

Todo vive en el bloque `opciones = ClaudeAgentOptions(…)` de `agente.py`.

| Quieres… | Cambia esta línea | Por |
|---|---|---|
| mejor juicio de tono (más caro) | `model=` | `"claude-sonnet-5-5"` |
| otras reglas | `SYSTEM_PROMPT` | tu propio texto |
| otros portales | `workspace/.claude/skills/comparar-portales/SKILL.md` | tu lista |
| que pueda dar más vueltas | `max_turns=` | un número mayor |
| subir el tope de gasto | `max_budget_usd=` | un número mayor |
| ver el error real cuando algo falla | `stderr=` | `lambda linea: print(linea, end="")` |

El porqué de cada decisión está en [decisiones.md](decisiones.md).

## Archivos

| Archivo | Qué es |
|---|---|
| `agente.py` | el agente: tus dos funciones, el system prompt, las decisiones y el lector del stream |
| `hooks.py` | el vigilante: `Write` solo dentro de `entregas/` |
| `credenciales.py` | lee `.env` y prepara la credencial |
| `pyproject.toml` | las dependencias |
| `workspace/datos/` | tus xlsx de entrada |
| `workspace/entregas/` | lo que produce el agente |
| `workspace/.claude/skills/comparar-portales/SKILL.md` | el procedimiento del comparativo |
| `decisiones.md` | las nueve respuestas, su código y dónde leer más |

## Si algo falla

| Lo que ves | Qué pasa |
|---|---|
| `invalid peer certificate: UnknownIssuer` | falta `$env:UV_SYSTEM_CERTS = "1"` |
| `[auth] Falta …` | `.env` incompleto: el mensaje dice qué falta |
| `Invalid API key` / `401` | la key de `.env` no sirve: pégala de nuevo |
| `Credit balance is too low` | la organización de la key no tiene créditos |
| `No existe la columna '…'` | pasa `--columnas` con los nombres reales del encabezado |
| `error_max_budget_usd` | archivo grande: sube `max_budget_usd` |

## App web (Streamlit)

1. Sube este repo a GitHub (sin `.env` ni `.streamlit/secrets.toml`: el `.gitignore` los excluye).
2. En share.streamlit.io: *Create app* → este repo → archivo principal `app.py`.
3. En *Settings → Secrets* pega (formato en `.streamlit/secrets.toml.example`):

```toml
ANTHROPIC_API_KEY = "sk-ant-..."
APP_PASSWORD = "elige-una-contraseña"
```

En local: copia `.streamlit/secrets.toml.example` a `.streamlit/secrets.toml`, llénalo y corre `uv run streamlit run app.py`.
