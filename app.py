# app.py — interfaz web de clausifier (Streamlit), con la interfaz de API_2_Grill.
#
# Secrets necesarios (Streamlit Cloud: Settings → Secrets; local: .streamlit/secrets.toml):
#     ANTHROPIC_API_KEY = "sk-ant-..."
#     APP_PASSWORD      = "..."
#     REGIONES_CSV_URL / INTERNET_CSV_URL   (opcionales: mapas de región)
# Correr en local:  uv run streamlit run app.py
#
# Flujo (igual que Grill): formulario → panel "Procesando dossier de noticias" → pantalla de
# resultados con métricas, duplicadas, tonos y costo. La IA es solo claude-haiku-5-5 (agente.py).

import asyncio
import contextlib
import hmac
import html
import io
import json
import os
import re
import time
from collections import Counter
from types import SimpleNamespace

import pandas as pd
import streamlit as st

from criterios import CRITERIOS_TONO
from perfil_cliente import (cargar_perfil, guardar_perfil, listar_perfiles, normalizar_perfil,
                            perfil_a_resumen, slugify, validar_perfil)

APP_VERSION = "v1.0 · Haiku"
MODELO = "claude-haiku-5-5"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="Limpieza y Análisis de Noticias", page_icon="◈", layout="wide",
                   initial_sidebar_state="collapsed")

# ======================================
# CSS · el de API_2_Grill con la paleta crema/terracota de clausifier
# ======================================
st.markdown("""
<style>
:root,[data-testid="stApp"]{
    color-scheme:light;
    --bg:#FAF9F5;--s1:#FFFFFF;--s2:#F5F3EC;--s3:#F0EEE6;
    --border:#E3DFD3;--border2:#D6D1C4;--border-input:#D6D1C4;
    --text:#141413;--text2:#4A4842;--text3:#6B6760;--text4:#A39E93;
    --acento:#D97757;--acento-hover:#C4643F;--acento-soft:rgba(217,119,87,0.09);--acento-line:rgba(217,119,87,0.35);
    --grad-br:linear-gradient(90deg,#D97757,#E59A7E);
    --green:#2E7D4F;--green-bdr:rgba(46,125,79,0.30);--amber:#9A6A14;--blue:#D97757;
    --success-bg:rgba(46,125,79,0.07);
    --r:8px;--r2:10px;--r3:14px;
    --shadow-sm:0 1px 2px rgba(20,20,19,0.05);
    --mono:ui-monospace,"SF Mono",Menlo,Consolas,monospace;
}
html,body,[data-testid="stApp"]{background:var(--bg)!important;color:var(--text)!important;}
#MainMenu,footer{visibility:hidden}.stDeployButton{display:none}
.block-container{padding-top:1.75rem!important;max-width:1080px!important}
h1,h2,h3{font-family:var(--mono)!important;letter-spacing:-0.01em}
div[data-testid="stTabs"] div[role="tablist"],div[data-baseweb="tab-list"]{gap:1.75rem!important;border-bottom:1px solid var(--border)!important;background:transparent!important;}
div[data-testid="stTab"] p,button[data-baseweb="tab"] p{color:var(--text3)!important;font-weight:500!important;}
div[data-testid="stTab"][aria-selected="true"] p,button[data-baseweb="tab"][aria-selected="true"] p{color:var(--acento)!important;font-weight:600!important;}
.app-header{background:var(--s1);border:1px solid var(--border);border-radius:var(--r3);padding:0.8rem 1.25rem;margin-bottom:1.1rem;display:flex;align-items:center;gap:0.85rem;box-shadow:var(--shadow-sm);}
.app-header-icon{width:36px;height:36px;background:var(--acento-soft);border:1px solid var(--acento-line);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.1rem;color:var(--acento);flex-shrink:0;}
.app-header-text{flex:1}
.app-header-title{font-family:var(--mono);font-size:1.15rem;font-weight:700;color:var(--text);line-height:1.25}
.app-header-title::before{content:"> ";color:var(--acento)}
.app-header-version{font-family:var(--mono);font-size:0.72rem;color:var(--text3);margin-top:0.2rem}
.app-header-badge{background:var(--acento-soft);border:1px solid var(--acento-line);color:var(--acento);font-family:var(--mono);font-size:0.66rem;font-weight:600;padding:0.28rem 0.75rem;border-radius:6px;letter-spacing:0.12em;text-transform:uppercase;white-space:nowrap;}
.metrics-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:0.75rem;margin:0.9rem 0}
.metric-card{background:var(--s1);border:1px solid var(--border);border-radius:var(--r2);padding:1rem 0.6rem;text-align:center;box-shadow:var(--shadow-sm);}
.metric-val{font-family:var(--mono);font-size:1.5rem;font-weight:700;line-height:1;margin-bottom:0.3rem;color:var(--text)}
.metric-lbl{font-size:0.64rem;color:var(--text3);text-transform:uppercase;letter-spacing:0.12em;font-weight:500}
[data-testid="stForm"]{background:var(--s1)!important;border:1px solid var(--border)!important;border-radius:var(--r3)!important;padding:1.5rem 1.75rem!important;box-shadow:var(--shadow-sm)!important;}
.sec-label{font-family:var(--mono);font-size:0.92rem;font-weight:600;color:var(--text);padding-bottom:0.45rem;border-bottom:1px solid var(--border);margin:1.2rem 0 0.8rem;display:flex;align-items:center;gap:0.6rem;}
.sec-label::before{content:'';display:inline-block;width:3px;height:15px;background:var(--acento);border-radius:2px}
.upload-zone-card{background:var(--s2);border:1px dashed var(--border2);border-radius:var(--r2);padding:0.65rem 0.85rem;display:flex;align-items:center;gap:0.65rem;margin:0.3rem 0}
.upload-zone-icon{width:34px;height:34px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:1rem;background:var(--s1);border:1px solid var(--border)}
.upload-zone-title{font-size:0.84rem;font-weight:600;color:var(--text);line-height:1.2}
.upload-zone-desc{font-size:0.72rem;color:var(--text3);line-height:1.35}
.stButton>button,[data-testid="stDownloadButton"]>button,[data-testid="stFormSubmitButton"]>button{border-radius:10px!important;font-family:var(--mono)!important;}
.success-banner{background:var(--success-bg);border:1px solid var(--green-bdr);border-radius:var(--r2);padding:0.9rem 1.2rem;margin:0.6rem 0 1rem;display:flex;align-items:center;gap:0.85rem;}
.success-icon{width:32px;height:32px;background:rgba(46,125,79,0.12);border:1px solid var(--green-bdr);border-radius:50%;display:flex;align-items:center;justify-content:center;color:var(--green);flex-shrink:0;}
.success-title{font-size:0.95rem;font-weight:700;color:var(--green)}
.success-sub{font-size:0.8rem;color:var(--text2)}
.auth-wrap{max-width:440px;margin:8vh auto 0;text-align:center}
.auth-icon{width:64px;height:64px;background:var(--acento-soft);border:1px solid var(--acento-line);border-radius:16px;display:inline-flex;align-items:center;justify-content:center;font-size:1.6rem;color:var(--acento);margin-bottom:1.25rem;}
.auth-eyebrow{font-family:var(--mono);font-size:0.68rem;font-weight:600;color:var(--acento);letter-spacing:0.18em;text-transform:uppercase;margin-bottom:0.85rem}
.auth-title{font-family:var(--mono);font-size:1.5rem;font-weight:700;color:var(--text);margin-bottom:0.45rem}
.auth-sub{font-size:0.92rem;color:var(--text2);margin-bottom:1.5rem}
.auth-foot{font-family:var(--mono);font-size:0.72rem;color:var(--text4);margin-top:1.2rem;text-align:center}
.config-badge{display:inline-flex;align-items:center;gap:0.4rem;background:var(--s2);border:1px solid var(--border);border-radius:6px;padding:0.25rem 0.7rem;font-family:var(--mono);font-size:0.66rem;color:var(--text3);margin-bottom:0.6rem;letter-spacing:0.06em;}
.live-panel{background:var(--s1);border:1px solid var(--border);border-radius:var(--r3);padding:1.4rem 1.6rem;margin:1.5rem auto;max-width:720px;box-shadow:var(--shadow-sm);}
.live-head{display:flex;align-items:center;gap:0.75rem;margin-bottom:0.8rem;}
.live-pulse{width:10px;height:10px;border-radius:50%;background:var(--acento);animation:livePulse 1.6s ease-out infinite;flex-shrink:0;}
@keyframes livePulse{0%{box-shadow:0 0 0 0 rgba(217,119,87,0.45)}70%{box-shadow:0 0 0 10px rgba(217,119,87,0)}100%{box-shadow:0 0 0 0 rgba(217,119,87,0)}}
.live-title{font-family:var(--mono);font-size:0.95rem;font-weight:600;color:var(--text)}
.live-sub{font-size:0.78rem;color:var(--text3);margin-top:0.25rem}
.live-metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:0.6rem;margin:0.5rem 0 0.8rem}
.live-metric{background:var(--s2);border:1px solid var(--border);border-radius:var(--r);padding:0.6rem 0.5rem;text-align:center}
.live-metric-val{font-family:var(--mono);font-size:1.1rem;font-weight:700;color:var(--text)}
.live-metrics .live-metric:nth-child(3) .live-metric-val{color:var(--acento)}
.live-metric-lbl{font-size:0.6rem;color:var(--text3);text-transform:uppercase;letter-spacing:0.12em;margin-top:0.2rem}
.step-list{display:flex;flex-direction:column;gap:0.3rem;margin:0.2rem 0 0.7rem}
.step-item{display:flex;align-items:center;gap:0.55rem;font-size:0.82rem;color:var(--text3);padding:0.22rem 0.1rem}
.step-item .dot{width:18px;height:18px;border-radius:50%;border:1.5px solid var(--border2);display:flex;align-items:center;justify-content:center;font-size:0.62rem;flex-shrink:0;background:var(--s1);color:transparent}
.step-item.is-done{color:var(--text2)}
.step-item.is-done .dot{background:rgba(46,125,79,0.14);border-color:var(--green);color:var(--green)}
.step-item.is-active{color:var(--text);font-weight:600}
.step-item.is-active .dot{border-color:var(--acento);color:var(--acento);animation:livePulse 1.6s ease-out infinite}
.live-hint{background:var(--acento-soft);border:1px solid var(--acento-line);color:#8A4127;border-radius:var(--r);padding:0.55rem 0.75rem;font-size:0.78rem;line-height:1.4}
.live-detail{font-size:0.8rem;color:var(--text2);margin-top:0.5rem;font-family:var(--mono)}
.footer{font-family:var(--mono);font-size:0.64rem;color:var(--text4);text-align:center;padding:0.9rem 0 0.6rem;letter-spacing:0.14em;text-transform:uppercase;border-top:1px solid var(--border);margin-top:1.2rem;}
[data-testid="stExpander"]{background:var(--s1)!important;border:1px solid var(--border)!important;border-radius:var(--r2)!important;}
@media(max-width:768px){.metrics-grid{grid-template-columns:repeat(2,1fr)}.app-header{flex-direction:column;text-align:center}}
</style>
""", unsafe_allow_html=True)


# ======================================
# Autenticación (como Grill, con comparación en tiempo constante)
# ======================================
def check_password() -> bool:
    if st.session_state.get("password_correct"):
        return True
    esperada = str(st.secrets.get("APP_PASSWORD", "") or "")
    st.markdown("""
    <div class="auth-wrap">
        <div class="auth-icon">◈</div>
        <div class="auth-eyebrow">Acceso privado</div>
        <div class="auth-title">Sistema de Limpieza y Análisis</div>
        <div class="auth-sub">Ingresa la contraseña para continuar con el análisis de noticias.</div>
    </div>""", unsafe_allow_html=True)
    if not esperada:
        st.warning("⚠️ No hay APP_PASSWORD en los Secrets: configura una contraseña antes de "
                   "publicar la app (Streamlit Cloud → Settings → Secrets).")
    _, col, _ = st.columns([1, 1.6, 1])
    with col:
        with st.form("pw"):
            pw = st.text_input("Contraseña", type="password", placeholder="Ingresa tu contraseña")
            if st.form_submit_button("Ingresar", width="stretch", type="primary"):
                if esperada and hmac.compare_digest(pw.encode(), esperada.encode()):
                    st.session_state["password_correct"] = True
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta")
        st.markdown(f'<div class="auth-foot">{APP_VERSION}</div>', unsafe_allow_html=True)
    return False


if not check_password():
    st.stop()

# ─── Credencial: solo la API key de los secrets ──────────────────────────
clave_api = st.secrets.get("ANTHROPIC_API_KEY", "")
if not clave_api:
    st.error("❌ Falta ANTHROPIC_API_KEY en los Secrets de Streamlit.")
    st.stop()
for otra in ("CLAUDE_CODE_OAUTH_TOKEN", "ANTHROPIC_PROFILE"):
    os.environ.pop(otra, None)
os.environ["ANTHROPIC_API_KEY"] = clave_api
for opcional in ("REGIONES_CSV_URL", "INTERNET_CSV_URL"):
    if st.secrets.get(opcional):
        os.environ[opcional] = str(st.secrets[opcional])

import agente  # noqa: E402 — después de fijar la credencial

from claude_agent_sdk import ClaudeSDKError, query  # noqa: E402

# ======================================
# Panel en vivo (pasos de Grill adaptados al pipeline de clausifier)
# ======================================
PIPELINE_STEPS = [
    ("config", "Cargar configuración"),
    ("read", "Leer el Excel"),
    ("norm", "Limpiar y normalizar"),
    ("dups", "Menciones y duplicadas"),
    ("ctx", "Contexto analizado (sin IA)"),
    ("ai", "Análisis IA (Tono, Subtema) con Haiku"),
    ("review", "Consistencia (revisor-subtemas)"),
    ("export", "Generar archivo de resultado"),
]
_PASO_A_STEP = {"Clasificando": "ai", "Revisando consistencia": "review", "Guardando": "export"}


def _fmt_elapsed(seconds: float) -> str:
    seconds = max(0, int(seconds))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def _render_live_html(pct, msg, elapsed, file_label, active_key, pasos=PIPELINE_STEPS) -> str:
    steps_html, reached = [], False
    for key, label in pasos:
        if active_key == "done":
            cls, mark = "is-done", "✓"
        elif key == active_key:
            cls, mark, reached = "is-active", "●", True
        elif not reached:
            cls, mark = "is-done", "✓"
        else:
            cls, mark = "", ""
        steps_html.append(f'<div class="step-item {cls}"><span class="dot">{mark}</span>{label}</div>')
    file_line = f" · {html.escape(file_label)}" if file_label else ""
    title = "Proceso completado" if active_key == "done" else "Procesando dossier de noticias"
    return f"""
    <div class="live-panel">
      <div class="live-head"><div class="live-pulse"></div><div>
        <div class="live-title">{title}</div>
        <div class="live-sub">El proceso sigue activo{file_line}. No cierres esta pestaña.</div>
      </div></div>
      <div class="live-metrics">
        <div class="live-metric"><div class="live-metric-val">{int(pct)}%</div><div class="live-metric-lbl">Avance</div></div>
        <div class="live-metric"><div class="live-metric-val">{elapsed}</div><div class="live-metric-lbl">Tiempo</div></div>
        <div class="live-metric"><div class="live-metric-val">en curso</div><div class="live-metric-lbl">Estado</div></div>
      </div>
      <div class="step-list">{''.join(steps_html)}</div>
      <div class="live-hint">Límite de seguridad: máximo 1.000 filas por archivo. Las duplicadas (URL + mención, o medio + hora + mención en radio/TV) no se envían a la IA: heredan el tono y subtema de su original.</div>
      <div class="live-detail">{html.escape(str(msg or ""))}</div>
    </div>"""


class Panel:
    """Tarjeta en vivo + barra de páginas, en placeholders fijos."""

    def __init__(self, file_label: str):
        self.t0, self.file_label = time.time(), file_label
        self.panel, self.barra = st.empty(), st.progress(0, text="Iniciando…")

    def paint(self, pct: float, msg: str, step: str) -> None:
        self.panel.markdown(_render_live_html(pct, msg, _fmt_elapsed(time.time() - self.t0),
                                              self.file_label, step), unsafe_allow_html=True)
        self.barra.progress(min(100, max(0, int(pct))), text=msg)

    def desde_progreso(self, p: dict) -> None:
        step = _PASO_A_STEP.get(p.get("paso", ""), "ai")
        total = p.get("total") or 0
        leidas = p.get("leidas", 0)
        if step == "ai":
            pct = 35 + (50 * leidas / total if total else 0)
        else:
            pct = {"review": 88, "export": 95}.get(step, 90)
        msg = (f"{p.get('paso', 'Clasificando')} · Página {p.get('pagina', 0)} de {p.get('paginas', 0)} · "
               f"{leidas} de {total} noticias") if total else p.get("paso", "Clasificando")
        self.paint(pct, msg, step)


# ======================================
# Agente
# ======================================
def _paso_por_herramienta(mensaje, modo: str) -> str | None:
    """Paso actual según la herramienta que pide el agente principal (no los ayudantes)."""
    if getattr(mensaje, "parent_tool_use_id", None) or not isinstance(mensaje, agente.AssistantMessage):
        return None
    for b in mensaje.content:
        if isinstance(b, agente.ToolUseBlock):
            if b.name in ("Agent", "Task"):
                return "Revisando consistencia" if modo == "xlsx" else "Buscando en portales"
            if b.name.endswith("candidatos_consistencia"):
                return "Revisando consistencia"
            if b.name == "Write":
                return "Guardando"
    return None


async def correr(prompt: str, vista, modo: str = "xlsx") -> dict:
    """Corre el agente mostrando solo un progreso limpio; el log completo queda en el resultado."""
    buffer = io.StringIO()
    agente.FALLO[0] = False
    agente.reiniciar_progreso()
    agente.PROGRESO["paso"] = "Clasificando" if modo == "xlsx" else "Buscando en portales"
    final = {"turnos": 0, "costo": 0.0}
    with contextlib.redirect_stdout(buffer):
        try:
            async for mensaje in query(prompt=prompt, options=agente.opciones):
                agente.imprimir(mensaje)
                agente.imprimir_error_de_tool(mensaje)
                paso = _paso_por_herramienta(mensaje, modo)
                if paso and not (agente.PROGRESO.get("revisado") and paso == "Clasificando"):
                    agente.PROGRESO["paso"] = paso
                if isinstance(mensaje, agente.ResultMessage):
                    final = {"turnos": final["turnos"] + (mensaje.num_turns or 0),
                             "costo": max(final["costo"], mensaje.total_cost_usd or 0.0)}
                vista(agente.PROGRESO)
        except ClaudeSDKError as error:
            print(f"\n⚠️  el agente paró: {error}")
            agente.FALLO[0] = True
    final["log"] = buffer.getvalue()
    return final


def contar_tonos(destino) -> tuple[int, Counter]:
    import openpyxl
    tonos: Counter = Counter()
    filas = 0
    if destino is not None and destino.exists():
        hoja = openpyxl.load_workbook(destino, read_only=True).active
        enc = list(next(hoja.iter_rows(max_row=1, values_only=True)))
        i_t = enc.index("Tono_IA") if "Tono_IA" in enc else None
        for f in hoja.iter_rows(min_row=2, values_only=True):
            if any(f):
                filas += 1
                if i_t is not None:
                    tonos[f[i_t] or "sin tono"] += 1
    return filas, tonos


def _lista(texto: str) -> list[str]:
    """Grill separa por coma o punto y coma."""
    return [x.strip() for x in re.split(r"[,;]", texto or "") if x.strip()]


# ======================================
# Proceso principal (equivalente a run_cleaning_process de Grill)
# ======================================
def run_cleaning_process(job: dict) -> None:
    nombre, blob = job["nombre"], job["blob"]
    size = len(blob)
    label = f"{nombre} ({size / 1048576:.1f} MB)" if size >= 104858 else f"{nombre} ({size / 1024:.0f} KB)"
    pasos_job = PIPELINE_STEPS if job["clasificar"] else [s for s in PIPELINE_STEPS if s[0] not in ("ai", "review")]
    panel = Panel(label)
    t0 = time.time()

    def paint(pct, msg, step):
        panel.panel.markdown(_render_live_html(pct, msg, _fmt_elapsed(time.time() - t0), label, step, pasos_job),
                             unsafe_allow_html=True)
        panel.barra.progress(min(100, max(0, int(pct))), text=msg)

    panel.paint = paint
    paint(2, "Cargando configuración…", "config")
    cliente, alias, voceros = job["brand"], _lista(job["alias"]), _lista(job["voceros"])
    datos = agente.WORKSPACE / "datos"
    datos.mkdir(parents=True, exist_ok=True)
    agente.ENTREGAS.mkdir(parents=True, exist_ok=True)
    (datos / nombre).write_bytes(blob)
    base = os.path.splitext(nombre)[0]
    limpio = agente.ENTREGAS / f"{base}_limpio.xlsx"
    destino = agente.ENTREGAS / f"{base}_clasificado.xlsx"
    destino.unlink(missing_ok=True)
    paint(10, "Leyendo el Excel…", "read")
    try:
        paint(20, "Limpiando, expandiendo menciones y marcando duplicadas…", "dups")
        r = agente.limpieza.limpiar_dossier(datos / nombre, limpio)
        paint(30, "Armando Contexto analizado…", "ctx")
        c = agente.contexto.agregar_contexto(limpio, cliente, alias, voceros)
    except Exception as e:  # noqa: BLE001 — p. ej. más de 1000 filas
        st.session_state["error_proceso"] = f"⚠️ {e}"
        return

    res = {"nombre": nombre, "brand": cliente, "r": r, "c": c, "clasificar": job["clasificar"],
           "criterio": job["criterio_texto"] and "personalizado" or job["criterio"], "final": None,
           "filas": 0, "tonos": {}}
    if job["clasificar"]:
        agente.UMBRAL_TITULO = job["umbral_titulo"] / 100
        agente.UMBRAL_CONTEXTO = job["umbral_cuerpo"] / 100
        args = SimpleNamespace(modo="xlsx", cliente=cliente, alias=", ".join(alias),
                               voceros=", ".join(voceros), archivo=f"entregas/{limpio.name}",
                               criterio=job["criterio"], criterio_texto=job["criterio_texto"])
        agente.opciones.max_budget_usd = agente.TOPE_CLASIFICAR
        (agente.ENTREGAS / f"{base}_trabajo.xlsx").unlink(missing_ok=True)
        prompt = agente.armar_prompt(args)
        st.session_state["ultimo_prompt"] = prompt
        paint(35, "Clasificando con Haiku…", "ai")
        final = asyncio.run(correr(prompt, panel.desde_progreso, "xlsx"))
        paint(96, "Generando archivo de resultado…", "export")
        res["final"] = final
        res["fallo"] = bool(agente.FALLO[0])
        res["revision"] = agente.PROGRESO.get("cambios_revision", 0) if agente.PROGRESO.get("revisado") else None
        filas, tonos = contar_tonos(destino)
        res["filas"], res["tonos"] = filas, dict(tonos.most_common())
        res["salida"] = destino.read_bytes() if destino.exists() else None
        res["salida_nombre"] = destino.name
    else:
        paint(90, "Generando archivo de resultado…", "export")
        entrega = agente.exportar_limpio(limpio)
        res["salida"] = entrega.read_bytes()
        res["salida_nombre"] = limpio.name
    res["duracion"] = _fmt_elapsed(time.time() - t0)
    paint(100, "Proceso completado", "done")
    st.session_state["resultado"] = res
    st.session_state["processing_complete"] = True


# ======================================
# Perfiles (como Grill: selector arriba, guardar al procesar)
# ======================================
_SIN_PERFIL = "Sin perfil (configuración manual)"


def selector_perfil() -> None:
    try:
        perfiles = listar_perfiles()
    except Exception:  # noqa: BLE001
        perfiles = []
    opciones = [_SIN_PERFIL] + [p["nombre"] for p in perfiles]
    sel = st.selectbox("Perfil de cliente", opciones, key="perfil_sel",
                       help="Carga la configuración guardada del cliente. Puedes editar los campos abajo "
                            "antes de procesar; también puedes guardar la configuración actual como un "
                            "perfil nuevo desde 'Ajustes finos'.")
    if sel != _SIN_PERFIL:
        pid = next((p["id"] for p in perfiles if p["nombre"] == sel), None)
        if pid:
            perfil = cargar_perfil(pid)
            if st.session_state.get("_perfil_aplicado") != sel:
                aplicar_perfil(perfil)
                st.session_state["_perfil_aplicado"] = sel
                st.rerun()
            st.caption("📋 " + perfil_a_resumen(perfil) + (f" · notas: {perfil['notas']}" if perfil.get("notas") else ""))
            st.download_button("⬇ Descargar este perfil (JSON)",
                               json.dumps({k: v for k, v in perfil.items() if k != "id"}, ensure_ascii=False, indent=2),
                               file_name=f"{pid}.json", mime="application/json", key="dl_perfil")
    else:
        st.session_state["_perfil_aplicado"] = None
    with st.expander("⬆ Cargar un perfil desde JSON"):
        up = st.file_uploader("Perfil JSON", type=["json"], key="perfil_json", label_visibility="collapsed")
        if up is not None and st.button("Aplicar perfil del JSON", key="aplicar_json"):
            try:
                data = json.loads(up.getvalue().decode("utf-8-sig"))
                errores = validar_perfil(data)
                if errores:
                    raise ValueError(" | ".join(errores))
                aplicar_perfil(normalizar_perfil(data))
                st.session_state["nombre_perfil_input"] = normalizar_perfil(data)["nombre"]
                st.rerun()
            except Exception as exc:  # noqa: BLE001
                st.error(f"El perfil JSON no es válido: {exc}")
    st.caption("Los perfiles se guardan en el disco del servidor: en Streamlit Cloud ese disco se borra "
               "al reiniciar la app. Descarga el JSON del perfil para conservarlo y súbelo cuando lo necesites.")


def aplicar_perfil(perfil: dict) -> None:
    st.session_state["brand_input"] = perfil.get("brand", "")
    st.session_state["alias_input"] = "; ".join(perfil.get("aliases", []))
    st.session_state["voceros_input"] = "; ".join(perfil.get("voceros", []))
    crit = perfil.get("criterio")
    st.session_state["criterio_input"] = crit if crit in CRITERIOS_TONO else list(CRITERIOS_TONO)[0]
    st.session_state["criterio_custom_input"] = perfil.get("criterio_custom", "")
    st.session_state["notas_input"] = perfil.get("notas", "")


def guardar_perfil_si_pide(datos: dict) -> bool:
    """Devuelve False si hay que detenerse (perfil existente sin sobrescribir o inválido)."""
    if not datos.get("guardar"):
        return True
    nombre = (datos.get("nombre_perfil") or "").strip() or datos["brand"]
    if any(p["id"] == slugify(nombre) for p in listar_perfiles()) and not datos.get("sobrescribir"):
        st.warning(f"Ya existe un perfil llamado '{nombre}'. Si quieres actualizarlo, marca "
                   "'Sobrescribir el perfil si ya existe' y procesa de nuevo; si es otro cliente, "
                   "cambia el nombre del perfil.")
        return False
    try:
        pid = guardar_perfil({"nombre": nombre, "brand": datos["brand"], "aliases": _lista(datos["alias"]),
                              "voceros": _lista(datos["voceros"]), "criterio": datos["criterio"],
                              "criterio_custom": datos["criterio_texto"], "taxonomia": None,
                              "notas": datos.get("notas", "")})
        st.toast(f"Perfil de cliente guardado: {pid}")
    except ValueError as exc:
        st.error(str(exc))
        return False
    return True


def ajustes_finos(sufijo: str = "") -> dict:
    with st.expander("⚙ Ajustes finos del análisis (opcional)"):
        ca, cb = st.columns(2)
        with ca:
            ut = st.slider("Similitud de titulares (%)", 75, 100, 92, 1, key=f"umbral_titulo{sufijo}",
                           help="Revisión de consistencia: títulos con esta similitud o más se revisan como "
                                "posible mismo hecho (mismo tono y subtema). Bájalo para unir más.")
        with cb:
            uc = st.slider("Similitud de resúmenes (%)", 70, 100, 85, 1, key=f"umbral_cuerpo{sufijo}",
                           help="Lo mismo con el Contexto analizado de cada noticia.")
        st.text_input("Modelo de IA", value=MODELO, disabled=True, key=f"modelo_fijo{sufijo}",
                      help="El agente y sus ayudantes usan siempre claude-haiku-5-5 con la API key de los Secrets.")
        st.markdown("**💾 Perfil de cliente**")
        g = st.checkbox("Guardar esta configuración como perfil al procesar", key=f"guardar_chk{sufijo}",
                        help="Crea o actualiza el JSON del cliente (marca, alias, voceros, criterio y notas).")
        s = st.checkbox("Sobrescribir el perfil si ya existe uno con ese nombre", key=f"sobrescribir_chk{sufijo}",
                        help="Por seguridad, si el perfil ya existe y no marcas esta casilla, el proceso se "
                             "detiene con un aviso en lugar de reemplazarlo.")
        n = st.text_input("Nombre del perfil", key=f"nombre_perfil_input{sufijo}",
                          placeholder="Por defecto se usa la marca")
        notas = st.text_input("Notas del perfil (opcional)", key=f"notas_input{sufijo}")
    return {"umbral_titulo": ut, "umbral_cuerpo": uc, "guardar": g, "sobrescribir": s,
            "nombre_perfil": n, "notas": notas}


def campos_analisis(sufijo: str = "") -> dict:
    c_brand, c_alias = st.columns(2)
    with c_brand:
        brand = st.text_input("Marca o Cliente Principal*", key=f"brand_input{sufijo}",
                              placeholder="Ej: Universidad de Antioquia, Ecopetrol, Bancolombia",
                              help="El tono se mide solo sobre esta marca, sus voceros y sus alias.")
    with c_alias:
        alias = st.text_input("Alias o términos relacionados (separados por coma o punto y coma)",
                              key=f"alias_input{sufijo}", placeholder="Ej: UdeA; Alma Mater; rectoría",
                              help="Variantes del nombre que deban atribuirse al cliente.")
    c_crit, c_voc = st.columns([3, 2])
    with c_crit:
        criterio = st.radio("Criterio del tono", list(CRITERIOS_TONO.keys()), index=0, key=f"criterio_input{sufijo}",
                            help="Aspectual estricto: la crítica dirigida a la marca es lo único Negativo "
                                 "(gobiernos, alcaldías, entidades públicas). Favorabilidad del sector: cuenta "
                                 "cómo queda parado el sector aunque la marca no sea el actor (gremios).")
    with c_voc:
        voceros = st.text_input("Vocero(s) de la marca (opcional)", key=f"voceros_input{sufijo}",
                                placeholder="Ej: Gonzalo Moreno; el rector",
                                help="Personas cuyo nombre se atribuye a la marca para el tono.")
    texto = st.text_area("Criterio de tono personalizado (opcional)", key=f"criterio_custom_input{sufijo}",
                         height=80,
                         placeholder="Si lo llenas, este texto reemplaza al criterio del catálogo para este cliente.",
                         help="Texto libre con la regla de tono propia del cliente. Tiene prioridad sobre el "
                              "criterio seleccionado arriba. Haiku lo recibe literal.")
    return {"brand": (brand or "").strip(), "alias": alias or "", "voceros": voceros or "",
            "criterio": criterio, "criterio_texto": (texto or "").strip()}


def encolar(datos: dict, nombre: str, blob: bytes, clasificar: bool) -> None:
    if not datos["brand"]:
        st.error("Por favor indica la Marca o Cliente Principal para realizar el análisis enfocado.")
        return
    if not guardar_perfil_si_pide(datos):
        return
    st.session_state["pending_job"] = {**datos, "nombre": os.path.basename(nombre), "blob": blob,
                                       "clasificar": clasificar}
    st.session_state["procesando"] = True
    st.rerun()


# ======================================
# Pantallas
# ======================================
def pantalla_formulario() -> None:
    st.markdown(f"""
    <div class="app-header">
        <div class="app-header-icon">◈</div>
        <div class="app-header-text">
            <div class="app-header-title">Limpieza y Análisis de Noticias</div>
            <div class="app-header-version">{APP_VERSION}</div>
        </div>
        <div class="app-header-badge">Estructurador + Haiku</div>
    </div>""", unsafe_allow_html=True)
    fuente = "Google Sheets (Regiones / Internet)" if os.environ.get("REGIONES_CSV_URL") else "regiones sin configurar (quedan N/A)"
    st.markdown(f'<span class="config-badge">⚙ Configuración: {fuente}</span>', unsafe_allow_html=True)
    if st.session_state.get("error_proceso"):
        st.error(st.session_state.pop("error_proceso"))

    tab_std, tab_custom, tab_portales = st.tabs(["📋 Dossier estándar", "🗂 Columnas personalizadas",
                                                 "🌐 Comparar portales"])
    with tab_std:
        selector_perfil()
        with st.form("main_form"):
            st.markdown('<div class="sec-label">1. Sube el archivo de entrada</div>', unsafe_allow_html=True)
            st.markdown("""
            <div class="upload-zone-card"><div class="upload-zone-icon">📋</div><div>
              <div class="upload-zone-title">Dossier de Noticias</div>
              <div class="upload-zone-desc">Sube el .xlsx con las columnas Título y Resumen - Aclaracion.</div>
            </div></div>""", unsafe_allow_html=True)
            f1 = st.file_uploader("Dossier", type=["xlsx"], label_visibility="collapsed", key="f1")
            st.markdown('<div class="sec-label">2. Configuración de Análisis IA (Tono, Subtema)</div>',
                        unsafe_allow_html=True)
            datos = campos_analisis()
            datos.update(ajustes_finos())
            b1, b2 = st.columns([2, 1])
            iniciar = b1.form_submit_button("▶ Iniciar Limpieza y Análisis", width="stretch", type="primary")
            solo = b2.form_submit_button("Solo limpiar ($0)", width="stretch")
            if iniciar or solo:
                if not f1:
                    st.error("Por favor, sube un archivo Excel.")
                else:
                    encolar(datos, f1.name, f1.getvalue(), clasificar=bool(iniciar))

    with tab_custom:
        st.markdown('<div class="sec-label">1. Sube el archivo y elige las columnas</div>', unsafe_allow_html=True)
        st.caption("Usa esta pestaña cuando las columnas de tu archivo no se llamen «Título» ni «Resumen - "
                   "Aclaracion». La limpieza, el Contexto analizado y la clasificación son los mismos; solo "
                   "cambia qué columnas se usan como título y cuerpo.")
        f2 = st.file_uploader("Dossier con columnas personalizadas", type=["xlsx"],
                              label_visibility="collapsed", key="f2_columnas")
        cols = agente.limpieza.leer_columnas_xlsx(f2.getvalue()) if f2 is not None else []
        sel_t = sel_c = None
        if f2 is not None:
            if not cols:
                st.error("No se pudieron leer las columnas del archivo.")
            else:
                sug_t = agente.limpieza.sugerir_columna(cols, ("título", "titulo", "headline", "titular"))
                sug_c = agente.limpieza.sugerir_columna(cols, ("resumen - aclaracion", "resumen", "cuerpoes",
                                                               "cuerpo", "body", "texto", "contenido"))
                c_t, c_c = st.columns(2)
                sel_t = c_t.selectbox("Columna de título", cols, key="custom_col_titulo",
                                      index=cols.index(sug_t) if sug_t in cols else 0)
                sel_c = c_c.selectbox("Columna de cuerpo (CuerpoEs)", cols, key="custom_col_cuerpo",
                                      index=cols.index(sug_c) if sug_c in cols else min(1, len(cols) - 1))
        st.markdown('<div class="sec-label">2. Configuración del análisis (igual que la pestaña estándar)</div>',
                    unsafe_allow_html=True)
        datos_c = campos_analisis("_custom")
        datos_c.update(ajustes_finos("_custom"))
        b1, b2 = st.columns([2, 1])
        iniciar_c = b1.button("▶ Iniciar Limpieza y Análisis", width="stretch", type="primary",
                              key="btn_custom")
        solo_c = b2.button("Solo limpiar ($0)", width="stretch", key="btn_custom_limpiar")
        if iniciar_c or solo_c:
            if f2 is None or not cols:
                st.error("Por favor, sube un archivo Excel.")
            elif sel_t == sel_c:
                st.error("La columna de título y la de cuerpo deben ser diferentes.")
            else:
                try:
                    blob = agente.limpieza.renombrar_columnas_xlsx(f2.getvalue(), sel_t, sel_c)
                except ValueError as exc:
                    st.error(str(exc))
                else:
                    encolar(datos_c, f2.name, blob, clasificar=bool(iniciar_c))

    with tab_portales:
        pantalla_portales()


def pantalla_portales() -> None:
    st.markdown('<div class="sec-label">Cobertura del cliente en 5 portales</div>', unsafe_allow_html=True)
    st.caption("Lanza 5 ayudantes Haiku en paralelo, uno por portal. Cuesta unas 5 veces más que clasificar.")
    c1, c2, c3 = st.columns(3)
    cliente = c1.text_input("Marca o Cliente Principal*", key="brand_portales")
    alias = c2.text_input("Alias (opcional)", key="alias_portales")
    voceros = c3.text_input("Vocero(s) (opcional)", key="voceros_portales")
    if st.button("▶ Comparar portales", type="primary", disabled=not cliente.strip(), key="btn_portales"):
        comparativo = agente.ENTREGAS / "comparativo.md"
        comparativo.unlink(missing_ok=True)
        args = SimpleNamespace(modo="portales", cliente=cliente.strip(), alias=", ".join(_lista(alias)),
                               voceros=", ".join(_lista(voceros)), archivo=None)
        agente.opciones.max_budget_usd = agente.TOPE_PORTALES
        pasos = [("config", "Cargar configuración"), ("ai", "Buscar en 5 portales (Haiku)"),
                 ("export", "Escribir comparativo")]
        t0, ph, barra = time.time(), st.empty(), st.progress(0, text="Buscando en portales…")

        def vista(p):
            step = "export" if p.get("paso") == "Guardando" else "ai"
            ph.markdown(_render_live_html(60 if step == "ai" else 90, p.get("paso", ""),
                                          _fmt_elapsed(time.time() - t0), cliente, step, pasos),
                        unsafe_allow_html=True)
        final = asyncio.run(correr(agente.armar_prompt(args), vista, "portales"))
        ph.empty()
        barra.progress(100, text="Listo")
        (st.error if agente.FALLO[0] else st.success)(
            f"{'Terminó con errores' if agente.FALLO[0] else 'Comparativo listo'} · "
            f"${final.get('costo', 0):.3f} · {final.get('turnos', 0)} turnos")
        with st.expander("Ver detalle técnico"):
            st.code(final["log"][-20000:], language=None)
        if comparativo.exists():
            texto = comparativo.read_text(encoding="utf-8")
            st.markdown(texto)
            st.download_button("⬇ Descargar comparativo.md", texto, file_name="comparativo.md")


def pantalla_resultados() -> None:
    res = st.session_state["resultado"]
    r, c, final = res["r"], res["c"], res.get("final")
    con_ia = res["clasificar"]
    if con_ia and res.get("fallo"):
        st.error("El análisis con IA terminó con errores: revisa el detalle técnico abajo.")
    st.markdown(
        '<div class="success-banner"><div class="success-icon">✓</div>'
        '<div><div class="success-title">Proceso completado</div>'
        f'<div class="success-sub">{"El archivo estructurado y analizado con IA" if con_ia else "El archivo estructurado (sin IA, $0)"} '
        'se encuentra listo para descargar</div></div></div>', unsafe_allow_html=True)

    if r["medios_sin_region"]:
        st.warning("⚠️ Medios sin región asignada (quedaron N/A): " + ", ".join(r["medios_sin_region"]) + ".")

    if con_ia:
        piezas = [f"{res['filas']} filas clasificadas"] if res["filas"] else []
        piezas += [f"{k}: {v}" for k, v in res["tonos"].items()]
        piezas.append(f"criterio: {res['criterio']}")
        if res.get("revision") is not None:
            piezas.append(f"{res['revision']} correcciones de consistencia (revisor-subtemas)")
        st.info("Análisis de Tono/Sub-tema · " + " · ".join(piezas))
    st.caption(f"Contexto analizado: {c['con_mencion']} con mención del cliente, {c['desde_lead']} desde el "
               f"lead, {c['solo_titulo']} solo título · {r['filas_expandidas']} filas nuevas por menciones.")

    if con_ia and final:
        costo = final.get("costo", 0.0)
        costo_txt = f"${costo:.2f}" if costo >= 0.01 else f"${costo:.4f}"
        det_costo = f"{final.get('turnos', 0)} turnos · {MODELO}"
    else:
        costo_txt, det_costo = "$0", ""
    st.markdown(f"""
    <div class="metrics-grid">
      <div class="metric-card m-total"><div class="metric-val" style="color:var(--text)">{r['total']}</div><div class="metric-lbl">Total Registros</div></div>
      <div class="metric-card m-unique"><div class="metric-val" style="color:var(--green)">{r['unicos']}</div><div class="metric-lbl">Únicos</div></div>
      <div class="metric-card m-dup"><div class="metric-val" style="color:var(--amber)">{r['duplicados']}</div><div class="metric-lbl">Duplicados</div></div>
      <div class="metric-card m-time"><div class="metric-val" style="color:var(--blue)">{res['duracion']}</div><div class="metric-lbl">Tiempo de Ejecución</div></div>
      <div class="metric-card m-cost" title="{det_costo}"><div class="metric-val" style="color:var(--amber)">{costo_txt}</div><div class="metric-lbl">Costo IA aprox.</div></div>
    </div>""", unsafe_allow_html=True)
    if det_costo:
        st.caption("Costo del análisis: " + det_costo + ".")

    dups = r.get("duplicadas") or []
    with st.expander(f"Noticias duplicadas ({len(dups)}) · ID duplicada → ID original", expanded=bool(dups)):
        if dups:
            st.caption("Cada fila duplicada lleva en «ID duplicada» el ID de su original y hereda su tono y "
                       "subtema. Regla de Grill: misma URL + misma mención (prensa, internet) o mismo medio + "
                       "misma hora + misma mención (radio y TV).")
            st.dataframe(pd.DataFrame(dups)[["ID duplicada", "ID original", "Título", "Medio", "Tipo de Medio",
                                             "Motivo", "Mención", "Fila"]],
                         hide_index=True, width="stretch")
        else:
            st.caption("No hay noticias duplicadas en este archivo.")

    c1, c2 = st.columns(2)
    if res.get("salida"):
        c1.download_button("⬇ Descargar Xlsx Estructurado con IA" if con_ia else "⬇ Descargar Xlsx limpio",
                           data=res["salida"], file_name=res["salida_nombre"], mime=XLSX_MIME,
                           width="stretch", type="primary")
    else:
        c1.error("No se generó el archivo clasificado.")
    if c2.button("Nuevo análisis", width="stretch"):
        pwd = st.session_state.get("password_correct")
        st.session_state.clear()
        st.session_state["password_correct"] = pwd
        st.rerun()
    if final:
        with st.expander("Ver detalle técnico"):
            st.code(final["log"][-20000:], language=None)


# ======================================
# Main
# ======================================
if st.session_state.get("pending_job"):
    job = st.session_state.pop("pending_job")
    try:
        run_cleaning_process(job)
    finally:
        st.session_state["procesando"] = False
    st.rerun()
elif st.session_state.get("processing_complete") and st.session_state.get("resultado"):
    pantalla_resultados()
else:
    pantalla_formulario()

st.markdown('<div class="footer">Estructuración y Limpieza · clausifier · Haiku</div>', unsafe_allow_html=True)
