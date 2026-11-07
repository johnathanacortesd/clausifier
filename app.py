# app.py — interfaz web de clausifier (Streamlit).
#
# Secrets necesarios (Streamlit Cloud: Settings → Secrets; local: .streamlit/secrets.toml):
#     ANTHROPIC_API_KEY = "sk-ant-..."
#     APP_PASSWORD      = "..."
# Correr en local:  uv run streamlit run app.py

import asyncio
import contextlib
import hmac
import io
import os
import shutil
from types import SimpleNamespace

import streamlit as st

st.set_page_config(page_title="clausifier", page_icon="📰", layout="wide")


# ─── Contraseña ───────────────────────────────────────────────────────────
def acceso_ok() -> bool:
    if st.session_state.get("acceso"):
        return True
    esperada = st.secrets.get("APP_PASSWORD", "")
    if not esperada:
        st.error("Falta APP_PASSWORD en los secrets.")
        return False
    clave = st.text_input("Contraseña", type="password")
    if clave:
        if hmac.compare_digest(clave.encode(), str(esperada).encode()):
            st.session_state["acceso"] = True
            st.rerun()
        st.error("Contraseña incorrecta.")
    return False


if not acceso_ok():
    st.stop()


# ─── Credencial: solo la API key de los secrets ──────────────────────────
clave_api = st.secrets.get("ANTHROPIC_API_KEY", "")
if not clave_api:
    st.error("Falta ANTHROPIC_API_KEY en los secrets.")
    st.stop()
for otra in ("CLAUDE_CODE_OAUTH_TOKEN", "ANTHROPIC_PROFILE"):
    os.environ.pop(otra, None)
os.environ["ANTHROPIC_API_KEY"] = clave_api
for opcional in ("REGIONES_CSV_URL", "INTERNET_CSV_URL"):   # mapas de región (opcionales)
    if st.secrets.get(opcional):
        os.environ[opcional] = str(st.secrets[opcional])

import agente  # noqa: E402 — después de fijar la credencial

from claude_agent_sdk import ClaudeSDKError, query  # noqa: E402


async def correr(prompt: str, salida) -> None:
    buffer = io.StringIO()
    agente.FALLO[0] = False
    with contextlib.redirect_stdout(buffer):
        try:
            async for mensaje in query(prompt=prompt, options=agente.opciones):
                agente.imprimir(mensaje)
                agente.imprimir_error_de_tool(mensaje)
                salida.code(buffer.getvalue()[-6000:], language=None)
        except ClaudeSDKError as error:
            print(f"\n⚠️  el agente paró: {error}")
            agente.FALLO[0] = True
    salida.code(buffer.getvalue()[-6000:], language=None)


# ─── Formulario ───────────────────────────────────────────────────────────
st.title("📰 clausifier")
st.caption("Clasifica noticias (Tono y Subtema) y compara la cobertura de un cliente en 5 portales.")

cliente = st.text_input("Cliente (marca) *")
alias = st.text_input("Alias (opcional, separados por coma)")
voceros = st.text_input("Voceros (opcional, separados por coma)")

tab_xlsx, tab_portales = st.tabs(["Clasificar xlsx", "Comparar portales"])

with tab_xlsx:
    subido = st.file_uploader("Archivo xlsx", type=["xlsx"])
    st.caption("Pipeline: limpieza de API_2_Grill → Contexto analizado y texto (sin IA, $0) → "
               "el agente clasifica Tono y Subtema leyendo solo el Contexto analizado.")

    def guardar_subido() -> str:
        datos = agente.WORKSPACE / "datos"
        datos.mkdir(parents=True, exist_ok=True)
        nombre = os.path.basename(subido.name)
        (datos / nombre).write_bytes(subido.getvalue())
        return nombre

    def limpiar_y_mostrar(nombre: str):
        try:
            limpio = agente.ENTREGAS / f"{os.path.splitext(nombre)[0]}_limpio.xlsx"
            r = agente.limpieza.limpiar_dossier(agente.WORKSPACE / "datos" / nombre, limpio)
            c = agente.contexto.agregar_contexto(limpio, cliente, agente.separar(alias),
                                                 agente.separar(voceros))
        except Exception as e:  # noqa: BLE001 — p. ej. más de 1000 filas
            st.error(f"No se pudo limpiar: {e}")
            return None
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Filas totales", r["total"], help=f"{r['filas_expandidas']} filas nuevas por menciones")
        c2.metric("Únicas", r["unicos"])
        c3.metric("Duplicadas", r["duplicados"])
        c4.metric("Medios sin región", len(r["medios_sin_region"]))
        if r["medios_sin_region"]:
            st.caption("Sin región: " + ", ".join(r["medios_sin_region"][:30]))
        st.caption(f"Contexto analizado: {c['con_mencion']} con mención del cliente, "
                   f"{c['desde_lead']} desde el lead, {c['solo_titulo']} solo título.")
        return limpio

    b1, b2 = st.columns(2)
    solo_limpiar = b2.button("Solo limpiar ($0)", disabled=not (cliente and subido))
    clasificar = b1.button("Clasificar", type="primary", disabled=not (cliente and subido))
    if (subido and not cliente):
        st.info("Escribe el cliente: el Contexto analizado se arma con sus menciones.")
    if solo_limpiar:
        limpio = limpiar_y_mostrar(guardar_subido())
        if limpio and limpio.exists():
            st.download_button("⬇️ Descargar xlsx limpio", limpio.read_bytes(), file_name=limpio.name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if clasificar:
        nombre = guardar_subido()
        destino = agente.ENTREGAS / f"{os.path.splitext(nombre)[0]}_clasificado.xlsx"
        destino.unlink(missing_ok=True)                    # no reutilizar una corrida anterior
        limpio = limpiar_y_mostrar(nombre)
        if limpio is None:
            st.stop()
        args = SimpleNamespace(modo="xlsx", cliente=cliente, alias=alias, voceros=voceros,
                               archivo=f"entregas/{limpio.name}")
        agente.opciones.max_budget_usd = 0.40
        with st.status("Clasificando…", expanded=True) as estado:
            asyncio.run(correr(agente.armar_prompt(args), st.empty()))
            estado.update(label="Terminó con errores" if agente.FALLO[0] else "Listo",
                          state="error" if agente.FALLO[0] else "complete")
        if destino.exists():
            st.download_button("⬇️ Descargar xlsx clasificado", destino.read_bytes(),
                               file_name=destino.name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab_portales:
    st.caption("Lanza 5 ayudantes en paralelo, uno por portal. Cuesta unas 5 veces más que clasificar.")
    if st.button("Comparar portales", type="primary", disabled=not cliente):
        comparativo = agente.ENTREGAS / "comparativo.md"
        comparativo.unlink(missing_ok=True)
        args = SimpleNamespace(modo="portales", cliente=cliente, alias=alias, voceros=voceros,
                               archivo=None)
        agente.opciones.max_budget_usd = agente.TOPE_PORTALES
        with st.status("Buscando en portales…", expanded=True) as estado:
            asyncio.run(correr(agente.armar_prompt(args), st.empty()))
            estado.update(label="Terminó con errores" if agente.FALLO[0] else "Listo",
                          state="error" if agente.FALLO[0] else "complete")
        if comparativo.exists():
            texto = comparativo.read_text(encoding="utf-8")
            st.markdown(texto)
            st.download_button("⬇️ Descargar comparativo.md", texto, file_name="comparativo.md")
