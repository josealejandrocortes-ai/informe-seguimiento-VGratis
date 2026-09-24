"""Aplicación web (Streamlit) para diligenciar el Informe Seguimiento Semanal.

Ejecutar:  streamlit run app.py
"""
import datetime as dt
import hashlib
import os

import streamlit as st

from informe_semanal.engine import (PROVEEDOR_DEFECTO, PROVEEDORES, Contexto, Libro, adivinar_fecha_corte,
                                    analizar, construir, contexto_para_ia, escanear)

st.set_page_config(page_title="Informe Seguimiento Semanal", page_icon="📊", layout="wide")
st.title("📊 Informe Seguimiento Semanal — cadena presupuestal")
st.caption("Sube la plantilla de Word y el Excel de soporte; la aplicación completa las cifras, tablas, "
           "gráficos y análisis conservando el formato.")

# ---------------- Barra lateral ----------------
with st.sidebar:
    st.header("Configuración")
    st.markdown("**Modelo de IA para los análisis <…>**")
    nombres = list(PROVEEDORES)
    proveedor = st.selectbox("Proveedor", nombres, index=nombres.index(PROVEEDOR_DEFECTO),
                             label_visibility="collapsed")
    cfg = PROVEEDORES[proveedor]
    st.caption(("🟢 Plan gratuito. " if cfg["gratis"] else "💲 ") + cfg["nota"])
    api_key = st.text_input(f"API key ({proveedor})", type="password",
                            value=os.environ.get(cfg["env"], "") if cfg["env"] else "",
                            disabled=not cfg["requiere_key"], key=f"key_{proveedor}")
    modelo = st.text_input("Modelo", value=cfg["modelo"], key=f"mod_{proveedor}")
    url_base = st.text_input("URL base de la API", value=cfg["url"], key=f"url_{proveedor}",
                             disabled=cfg["tipo"] == "anthropic")
    st.markdown("**Reemplazos de texto** (uno por línea: `original => nuevo`)")
    glos_txt = st.text_area(
        "glosario", label_visibility="collapsed", height=110,
        value="Despacho del Ministro => Despacho de la Ministra\nDESPACHO DEL MINISTRO => DESPACHO DE LA MINISTRA")
    if proveedor.startswith("Ollama"):
        st.info("Con Ollama los datos no salen de tu computador.")
    else:
        st.info("Los análisis <…> envían al proveedor elegido los rangos citados en la instrucción y un resumen "
                "calculado (no el archivo completo). Sin clave, esos apartados quedan pendientes.")

glosario = {}
for linea in glos_txt.splitlines():
    if "=>" in linea:
        a, b = linea.split("=>", 1)
        if a.strip():
            glosario[a.strip()] = b.strip()

# ---------------- 1. Archivos ----------------
c1, c2 = st.columns(2)
f_word = c1.file_uploader("1️⃣ Plantilla del informe (Word .docx)", type=["docx"])
f_xl = c2.file_uploader("2️⃣ Excel de soporte (.xlsm / .xlsx)", type=["xlsm", "xlsx"])
if not (f_word and f_xl):
    st.info("Sube ambos archivos para continuar. Recuerda actualizar las tablas dinámicas de 'TD Maestro' y "
            "**guardar** el Excel antes de subirlo: la aplicación lee los valores guardados.")
    st.stop()

b_word, b_xl = f_word.getvalue(), f_xl.getvalue()


@st.cache_resource(show_spinner="Leyendo el Excel…")
def cargar(h: str, datos: bytes):
    return Libro(datos)


lib = cargar(hashlib.md5(b_xl).hexdigest(), b_xl)
fecha = st.date_input("Fecha de corte", value=adivinar_fecha_corte(lib), format="DD/MM/YYYY",
                      help="Se deduce de 'Sem dd Mmm' en la hoja de ejecución; corrígela si no coincide.")
ctx = Contexto(lib, fecha, glosario)
esc = escanear(b_word, ctx)

# ---------------- 2. Revisión de marcadores ----------------
malos = [v for v in esc["valores"] if v["error"]]
st.subheader("3️⃣ Marcadores {…}")
m1, m2, m3 = st.columns(3)
m1.metric("Marcadores de valor", len(esc["valores"]))
m2.metric("Con error", len(malos), delta_color="inverse")
m3.metric("Análisis <…>", len(esc["analisis"]))
if malos:
    st.error("Hay marcadores que no se pudieron resolver (se dejarán tal cual en el Word):")
    st.dataframe([{"Marcador": v["marcador"], "Error": v["error"]} for v in malos], use_container_width=True)
with st.expander("Ver todos los valores que se van a insertar"):
    st.dataframe([{"Marcador": v["marcador"], "Valor": v["valor"] if not v["error"] else "⚠ " + v["error"]}
                  for v in esc["valores"]], use_container_width=True, height=320)

# ---------------- 3. Análisis ----------------
analisis = {}
if esc["analisis"]:
    st.subheader("4️⃣ Análisis <…>")
    estado = st.session_state.setdefault("analisis", {})
    listo = api_key or not cfg["requiere_key"]
    if st.button(f"✨ Generar análisis con {proveedor.split(' (')[0]}", disabled=not listo, type="secondary"):
        with st.spinner("Redactando análisis…"):
            for it in esc["analisis"]:
                try:
                    estado[it["instruccion"]] = analizar(ctx, it["instruccion"], proveedor, api_key, modelo, url_base)
                except Exception as e:  # noqa: BLE001
                    st.error(f"No se pudo generar: {e}")
                    break
    if not listo:
        st.warning("Ingresa la API key en la barra lateral para generar los análisis (o escríbelos a mano abajo).")
    else:
        st.caption("Revisa siempre los textos: los modelos gratuitos o pequeños pueden equivocarse con las cifras.")
    for i, it in enumerate(esc["analisis"]):
        with st.container(border=True):
            st.markdown(f"**Instrucción:** {it['instruccion']}")
            estado.setdefault(it["instruccion"], "")
            txt = st.text_area("Texto que se insertará (editable)", value=estado[it["instruccion"]],
                               key=f"an_{i}_{hash(it['instruccion'])}", height=110)
            estado[it["instruccion"]] = txt
            analisis[it["instruccion"]] = txt
            with st.expander("Datos que se envían como contexto"):
                st.code(contexto_para_ia(ctx, it["instruccion"]))

# ---------------- 4. Generar ----------------
st.subheader("5️⃣ Generar informe")
pend = [i for i, t in analisis.items() if not t.strip()]
if pend:
    st.warning(f"{len(pend)} análisis sin texto: quedarán como marcador <…> en el Word.")
if st.button("📄 Generar informe", type="primary"):
    datos, info = construir(b_word, ctx, {k: v for k, v in analisis.items() if v.strip()})
    st.success(f"Informe generado: {info['reemplazados']} reemplazos y {len(info['graficos'])} gráficos actualizados.")
    for s in info["sin_resolver"]:
        st.error(f"Sin resolver: {s}")
    st.download_button("⬇️ Descargar informe (.docx)", data=datos, mime=
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       file_name=f"{fecha.strftime('%d%m%Y')}_Informe_Seguimiento_Semanal.docx")
