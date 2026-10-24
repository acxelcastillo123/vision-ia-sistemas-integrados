# -*- coding: utf-8 -*-
"""
TALLER 5 - Inspección de contenedores tapados · Dashboard tipo MES
Ejecutar con:   streamlit run app.py
Origen único de imágenes: carpeta «imagenes/» de esta misma carpeta.
"""
import csv
import io
import random
import textwrap
import time
from datetime import datetime
from pathlib import Path

import cv2
import matplotlib
import numpy as np
import pandas as pd

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import streamlit as st  # noqa: E402

import motor_vision as motor  # noqa: E402

BASE = Path(__file__).resolve().parent
CARPETA_IMAGENES = BASE / "imagenes"
EXCEL = BASE / "historico_inspecciones.xlsx"
CSV_OP = BASE / "etiquetas_operador.csv"
EXTENSIONES = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}
VERDE, ROJO, AMBAR, GRIS, AZUL = "#2e7d32", "#c62828", "#ef6c00", "#455a64", "#1f4e78"

st.set_page_config(page_title="Taller 5 · Inspección de contenedores", layout="wide")

st.markdown("""
<style>
.mes-cab{background:linear-gradient(90deg,#12344d,#1f4e78);color:#fff;padding:14px 20px;
border-radius:8px;margin-bottom:10px}
.mes-cab h2{margin:0;color:#fff;font-size:1.5rem}
.mes-cab small{color:#cfe3f3}
.isa{display:flex;gap:6px;margin:6px 0 14px 0}
.isa div{flex:1;padding:7px 10px;border-radius:6px;background:#eceff1;color:#455a64;
font-size:.78rem;border:1px solid #cfd8dc}
.isa div.act{background:#1f4e78;color:#fff;border-color:#1f4e78;font-weight:600}
.kpi{background:#f4f6f8;border-left:6px solid #455a64;border-radius:6px;padding:10px 14px;
min-height:96px}
.kpi .t{color:#546e7a;font-size:.78rem;text-transform:uppercase;letter-spacing:.04em}
.kpi .v{color:#102027;font-size:1.9rem;font-weight:700;line-height:1.15}
.kpi .s{color:#607d8b;font-size:.78rem}
.andon{border-radius:8px;padding:12px 18px;color:#fff;font-weight:700;font-size:1.05rem;
margin:10px 0}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Estado de sesión
# ---------------------------------------------------------------------------
ss = st.session_state
ss.setdefault("resultados", [])
ss.setdefault("vistas", {})
ss.setdefault("lote", None)
ss.setdefault("msg_excel", None)
ss.setdefault("excel_ok", True)
ss.setdefault("ent_orden", [])
ss.setdefault("ent_resp", {})
ss.setdefault("ent_causas", {})
ss.setdefault("ent_desc", {})
ss.setdefault("ent_pila", [])
ss.setdefault("ent_calib", None)
ss.setdefault("validacion", None)
ss.setdefault("val_op", None)

# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def cargar_items():
    """Lee únicamente las imágenes de la carpeta «imagenes/» del proyecto."""
    items = {}
    if CARPETA_IMAGENES.exists():
        for p in sorted(CARPETA_IMAGENES.iterdir()):
            if p.suffix.lower() in EXTENSIONES:
                items[p.name] = p.read_bytes()
    return items

def rgb(bgr):
    return cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)

def guardar_en_excel(registros):
    """Agrega los registros al histórico SIN sobrescribir lo anterior."""
    try:
        motor.guardar_excel(EXCEL, registros)
        ss.msg_excel = (f"{len(registros)} registros agregados a «{EXCEL.name}» "
                        f"(el histórico anterior se conserva).")
        ss.excel_ok = True
    except PermissionError:
        ss.msg_excel = ("No se pudo escribir en el Excel porque está abierto. Ciérrelo y pulse "
                        "«Guardar lote en Excel».")
        ss.excel_ok = False

def estilo_resultado(col):
    return ["background-color:#C6EFCE;color:#006100;font-weight:bold" if v == "APROBADO"
            else "background-color:#FFC7CE;color:#9C0006;font-weight:bold" for v in col]

def kpi(titulo, valor, sub=" ", color=GRIS):
    return (f'<div class="kpi" style="border-left-color:{color}"><div class="t">{titulo}</div>'
            f'<div class="v">{valor}</div><div class="s">{sub}</div></div>')

def guardar_csv_operador():
    with open(CSV_OP, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["imagen", "etiqueta", "causa", "descripcion"])
        for n in sorted(ss.ent_resp):
            w.writerow([n, "APROBADO" if ss.ent_resp[n] else "RECHAZADO",
                        ss.ent_causas.get(n, ""), ss.ent_desc.get(n, "")])

def leer_csv_operador():
    out = {}
    if CSV_OP.exists():
        with open(CSV_OP, newline="", encoding="utf-8-sig") as f:
            for fila in csv.DictReader(f):
                v = (fila.get("etiqueta") or "").strip().upper()
                if v in ("APROBADO", "RECHAZADO"):
                    out[fila["imagen"].strip()] = {
                        "apta": v == "APROBADO",
                        "causa": (fila.get("causa") or "").strip(),
                        "descripcion": (fila.get("descripcion") or "").strip()}
    return out

ITEMS = cargar_items()

# ---------------------------------------------------------------------------
# Barra lateral
# ---------------------------------------------------------------------------
st.sidebar.title("Configuración")
st.sidebar.info("Se evaluarán todas las imágenes contenidas en la carpeta «imagenes/» "
                "de este proyecto (formatos png, jpg, bmp, tif). Copie allí las "
                "fotografías de la cámara industrial antes de iniciar.")
st.sidebar.caption(f"{len(ITEMS)} imágenes disponibles en «imagenes/»")

st.sidebar.subheader("Identificación de la evaluación")
estacion = st.sidebar.text_input(
    "Nombre de la estación o línea evaluada", value="EST-INSP-01",
    help="Escriba aquí el nombre de la estación que se está evaluando para que el "
         "dashboard y el Excel histórico muestren claramente a qué estación "
         "pertenecen los resultados.")
operador = st.sidebar.text_input(
    "Operador", value="Operador 1",
    help="Escribe el nombre del operador de turno; queda registrado en el histórico "
         "de inspecciones.")
meta = st.sidebar.slider("Meta de calidad: rendimiento (%)", 50.0, 100.0, 95.0, 1.0,
                         help="Porcentaje mínimo de piezas aprobadas que se espera del "
                              "proceso de tapado.")

st.sidebar.subheader("Ejecución")
pausa = st.sidebar.slider("Pausa entre imágenes (s)", 0.0, 2.0, 0.4, 0.1,
                          help="Pausa visual para seguir la inspección pieza por pieza; "
                               "no afecta el resultado.")
semilla = st.sidebar.number_input(
    "Repetir el mismo orden de imágenes", 0, 99999, 0,
    help="Escriba un número fijo (por ejemplo 7) para que las imágenes se procesen "
         "siempre en el mismo orden, útil para comparar cambios. Con 0 el orden es "
         "aleatorio en cada ejecución, como exige el taller.")

_cal = motor.cargar_calibracion()
if _cal:
    st.sidebar.success(f"Calibración activa: umbrales ajustados con respuestas del "
                       f"operador (exactitud {100 * _cal['exactitud_loocv']:.0f} %).")
else:
    st.sidebar.info("El sistema trabaja con los umbrales de fábrica, calibrados con las "
                    "imágenes de referencia del taller.")

def _lim(v, lo, hi):
    return float(min(max(v, lo), hi))

with st.sidebar.expander("Umbrales de inspección"):
    st.sidebar.caption("Valores máximos que debe cumplir una pieza bien tapada. "
                       "Modifíquelos solo si conoce el proceso o después de calibrar en "
                       "la pestaña Entrenamiento asistido.")
    umbrales = {
        "inclinacion": st.slider("Inclinación máx. de la tapa (°)", 0.5, 15.0,
                                 _lim(motor.UMBRALES["inclinacion"], 0.5, 15.0), 0.1),
        "altura": st.slider("Altura máx. de tapa sellada (px)", 30.0, 130.0,
                            _lim(motor.UMBRALES["altura"], 30.0, 130.0), 0.5),
        "desalineacion": st.slider("Desalineación máx. tapa/contenedor (px)", 10.0, 100.0,
                                   _lim(motor.UMBRALES["desalineacion"], 10.0, 100.0), 0.5),
    }

iniciar = st.sidebar.button("Iniciar inspección", type="primary")
if st.sidebar.button("Limpiar resultados de la sesión"):
    ss.resultados, ss.vistas, ss.msg_excel, ss.lote = [], {}, None, None
    st.rerun()

# ---------------------------------------------------------------------------
# Encabezado y pestañas
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="mes-cab"><h2>Control de calidad por visión de computadora · Taller 5</h2>'
    '<small>Inspección de tapado de contenedores</small></div>', unsafe_allow_html=True)

tab_insp, tab_mes, tab_ent, tab_val = st.tabs(
    ["Inspección", "Dashboard MES", "Entrenamiento asistido", "Validación del modelo"])

# ---------------------------------------------------------------------------
# PESTAÑA 1 · Inspección automática
# ---------------------------------------------------------------------------
def ejecutar_lote(items, ph_img, ph_info, barra):
    """Procesa TODAS las imágenes en orden aleatorio, sin intervención manual."""
    if semilla:
        random.seed(int(semilla))
    orden = list(items)
    random.shuffle(orden)
    ss.resultados, ss.vistas, ss.msg_excel = [], {}, None
    ss.lote = datetime.now().strftime("L%Y%m%d-%H%M%S")
    for n, nombre in enumerate(orden, 1):
        try:
            img = motor.decodificar(items[nombre])
        except ValueError:
            continue
        reg, vista = motor.inspeccionar(nombre, img, ss.lote, umbrales,
                                        estacion=estacion, operador=operador)
        ss.resultados.append(reg)
        ss.vistas[nombre] = rgb(vista)
        ph_img.image(ss.vistas[nombre], caption=f"{n}/{len(orden)} · {nombre}")
        with ph_info.container():
            (st.success if reg["Resultado"] == "APROBADO" else st.error)(reg["Resultado"])
            st.write(f"**Lote:** {ss.lote} · **Imagen:** {nombre}")
            st.write(f"**Criterios evaluados:** contenedor = {reg['Contenedor']}, "
                     f"tapa = {reg['Tapa']}, inclinación = {reg['Inclinación (°)']}°, "
                     f"altura = {reg['Altura (px)']} px, "
                     f"desalineación = {reg['Desalineación (px)']} px")
            st.write(f"**Justificación técnica:** {reg['Justificación Técnica']}")
        barra.progress(n / len(orden), text=f"Procesando {n} de {len(orden)}")
        time.sleep(pausa)
    guardar_en_excel(ss.resultados)
    st.rerun()

with tab_insp:
    st.caption("Pulse «Iniciar inspección» en la barra lateral para procesar todas las "
               "imágenes de la carpeta «imagenes/» en orden aleatorio. Para cada pieza "
               "verá la fotografía con sus recuadros de detección, el resultado, los "
               "criterios medidos y la justificación técnica.")
    barra = st.empty()
    col_img, col_info = st.columns([3, 2])
    ph_img, ph_info = col_img.empty(), col_info.empty()
    if iniciar:
        if not ITEMS:
            st.warning("No hay imágenes para procesar. Copie las fotografías en la "
                       "carpeta «imagenes/» de este proyecto.")
        else:
            barra = barra.progress(0.0, text="Iniciando...")
            ejecutar_lote(ITEMS, ph_img, ph_info, barra)
    elif ss.resultados:
        ult = ss.resultados[-1]
        ph_img.image(ss.vistas[ult["Imagen"]], caption=f"Última pieza: {ult['Imagen']}")
        with ph_info.container():
            (st.success if ult["Resultado"] == "APROBADO" else st.error)(ult["Resultado"])
            st.write(f"**Justificación técnica:** {ult['Justificación Técnica']}")
        if ss.msg_excel:
            (st.success if ss.excel_ok else st.error)(ss.msg_excel)
        if not ss.excel_ok and st.button("Guardar lote en Excel"):
            guardar_en_excel(ss.resultados)
            st.rerun()
        st.info("Para visualizar el resumen estadístico completo ve a la pestaña "
                "Dashboard MES.")
        with st.expander("Ver todas las piezas inspeccionadas"):
            cols = st.columns(4)
            for i, reg in enumerate(ss.resultados):
                cols[i % 4].image(ss.vistas[reg["Imagen"]], caption=reg["Imagen"])
    else:
        st.info("Para comenzar, pulse «Iniciar inspección» en la barra lateral. Se "
                "procesarán todas las imágenes en orden aleatorio, sin intervención "
                "manual.")

# ---------------------------------------------------------------------------
# PESTAÑA 2 · Dashboard MES
# ---------------------------------------------------------------------------
def preparar(df):
    df = df.copy()
    for c in ("Inclinación (°)", "Altura (px)", "Desalineación (px)"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["_ts"] = pd.to_datetime(df["Fecha"].astype(str) + " " + df["Hora"].astype(str),
                               errors="coerce")
    df = df.sort_values("_ts", kind="stable").reset_index(drop=True)
    df["N°"] = np.arange(1, len(df) + 1)
    if "Causa de rechazo" not in df:
        df["Causa de rechazo"] = "—"
    df["Causa de rechazo"] = df["Causa de rechazo"].fillna("—").astype(str)
    return df

def disposicion(fila):
    if fila["Resultado"] == "APROBADO":
        return "LIBERAR"
    causa = fila["Causa de rechazo"].split(",")[0].strip()
    return "RETENER · " + motor.ACCIONES.get(causa, "Revisión manual")

def _base_ax(ax):
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.grid(axis="y", alpha=.25)

def fig_donut(ap, re_, fpy):
    fig, ax = plt.subplots(figsize=(4.4, 3.6))
    ax.pie([ap, re_], colors=[VERDE, ROJO], startangle=90, counterclock=False,
           wedgeprops=dict(width=.36, edgecolor="white"))
    ax.text(0, .06, f"{fpy:.1f}%", ha="center", va="center", fontsize=20, fontweight="bold")
    ax.text(0, -.2, "rendimiento", ha="center", va="center", fontsize=9, color=GRIS)
    ax.legend([f"Aprobadas ({ap})", f"Rechazadas ({re_})"], loc="lower center",
              bbox_to_anchor=(.5, -.12), ncol=2, frameon=False, fontsize=8)
    ax.set_title("Resultado de la inspección", fontsize=11)
    return fig

def fig_pareto(df):
    causas = (df.loc[df["Resultado"] == "RECHAZADO", "Causa de rechazo"]
              .str.split(",").explode().str.strip().replace("—", np.nan)
              .dropna().value_counts())
    fig, ax = plt.subplots(figsize=(7.4, 3.6))
    if causas.empty:
        ax.text(.5, .5, "Sin rechazos", ha="center", va="center")
        ax.axis("off")
        return fig
    etiquetas = [textwrap.fill(c, 14) for c in causas.index]
    ax.bar(etiquetas, causas.values, color=AMBAR)
    for i, v in enumerate(causas.values):
        ax.text(i, v + .1, str(v), ha="center", fontsize=9)
    ax2 = ax.twinx()
    ax2.plot(etiquetas, causas.cumsum() / causas.sum() * 100, color=AZUL, marker="o")
    ax2.set_ylim(0, 105)
    ax2.set_ylabel("% acumulado")
    ax2.spines["top"].set_visible(False)
    _base_ax(ax)
    ax.set_ylabel("Piezas")
    ax.set_title("Pareto de causas de rechazo", fontsize=11)
    ax.tick_params(axis="x", labelsize=8)
    return fig

def fig_spc(df, col, limite, titulo, unidad):
    fig, ax = plt.subplots(figsize=(7.4, 3.3))
    d = df[df[col].notna()]
    if d.empty:
        ax.text(.5, .5, "Sin mediciones", ha="center", va="center")
        ax.axis("off")
        return fig
    top = max(float(d[col].max()), limite) * 1.18
    ax.axhspan(0, limite, color=VERDE, alpha=.08)
    ax.axhline(limite, color=ROJO, ls="--", lw=1.4)
    ax.text(d["N°"].max(), limite, f" LSE {limite:g} {unidad}", color=ROJO, va="bottom",
            ha="right", fontsize=8)
    ax.plot(d["N°"], d[col], color="#b0bec5", lw=1, zorder=1)
    ax.scatter(d["N°"], d[col], s=34, zorder=3, edgecolor="white", linewidth=.6,
               c=np.where(d["Resultado"] == "APROBADO", VERDE, ROJO))
    ax.set_ylim(0, top)
    ax.set_xlabel("N° de pieza (orden de inspección)")
    ax.set_ylabel(unidad)
    ax.set_title(titulo, fontsize=11)
    _base_ax(ax)
    return fig

def fig_rendimiento(df, meta_):
    fig, ax = plt.subplots(figsize=(7.4, 3.3))
    ok = (df["Resultado"] == "APROBADO").cumsum() / df["N°"] * 100
    ax.plot(df["N°"], ok, color=AZUL, lw=2)
    ax.axhline(meta_, color=VERDE, ls="--")
    ax.text(df["N°"].max(), meta_, f" meta {meta_:g}%", color=VERDE, ha="right",
            va="bottom", fontsize=8)
    ax.set_ylim(0, 105)
    ax.set_xlabel("N° de pieza")
    ax.set_ylabel("% acumulado")
    ax.set_title("Rendimiento acumulado (First Pass Yield)", fontsize=11)
    _base_ax(ax)
    return fig

def fig_lotes(df, meta_):
    g = df.groupby("Lote", sort=False)["Resultado"].agg(
        lambda s: (s == "APROBADO").mean() * 100)
    fig, ax = plt.subplots(figsize=(7.4, 3.3))
    colores = [VERDE if v >= meta_ else (AMBAR if v >= meta_ - 10 else ROJO) for v in g.values]
    ax.bar(range(len(g)), g.values, color=colores)
    ax.set_xticks(range(len(g)))
    ax.set_xticklabels([str(i) for i in g.index], rotation=30, ha="right", fontsize=7)
    ax.axhline(meta_, color=GRIS, ls="--")
    ax.set_ylim(0, 105)
    ax.set_ylabel("% aprobadas")
    ax.set_title("Rendimiento por lote", fontsize=11)
    _base_ax(ax)
    return fig

def mostrar(fig):
    st.pyplot(fig)
    plt.close(fig)

def mostrar_mes(df, es_historico):
    total = len(df)
    ap = int((df["Resultado"] == "APROBADO").sum())
    re_ = total - ap
    fpy = 100 * ap / total
    color_fpy = VERDE if fpy >= meta else (AMBAR if fpy >= meta - 10 else ROJO)
    lote_txt = ", ".join(df["Lote"].astype(str).unique()[:3]) + ("…" if df["Lote"].nunique() > 3
                                                                 else "")
    est_txt = ", ".join(df["Estación"].astype(str).unique()[:3]) if "Estación" in df else estacion
    op_txt = ", ".join(df["Operador"].astype(str).unique()[:3]) if "Operador" in df else operador
    st.markdown(
        '<div class="isa"><div>N0 · Proceso<br>Contenedores + tapas</div>'
        '<div>N1 · Sensor<br>Cámara industrial</div>'
        '<div>N2 · Control<br>Algoritmo de visión (OpenCV)</div>'
        '<div class="act">N3 · MES<br>Calidad · KPI · Trazabilidad</div>'
        '<div>N4 · Empresa<br>ERP (fuera de alcance)</div></div>', unsafe_allow_html=True)
    st.caption(f"Área: Envasado · Línea 1 · Estación evaluada: **{est_txt}** · "
               f"Operador: **{op_txt}** · Lote(s): **{lote_txt}** · "
               f"Actualizado: {datetime.now():%Y-%m-%d %H:%M:%S}")
    c = st.columns(5)
    c[0].markdown(kpi("Producción inspeccionada", total, "piezas"), unsafe_allow_html=True)
    c[1].markdown(kpi("Aprobadas (liberadas)", ap, f"{fpy:.1f} %", VERDE), unsafe_allow_html=True)
    c[2].markdown(kpi("Rechazadas (retenidas)", re_, f"{100 - fpy:.1f} % · "
                                                     f"{re_ / total * 1e6:,.0f} PPM", ROJO),
                  unsafe_allow_html=True)
    c[3].markdown(kpi("Rendimiento (FPY)", f"{fpy:.1f}%", f"meta {meta:g} %", color_fpy),
                  unsafe_allow_html=True)
    if "_t_ms" in df and df["_t_ms"].notna().any():
        c[4].markdown(kpi("Tiempo de ciclo medio", f"{df['_t_ms'].mean():.0f} ms",
                          "por pieza (visión)"), unsafe_allow_html=True)
    else:
        c[4].markdown(kpi("Lotes analizados", df["Lote"].nunique(), "en el histórico"),
                      unsafe_allow_html=True)
    if fpy >= meta:
        est, col = "ANDON VERDE · Proceso dentro de la meta de calidad", VERDE
    elif fpy >= meta - 10:
        est, col = "ANDON AMARILLO · Rendimiento por debajo de la meta: atención", AMBAR
    else:
        est, col = "ANDON ROJO · Rendimiento crítico: revisar el proceso de tapado", ROJO
    st.markdown(f'<div class="andon" style="background:{col}">{est}</div>',
                unsafe_allow_html=True)
    g1, g2 = st.columns([2, 3])
    with g1:
        mostrar(fig_donut(ap, re_, fpy))
    with g2:
        mostrar(fig_pareto(df))
    st.markdown("#### Control estadístico de proceso (SPC)")
    s1, s2 = st.columns(2)
    with s1:
        mostrar(fig_spc(df, "Altura (px)", umbrales["altura"],
                        "Altura visible de la tapa (sellado)", "px"))
    with s2:
        mostrar(fig_spc(df, "Inclinación (°)", umbrales["inclinacion"],
                        "Inclinación de la tapa", "°"))
    s3, s4 = st.columns(2)
    with s3:
        mostrar(fig_spc(df, "Desalineación (px)", umbrales["desalineacion"],
                        "Desalineación tapa / contenedor", "px"))
    with s4:
        mostrar(fig_lotes(df, meta) if es_historico and df["Lote"].nunique() > 1
                else fig_rendimiento(df, meta))
    st.markdown("#### Trazabilidad y disposición de producto")
    f = st.multiselect("Filtrar por resultado", ["APROBADO", "RECHAZADO"],
                       default=["APROBADO", "RECHAZADO"])
    vista = df[df["Resultado"].isin(f)].copy()
    vista["Disposición"] = vista.apply(disposicion, axis=1)
    cols = ["N°", "Lote", "Imagen", "Fecha", "Hora", "Contenedor", "Tapa",
            "Inclinación (°)", "Altura (px)", "Desalineación (px)", "Resultado",
            "Justificación Técnica", "Disposición"]
    cols = [x for x in cols if x in vista.columns]
    if "Estación" in vista.columns:
        cols += ["Estación", "Operador"]
    st.dataframe(vista[cols].style.apply(estilo_resultado, subset=["Resultado"]),
                 hide_index=True)
    d1, d2 = st.columns(2)
    d1.download_button("Descargar vista (CSV)",
                       vista[cols].to_csv(index=False).encode("utf-8-sig"),
                       "trazabilidad.csv", "text/csv")
    if EXCEL.exists():
        d2.download_button("Descargar histórico completo (Excel)", EXCEL.read_bytes(),
                           EXCEL.name,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab_mes:
    st.caption("Resumen estadístico del lote actual o del histórico acumulado: "
               "rendimiento, semáforo Andon, causas de rechazo, cartas de control y "
               "trazabilidad con disposición de cada pieza.")
    origen = st.radio("Fuente de datos", ["Lote actual (sesión)",
                                          "Histórico acumulado (Excel)"], horizontal=True)
    es_hist = origen.startswith("Histórico")
    datos = None
    if es_hist:
        if not EXCEL.exists():
            st.info("Aún no existe el histórico: se crea al terminar la primera inspección.")
        else:
            try:
                datos = pd.read_excel(EXCEL, sheet_name="Historico")
            except (PermissionError, ValueError) as e:
                st.error(f"No se pudo leer el Excel (¿está abierto?): {e}")
        if datos is not None and not datos.empty:
            lotes = list(dict.fromkeys(datos["Lote"].astype(str)))
            elegidos = st.multiselect("Lotes a analizar", lotes, default=lotes)
            datos = datos[datos["Lote"].astype(str).isin(elegidos)]
    elif ss.resultados:
        datos = pd.DataFrame(ss.resultados)
    if datos is None or datos.empty:
        if not es_hist:
            st.info("Aún no hay resultados en esta sesión. Ejecute una inspección.")
    else:
        mostrar_mes(preparar(datos), es_hist)

# ---------------------------------------------------------------------------
# PESTAÑA 3 · Entrenamiento y calibración asistida por el operador
# ---------------------------------------------------------------------------
CAUSAS_OPC = [motor.C_SIN_CONT, motor.C_SIN_TAPA, motor.C_INCL, motor.C_DESAL,
              motor.C_SELLADO, "Otro"]

with tab_ent:
    st.subheader("Entrenamiento y calibración asistida por el operador")
    st.write("Observe la pieza con los recuadros de detección y responda manualmente si es "
             "**APTA** o **RECHAZADA**; si la rechaza, indique la causa en el selector. "
             "Con estas respuestas el algoritmo ajusta los umbrales sin sobreajuste y su "
             "criterio queda registrado para comparar en la pestaña Validación del modelo.")
    with st.expander("¿Cómo se evita el sobreajuste?"):
        st.markdown(
            "- Los umbrales actuales solo se mueven si sus respuestas lo exigen; si ya son "
            "compatibles se conservan.\n"
            "- Cuando deben moverse, se ubican a 1/3 del margen desde el lado aprobado "
            "(criterio conservador: es peor aceptar una pieza mala que rechazar una buena).\n"
            "- Siempre quedan dentro de rangos físicos plausibles.\n"
            "- La calidad se mide con validación cruzada leave-one-out: cada pieza se predice "
            "sin haberla usado para aprender.\n"
            "- Se exigen mínimo 6 respuestas con piezas APTAS y RECHAZADAS.")
    if not ITEMS:
        st.info("Copie imágenes en la carpeta «imagenes/» para poder entrenar.")
    else:
        csv_ref = CARPETA_IMAGENES / "etiquetas.csv"
        b1, b2 = st.columns(2)
        if b1.button("Iniciar / reiniciar entrenamiento"):
            orden = list(ITEMS)
            random.shuffle(orden)
            ss.ent_orden, ss.ent_resp = orden, {}
            ss.ent_causas, ss.ent_desc = {}, {}
            ss.ent_pila, ss.ent_calib = [], None
            st.rerun()
        if csv_ref.exists():
            if b2.button("Cargar respuestas de referencia del taller"):
                ref = motor.leer_referencia(csv_ref)
                ss.ent_resp = {n: r["apta"] for n, r in ref.items() if n in ITEMS}
                ss.ent_causas = {n: ("; ".join(sorted(r["causas"])) or "Ninguna")
                                 for n, r in ref.items() if n in ITEMS}
                ss.ent_desc = {n: r["descripcion"] for n, r in ref.items() if n in ITEMS}
                ss.ent_orden = list(ss.ent_resp)
                ss.ent_pila, ss.ent_calib = [], None
                guardar_csv_operador()
                st.rerun()
        orden = [n for n in ss.ent_orden if n in ITEMS]
        resp, causas, desc = ss.ent_resp, ss.ent_causas, ss.ent_desc
        pend = [n for n in orden if n not in resp]
        if orden:
            st.progress(len(resp) / len(orden),
                        text=f"Respondidas {len(resp)} de {len(orden)} · "
                             f"APTAS {sum(resp.values())} · "
                             f"RECHAZADAS {len(resp) - sum(resp.values())}")
        if pend:
            actual = pend[0]
            img = motor.decodificar(ITEMS[actual])
            m = motor.medir(img, umbrales)
            ci, cd = st.columns([3, 2])
            ci.image(rgb(motor.dibujar(img, m)), caption=f"{actual} · recuadros de detección")
            cd.markdown("**Mediciones del sistema**")
            cd.write(f"Contenedor: **{'Presente' if m['contenedor'] else 'Ausente'}** · "
                     f"Tapa: **{'Presente' if m['tapa'] else 'Ausente'}**")
            if m["inclinacion"] is not None:
                cd.write(f"Inclinación: **{m['inclinacion']} °** · "
                         f"Altura: **{m['altura']} px** · "
                         f"Desalineación: **{m['desalineacion']} px**")
            cd.markdown("**¿La pieza está bien tapada y sellada?**")
            causa_sel = cd.selectbox(
                "Si la rechaza, ¿cuál es la causa?", CAUSAS_OPC,
                help="Seleccione el motivo de rechazo; queda registrado como criterio del "
                     "operador y se usa en la validación del modelo.")
            desc_sel = ""
            if causa_sel == "Otro":
                desc_sel = cd.text_input("Describa la causa elegida como «Otro»")
            if cd.button("APTA", type="primary"):
                resp[actual] = True
                causas[actual], desc[actual] = "Ninguna", ""
                ss.ent_pila.append(actual)
                guardar_csv_operador()
                st.rerun()
            if cd.button("RECHAZADA"):
                resp[actual] = False
                causas[actual] = causa_sel if causa_sel != "Otro" else ("Otro: " + desc_sel
                                                                        if desc_sel else "Otro")
                desc[actual] = desc_sel
                ss.ent_pila.append(actual)
                guardar_csv_operador()
                st.rerun()
            if len(pend) > 1 and cd.button("Omitir"):
                ss.ent_orden.remove(actual)
                ss.ent_orden.append(actual)
                st.rerun()
            if ss.ent_pila and cd.button("Deshacer la última respuesta"):
                ant = ss.ent_pila.pop()
                resp.pop(ant, None)
                causas.pop(ant, None)
                desc.pop(ant, None)
                guardar_csv_operador()
                st.rerun()
        elif orden:
            st.success("Todas las piezas fueron respondidas.")
            if ss.ent_pila and st.button("Deshacer la última respuesta"):
                ant = ss.ent_pila.pop()
                resp.pop(ant, None)
                causas.pop(ant, None)
                desc.pop(ant, None)
                guardar_csv_operador()
                st.rerun()
        else:
            st.info("Pulse «Iniciar / reiniciar entrenamiento» para comenzar.")
        if resp:
            rech = {n: c for n, c in causas.items() if resp.get(n) is False}
            if rech:
                st.markdown("**Causas registradas por el operador:** " +
                            ", ".join(f"{n}: {c}" for n, c in sorted(rech.items())))
            st.divider()
            usar_hsv = st.checkbox("Calibrar también los umbrales de color (HSV)", value=True)
            if st.button("Calibrar con mis respuestas", type="primary"):
                try:
                    refs = {n: motor.decodificar(ITEMS[n]) for n in resp}
                    barra_c = st.progress(0.0, text="Calibrando...")
                    ss.ent_calib = motor.calibrar(
                        refs, resp, calibrar_hsv=usar_hsv,
                        progreso=lambda p: barra_c.progress(min(p, 1.0), text="Calibrando..."))
                    barra_c.empty()
                except ValueError as e:
                    st.error(str(e))
            r = ss.ent_calib
            if r:
                u, act = r["umbrales"], motor.UMBRALES
                st.success("Calibración calculada. Revise los resultados y aplíquela.")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Respuestas", r["n_respuestas"])
                m2.metric("Exactitud con umbrales por defecto",
                          f"{100 * r['exactitud_defecto']:.1f} %")
                m3.metric("Exactitud calibrada", f"{100 * r['exactitud']:.1f} %")
                m4.metric("Validación cruzada (LOO)", f"{100 * r['exactitud_loocv']:.1f} %")
                st.dataframe(pd.DataFrame({
                    "Umbral": ["Inclinación máx. (°)", "Altura máx. de tapa sellada (px)",
                               "Desalineación máx. (px)", "HSV azul: S mín. / V mín.",
                               "HSV blanco: S máx. / V mín."],
                    "Actual": [act["inclinacion"], act["altura"], act["desalineacion"],
                               f"{act['azul_s_min']} / {act['azul_v_min']}",
                               f"{act['blanco_s_max']} / {act['blanco_v_min']}"],
                    "Calibrado": [u["inclinacion"], u["altura"], u["desalineacion"],
                                  f"{u['azul_s_min']} / {u['azul_v_min']}",
                                  f"{u['blanco_s_max']} / {u['blanco_v_min']}"]}),
                    hide_index=True)
                st.write("**Matriz de confusión:** " + " · ".join(
                    f"{k}: {v}" for k, v in r["matriz_confusion"].items()))
                if r["errores"]:
                    st.warning("Piezas que el sistema aún clasifica distinto a su respuesta: "
                               + ", ".join(r["errores"]))
                a1, a2 = st.columns(2)
                if a1.button("Aplicar y guardar calibración"):
                    motor.guardar_calibracion(r)
                    ss.ent_calib = None
                    st.rerun()
                if a2.button("Restablecer valores por defecto"):
                    motor.restablecer_calibracion()
                    ss.ent_calib = None
                    st.rerun()

# ---------------------------------------------------------------------------
# PESTAÑA 4 · Validación del modelo
# ---------------------------------------------------------------------------
def resumen_grupos(df):
    filas = []
    for g, sub in df.groupby("Causas esperadas"):
        med = sub.dropna(subset=["Altura (px)"])
        fila = {"Grupo de referencia": "APROBADAS" if g == "Ninguna" else g,
                "Piezas": len(sub)}
        for c in ("Inclinación (°)", "Altura (px)", "Desalineación (px)"):
            fila[c] = f"{med[c].min():g} – {med[c].max():g}" if len(med) else "sin medición"
        filas.append(fila)
    filas.append({"Grupo de referencia": "UMBRAL ACTIVO (máx.)", "Piezas": " ",
                  "Inclinación (°)": f"{umbrales['inclinacion']:g}",
                  "Altura (px)": f"{umbrales['altura']:g}",
                  "Desalineación (px)": f"{umbrales['desalineacion']:g}"})
    return pd.DataFrame(filas)

with tab_val:
    st.caption("Compara la decisión y la causa del sistema con la referencia etiquetada del "
               "taller (imagenes/etiquetas.csv) y, si existe, con el criterio registrado por "
               "el operador en la pestaña Entrenamiento asistido.")
    csv_val = CARPETA_IMAGENES / "etiquetas.csv"
    if not csv_val.exists():
        st.info("No se encontró «imagenes/etiquetas.csv».")
    elif st.button("Ejecutar validación contra la referencia", type="primary"):
        referencia = motor.leer_referencia(csv_val)
        imagenes = {n: motor.decodificar(ITEMS[n]) for n in referencia if n in ITEMS}
        ss.validacion = motor.validar(imagenes, referencia, umbrales)
    if ss.validacion is not None and not ss.validacion[0].empty:
        dfv, met = ss.validacion
        st.subheader("Validación contra la referencia del taller")
        c = st.columns(5)
        c[0].markdown(kpi("Exactitud", f"{100 * met['exactitud']:.1f}%",
                          f"{met['VP'] + met['VN']} de {met['n']} piezas",
                          VERDE if met["exactitud"] == 1 else AMBAR), unsafe_allow_html=True)
        c[1].markdown(kpi("Sensibilidad", f"{100 * (met['sensibilidad'] or 0):.1f}%",
                          "aptas detectadas como aptas"), unsafe_allow_html=True)
        c[2].markdown(kpi("Especificidad", f"{100 * (met['especificidad'] or 0):.1f}%",
                          "defectos detectados como defecto"), unsafe_allow_html=True)
        c[3].markdown(kpi("Precisión", f"{100 * (met['precision'] or 0):.1f}%",
                          "aprobadas que eran aptas"), unsafe_allow_html=True)
        c[4].markdown(kpi("Causa correcta", f"{100 * met['exactitud_causa']:.1f}%",
                          "causa igual a la referencia"), unsafe_allow_html=True)
        st.markdown("**Matriz de confusión**")
        st.dataframe(pd.DataFrame(
            {"Sistema: APROBADO": [met["VP"], met["FP"]],
             "Sistema: RECHAZADO": [met["FN"], met["VN"]]},
            index=["Referencia: APTA", "Referencia: RECHAZADA"]))
        st.markdown("**Margen de separación por grupo** (medidas por clase frente al umbral)")
        st.dataframe(resumen_grupos(dfv), hide_index=True)
        malos = dfv[~(dfv["Resultado correcto"] & dfv["Causa correcta"])]
        if malos.empty:
            st.success("Sin discrepancias: todas las piezas coinciden con la referencia "
                       "(resultado y causa).")
        else:
            st.warning(f"{len(malos)} pieza(s) con discrepancias:")
            st.dataframe(malos, hide_index=True)
        with st.expander("Detalle por imagen"):
            st.dataframe(dfv, hide_index=True)
        buf = io.BytesIO()
        dfv.to_excel(buf, index=False, sheet_name="Validacion")
        st.download_button("Descargar informe de validación (Excel)", buf.getvalue(),
                           "validacion_modelo.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.divider()
    st.subheader("Validación contra el criterio del operador")
    op = leer_csv_operador()
    if not op:
        st.caption("Aún no hay criterio del operador registrado. Respuesta las piezas en la "
                   "pestaña Entrenamiento asistido (las causas elegidas se guardan "
                   "automáticamente en «etiquetas_operador.csv»).")
    else:
        if st.button("Validar contra el criterio del operador"):
            ref_op = {n: {"apta": v["apta"],
                          "causas": frozenset(c.strip() for c in v["causa"].split(";")
                                              if c.strip() and c.strip() != "Ninguna"),
                          "descripcion": v["descripcion"]} for n, v in op.items()}
            imgs_op = {n: motor.decodificar(ITEMS[n]) for n in ref_op if n in ITEMS}
            ss.val_op = motor.validar(imgs_op, ref_op, umbrales)
        if ss.val_op is not None and not ss.val_op[0].empty:
            dfo, meto = ss.val_op
            c = st.columns(3)
            c[0].markdown(kpi("Exactitud vs. operador", f"{100 * meto['exactitud']:.1f}%",
                              f"{meto['VP'] + meto['VN']} de {meto['n']} piezas"),
                          unsafe_allow_html=True)
            c[1].markdown(kpi("Causa coincidente", f"{100 * meto['exactitud_causa']:.1f}%",
                              "según causa registrada"), unsafe_allow_html=True)
            c[2].markdown(kpi("Discrepancias", meto["FP"] + meto["FN"],
                              "piezas por revisar"), unsafe_allow_html=True)
            malos_op = dfo[~(dfo["Resultado correcto"] & dfo["Causa correcta"])]
            if malos_op.empty:
                st.success("El sistema coincide con el criterio del operador en resultado y "
                           "causa.")
            else:
                st.warning("Piezas donde el sistema difiere del operador:")
                st.dataframe(malos_op[["Imagen", "Esperado", "Obtenido", "Causas esperadas",
                                       "Causas obtenidas"]], hide_index=True)
