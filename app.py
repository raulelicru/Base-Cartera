"""App Streamlit: Generación Automática de Base de Cartera por Campaña de Trabajo.

Ejecutar con:  streamlit run app.py
"""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

import procesamiento as proc

DATA_DIR = Path(os.environ.get("BASE_CARTERA_DATA_DIR", Path(__file__).parent / "data"))
RUTA_ESTRUCTURA = DATA_DIR / "Estructura_General_de_Bases.xlsx"
RUTA_CATALOGO = DATA_DIR / "catalogo_cp.parquet"

st.set_page_config(page_title="Base de Cartera", page_icon="📋", layout="wide")


# --------------------------------------------------------------------------
# Archivos de referencia persistentes
# --------------------------------------------------------------------------


def fecha_archivo(ruta: Path) -> str:
    return datetime.fromtimestamp(ruta.stat().st_mtime).strftime("%d/%m/%Y %H:%M")


@st.cache_data(show_spinner=False)
def cargar_estructura(ruta: str, _mtime: float) -> proc.EstructuraGeneral:
    return proc.leer_estructura(Path(ruta).read_bytes())


@st.cache_data(show_spinner=False)
def cargar_catalogo(ruta: str, _mtime: float) -> pd.DataFrame:
    return pd.read_parquet(ruta)


def estructura_actual() -> proc.EstructuraGeneral | None:
    if not RUTA_ESTRUCTURA.exists():
        return None
    return cargar_estructura(str(RUTA_ESTRUCTURA), RUTA_ESTRUCTURA.stat().st_mtime)


def catalogo_actual() -> pd.DataFrame | None:
    if not RUTA_CATALOGO.exists():
        return None
    return cargar_catalogo(str(RUTA_CATALOGO), RUTA_CATALOGO.stat().st_mtime)


with st.sidebar:
    st.header("Archivos de referencia")
    st.caption("Se guardan en la app y se reutilizan en cada corrida. Reemplácelos cuando cambien.")

    # ---- Estructura General de Bases
    st.subheader("Estructura General de Bases")
    estructura = None
    if RUTA_ESTRUCTURA.exists():
        try:
            estructura = estructura_actual()
            st.success(f"Cargada ({fecha_archivo(RUTA_ESTRUCTURA)})")
            st.caption(
                f"{len(estructura.zonas):,} zonas · Campañas de trabajo: "
                f"{', '.join(map(str, estructura.campanias_disponibles())) or '—'}"
            )
            for adv in estructura.advertencias:
                st.warning(adv)
        except Exception as e:  # noqa: BLE001
            st.error(f"El archivo guardado no es válido: {e}")
    else:
        st.info("Aún no se ha cargado.")

    archivo_estructura = st.file_uploader(
        "Reemplazar Estructura General (.xlsx)", type=["xlsx", "xlsm", "xls"], key="up_estructura"
    )
    if archivo_estructura is not None and st.button("Guardar Estructura General", use_container_width=True):
        contenido = archivo_estructura.getvalue()
        try:
            nueva = proc.leer_estructura(contenido)
        except Exception as e:  # noqa: BLE001
            st.error(f"No se pudo leer el archivo: {e}")
        else:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            RUTA_ESTRUCTURA.write_bytes(contenido)
            st.cache_data.clear()
            st.toast(f"Estructura guardada: {len(nueva.zonas):,} zonas.")
            st.rerun()

    st.divider()

    # ---- Catálogo SEPOMEX
    st.subheader("Catálogo de Códigos Postales")
    catalogo = None
    if RUTA_CATALOGO.exists():
        catalogo = catalogo_actual()
        st.success(f"Cargado ({fecha_archivo(RUTA_CATALOGO)})")
        st.caption(f"{len(catalogo):,} códigos postales")
    else:
        st.info("Aún no se ha cargado. Descárguelo de correosdemexico.gob.mx (TXT, XLS o ZIP).")

    archivo_cp = st.file_uploader(
        "Reemplazar catálogo SEPOMEX", type=["txt", "csv", "xls", "xlsx", "zip"], key="up_cp"
    )
    if archivo_cp is not None and st.button("Guardar catálogo", use_container_width=True):
        try:
            with st.spinner("Procesando catálogo…"):
                nuevo = proc.leer_catalogo_cp(archivo_cp.getvalue(), archivo_cp.name)
            if nuevo.empty:
                raise ValueError("El catálogo no contiene códigos postales válidos.")
        except Exception as e:  # noqa: BLE001
            st.error(f"No se pudo leer el catálogo: {e}")
        else:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            nuevo.to_parquet(RUTA_CATALOGO, index=False)
            st.cache_data.clear()
            st.toast(f"Catálogo guardado: {len(nuevo):,} códigos postales.")
            st.rerun()


# --------------------------------------------------------------------------
# Corrida
# --------------------------------------------------------------------------

st.title("📋 Base de Gestión por Campaña de Trabajo")
st.write(
    "Suba la **Cartera de la campaña**, confirme el número de **Campaña de Trabajo** y genere la base. "
    "Las filas con información faltante se marcan en amarillo para revisión manual; nunca se eliminan "
    "ni se completan con datos inventados."
)

if estructura is None:
    st.warning("Primero cargue el archivo **Estructura General de Bases** en la barra lateral.")
    st.stop()
if catalogo is None:
    st.warning(
        "No hay catálogo de códigos postales. Puede generar la base, pero Municipio, Estado y "
        "Zona (Urbano/Rural) quedarán vacíos y todas las filas se marcarán para revisión."
    )

archivo_cartera = st.file_uploader(
    "Cartera de la campaña (ej. Cartera_Campaña_19.xlsx)", type=["xlsx", "xlsm", "xls", "csv"]
)

inferida = proc.inferir_campania(archivo_cartera.name) if archivo_cartera else None
col1, col2 = st.columns([1, 3])
with col1:
    campania = st.number_input(
        "Campaña de Trabajo (N)",
        min_value=1,
        max_value=99,
        value=inferida if inferida else None,
        step=1,
        placeholder="Ej. 19",
        key=f"campania_{archivo_cartera.name if archivo_cartera else ''}",
    )
with col2:
    if inferida:
        st.caption(f"Inferida del nombre del archivo: **{inferida}**. Verifique antes de generar.")
    disponibles = estructura.campanias_disponibles()
    if campania and disponibles and int(campania) not in disponibles:
        st.error(
            f"La Campaña de Trabajo {int(campania)} no existe en la Estructura General. "
            f"Disponibles: {', '.join(map(str, disponibles))}."
        )

generar = st.button(
    "Generar base", type="primary", disabled=archivo_cartera is None or not campania
)

if generar:
    try:
        with st.spinner("Leyendo cartera…"):
            cartera = proc.leer_cartera(archivo_cartera.getvalue(), archivo_cartera.name)
        with st.spinner(f"Procesando {len(cartera):,} cuentas…"):
            resultado = proc.generar_base(cartera, estructura, catalogo, int(campania))
            excel = proc.exportar_excel(resultado, int(campania))
            csv = proc.exportar_csv(resultado)
    except Exception as e:  # noqa: BLE001
        st.error(f"No se pudo generar la base: {e}")
        st.stop()
    st.session_state["resultado"] = {
        "resultado": resultado,
        "excel": excel,
        "csv": csv,
        "campania": int(campania),
    }

if "resultado" in st.session_state:
    datos = st.session_state["resultado"]
    resultado: proc.Resultado = datos["resultado"]
    n = datos["campania"]
    r = resultado.resumen
    total = r["Total de filas procesadas"]

    st.divider()
    st.subheader(f"Resumen — Campaña de Trabajo {n}")
    for adv in resultado.advertencias:
        st.warning(adv)

    def pct(v):
        return f"{v / total:.1%}" if total else "—"

    m = st.columns(6)
    m[0].metric("Filas procesadas", f"{total:,}")
    m[1].metric("CP identificado", f"{r['CP encontrado en catálogo']:,}", pct(r["CP encontrado en catálogo"]))
    m[2].metric("Zona encontrada", f"{r['Zona encontrada']:,}", pct(r["Zona encontrada"]))
    m[3].metric("Fecha de cierre", f"{r['Fecha de cierre calculada']:,}", pct(r["Fecha de cierre calculada"]))
    m[4].metric("Morosidad", f"{r['Morosidad calculada']:,}", pct(r["Morosidad calculada"]))
    m[5].metric(
        "Requieren revisión",
        f"{r['Filas que requieren revisión']:,}",
        pct(r["Filas que requieren revisión"]),
        delta_color="inverse",
    )

    with st.expander("Detalle del resumen"):
        st.table(pd.DataFrame(list(r.items()), columns=["Concepto", "Filas"]).set_index("Concepto"))

    d1, d2, _ = st.columns([1, 1, 2])
    d1.download_button(
        "⬇️ Descargar Excel",
        data=datos["excel"],
        file_name=f"Base_Gestion_Campaña_{n}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    d2.download_button(
        "⬇️ Descargar CSV",
        data=datos["csv"],
        file_name=f"Base_Gestion_Campaña_{n}.csv",
        mime="text/csv",
        use_container_width=True,
    )

    base = resultado.base
    tab_base, tab_rev = st.tabs(["Base de gestión", f"Revisión ({int(base['requiere_revision'].sum()):,})"])
    with tab_base:
        solo_rev = st.toggle("Mostrar solo filas que requieren revisión")
        vista = base[base["requiere_revision"]] if solo_rev else base
        LIMITE = 2000
        if len(vista) > LIMITE:
            st.caption(f"Vista previa de las primeras {LIMITE:,} de {len(vista):,} filas. Descargue el archivo para verlas todas.")
        vista = vista.head(LIMITE)
        estilo = vista.style.apply(
            lambda fila: ["background-color: #fff3a0; color: black" if fila["requiere_revision"] else "" for _ in fila],
            axis=1,
        ).format({"Fecha de cierre": lambda v: v.strftime("%d/%m/%Y") if hasattr(v, "strftime") else ("" if v is None else v)})
        st.dataframe(estilo, use_container_width=True, hide_index=True)
    with tab_rev:
        rev = base[base["requiere_revision"]]
        if rev.empty:
            st.success("Todas las filas se completaron sin observaciones.")
        else:
            conteo = (
                rev["motivo_revision"].str.split("; ").explode().value_counts().rename_axis("Motivo").reset_index(name="Filas")
            )
            st.dataframe(conteo, hide_index=True, use_container_width=True)
            st.dataframe(
                rev[["NoDama", "ZONA", "RUTA", "CampaniaSaldo", "Direccion", "Cp", "motivo_revision"]],
                hide_index=True,
                use_container_width=True,
            )
