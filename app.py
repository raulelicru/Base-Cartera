"""App Streamlit: Generación Automática de Base de Cartera por Campaña de Trabajo.

Ejecutar con:  streamlit run app.py
"""

from __future__ import annotations

import os
import uuid
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

import almacenamiento as alm
import procesamiento as proc

VERSION = "23/09/2026 · visitas v2 (SaldoDama, TelefonoCelular)"
DATA_DIR = Path(os.environ.get("BASE_CARTERA_DATA_DIR", Path(__file__).parent / "data"))

st.set_page_config(page_title="Base de Cartera", page_icon="📋", layout="wide")


# --------------------------------------------------------------------------
# Almacenamiento: Supabase si hay credenciales, si no archivos locales
# --------------------------------------------------------------------------


def credenciales_supabase() -> tuple[str | None, str | None]:
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_KEY")
    try:
        sec = st.secrets.get("supabase", {})
        url = sec.get("url", url)
        key = sec.get("key", key)
    except Exception:  # noqa: BLE001 - sin secrets.toml
        pass
    return url, key


@st.cache_resource(show_spinner=False)
def obtener_almacen(url: str | None, key: str | None):
    if url and key:
        return alm.AlmacenSupabase(url, key)
    return alm.AlmacenLocal(DATA_DIR)


@st.cache_data(show_spinner="Cargando Estructura General…")
def cargar_estructura(_almacen, _version) -> proc.EstructuraGeneral | None:
    contenido = _almacen.leer_estructura()
    return proc.leer_estructura(contenido) if contenido else None


@st.cache_data(show_spinner="Cargando catálogo de códigos postales…")
def cargar_catalogo(_almacen, _version) -> pd.DataFrame | None:
    return _almacen.leer_catalogo()


def fmt_fecha(f: datetime | None) -> str:
    if f is None:
        return "—"
    if f.tzinfo is not None:
        f = f.astimezone()
    return f.strftime("%d/%m/%Y %H:%M")


try:
    almacen = obtener_almacen(*credenciales_supabase())
except Exception as e:  # noqa: BLE001
    st.error(f"No se pudo conectar a Supabase: {e}")
    st.stop()


with st.sidebar:
    st.caption(f"Versión de la app: {VERSION}")
    st.header("Archivos de referencia")
    if almacen.nombre == "Supabase":
        st.caption("🟢 Conectado a Supabase. Los archivos y las bases generadas se guardan en la nube.")
    else:
        st.caption("⚪ Modo local (sin Supabase). Configure `.streamlit/secrets.toml` para conectarlo.")

    # ---- Estructura General de Bases
    st.subheader("Estructura General de Bases")
    estructura = None
    try:
        fecha_est = almacen.fecha_estructura()
        if fecha_est is not None:
            estructura = cargar_estructura(almacen, str(fecha_est))
    except Exception as e:  # noqa: BLE001
        st.error(f"No se pudo leer la Estructura guardada: {e}")
        fecha_est = None
    if estructura is not None:
        st.success(f"Cargada ({fmt_fecha(fecha_est)})")
        st.caption(
            f"{len(estructura.zonas):,} zonas · Campañas de trabajo: "
            f"{', '.join(map(str, estructura.campanias_disponibles())) or '—'}"
        )
        for adv in estructura.advertencias:
            st.warning(adv)
    elif fecha_est is None:
        st.info("Aún no se ha cargado.")

    archivo_estructura = st.file_uploader(
        "Reemplazar Estructura General (.xlsx)", type=["xlsx", "xlsm"], key="up_estructura"
    )
    if archivo_estructura is not None and st.button("Guardar Estructura General", width="stretch"):
        contenido = archivo_estructura.getvalue()
        try:
            nueva = proc.leer_estructura(contenido)
            with st.spinner("Guardando…"):
                almacen.guardar_estructura(contenido)
        except Exception as e:  # noqa: BLE001
            st.error(f"No se pudo guardar el archivo: {e}")
        else:
            st.cache_data.clear()
            st.toast(f"Estructura guardada: {len(nueva.zonas):,} zonas.")
            st.rerun()

    st.divider()

    # ---- Catálogo SEPOMEX
    st.subheader("Catálogo de Códigos Postales")
    catalogo = None
    try:
        n_cp, fecha_cp = almacen.info_catalogo()
        if n_cp:
            catalogo = cargar_catalogo(almacen, f"{n_cp}-{fecha_cp}")
    except Exception as e:  # noqa: BLE001
        st.error(f"No se pudo leer el catálogo guardado: {e}")
        n_cp, fecha_cp = 0, None
    if catalogo is not None:
        st.success(f"Cargado ({fmt_fecha(fecha_cp)})")
        st.caption(f"{len(catalogo):,} códigos postales")
    else:
        st.info("Aún no se ha cargado. Descárguelo de correosdemexico.gob.mx (TXT, XLS o ZIP).")

    archivo_cp = st.file_uploader(
        "Reemplazar catálogo SEPOMEX", type=["txt", "csv", "xls", "xlsx", "zip"], key="up_cp"
    )
    if archivo_cp is not None and st.button("Guardar catálogo", width="stretch"):
        try:
            with st.spinner("Procesando catálogo…"):
                nuevo = proc.leer_catalogo_cp(archivo_cp.getvalue(), archivo_cp.name)
            if nuevo.empty:
                raise ValueError("El catálogo no contiene códigos postales válidos.")
            barra = st.progress(0.0, text="Guardando catálogo…")
            almacen.guardar_catalogo(nuevo, progreso=lambda x: barra.progress(x, text="Guardando catálogo…"))
        except Exception as e:  # noqa: BLE001
            st.error(f"No se pudo guardar el catálogo: {e}")
        else:
            st.cache_data.clear()
            st.toast(f"Catálogo guardado: {len(nuevo):,} códigos postales.")
            st.rerun()

    st.divider()

    # ---- Formato de salida
    st.subheader("Formato de salida")
    st.caption(
        "El Excel sale igual que la cartera enviada (mismas columnas, orden y colores); "
        "el sistema llena las columnas vacías y anexa al final las que no existen."
    )
    color_sistema = st.color_picker(
        "Color de las columnas que agrega el sistema", value=proc.COLOR_SISTEMA, key="color_sistema"
    )
    colorear_llenadas = st.checkbox(
        "Pintar también las columnas que ya venían vacías en la cartera y llena el sistema",
        value=True,
        key="colorear_llenadas",
        help="REGION, RUTA, DIVISION, ID COBRADOR, Concatenado, Fecha de cierre, Morosidad, "
        "Campaña de trabajo y Referencia de Pago, cuando ya vienen como columnas vacías en la cartera.",
    )
    st.caption("🟨 Las celdas anexadas de las filas que requieren revisión se marcan en amarillo.")


# --------------------------------------------------------------------------
# Presentación de resultados
# --------------------------------------------------------------------------


@st.cache_data(show_spinner="Preparando archivos…", max_entries=6)
def archivos_descarga(clave: str, _resultado, _original, n: int, color: str, colorear: bool):
    """Excel sobre la cartera original (si es .xlsx/.xlsm), Excel estándar y CSV."""
    original = None
    if _original is not None and proc.es_excel_openpyxl(_original[1]):
        try:
            original = proc.exportar_excel_original(
                _original[0], _original[1], _resultado, n, color_sistema=color, colorear_llenadas=colorear
            )
        except Exception as e:  # noqa: BLE001
            original = e
    return original, proc.exportar_excel(_resultado, n), proc.exportar_csv(_resultado)


@st.cache_data(show_spinner=False, max_entries=6)
def cartera_original(clave: str, _original):
    """Cartera completa (todas sus columnas) leída del archivo original, o None."""
    if _original is None:
        return None
    try:
        return proc.leer_cartera(_original[0], _original[1])
    except Exception:  # noqa: BLE001
        return None


@st.cache_data(show_spinner="Preparando base de visitas…", max_entries=6)
def archivo_visitas(clave: str, _resultado, _cartera, fecha, zonas: tuple):
    tabla = proc.tabla_visitas(_resultado, _cartera, fecha)
    if zonas:
        tabla = tabla[tabla["ZONA"].map(proc.clave).isin(zonas)]
    return proc.exportar_visitas(tabla, fecha), len(tabla)


def seccion_visitas(resultado: proc.Resultado, clave: str, original) -> None:
    st.markdown("#### 🚶 Base para visitas de gestores")
    st.caption(
        "Mismo formato de la base de visitas. Todas las columnas salen llenas con los datos de la base "
        "(DIRECCION = dirección generada por el sistema, IMPORTE NETO FACTURA = SaldoDama, "
        "TELEFONO CELULAR = TelefonoCelular); sólo **ASIGNACION** va en blanco para asignar al gestor."
    )
    cartera = cartera_original(clave, original)
    faltantes = proc.columnas_visitas_faltantes(cartera)
    if faltantes:
        st.warning(
            "La cartera subida no trae las columnas **"
            + ", ".join(proc.FUENTE_VISITAS[f] for f in faltantes)
            + "**, por eso "
            + ", ".join(faltantes)
            + " saldrán vacías en la base de visitas. Inclúyalas en la cartera (con esos nombres) para que se llenen."
        )
    v1, v2, v3 = st.columns([1, 2, 2])
    fecha = v1.date_input("Fecha de asignación", value=date.today(), format="DD/MM/YYYY", key=f"fv_{clave}")
    zonas_disp = sorted({proc.clave(z) for z in resultado.base["ZONA"] if proc.clave(z)}, key=lambda z: (len(z), z))
    zonas = v2.multiselect("Zonas", zonas_disp, key=f"zv_{clave}", placeholder="Todas las zonas")
    contenido, n_filas = archivo_visitas(clave, resultado, cartera, fecha, tuple(zonas))
    v3.markdown("<div style='height:1.75rem'></div>", unsafe_allow_html=True)
    v3.download_button(
        f"⬇️ Descargar base para visitas ({n_filas:,} cuentas)",
        data=contenido,
        file_name=f"Base_para_visitas_{fecha:%d-%m-%Y}.xlsx",
        mime=alm.MIME_XLSX,
        width="stretch",
        key=f"vis_{clave}",
        disabled=n_filas == 0,
    )
    st.divider()


def mostrar_resultado(resultado: proc.Resultado, n: int, clave: str, original: tuple[bytes, str] | None) -> None:
    r = resultado.resumen
    total = r["Total de filas procesadas"]
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

    excel_original, excel_estandar, csv = archivos_descarga(
        clave, resultado, original, n, color_sistema, colorear_llenadas
    )
    if isinstance(excel_original, bytes):
        base_nombre, _, ext = original[1].rpartition(".")
        st.download_button(
            "⬇️ Descargar Excel (su cartera + columnas del sistema)",
            data=excel_original,
            file_name=f"{base_nombre}_Base_Gestion.{ext}",
            mime=alm.MIME_XLSX if ext.lower() == "xlsx" else "application/vnd.ms-excel.sheet.macroEnabled.12",
            type="primary",
            key=f"orig_{clave}",
        )
        st.caption(
            "Sale igual que la cartera enviada (mismas columnas, orden y colores). Las columnas que pone el "
            "sistema van en el color elegido en **Formato de salida**; las filas a revisar, en amarillo."
        )
    elif isinstance(excel_original, Exception):
        st.error(f"No se pudo conservar el formato de la cartera: {excel_original}")
    elif original is None:
        st.warning("Esta corrida no tiene guardado el archivo original; sólo está disponible el formato estándar.")
    else:
        st.warning(
            "Para que el Excel salga igual que la cartera enviada, súbala como **.xlsx** "
            "(ábrala en Excel y use *Guardar como → Libro de Excel (.xlsx)*)."
        )

    with st.expander("Otros formatos (formato estándar de 22 columnas, CSV)", expanded=not isinstance(excel_original, bytes)):
        o1, o2, _ = st.columns([2, 1, 2])
        o1.download_button(
            "Excel formato estándar (22 columnas)",
            data=excel_estandar,
            file_name=f"Base_Gestion_Estandar_Campaña_{n}.xlsx",
            mime=alm.MIME_XLSX,
            width="stretch",
            key=f"xlsx_{clave}",
        )
        o2.download_button(
            "CSV",
            data=csv,
            file_name=f"Base_Gestion_Estandar_Campaña_{n}.csv",
            mime="text/csv",
            width="stretch",
            key=f"csv_{clave}",
        )

    seccion_visitas(resultado, clave, original)

    base = resultado.base.drop(columns=["_fila_excel"], errors="ignore")
    tab_base, tab_rev = st.tabs(["Vista previa de datos calculados", f"Revisión ({int(base['requiere_revision'].sum()):,})"])
    with tab_base:
        solo_rev = st.toggle("Mostrar solo filas que requieren revisión", key=f"solo_{clave}")
        vista = base[base["requiere_revision"]] if solo_rev else base
        limite = 2000
        if len(vista) > limite:
            st.caption(
                f"Vista previa de las primeras {limite:,} de {len(vista):,} filas. "
                "Descargue el archivo para verlas todas."
            )
        vista = vista.head(limite)
        estilo = vista.style.apply(
            lambda fila: ["background-color: #fff3a0; color: black" if fila["requiere_revision"] else "" for _ in fila],
            axis=1,
        ).format(
            {"Fecha de cierre": lambda v: v.strftime("%d/%m/%Y") if hasattr(v, "strftime") else ("" if v is None else v)}
        )
        st.dataframe(estilo, width="stretch", hide_index=True)
    with tab_rev:
        rev = base[base["requiere_revision"]]
        if rev.empty:
            st.success("Todas las filas se completaron sin observaciones.")
        else:
            conteo = (
                rev["motivo_revision"].str.split("; ").explode().value_counts()
                .rename_axis("Motivo").reset_index(name="Filas")
            )
            st.dataframe(conteo, hide_index=True, width="stretch")
            st.dataframe(
                rev[["NoDama", "ZONA", "RUTA", "CampaniaSaldo", "Direccion", "Cp", "motivo_revision"]],
                hide_index=True,
                width="stretch",
            )


# --------------------------------------------------------------------------
# Generar base
# --------------------------------------------------------------------------


def pantalla_generar() -> None:
    st.write(
        "Suba la **Cartera de la campaña**, confirme el número de **Campaña de Trabajo** y genere la base. "
        "Las filas con información faltante se marcan en amarillo para revisión manual; nunca se eliminan "
        "ni se completan con datos inventados."
    )
    if estructura is None:
        st.warning("Primero cargue el archivo **Estructura General de Bases** en la barra lateral.")
        return
    if catalogo is None:
        st.warning(
            "No hay catálogo de códigos postales. Puede generar la base, pero Municipio, Estado y "
            "Zona (Urbano/Rural) quedarán vacíos y todas las filas se marcarán para revisión."
        )

    archivo_cartera = st.file_uploader(
        "Cartera de la campaña (ej. Cartera_Campaña_19.xlsx)", type=["xlsx", "xlsm", "xls", "csv"]
    )
    if archivo_cartera is not None and not proc.es_excel_openpyxl(archivo_cartera.name):
        st.warning(
            "Este archivo no es .xlsx: la base se puede generar, pero **no** podrá salir con el mismo formato y "
            "colores de la cartera. Para conservarlos, guárdela en Excel como *Libro de Excel (.xlsx)*."
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

    guardar = False
    if almacen.guarda_historial:
        guardar = st.checkbox("Guardar esta base en Supabase", value=True)

    if st.button("Generar base", type="primary", disabled=archivo_cartera is None or not campania):
        try:
            with st.spinner("Leyendo cartera…"):
                cartera = proc.leer_cartera(archivo_cartera.getvalue(), archivo_cartera.name)
            with st.spinner(f"Procesando {len(cartera):,} cuentas…"):
                resultado = proc.generar_base(cartera, estructura, catalogo, int(campania))
        except Exception as e:  # noqa: BLE001
            st.error(f"No se pudo generar la base: {e}")
            return
        corrida_id = None
        if guardar:
            barra = st.progress(0.0, text="Guardando en Supabase…")
            try:
                corrida_id = almacen.guardar_corrida(
                    resultado,
                    int(campania),
                    archivo_cartera.name,
                    progreso=lambda x: barra.progress(x, text="Guardando en Supabase…"),
                    contenido=archivo_cartera.getvalue(),
                )
            except Exception as e:  # noqa: BLE001
                st.error(f"La base se generó, pero no se pudo guardar en Supabase: {e}")
            finally:
                barra.empty()
        st.session_state["resultado"] = {
            "resultado": resultado,
            "original": (archivo_cartera.getvalue(), archivo_cartera.name),
            "campania": int(campania),
            "corrida_id": corrida_id,
            "clave": f"actual_{uuid.uuid4().hex}",
        }

    if "resultado" in st.session_state:
        datos = st.session_state["resultado"]
        st.divider()
        st.subheader(f"Resumen — Campaña de Trabajo {datos['campania']}")
        if datos.get("corrida_id"):
            st.caption(f"✅ Guardada en Supabase como corrida #{datos['corrida_id']}.")
        mostrar_resultado(datos["resultado"], datos["campania"], datos["clave"], datos["original"])


# --------------------------------------------------------------------------
# Historial (sólo Supabase)
# --------------------------------------------------------------------------


@st.cache_data(show_spinner="Descargando base guardada…", max_entries=5)
def cargar_corrida(_almacen, corrida_id: int):
    resultado, n = _almacen.leer_corrida(corrida_id)
    return resultado, n, _almacen.leer_archivo_original(corrida_id)


def pantalla_historial() -> None:
    try:
        corridas = almacen.listar_corridas()
    except Exception as e:  # noqa: BLE001
        st.error(f"No se pudo leer el historial: {e}")
        return
    if corridas.empty:
        st.info("Todavía no hay bases guardadas.")
        return

    corridas["creado_en"] = pd.to_datetime(corridas["creado_en"]).dt.tz_convert(None)
    campanias = sorted(corridas["campania_trabajo"].unique(), reverse=True)
    filtro = st.selectbox("Campaña de Trabajo", ["Todas"] + [int(c) for c in campanias])
    if filtro != "Todas":
        corridas = corridas[corridas["campania_trabajo"] == filtro]

    columnas = ["id", "campania_trabajo", "archivo_cartera", "total_filas", "filas_revision", "creado_en"]
    tabla = corridas[columnas].rename(
        columns={
            "id": "Corrida",
            "campania_trabajo": "Campaña",
            "archivo_cartera": "Archivo",
            "total_filas": "Filas",
            "filas_revision": "En revisión",
            "creado_en": "Generada (UTC)",
        }
    )
    st.dataframe(tabla, hide_index=True, width="stretch")

    opciones = corridas["id"].tolist()
    elegido = st.selectbox(
        "Abrir corrida",
        opciones,
        format_func=lambda i: (
            lambda c: f"#{i} · Campaña {c.campania_trabajo} · {c.archivo_cartera or ''} · "
            f"{c.creado_en:%d/%m/%Y %H:%M}"
        )(corridas.set_index("id").loc[i]),
    )
    c1, c2, _ = st.columns([1, 1, 3])
    abrir = c1.button("Ver / descargar", type="primary", width="stretch")
    if c2.button("🗑️ Borrar corrida", width="stretch"):
        st.session_state["confirmar_borrado"] = elegido
    if st.session_state.get("confirmar_borrado") == elegido:
        st.warning(f"¿Borrar definitivamente la corrida #{elegido} y todas sus filas?")
        b1, b2, _ = st.columns([1, 1, 3])
        if b1.button("Sí, borrar", type="primary"):
            almacen.borrar_corrida(int(elegido))
            st.session_state.pop("confirmar_borrado", None)
            st.cache_data.clear()
            st.rerun()
        if b2.button("Cancelar"):
            st.session_state.pop("confirmar_borrado", None)
            st.rerun()

    if abrir:
        st.session_state["historial_abierto"] = int(elegido)
    abierto = st.session_state.get("historial_abierto")
    if abierto in opciones:
        try:
            resultado, n, original = cargar_corrida(almacen, abierto)
        except Exception as e:  # noqa: BLE001
            st.error(f"No se pudo leer la corrida: {e}")
            return
        st.divider()
        st.subheader(f"Corrida #{abierto} — Campaña de Trabajo {n}")
        mostrar_resultado(resultado, n, f"hist_{abierto}", original)


st.title("📋 Base de Gestión por Campaña de Trabajo")
if almacen.guarda_historial:
    tab_generar, tab_historial = st.tabs(["Generar base", "Historial"])
    with tab_generar:
        pantalla_generar()
    with tab_historial:
        pantalla_historial()
else:
    pantalla_generar()
