# Base de Cartera — Generación automática de la Base de Gestión

App en **Streamlit** que genera la *Base de Gestión* de una Campaña de Trabajo a partir de:

1. **Cartera de la campaña** (ej. `Cartera_Campaña_19.xlsx`, hoja `BASE`).
2. **Estructura General de Bases** (`Estructura_General_de_Bases.xlsx`, hojas *Base de Zonas*, *Calendario de Cierre* y *Campaña de Trabajo*).
3. **Catálogo Nacional de Códigos Postales** (SEPOMEX / Correos de México).

## Instalación y ejecución

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

La app abre en `http://localhost:8501`.

## Conexión a Supabase

La app guarda todo en el proyecto de Supabase **base-cartera** (`https://irqpuydmkqgvpcbdiydc.supabase.co`):

| Qué | Dónde |
|---|---|
| Estructura General de Bases | Storage, bucket privado `referencias` (+ copias en `historial/`) |
| Catálogo de códigos postales | Tabla `catalogo_cp` |
| Cada base generada | Tablas `corridas` (resumen) y `base_gestion` (filas, con `requiere_revision`) |

El esquema está en `supabase/migrations/`. Las tablas tienen RLS activo sin políticas: sólo el servidor
con la llave secreta puede leer o escribir.

**Configurar la llave (una sola vez):**

1. En Supabase: *Project Settings → API Keys* y copie la llave **secret** (`sb_secret_…`) o la legacy **service_role**.
2. Local: copie `.streamlit/secrets.toml.example` como `.streamlit/secrets.toml` y pegue la llave.
   En Streamlit Cloud: péguelo en *App settings → Secrets*.
   (También puede usar las variables de entorno `SUPABASE_URL` y `SUPABASE_KEY`.)

Sin llave, la app funciona en **modo local**: guarda la Estructura y el catálogo en `data/` y no guarda historial.

> El plan Free de Supabase pausa el proyecto tras ~7 días sin actividad; se reactiva desde el panel de Supabase.

## Uso

1. **Barra lateral → Estructura General de Bases**: suba el archivo y pulse *Guardar*. Se guarda (en Supabase, o en `data/` en modo local) y se reutiliza en cada corrida; reemplácelo cuando cambien rutas, zonas, divisiones o el calendario (normalmente cada 14 días).
2. **Barra lateral → Catálogo de Códigos Postales**: descargue el catálogo oficial de Correos de México
   (correosdemexico.gob.mx, formato TXT, XLS o el ZIP tal cual) y súbalo. Se convierte a una tabla
   de 1 fila por CP (`catalogo_cp` en Supabase o `data/catalogo_cp.parquet` en modo local). Actualícelo al menos una vez al año.
3. **Pantalla principal**: suba la Cartera; el número de Campaña de Trabajo se infiere del nombre del archivo
   (`…Campaña_19…` → 19) y puede corregirse. Pulse **Generar base**.
4. Revise el resumen y descargue el **Excel** (hojas *Base de Gestion*, *Revision*, *Resumen*; filas a revisar
   en amarillo) o el **CSV** (incluye las columnas `requiere_revision` y `motivo_revision`).

La carpeta de datos puede cambiarse con la variable de entorno `BASE_CARTERA_DATA_DIR`.

Con Supabase conectado aparece la pestaña **Historial**: permite volver a abrir, descargar o borrar
cualquier base generada anteriormente, filtrando por Campaña de Trabajo.

## Lógica implementada (según la especificación)

| Paso | Descripción |
|---|---|
| 4.1 | CP con `CP\s*(\d{1,5})\s*$`, relleno a 5 dígitos; `00000` o sin match → vacío y revisión. |
| 4.2 | Cruce con SEPOMEX: Municipio, Estado y Zona (`d_zona`; si no existe, se infiere del tipo de asentamiento). |
| 4.3–4.4 | Segmentación por palabras clave `No`, `Mza`, `Int`, `Lt` → *Direccion Calle* (`{Vialidad} No. {x} Interior {x} Mz {x} L- {x}`) y *Colonia*. |
| 4.5 | `ZONA` → REGION, RUTA, DIVISION, ID COBRADOR (hoja *Base de Zonas*). |
| 4.6 | `Concatenado = NoDama-CampaniaSaldo`. |
| 4.7 | Fecha de cierre: bloque `Campaña de Trabajo N` de *Calendario de Cierre*, por RUTA. |
| 4.8 | Morosidad: bloque `Campaña de Trabajo N` de *Campaña de Trabajo*, por CampaniaSaldo. |
| 4.9 | Campaña de trabajo = N. |
| 4.10 | `Referencia de Pago = NoDama-DigitoVerificador`. |
| 6 | Ninguna fila se elimina ni se inventa; las incompletas se marcan con su motivo. |

## Estructura del código

- `app.py` — interfaz Streamlit.
- `procesamiento.py` — toda la lógica (independiente de Streamlit, reutilizable).
- `almacenamiento.py` — persistencia en Supabase o en disco local.
- `tests/` — pruebas: `python -m pytest`.
