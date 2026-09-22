"""Persistencia de archivos de referencia y bases generadas.

- ``AlmacenSupabase``: Estructura General en Storage (bucket ``referencias``),
  catálogo de CP en la tabla ``catalogo_cp`` y cada base generada en
  ``corridas`` + ``base_gestion``.
- ``AlmacenLocal``: archivos en disco (modo sin Supabase; no guarda historial).
"""

from __future__ import annotations

import math
from datetime import date, datetime, timezone
from pathlib import Path

import pandas as pd

import procesamiento as proc

BUCKET = "referencias"
ARCHIVO_ESTRUCTURA = "Estructura_General_de_Bases.xlsx"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
LOTE = 1000  # máximo de filas por petición (límite por defecto de PostgREST)

# Columna de salida → columna en la tabla base_gestion
COLUMNAS_BD = {
    "ZONA": "zona",
    "REGION": "region",
    "RUTA": "ruta",
    "DIVISION": "division",
    "ID COBRADOR": "id_cobrador",
    "NoDama": "no_dama",
    "Direccion": "direccion",
    "Direccion Calle": "direccion_calle",
    "Colonia": "colonia",
    "Municipio / Poblacion": "municipio_poblacion",
    "Cp": "cp",
    "Estado": "estado",
    "Zona (Urbano/Rural)": "zona_urbano_rural",
    "Referencia": "referencia",
    "AnioSaldo": "anio_saldo",
    "CampaniaSaldo": "campania_saldo",
    "DigitoVerificador": "digito_verificador",
    "Concatenado": "concatenado",
    "Fecha de cierre": "fecha_cierre",
    "Morosidad": "morosidad",
    "Campaña de trabajo": "campania_trabajo",
    "Referencia de Pago": "referencia_pago",
    "requiere_revision": "requiere_revision",
    "motivo_revision": "motivo_revision",
}
_A_SALIDA = {v: k for k, v in COLUMNAS_BD.items()}


def _texto(valor) -> str | None:
    v = proc.valor_limpio(valor)
    return None if v is None else str(v)


def _fecha(valor) -> str | None:
    v = proc.valor_limpio(valor)
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    try:
        return pd.to_datetime(v, dayfirst=True).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return None


def _entero_o_texto(valor):
    """'12' → 12 al leer de la BD (se guardan como texto para no perder ceros ni letras)."""
    if valor is None:
        return None
    s = str(valor)
    return int(s) if s.isdigit() and not (len(s) > 1 and s.startswith("0")) else s


def fila_a_bd(fila: dict, corrida_id: int, n: int) -> dict:
    registro = {"corrida_id": corrida_id, "fila": n}
    for salida, col in COLUMNAS_BD.items():
        v = fila.get(salida)
        if col == "fecha_cierre":
            registro[col] = _fecha(v)
        elif col == "campania_trabajo":
            registro[col] = int(v)
        elif col == "requiere_revision":
            registro[col] = bool(v)
        else:
            registro[col] = _texto(v)
    return registro


def bd_a_fila(registro: dict) -> dict:
    fila = {}
    for col, salida in _A_SALIDA.items():
        v = registro.get(col)
        if col == "fecha_cierre":
            v = datetime.strptime(v[:10], "%Y-%m-%d") if v else None
        elif col in ("requiere_revision", "campania_trabajo", "motivo_revision"):
            pass
        elif col in ("zona", "ruta", "id_cobrador", "no_dama", "anio_saldo", "campania_saldo",
                     "digito_verificador", "morosidad"):
            v = _entero_o_texto(v)
        fila[salida] = v
    fila["motivo_revision"] = fila.get("motivo_revision") or ""
    return fila


def _lotes(elementos: list, tamanio: int = LOTE):
    for i in range(0, len(elementos), tamanio):
        yield elementos[i:i + tamanio]


# --------------------------------------------------------------------------
# Supabase
# --------------------------------------------------------------------------


class AlmacenSupabase:
    nombre = "Supabase"
    guarda_historial = True

    def __init__(self, url: str, key: str, cliente=None):
        if cliente is None:
            from supabase import create_client

            cliente = create_client(url, key)
        self.db = cliente

    # ---- Estructura General (Storage)
    def fecha_estructura(self) -> datetime | None:
        archivos = self.db.storage.from_(BUCKET).list("", {"search": ARCHIVO_ESTRUCTURA})
        for a in archivos or []:
            if a.get("name") == ARCHIVO_ESTRUCTURA:
                f = a.get("updated_at") or a.get("created_at")
                return pd.to_datetime(f).to_pydatetime() if f else datetime.now(timezone.utc)
        return None

    def leer_estructura(self) -> bytes | None:
        if self.fecha_estructura() is None:
            return None
        return self.db.storage.from_(BUCKET).download(ARCHIVO_ESTRUCTURA)

    def guardar_estructura(self, contenido: bytes) -> None:
        bucket = self.db.storage.from_(BUCKET)
        opciones = {"content-type": MIME_XLSX, "upsert": "true"}
        bucket.upload(ARCHIVO_ESTRUCTURA, contenido, dict(opciones))
        # Copia histórica para trazabilidad
        marca = datetime.now().strftime("%Y%m%d_%H%M%S")
        bucket.upload(f"historial/Estructura_{marca}.xlsx", contenido, dict(opciones))

    # ---- Catálogo de CP (tabla catalogo_cp)
    def info_catalogo(self) -> tuple[int, datetime | None]:
        r = self.db.table("catalogo_cp").select("actualizado_en", count="exact").order(
            "actualizado_en", desc=True
        ).limit(1).execute()
        fecha = pd.to_datetime(r.data[0]["actualizado_en"]).to_pydatetime() if r.data else None
        return (r.count or 0), fecha

    def leer_catalogo(self) -> pd.DataFrame | None:
        filas, inicio = [], 0
        while True:
            r = self.db.table("catalogo_cp").select("cp,municipio,estado,zona").order("cp").range(
                inicio, inicio + LOTE - 1
            ).execute()
            filas.extend(r.data or [])
            if len(r.data or []) < LOTE:
                break
            inicio += LOTE
        if not filas:
            return None
        df = pd.DataFrame(filas).rename(columns={"cp": "Cp", "municipio": "Municipio", "estado": "Estado", "zona": "Zona"})
        return df

    def guardar_catalogo(self, catalogo: pd.DataFrame, progreso=None) -> None:
        marca = datetime.now(timezone.utc).isoformat()
        registros = [
            {
                "cp": r.Cp,
                "municipio": _texto(r.Municipio),
                "estado": _texto(r.Estado),
                "zona": _texto(r.Zona),
                "actualizado_en": marca,
            }
            for r in catalogo.itertuples(index=False)
        ]
        total = max(1, math.ceil(len(registros) / LOTE))
        for i, lote in enumerate(_lotes(registros), start=1):
            self.db.table("catalogo_cp").upsert(lote, on_conflict="cp", returning="minimal").execute()
            if progreso:
                progreso(i / total)
        # Quita los CP que ya no vienen en el catálogo nuevo
        self.db.table("catalogo_cp").delete().lt("actualizado_en", marca).execute()

    # ---- Bases generadas
    def guardar_corrida(self, resultado: proc.Resultado, campania: int, archivo: str, progreso=None) -> int:
        r = self.db.table("corridas").insert(
            {
                "campania_trabajo": int(campania),
                "archivo_cartera": archivo,
                "total_filas": int(resultado.resumen["Total de filas procesadas"]),
                "filas_revision": int(resultado.resumen["Filas que requieren revisión"]),
                "resumen": {k: int(v) for k, v in resultado.resumen.items()},
            }
        ).execute()
        corrida_id = r.data[0]["id"]
        try:
            registros = [
                fila_a_bd(f, corrida_id, n) for n, f in enumerate(resultado.base.to_dict("records"), start=1)
            ]
            total = max(1, math.ceil(len(registros) / LOTE))
            for i, lote in enumerate(_lotes(registros), start=1):
                self.db.table("base_gestion").insert(lote, returning="minimal").execute()
                if progreso:
                    progreso(i / total)
        except Exception:
            # No dejar corridas incompletas (base_gestion se borra en cascada)
            self.db.table("corridas").delete().eq("id", corrida_id).execute()
            raise
        return corrida_id

    def listar_corridas(self, limite: int = 100) -> pd.DataFrame:
        r = self.db.table("corridas").select(
            "id,campania_trabajo,archivo_cartera,total_filas,filas_revision,creado_en"
        ).order("creado_en", desc=True).limit(limite).execute()
        return pd.DataFrame(r.data or [])

    def leer_corrida(self, corrida_id: int) -> tuple[proc.Resultado, int]:
        c = self.db.table("corridas").select("*").eq("id", corrida_id).limit(1).execute()
        if not c.data:
            raise ValueError(f"No existe la corrida {corrida_id}.")
        corrida = c.data[0]
        filas, inicio = [], 0
        while True:
            r = self.db.table("base_gestion").select("*").eq("corrida_id", corrida_id).order("fila").range(
                inicio, inicio + LOTE - 1
            ).execute()
            filas.extend(r.data or [])
            if len(r.data or []) < LOTE:
                break
            inicio += LOTE
        base = pd.DataFrame(
            [bd_a_fila(f) for f in filas], columns=proc.COLUMNAS_SALIDA + ["requiere_revision", "motivo_revision"],
            dtype=object,
        )
        base["requiere_revision"] = base["requiere_revision"].astype(bool)
        return proc.Resultado(base=base, resumen=corrida["resumen"], advertencias=[]), corrida["campania_trabajo"]

    def borrar_corrida(self, corrida_id: int) -> None:
        self.db.table("corridas").delete().eq("id", corrida_id).execute()


# --------------------------------------------------------------------------
# Local (sin Supabase)
# --------------------------------------------------------------------------


class AlmacenLocal:
    nombre = "Local"
    guarda_historial = False

    def __init__(self, carpeta: Path):
        self.carpeta = Path(carpeta)
        self.ruta_estructura = self.carpeta / ARCHIVO_ESTRUCTURA
        self.ruta_catalogo = self.carpeta / "catalogo_cp.parquet"

    @staticmethod
    def _fecha(ruta: Path) -> datetime | None:
        return datetime.fromtimestamp(ruta.stat().st_mtime) if ruta.exists() else None

    def fecha_estructura(self):
        return self._fecha(self.ruta_estructura)

    def leer_estructura(self):
        return self.ruta_estructura.read_bytes() if self.ruta_estructura.exists() else None

    def guardar_estructura(self, contenido: bytes) -> None:
        self.carpeta.mkdir(parents=True, exist_ok=True)
        self.ruta_estructura.write_bytes(contenido)

    def info_catalogo(self):
        if not self.ruta_catalogo.exists():
            return 0, None
        return len(pd.read_parquet(self.ruta_catalogo, columns=["Cp"])), self._fecha(self.ruta_catalogo)

    def leer_catalogo(self):
        return pd.read_parquet(self.ruta_catalogo) if self.ruta_catalogo.exists() else None

    def guardar_catalogo(self, catalogo: pd.DataFrame, progreso=None) -> None:
        self.carpeta.mkdir(parents=True, exist_ok=True)
        catalogo.to_parquet(self.ruta_catalogo, index=False)
        if progreso:
            progreso(1.0)
