from datetime import datetime

import pandas as pd
import pytest

import almacenamiento as alm
import procesamiento as proc
from tests.fake_supabase import FakeSupabase
from tests.test_procesamiento import _cartera, _estructura_xlsx


@pytest.fixture
def almacen():
    return alm.AlmacenSupabase("http://x", "k", cliente=FakeSupabase())


def _resultado():
    estructura = proc.leer_estructura(_estructura_xlsx())
    catalogo = pd.DataFrame({"Cp": ["05050"], "Municipio": ["Cuajimalpa"], "Estado": ["CDMX"], "Zona": ["Urbano"]})
    return proc.generar_base(proc._normalizar_columnas_cartera(_cartera()), estructura, catalogo, 19)


def test_estructura(almacen):
    assert almacen.fecha_estructura() is None and almacen.leer_estructura() is None
    almacen.guardar_estructura(b"v1")
    almacen.guardar_estructura(b"v2")
    assert almacen.leer_estructura() == b"v2"
    assert isinstance(almacen.fecha_estructura(), datetime)
    assert sum(k.startswith("historial/") for k in almacen.db.objetos) >= 1


def test_catalogo_reemplazo(almacen, monkeypatch):
    monkeypatch.setattr(alm, "LOTE", 2)
    almacen.guardar_catalogo(pd.DataFrame({"Cp": ["01000", "02000", "03000"], "Municipio": ["A", "B", "C"],
                                           "Estado": ["X"] * 3, "Zona": ["Urbano", None, "Rural"]}))
    assert almacen.info_catalogo()[0] == 3
    almacen.guardar_catalogo(pd.DataFrame({"Cp": ["01000", "04000"], "Municipio": ["A2", "D"],
                                           "Estado": ["X"] * 2, "Zona": [None, None]}))
    cat = almacen.leer_catalogo()
    assert sorted(cat["Cp"]) == ["01000", "04000"]
    assert cat.set_index("Cp").loc["01000", "Municipio"] == "A2"


def test_corrida_ida_y_vuelta(almacen, monkeypatch):
    monkeypatch.setattr(alm, "LOTE", 2)
    res = _resultado()
    cid = almacen.guardar_corrida(res, 19, "Cartera_Campaña_19.xlsx")
    assert len(almacen.db.tablas["base_gestion"]) == 3
    fila = almacen.db.tablas["base_gestion"][0]
    assert fila["fecha_cierre"] == "2026-09-14" and fila["no_dama"] == "7601974" and fila["campania_trabajo"] == 19

    leido, n = almacen.leer_corrida(cid)
    assert n == 19
    pd.testing.assert_frame_equal(leido.base.reset_index(drop=True), res.base.reset_index(drop=True), check_dtype=False)
    assert leido.resumen == res.resumen
    proc.exportar_excel(leido, n)  # se puede volver a exportar

    lista = almacen.listar_corridas()
    assert lista.iloc[0]["id"] == cid and lista.iloc[0]["filas_revision"] == 2

    almacen.borrar_corrida(cid)
    assert almacen.listar_corridas().empty and not almacen.db.tablas["base_gestion"]


def test_corrida_fallida_no_deja_residuos(almacen):
    almacen.db.fallar_en = "base_gestion"
    with pytest.raises(RuntimeError):
        almacen.guardar_corrida(_resultado(), 19, "x.xlsx")
    assert almacen.listar_corridas().empty


def test_corrida_guarda_archivo_original(almacen):
    res = _resultado()
    cid = almacen.guardar_corrida(res, 19, "Cartera_Campaña_19.xlsx", contenido=b"XLSX")
    contenido, nombre = almacen.leer_archivo_original(cid)
    assert contenido == b"XLSX" and nombre == "Cartera_Campaña_19.xlsx"
    ruta = almacen.db.tablas["corridas"][0]["archivo_original"]
    assert ruta == f"carteras/{cid}/Cartera_Campana_19.xlsx"
    almacen.borrar_corrida(cid)
    assert ruta not in almacen.db.objetos
