"""Cliente Supabase en memoria que imita la parte de la API usada por AlmacenSupabase."""
import copy
from datetime import datetime, timezone
from types import SimpleNamespace


class _Consulta:
    def __init__(self, db, tabla):
        self.db, self.tabla = db, tabla
        self.op, self.payload, self.filtros = "select", None, []
        self.orden, self.rango, self.limite, self.cuenta, self.conflicto = None, None, None, None, None

    def select(self, *cols, count=None):
        self.op, self.cuenta = "select", count
        return self

    def insert(self, filas, returning=None):
        self.op, self.payload = "insert", filas
        return self

    def upsert(self, filas, on_conflict="", returning=None):
        self.op, self.payload, self.conflicto = "upsert", filas, on_conflict
        return self

    def delete(self):
        self.op = "delete"
        return self

    def eq(self, c, v):
        self.filtros.append(lambda f: f.get(c) == v)
        return self

    def lt(self, c, v):
        self.filtros.append(lambda f: f.get(c) is not None and f.get(c) < v)
        return self

    def order(self, c, desc=False):
        self.orden = (c, desc)
        return self

    def range(self, a, b):
        self.rango = (a, b)
        return self

    def limit(self, n):
        self.limite = n
        return self

    def execute(self):
        filas = self.db.tablas.setdefault(self.tabla, [])
        if self.op in ("insert", "upsert"):
            nuevos = self.payload if isinstance(self.payload, list) else [self.payload]
            salida = []
            for n in nuevos:
                n = copy.deepcopy(n)
                if self.op == "upsert" and self.conflicto:
                    existente = next((f for f in filas if f[self.conflicto] == n[self.conflicto]), None)
                    if existente:
                        existente.update(n)
                        salida.append(existente)
                        continue
                if self.tabla == "corridas":
                    self.db.seq += 1
                    n.setdefault("id", self.db.seq)
                    n.setdefault("creado_en", datetime.now(timezone.utc).isoformat())
                if self.db.fallar_en == self.tabla:
                    raise RuntimeError("fallo simulado")
                filas.append(n)
                salida.append(n)
            return SimpleNamespace(data=salida, count=None)
        sel = [f for f in filas if all(p(f) for p in self.filtros)]
        if self.op == "delete":
            self.db.tablas[self.tabla] = [f for f in filas if f not in sel]
            if self.tabla == "corridas":
                ids = {f["id"] for f in sel}
                self.db.tablas["base_gestion"] = [
                    f for f in self.db.tablas.get("base_gestion", []) if f["corrida_id"] not in ids
                ]
            return SimpleNamespace(data=sel, count=None)
        total = len(sel)
        if self.orden:
            sel = sorted(sel, key=lambda f: f.get(self.orden[0]), reverse=self.orden[1])
        if self.rango:
            sel = sel[self.rango[0]: self.rango[1] + 1]
        if self.limite is not None:
            sel = sel[: self.limite]
        return SimpleNamespace(data=copy.deepcopy(sel), count=total if self.cuenta else None)


class _Bucket:
    def __init__(self, objetos):
        self.objetos = objetos

    def list(self, path="", options=None):
        return [{"name": k, "updated_at": v[1]} for k, v in self.objetos.items() if "/" not in k]

    def upload(self, path, contenido, opciones=None):
        if path in self.objetos and (opciones or {}).get("upsert") != "true":
            raise RuntimeError("ya existe")
        self.objetos[path] = (contenido, datetime.now(timezone.utc).isoformat())

    def download(self, path):
        return self.objetos[path][0]


class FakeSupabase:
    def __init__(self):
        self.tablas, self.objetos, self.seq, self.fallar_en = {}, {}, 0, None
        self.storage = SimpleNamespace(from_=lambda bucket: _Bucket(self.objetos))

    def table(self, nombre):
        return _Consulta(self, nombre)
