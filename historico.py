# -*- coding: utf-8 -*-
"""Historico de los articulos que DON ARTICULO ha importado al ERP (SQLite, junto al programa).

Sirve para recordar con que codigo entro cada articulo en el ERP:
  * en el banner de eleccion, la opcion ya importada sale en gris ("YA IMPORTADO" + fecha);
  * si se importa OTRO articulo cuyo codigo ya esta usado en el ERP, entra como "x.1"
    (o "x.2" si "x.1" tambien existe), porque el ERP no admite dos articulos con el mismo codigo.
Solo se apuntan las importaciones que terminan bien.
La pestaña REINICIAR REFERENCIA de la web olvida lo apuntado de una referencia (reiniciar): util cuando se borro el
articulo en el ERP y el programa sigue creyendo que ya esta metido.
"""
import os
import sqlite3
import threading
from contextlib import contextmanager
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'historico_importaciones.db')
_LOCK = threading.Lock()

_SCHEMA = """
CREATE TABLE IF NOT EXISTS importaciones (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha           TEXT NOT NULL,      -- ISO: 2030-01-01T00:00:00
    codigo_tecleado TEXT NOT NULL,      -- lo que se tecleo en DON ARTICULO (normalizado)
    codigo_erp      TEXT NOT NULL,      -- codigo con el que entro en el ERP (x, x.1, x.2...)
    proveedor       TEXT NOT NULL,      -- carpeta de DATABASE
    num_proveedor   TEXT,
    fichero         TEXT,
    hoja            TEXT,
    fila            INTEGER,            -- fila de Excel de la tarifa
    descripcion     TEXT,               -- descripcion importada (columna B)
    precio_compra   REAL,               -- precio de compra importado (columna K)
    firma           TEXT                -- lo que importaba esa fila (descripcion|precio): reconoce el MISMO articulo
);
CREATE INDEX IF NOT EXISTS ix_imp_tecleado ON importaciones (codigo_tecleado);
CREATE INDEX IF NOT EXISTS ix_imp_erp ON importaciones (codigo_erp);

-- Lo que borra la pestaña REINICIAR REFERENCIA se guarda aqui (copia de la fila + cuando y de que referencia): se puede recuperar.
CREATE TABLE IF NOT EXISTS reinicios (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha_reinicio  TEXT NOT NULL,
    referencia      TEXT NOT NULL,
    id_original     INTEGER,
    fecha           TEXT,
    codigo_tecleado TEXT,
    codigo_erp      TEXT,
    proveedor       TEXT,
    num_proveedor   TEXT,
    fichero         TEXT,
    hoja            TEXT,
    fila            INTEGER,
    descripcion     TEXT,
    precio_compra   REAL,
    firma           TEXT
);
"""


@contextmanager
def _db():
    with _LOCK:
        con = sqlite3.connect(DB_PATH, timeout=10)
        try:
            con.row_factory = sqlite3.Row
            con.executescript(_SCHEMA)
            yield con
            con.commit()
        finally:
            con.close()


def _norm(code):
    return ' '.join(str(code or '').split()).upper()


def fecha_corta(iso):
    try:
        return datetime.fromisoformat(iso).strftime('%d/%m/%Y %H:%M')
    except (TypeError, ValueError):
        return str(iso or '')


def registrar(codigo_tecleado, codigo_erp, proveedor, num_proveedor=None, fichero=None, hoja=None,
              fila=None, descripcion=None, precio_compra=None, firma=None):
    """Apunta una importacion terminada con exito."""
    with _db() as con:
        con.execute(
            "INSERT INTO importaciones (fecha, codigo_tecleado, codigo_erp, proveedor, num_proveedor,"
            " fichero, hoja, fila, descripcion, precio_compra, firma) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (datetime.now().isoformat(timespec='seconds'), _norm(codigo_tecleado), str(codigo_erp).strip(),
             proveedor, None if num_proveedor is None else str(num_proveedor), fichero, hoja,
             None if fila is None else int(fila), descripcion, precio_compra, firma))


def importaciones(codigo_tecleado, proveedor=None):
    """Importaciones hechas con ese codigo tecleado (opcionalmente de un proveedor), la mas antigua primero."""
    sql = "SELECT * FROM importaciones WHERE codigo_tecleado = ?"
    args = [_norm(codigo_tecleado)]
    if proveedor is not None:
        sql += " AND proveedor = ?"
        args.append(proveedor)
    with _db() as con:
        return [dict(r) for r in con.execute(sql + " ORDER BY fecha, id", args)]


def mismo_articulo(codigo_tecleado, proveedor, firma):
    """Ultima importacion de ESE MISMO articulo (mismo codigo tecleado, proveedor y firma), o None."""
    if not firma:
        return None
    previas = [i for i in importaciones(codigo_tecleado, proveedor) if i['firma'] == firma]
    return previas[-1] if previas else None


def uso_codigo_erp(codigo_erp):
    """Primera importacion que uso ese codigo en el ERP, o None si esta libre."""
    with _db() as con:
        r = con.execute("SELECT * FROM importaciones WHERE UPPER(codigo_erp) = ? ORDER BY fecha, id LIMIT 1",
                        (_norm(codigo_erp),)).fetchone()
        return dict(r) if r else None


def usos_codigo_erp(codigo_erp):
    """Todas las importaciones que usaron ese codigo en el ERP (la mas antigua primero)."""
    with _db() as con:
        return [dict(r) for r in con.execute(
            "SELECT * FROM importaciones WHERE UPPER(codigo_erp) = ? ORDER BY fecha, id", (_norm(codigo_erp),))]


def siguiente_codigo_libre(codigo_erp):
    """Primer codigo libre en el ERP para ese codigo: x.1, x.2, ..."""
    base = str(codigo_erp).strip()
    n = 1
    while uso_codigo_erp('%s.%d' % (base, n)):
        n += 1
    return '%s.%d' % (base, n)


_COLUMNAS_FILA = ('fecha', 'codigo_tecleado', 'codigo_erp', 'proveedor', 'num_proveedor', 'fichero', 'hoja', 'fila',
                  'descripcion', 'precio_compra', 'firma')


def _filas_de(con, referencia):
    return [dict(r) for r in con.execute(
        "SELECT * FROM importaciones WHERE codigo_tecleado = ? OR UPPER(codigo_erp) = ? ORDER BY fecha, id",
        (_norm(referencia), _norm(referencia)))]


def coincidencias(referencia):
    """Todo lo apuntado de una referencia: lo que se tecleo con ese codigo o lo que entro en el ERP con ese codigo
    (asi tambien se encuentra "x.1" tecleando "x.1"). No modifica nada."""
    with _db() as con:
        return _filas_de(con, referencia)


class HistorialCambiado(Exception):
    """Lo apuntado de la referencia ya no es lo que vio la persona al buscar (otro PC importo o reinicio mientras tanto)."""

    def __init__(self, hay, esperadas):
        super().__init__("hay %d filas y se esperaban %d" % (hay, esperadas))
        self.hay, self.esperadas = hay, esperadas


def reiniciar(referencia, esperadas=None):
    """Olvida lo apuntado de esa referencia, como si nunca se hubiera metido (el ERP no se toca): la proxima vez entra con
    su codigo original. Las filas borradas se copian antes a 'reinicios'; todo en una sola transaccion. Devuelve las filas.
    Si se indica 'esperadas' (las que vio la persona) y ahora hay otro numero, no borra nada y lanza HistorialCambiado."""
    with _db() as con:
        filas = _filas_de(con, referencia)
        if esperadas is not None and len(filas) != esperadas:
            raise HistorialCambiado(len(filas), esperadas)
        ahora = datetime.now().isoformat(timespec='seconds')
        for f in filas:
            con.execute(
                "INSERT INTO reinicios (fecha_reinicio, referencia, id_original, " + ", ".join(_COLUMNAS_FILA) + ")"
                " VALUES (?,?,?," + ",".join("?" * len(_COLUMNAS_FILA)) + ")",
                (ahora, _norm(referencia), f['id']) + tuple(f[c] for c in _COLUMNAS_FILA))
            con.execute("DELETE FROM importaciones WHERE id = ?", (f['id'],))
        return filas
