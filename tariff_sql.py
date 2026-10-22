# -*- coding: utf-8 -*-
"""Tarifas grandes de DATABASE leidas como SQL (SQLite) en vez de releer el Excel entero en cada busqueda.

Por que: DON ARTICULO abre con pandas el .xlsx del proveedor en CADA busqueda (una tarifa de miles de filas y decenas de columnas,
tarda decenas de segundos). Aqui el Excel se convierte UNA vez en un fichero .sqlite junto a el (mismo nombre) con un indice de
referencias; las busquedas consultan el indice y el DataFrame sale de SQLite (o de memoria) con EXACTAMENTE el mismo contenido.

Garantias:
  * Solo para los proveedores de SQL_PROVIDERS (hoy solo uno de ejemplo). Vaciar el conjunto devuelve todo al Excel.
  * El .sqlite se reconstruye solo si cambia el .xlsx (fecha de modificacion, tamano) -> poner una tarifa nueva en DATABASE basta.
  * Al construirlo se comprueba que releer cada hoja da un DataFrame IDENTICO al del Excel; si no lo es, NO se usa SQL
    (se sigue con el Excel y no se reintenta hasta que cambie el fichero). Nada se pierde por el camino.
  * El indice usa la MISMA regla de coincidencia que find_reference_hit (columnas candidatas + _norm_cell).
No toca red, ERP ni Selenium.
"""
from __future__ import annotations

import os
import re
import sqlite3
import threading
import time
import uuid
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Tuple

import pandas as pd

SQL_PROVIDERS = {'PROVEEDOR_EJEMPLO'}            # proveedor normalizado (mayusculas, sin espacios, como _normalize_provider)
VERSION = 1

_LOCK = threading.RLock()
_MEM: Dict[str, Tuple[tuple, Dict[str, pd.DataFrame]]] = {}      # ruta -> (firma, hojas)
_SIN_SQL: Dict[str, tuple] = {}                                  # ruta -> firma para la que no se pudo usar SQL (se sigue con Excel)
_ULTIMO_MOTIVO: Dict[str, str] = {}
_REINTENTO: Dict[str, float] = {}                                # ruta -> instante hasta el que no se reintenta tras un fallo transitorio
REINTENTO_SEGUNDOS = 120


class NoSql(Exception):
    """Esta tarifa no se puede servir desde SQL: el llamador sigue con el Excel."""


def normaliza_proveedor(nombre: object) -> str:
    return re.sub(r'\s+', '', str(nombre or '').strip().upper())


def usa_sql(proveedor: object) -> bool:
    return normaliza_proveedor(proveedor) in SQL_PROVIDERS


def ruta_sql(ruta_xlsx: str) -> str:
    return os.path.splitext(ruta_xlsx)[0] + '.sqlite'


def _firma(ruta_xlsx: str) -> tuple:
    st = os.stat(ruta_xlsx)
    try:
        pandas_mayor = int(str(pd.__version__).split('.')[0])
    except ValueError:
        pandas_mayor = 0
    return (st.st_mtime_ns, st.st_size, VERSION, pandas_mayor)       # otra version de pandas -> se rehace (los tipos pueden cambiar)


def motivo_sin_sql(ruta_xlsx: str) -> str:
    return _ULTIMO_MOTIVO.get(ruta_xlsx, '')


# ---------------------------------------------------------------- lectura / comparacion
def _leer_tabla(con: sqlite3.Connection, tabla: str) -> pd.DataFrame:
    df = pd.read_sql_query(f'SELECT * FROM "{tabla}" ORDER BY "__idx"', con, index_col='__idx')
    df = df.where(df.notna(), float('nan'))              # NULL de SQLite -> NaN, como deja read_excel las celdas vacias
    for col, dtype in con.execute('SELECT col, dtype FROM dtypes WHERE tabla = ?', (tabla,)):
        if col in df.columns and str(df[col].dtype) != dtype:      # p. ej. columnas enteras vacias: float64 en el Excel, object al volver de SQL
            df[col] = df[col].astype(dtype)
    df.index.name = None
    if list(df.index) == list(range(len(df))):
        df.index = pd.RangeIndex(len(df))
    return df


def _identicos(original: pd.DataFrame, leido: pd.DataFrame) -> bool:
    try:
        pd.testing.assert_frame_equal(original, leido, check_dtype=True, check_exact=True, check_names=False, check_index_type='equiv')
        return True
    except AssertionError:
        return False


def _hojas_desde_sql(ruta_sql_: str) -> Dict[str, pd.DataFrame]:
    con = sqlite3.connect(ruta_sql_)
    try:
        orden = con.execute('SELECT nombre, tabla FROM hojas ORDER BY pos').fetchall()
        return {nombre: _leer_tabla(con, tabla) for nombre, tabla in orden}
    finally:
        con.close()


def _firma_guardada(ruta_sql_: str) -> Optional[tuple]:
    if not os.path.isfile(ruta_sql_):
        return None
    try:
        con = sqlite3.connect(ruta_sql_)
        try:
            fila = con.execute("SELECT v FROM meta WHERE k='firma'").fetchone()
        finally:
            con.close()
        return tuple(int(x) for x in fila[0].split(',')) if fila else None
    except (sqlite3.Error, ValueError):
        return None


# ---------------------------------------------------------------- construccion
def _construir(ruta_xlsx: str, leer_excel: Callable[[], Dict[str, pd.DataFrame]],
               cols_ref: Callable[[pd.DataFrame], Sequence], norm: Callable[[object], str],
               log: Optional[Callable[[str], None]]) -> Dict[str, pd.DataFrame]:
    t0 = time.time()
    firma = _firma(ruta_xlsx)
    hojas = leer_excel()
    destino = ruta_sql(ruta_xlsx)
    tmp = f'{destino}.{os.getpid()}.{uuid.uuid4().hex[:8]}.tmp'        # nombre unico: dos procesos construyendo a la vez no se pisan
    con = sqlite3.connect(tmp)
    try:
        con.execute('CREATE TABLE meta (k TEXT PRIMARY KEY, v TEXT)')
        con.execute('CREATE TABLE hojas (pos INTEGER PRIMARY KEY, nombre TEXT, tabla TEXT)')
        con.execute('CREATE TABLE refs (hoja INTEGER, ref TEXT, colpos INTEGER, col TEXT, idx INTEGER)')
        con.execute('CREATE TABLE dtypes (tabla TEXT, col TEXT, dtype TEXT)')
        for pos, (nombre, df) in enumerate(hojas.items()):
            if df is None or df.empty:
                raise NoSql(f"la hoja '{nombre}' esta vacia")
            if not all(isinstance(c, str) for c in df.columns) or len(set(df.columns)) != len(df.columns):
                raise NoSql(f"la hoja '{nombre}' tiene cabeceras no textuales o repetidas")
            tabla = f'h{pos}'
            df.to_sql(tabla, con, index=True, index_label='__idx', if_exists='replace', chunksize=500)
            con.execute('INSERT INTO hojas VALUES (?,?,?)', (pos, nombre, tabla))
            con.executemany('INSERT INTO dtypes VALUES (?,?,?)', [(tabla, c, str(df[c].dtype)) for c in df.columns])
            filas: List[tuple] = []
            for colpos, c in enumerate(cols_ref(df)):
                try:
                    s = df[c].map(norm)
                except Exception:           # igual que find_reference_hit: una columna ilegible se salta
                    continue
                for idx, valor in s[s != ''].items():
                    filas.append((pos, valor, colpos, str(c), int(idx)))
            con.executemany('INSERT INTO refs VALUES (?,?,?,?,?)', filas)
            if not _identicos(df, _leer_tabla(con, tabla)):
                raise NoSql(f"releer la hoja '{nombre}' desde SQL no da el mismo DataFrame que el Excel")
        con.execute('CREATE INDEX ix_refs ON refs (ref)')
        con.execute("INSERT INTO meta VALUES ('firma', ?)", (','.join(str(x) for x in firma),))
        con.commit()
        con.close()
        con = None
        os.replace(tmp, destino)                                      # si otro lector tiene abierto el .sqlite, Windows no deja: se limpia el .tmp
    finally:
        if con is not None:
            con.close()
        for resto in (tmp, tmp + '-journal'):
            if os.path.exists(resto):
                try:
                    os.remove(resto)
                except OSError:
                    pass
    if log:
        log(f'SQL: {os.path.basename(ruta_xlsx)} convertido a {os.path.basename(destino)} en {round(time.time() - t0, 1)} s')
    return hojas


def _copias(hojas: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    return {nombre: df.copy(deep=True) for nombre, df in hojas.items()}


def hojas(ruta_xlsx: str, leer_excel: Callable[[], Dict[str, pd.DataFrame]], cols_ref: Callable[[pd.DataFrame], Sequence],
          norm: Callable[[object], str], log: Optional[Callable[[str], None]] = None) -> Dict[str, pd.DataFrame]:
    """Hojas de la tarifa (dict nombre -> DataFrame) servidas desde SQL/memoria; NoSql si hay que usar el Excel.
    Siempre devuelve COPIAS: quien las reciba puede modificarlas."""
    with _LOCK:
        firma = _firma(ruta_xlsx)
        if _SIN_SQL.get(ruta_xlsx) == firma:
            raise NoSql(_ULTIMO_MOTIVO.get(ruta_xlsx, 'sin SQL'))
        en_memoria = _MEM.get(ruta_xlsx)
        if en_memoria and en_memoria[0] == firma:
            return _copias(en_memoria[1])
        if time.time() < _REINTENTO.get(ruta_xlsx, 0):                   # fallo transitorio reciente: no se reconstruye en cada busqueda
            raise NoSql(_ULTIMO_MOTIVO.get(ruta_xlsx, 'sin SQL de momento'))
        destino = ruta_sql(ruta_xlsx)
        try:
            if os.path.isfile(destino) and _firma_guardada(destino) == firma:
                datos = _hojas_desde_sql(destino)
            else:
                datos = _construir(ruta_xlsx, leer_excel, cols_ref, norm, log)
        except NoSql as e:                                               # estructural: no cambia hasta que cambie el Excel
            _SIN_SQL[ruta_xlsx] = firma
            _ULTIMO_MOTIVO[ruta_xlsx] = str(e)
            if log:
                log(f'SQL no aplicable a {os.path.basename(ruta_xlsx)} ({e}); se sigue con el Excel')
            raise
        except Exception as e:                                           # transitorio (bloqueo de fichero, otro proceso construyendo...)
            _ULTIMO_MOTIVO[ruta_xlsx] = f'{type(e).__name__}: {e}'
            try:
                if os.path.isfile(destino) and _firma_guardada(destino) == _firma(ruta_xlsx):      # quiza otro proceso ya lo dejo hecho
                    datos = _hojas_desde_sql(destino)
                    _MEM[ruta_xlsx] = (firma, datos)
                    return _copias(datos)
            except Exception:
                pass
            _REINTENTO[ruta_xlsx] = time.time() + REINTENTO_SEGUNDOS
            if log:
                log(f'SQL fallo en {os.path.basename(ruta_xlsx)} ({_ULTIMO_MOTIVO[ruta_xlsx]}); se lee el Excel y se reintenta en {REINTENTO_SEGUNDOS} s')
            raise NoSql(_ULTIMO_MOTIVO[ruta_xlsx])
        _REINTENTO.pop(ruta_xlsx, None)
        _MEM[ruta_xlsx] = (firma, datos)
        return _copias(datos)


def buscar(ruta_xlsx: str, ref_norm: str) -> Optional[List[Tuple[str, str, int, List[int]]]]:
    """Coincidencias EXACTAS de la referencia normalizada, por hoja y en el orden de las hojas:
    [(hoja, primera columna con coincidencia, primera fila de esa columna, [todas las filas de la hoja donde casa])].
    Lista vacia = no esta. None = no hay SQL valido para esta tarifa (llamar antes a hojas())."""
    destino = ruta_sql(ruta_xlsx)
    with _LOCK:
        if _SIN_SQL.get(ruta_xlsx) == _firma(ruta_xlsx) or _firma_guardada(destino) != _firma(ruta_xlsx):
            return None
    con = sqlite3.connect(Path(destino).resolve().as_uri() + '?mode=ro', uri=True)         # solo lectura: nunca crea un .sqlite vacio
    try:
        filas = con.execute(
            'SELECT h.nombre, r.colpos, r.col, r.idx FROM refs r JOIN hojas h ON h.pos = r.hoja WHERE r.ref = ? '
            'ORDER BY h.pos, r.colpos, r.idx', (ref_norm,)).fetchall()
    finally:
        con.close()
    por_hoja: Dict[str, List[tuple]] = {}
    for nombre, colpos, col, idx in filas:
        por_hoja.setdefault(nombre, []).append((colpos, col, idx))
    salida = []
    for nombre, lista in por_hoja.items():
        primera_col = lista[0][1]                                    # la primera columna candidata (por orden) con coincidencia
        primera_fila = min(i for _, c, i in lista if c == primera_col)
        salida.append((nombre, primera_col, primera_fila, sorted({i for _, _, i in lista})))
    return salida
