# -*- coding: utf-8 -*-
"""Indice en memoria referencia -> proveedores de DATABASE.

Sirve para saber, sin releer los Excel en cada busqueda, en que proveedores existe una
referencia. Usa la MISMA regla de coincidencia que DON_ARTICULO.find_reference_anysheet
(columnas candidatas de referencia + _norm_cell), de modo que "esta en el indice" equivale
a "find_reference_anysheet la encontraria". Se reconstruye solo si cambia algun .xlsx.
"""
import os
import threading
import time

import pandas as pd

from DON_ARTICULO import (_candidate_ref_columns, _is_converted_template, _iter_tariff_files, _norm_cell,
                          _read_tariff_sheets)

_LOCK = threading.Lock()
_INDEX = {}         # ref normalizada -> proveedor (str) o tupla de proveedores en orden de busqueda
_ERP_CODES = {}     # codigo normalizado -> tupla de proveedores cuya tarifa en formato ERP (articulos ya cargados
                    # en el ERP) lo trae en la columna Codigo
_SIGNATURE = None   # (ruta, mtime, tamano) de cada tarifa cuando se construyo el indice
_BUILD_SECONDS = 0.0


def _signature(database_root):
    sig = []
    for path, _fichero, _proveedor in _iter_tariff_files(database_root):
        try:
            st = os.stat(path)
            sig.append((path, st.st_mtime_ns, st.st_size))
        except OSError:
            sig.append((path, 0, 0))
    return tuple(sig)


def _erp_code_column(df):
    """Columna 'Codigo' de una tarifa en formato ERP, o None."""
    for c in df.columns:
        if _norm_cell(c) == "CODIGO":
            return c
    return None


def _build(database_root):
    """(indice referencia -> proveedores, codigos ya cargados en el ERP -> proveedores) en una sola pasada."""
    index = {}
    erp = {}
    for path, _fichero, proveedor in _iter_tariff_files(database_root):
        try:
            sheets = _read_tariff_sheets(path)
        except Exception:
            continue    # mismo criterio que find_reference_anysheet: una tarifa ilegible se salta
        for df in sheets.values():
            if df is None or df.empty:
                continue
            if _is_converted_template(df):      # tarifa en formato ERP = articulos que YA estan en el ERP
                col = _erp_code_column(df)
                if col is not None:
                    for codigo in df[col].map(_norm_cell).unique():
                        if codigo:
                            erp.setdefault(codigo, set()).add(proveedor)
            for col in _candidate_ref_columns(df):
                try:
                    valores = df[col].map(_norm_cell)
                except Exception:
                    continue
                for ref in valores[valores != ""].unique():
                    actual = index.get(ref)
                    if actual is None:
                        index[ref] = proveedor
                    elif isinstance(actual, str):
                        if actual != proveedor:
                            index[ref] = (actual, proveedor)
                    elif proveedor not in actual:
                        index[ref] = actual + (proveedor,)
    return index, {codigo: tuple(sorted(provs)) for codigo, provs in erp.items()}


def _ensure_built(database_root, logger=None):
    """Construye los indices si es la primera vez o ha cambiado alguna tarifa."""
    global _INDEX, _ERP_CODES, _SIGNATURE, _BUILD_SECONDS
    sig = _signature(database_root)
    if sig != _SIGNATURE and logger:
        logger("Indexando DATABASE (solo ocurre tras arrancar o al cambiar una tarifa; con tarifas grandes puede tardar)...")
    elif _LOCK.locked() and logger:
        logger("Esperando a que termine la indexacion de DATABASE...")
    with _LOCK:
        if sig != _SIGNATURE:
            t0 = time.time()
            _INDEX, _ERP_CODES = _build(database_root)
            _SIGNATURE = sig
            _BUILD_SECONDS = time.time() - t0


def providers_for(database_root, reference, logger=None):
    """Proveedores (nombre de carpeta) que tienen la referencia, en orden de busqueda.
    Lista vacia si no esta en ninguno. La primera llamada tras arrancar (o tras cambiar
    una tarifa) construye el indice y puede tardar; las siguientes son inmediatas."""
    ref_norm = _norm_cell(reference)
    if not ref_norm:
        return []
    _ensure_built(database_root, logger)
    hit = _INDEX.get(ref_norm)
    if hit is None:
        return []
    return [hit] if isinstance(hit, str) else list(hit)


def erp_providers_for(database_root, codigo, logger=None):
    """Proveedores cuya tarifa en formato ERP (articulos ya cargados en el ERP) trae ese codigo en la
    columna Codigo. Lista vacia si ninguno."""
    codigo_norm = _norm_cell(codigo)
    if not codigo_norm:
        return []
    _ensure_built(database_root, logger)
    return list(_ERP_CODES.get(codigo_norm, ()))


def stats():
    """(referencias indexadas, segundos que tardo la ultima construccion)."""
    return len(_INDEX), round(_BUILD_SECONDS, 1)
