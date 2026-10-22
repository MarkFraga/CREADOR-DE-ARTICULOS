# -*- coding: utf-8 -*-
"""Genera una configuracion y una carpeta DATABASE de EJEMPLO, con datos ficticios, para probar DON ARTICULO.

    python examples/make_example_config.py            # escribe en la carpeta del proyecto
    python examples/make_example_config.py --out ./prueba

Crea:
    ayuda.xlsx                      proveedor, descuentos, nº de proveedor, D.PROFESIONAL, D.PARTICULAR
    PROVS.xlsx                      categoria principal de cada proveedor
    exceltodo.xlsx                  descuentos por categoria y proveedor
    overrides.xlsx                  categorias forzadas por referencia o patron
    conjuntos_ejemplo.xlsx          tabla del aviso "esta referencia es parte de un conjunto"
    DATABASE/<proveedor>/TARIFA_*.xlsx   tarifas de ejemplo (una carpeta por proveedor)

Todo es inventado: ningun proveedor, precio ni referencia es real.
"""
import argparse
import os

from openpyxl import Workbook

PROVEEDOR_PERFIL = "PROVEEDOR_EJEMPLO"                # tiene perfil en ROBOT.PROVIDER_PROFILES
PROVEEDOR_CERAMICA = "PROVEEDOR_001"                  # ceramica "pura" (ROBOT.FORCED_CERAMIC_PROVIDERS)

AYUDA_HEADERS = ["PROVEEDOR", "DESCUENTOS", "NÚMERO PROVEEDOR", "D.PROFESIONAL", "D.PARTICULAR"]
PRO_DISCOUNT = 28.0        # % de Tarifa 2 (profesional) sobre Tarifa 1
PRIVATE_DISCOUNT = 17.0    # % de Tarifa 3 (particular) sobre Tarifa 1
AYUDA_ROWS = [
    [PROVEEDOR_PERFIL, 38, 1000, PRO_DISCOUNT, PRIVATE_DISCOUNT],
    # "1,8+4,5" = formula ceramica (x1,8 + 4,5), no un descuento
    [PROVEEDOR_CERAMICA, "1,8+4,5", 1001, PRO_DISCOUNT, PRIVATE_DISCOUNT],
]

PROVS_HEADERS = ["PROVEEDOR", "TIPO"]
PROVS_ROWS = [[PROVEEDOR_PERFIL, "SANITARIO"], [PROVEEDOR_CERAMICA, "CERAMICA"]]

EXCELTODO_HEADERS = ["TIPO", "PROVEEDOR", "DESCUENTO PROFESINAL", "DESCUENTO PARTICULAR"]
EXCELTODO_ROWS = [
    ["GRIFERIA", "TODOS", 21, 9],
    ["ACCESORIOS", "TODOS", PRO_DISCOUNT, PRIVATE_DISCOUNT],
    ["RESTO", "TODOS", PRO_DISCOUNT, PRIVATE_DISCOUNT],
]

OVERRIDES_HEADERS = ["PROVEEDOR", "REFERENCIA", "PATRON", "CATEGORIA"]
OVERRIDES_ROWS = [[PROVEEDOR_PERFIL, "A300", "", "ACCESORIOS"]]

CONJUNTOS_HEADERS = ["Codigo", "Referencia conjunta", "Conjuntos", "Origen", "Tambien suelta"]
CONJUNTOS_ROWS = [["A500", "A499", "A499+A500", "ejemplo", "NO"]]

TARIFA_PERFIL_HEADERS = ["Cod. Articulo", "Articulo", "Precio", "EAN"]
TARIFA_PERFIL_ROWS = [
    ["A100", "MEZCLADOR LAVABO (EXPOSICION)", 120.0, "840"],
    ["A200", "INODORO SUSPENDIDO RIMLESS BLANCO", 210.5, ""],
    ["A300", "TOALLERO BARRA 60 CM CROMO", 35.9, ""],
    ["A499", "PLATO DE DUCHA RESINA 120X80", 180.0, ""],
    ["A500", "SIFON PARA PLATO DE DUCHA", 22.0, ""],
    ["CER001", "REFERENCIA REPETIDA EN OTRO PROVEEDOR", 99.0, ""],
]

TARIFA_CERAMICA_HEADERS = ["Referencia", "Descripcion", "Serie", "Color", "Precio", "Unidad", "M2/CAJA", "Piezas/caja",
                           "Cajas/palet", "Kg/caja", "EAN"]
TARIFA_CERAMICA_ROWS = [
    ["CER001", "PAVIMENTO PORCELANICO 60X120", "SERIE A", "BLANCO MATE", 10.5, "M2", 1.44, 2, 40, 30.0, "8400000000024"],
    ["CER002", "REVESTIMIENTO CERAMICO 30X90", "SERIE B", "GRIS", 8.2, "M2", 1.35, 5, 48, 24.0, ""],
    ["CER003", "RODAPIE PORCELANICO 8X60", "SERIE A", "BLANCO MATE", 3.1, "PZ", "", 10, 60, 18.0, ""],
    # Misma referencia dos veces con distinto color y precio: el programa pregunta cual es.
    ["CER004", "PAVIMENTO PORCELANICO 80X80", "SERIE C", "ARENA", 12.0, "M2", 1.28, 2, 40, 31.0, ""],
    ["CER004", "PAVIMENTO PORCELANICO 80X80", "SERIE C", "ANTRACITA", 12.9, "M2", 1.28, 2, 40, 31.0, ""],
]


def _write(path, headers, rows):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)


def build(out_dir):
    """Escribe todos los ficheros de ejemplo bajo `out_dir` y devuelve la ruta de la carpeta DATABASE."""
    _write(os.path.join(out_dir, "ayuda.xlsx"), AYUDA_HEADERS, AYUDA_ROWS)
    _write(os.path.join(out_dir, "PROVS.xlsx"), PROVS_HEADERS, PROVS_ROWS)
    _write(os.path.join(out_dir, "exceltodo.xlsx"), EXCELTODO_HEADERS, EXCELTODO_ROWS)
    _write(os.path.join(out_dir, "overrides.xlsx"), OVERRIDES_HEADERS, OVERRIDES_ROWS)
    _write(os.path.join(out_dir, "conjuntos_ejemplo.xlsx"), CONJUNTOS_HEADERS, CONJUNTOS_ROWS)
    database = os.path.join(out_dir, "DATABASE")
    _write(os.path.join(database, PROVEEDOR_PERFIL, "TARIFA_EJEMPLO_2026.xlsx"), TARIFA_PERFIL_HEADERS, TARIFA_PERFIL_ROWS)
    _write(os.path.join(database, PROVEEDOR_CERAMICA, "TARIFA_CERAMICA_2026.xlsx"), TARIFA_CERAMICA_HEADERS, TARIFA_CERAMICA_ROWS)
    return database


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--out", default=os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        help="carpeta de destino (por defecto, la carpeta del proyecto)")
    args = parser.parse_args()
    db = build(args.out)
    print("Configuracion de ejemplo escrita en:", args.out)
    print("DATABASE de ejemplo:", db)
    print("Para usarla:  set DONARTICULO_DATABASE=%s" % db)
