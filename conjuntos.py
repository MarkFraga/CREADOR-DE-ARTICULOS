"""Aviso "esta referencia es parte de un conjunto" para DON ARTICULO.

No lee red, ERP ni Selenium: solo una tabla Excel local por proveedor.

Que hace
    Cuando alguien teclea una referencia INDIVIDUAL de un proveedor configurado, por ejemplo A500, que realmente
    forma parte de un conjunto (A499+A500), el programa debe avisar de que normalmente se compra junto con la otra
    referencia (A499) y preguntar. Hay dos casos (columna "Tambien suelta"):
        NO  -> la pieza solo existe como parte de un conjunto.
        SI  -> la pieza tambien se vende suelta por si sola en la tarifa de fabrica; el aviso lo dice (valora las dos opciones).
    Pregunta:
        ACEPTAR -> se introduce solo la individual (flujo normal, sin cambios).
        NO      -> se introduce el conjunto (referencia x+y-z).
    Las referencias que no estan en la tabla (cualquier otro codigo, cualquier otro proveedor) NO avisan nunca.

Como se amplia a otro proveedor:
    1. Crear su tabla con las mismas columnas que la de ejemplo (Codigo, Referencia conjunta, Conjuntos, Origen, Tambien suelta).
    2. Añadir una linea en PROVEEDORES_CON_AVISO: 'NOMBRENORMALIZADO': 'TABLA.xlsx'.
    Para desactivar el aviso entero basta con vaciar ese diccionario.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Dict, Mapping, Optional, Tuple

# Proveedor (clave normalizada como _normalize_provider: mayusculas y sin espacios) -> tabla Excel.
# Aqui solo hay un proveedor ficticio de ejemplo. Para desactivar el aviso entero basta con vaciar este diccionario.
PROVEEDORES_CON_AVISO: Mapping[str, str] = {'PROVEEDOR_EJEMPLO': 'conjuntos_ejemplo.xlsx'}

MAX_COMPANERAS_EN_TEXTO = 6          # en el mensaje se listan como mucho estas; el resto "y N mas"
SEPARADOR = ';'

TITULO = 'ATENCIÓN: ESTA REFERENCIA ES PARTE DE UN CONJUNTO'
BOTON_ACEPTAR = 'ACEPTAR: INTRODUCIR SOLO ESTA'
BOTON_CONJUNTO = 'NO: INTRODUCIR EL CONJUNTO'
BOTON_CANCELAR = 'NO: CANCELAR (EL CONJUNTO NO ESTÁ EN LA TARIFA)'      # cuando la tarifa no trae ninguna referencia del conjunto


@dataclass(frozen=True)
class Conjunto:
    """Una referencia individual y los conjuntos de los que sale."""
    codigo: str
    companeras: Tuple[str, ...]      # la otra parte de cada x+y (la que NO va en el codigo individual)
    conjuntos: Tuple[str, ...]       # codigos suma x+y-z de los que sale
    origen: str
    tambien_suelta: bool = False     # la tarifa de fabrica tambien la vende suelta (no solo dentro de un conjunto)


@dataclass(frozen=True)
class Aviso:
    """Todo lo que la pantalla necesita para preguntar. Texto ya redactado."""
    proveedor: str
    referencia: str
    companeras: Tuple[str, ...]
    conjuntos: Tuple[str, ...]
    tambien_suelta: bool
    titulo: str
    mensaje: str
    boton_aceptar: str
    boton_conjunto: str

    @property
    def hay_conjunto(self) -> bool:
        return bool(self.conjuntos)

    @property
    def conjunto_unico(self) -> Optional[str]:
        """Si "NO" tiene un solo conjunto posible, es ese; si hay varios, la persona elige uno de `conjuntos`."""
        return self.conjuntos[0] if len(self.conjuntos) == 1 else None


def normaliza_proveedor(nombre: object) -> str:
    """Igual que _normalize_provider de DON_ARTICULO.py (sin alias): mayusculas y sin espacios."""
    return re.sub(r'\s+', '', str(nombre or '').strip().upper())


def normaliza_referencia(referencia: object) -> str:
    return re.sub(r'\s+', '', str(referencia or '').strip().upper())


def _lista(celda: object) -> Tuple[str, ...]:
    return tuple(p.strip() for p in str(celda or '').split(SEPARADOR) if p.strip())


def cargar_tabla(ruta: str) -> Dict[str, Conjunto]:
    """Lee la tabla Excel (cabecera: Codigo | Referencia conjunta | Conjuntos | Origen | Tambien suelta)."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    tabla: Dict[str, Conjunto] = {}
    try:
        for fila in wb.active.iter_rows(min_row=2, values_only=True):
            if not fila or fila[0] in (None, ''):
                continue
            codigo = normaliza_referencia(fila[0])
            suelta = len(fila) > 4 and str(fila[4] or '').strip().upper() in ('SI', 'SÍ', 'S', '1', 'TRUE')
            tabla[codigo] = Conjunto(codigo, _lista(fila[1]), _lista(fila[2]), str(fila[3] or ''), suelta)
    finally:
        wb.close()                    # en Windows, si no se cierra el libro el Excel queda bloqueado
    return tabla


def redacta_mensaje(referencia: str, companeras: Tuple[str, ...], conjuntos: Tuple[str, ...], tambien_suelta: bool = False) -> str:
    if len(companeras) == 1:
        junto = f'con la referencia {companeras[0]}'
    elif len(companeras) > MAX_COMPANERAS_EN_TEXTO:       # hay muchas piezas compatibles: una lista asi no ayuda
        junto = f'con una pieza compañera compatible (hay {len(companeras)} opciones)'
    else:
        junto = f'con una de estas referencias: {", ".join(companeras)}'
    if not conjuntos:
        conj = ' (el conjunto no está en la tarifa)'
    elif len(companeras) > MAX_COMPANERAS_EN_TEXTO:       # ya dice "hay N opciones": no se repite
        conj = ''
    elif len(conjuntos) == 1:
        conj = f' (conjunto {conjuntos[0]})'
    else:
        conj = f' ({len(conjuntos)} conjuntos posibles)'
    if tambien_suelta:
        inicio = f'La referencia {referencia} también se vende suelta, pero forma parte de conjuntos y normalmente se compra junto {junto}{conj}. '
    else:
        inicio = f'La referencia {referencia} es una parte que normalmente se compra junto {junto}{conj}. '
    return inicio + f'Si aceptas, se introduce SOLO la {referencia}, de forma individualizada.'


class GestorAvisos:
    """Carga perezosa de las tablas (una por proveedor) desde `carpeta` y contesta si hay que avisar."""

    def __init__(self, carpeta: str, proveedores: Optional[Mapping[str, str]] = None) -> None:
        self._carpeta = carpeta
        self._proveedores = dict(PROVEEDORES_CON_AVISO if proveedores is None else proveedores)
        self._tablas: Dict[str, Dict[str, Conjunto]] = {}

    def aplica(self, proveedor: object) -> bool:
        return normaliza_proveedor(proveedor) in self._proveedores

    def _tabla(self, clave: str) -> Dict[str, Conjunto]:
        if clave not in self._tablas:
            ruta = os.path.join(self._carpeta, self._proveedores[clave])
            if not os.path.isfile(ruta):
                return {}                          # sin tabla no hay aviso; no se recuerda, por si aparece despues
            self._tablas[clave] = cargar_tabla(ruta)
        return self._tablas[clave]

    def aviso(self, proveedor: object, referencia: object) -> Optional[Aviso]:
        """None = no hay que avisar (otro proveedor, o la referencia no es una individual de conjunto)."""
        clave = normaliza_proveedor(proveedor)
        if clave not in self._proveedores:
            return None
        ref = normaliza_referencia(referencia)
        info = self._tabla(clave).get(ref)
        if info is None or not info.companeras:
            return None
        return Aviso(clave, ref, info.companeras, info.conjuntos, info.tambien_suelta, TITULO,
                     redacta_mensaje(ref, info.companeras, info.conjuntos, info.tambien_suelta), BOTON_ACEPTAR,
                     BOTON_CONJUNTO if info.conjuntos else BOTON_CANCELAR)
