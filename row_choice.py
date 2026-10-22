# -*- coding: utf-8 -*-
"""Banner de eleccion cuando una referencia sale VARIAS VECES en la misma hoja de la tarifa
de un proveedor (formatos, colores, acabados o precios distintos con el mismo codigo).

Manda el CODIGO/REFERENCIA: dos filas con la misma referencia son el mismo articulo si el
programa importaria lo mismo de ambas; solo entonces se colapsan en una y NO se pregunta.
Si importarian algo distinto se pregunta, mostrando que cambia para que la persona elija.
Las filas que ya se importaron (historico) salen en gris y no se pueden volver a elegir.
"""
from DON_ARTICULO import _get_first_number_from_row, _norm_cell, _pick_description, _to_float

MAX_OPCIONES = 60        # mas opciones que esto = la referencia no es de un articulo (codigo de familia, color...)
                         # (hay codigos con muchas variantes reales de color/envase)
MAX_FILAS_ANALIZAR = 500  # por encima, ni se comparan las filas: es claramente un codigo generico
MAX_A_COMPARAR = 100     # filas distintas que se comparan por lo que importaria el programa (cada una = 1 conversion)
MAX_COLUMNAS_DIFF = 3    # columnas que cambian que se muestran por opcion
_PRICE_COLUMNS = ('PRECIO', 'PRECIOS', 'PVP', 'PRICE', 'PRECIO COMPRA', 'PRECIO EURO',
                  'PRECIO TARIFA', 'PRECIO PVP', 'PVP TARIFA', 'TARIFA', 'TARIFA 1', 'NETO',
                  'PRECIO NETO')
_PRICE_HINTS = ('PRECIO', 'PVP', 'PRICE', 'TARIFA')
_SHOWN_HINTS = ('DESC', 'PRECIO', 'PVP', 'PRICE')   # columnas que ya salen en la opcion


def price_for_display(row):
    """Precio de tarifa de una fila (columna conocida o, si no, cualquier columna de precio)."""
    price = _get_first_number_from_row(row, _PRICE_COLUMNS)
    if price is not None:
        return price
    for key, value in (row or {}).items():
        if any(h in _norm_cell(key) for h in _PRICE_HINTS):
            number = _to_float(value)
            if number is not None and number > 0:
                return number
    return None


def money(value):
    return ('%.2f' % value).replace('.', ',') + ' €'


def _signature(row):
    return tuple(sorted((_norm_cell(k), _norm_cell(v)) for k, v in row.items()))


def distinct_rows(df, rows):
    """Indices de `rows` quitando las filas 100% identicas a una anterior (todas las columnas)."""
    seen = set()
    distinct = []
    for r in rows:
        sig = _signature(df.loc[r].to_dict())
        if sig not in seen:
            seen.add(sig)
            distinct.append(r)
    return distinct


def _distinct_by_import(df, rows, signature):
    """Quita las filas que el programa importaria igual que una anterior (signature(idx, row) igual).
    Si no se pudo calcular (None) la fila se conserva: no se puede probar que sea la misma."""
    seen = set()
    kept = []
    for r in rows:
        sig = signature(r, df.loc[r].to_dict())
        if sig is not None:
            if sig in seen:
                continue
            seen.add(sig)
        kept.append(r)
    return kept


def _differing_columns(rows_data):
    """Columnas (en el orden de la hoja) cuyo valor cambia entre las filas, sin las que ya se muestran."""
    columns = list(next(iter(rows_data.values())).keys())
    diff = []
    for col in columns:
        name = _norm_cell(col)
        if any(h in name for h in _SHOWN_HINTS):
            continue
        values = {_norm_cell(row.get(col)) for row in rows_data.values()}
        if len(values) > 1:
            diff.append(col)
        if len(diff) == MAX_COLUMNAS_DIFF:
            break
    return diff


def _fmt(value):
    if isinstance(value, float) and value == value and value.is_integer():
        return str(int(value))      # 8400000000017.0 -> 8400000000017
    return str(value).strip()


def _short(value, size=40):
    text = ' '.join(_fmt(value).split())      # sin saltos de linea ni espacios repetidos (cabeceras largas)
    return text if len(text) <= size else text[:size - 1] + '…'


def _too_many(hit, reference, count):
    return {'error': "La referencia '%s' coincide con %d filas de la tarifa de %s: parece un código de familia, "
                     "colección o color, no la referencia de un artículo. Teclea la referencia completa del artículo."
                     % (reference, count, hit['provider_name'])}


def subtitle_for(options, ninguna_importada):
    """Texto bajo el titulo segun cuantas opciones esten ya importadas (en gris)."""
    enabled = [o for o in options if not o.get('disabled')]
    if len(enabled) == len(options):
        return ninguna_importada
    if not enabled:
        return 'YA ESTÁN IMPORTADAS TODAS (SALEN EN GRIS)'
    if len(enabled) == 1:
        nota = enabled[0].get('note')
        return '¿QUIERES IMPORTAR LA OTRA? ES «%s»%s' % (enabled[0]['title'][:60], (' · ' + nota) if nota else '')
    return '¿QUIERES IMPORTAR OTRA? LAS YA IMPORTADAS SALEN EN GRIS'


def row_choice_payload(hit, reference, signature=None, annotate=None):
    """Datos del banner si la referencia sale varias veces (y distintas) en la hoja ganadora;
    None si no hay duda (una sola fila o todas importarian lo mismo).
      signature(idx, row) -> clave de lo que importaria el programa de esa fila (None si no se pudo).
      annotate(idx, row)  -> {'disabled': bool, 'badge': str, 'note': str} o None (historico)."""
    df, rows = hit['df'], hit['rows']
    if len(rows) > MAX_FILAS_ANALIZAR:
        return _too_many(hit, reference, len(rows))
    distinct = distinct_rows(df, rows)
    if len(distinct) < 2:
        return None
    if len(distinct) > MAX_A_COMPARAR:
        return _too_many(hit, reference, len(distinct))
    if signature is not None:
        distinct = _distinct_by_import(df, distinct, signature)
        if len(distinct) < 2:
            return None
    if len(distinct) > MAX_OPCIONES:
        return _too_many(hit, reference, len(distinct))
    data = {r: df.loc[r].to_dict() for r in distinct}
    diff_cols = _differing_columns(data)
    options = []
    for n, r in enumerate(distinct, 1):
        row = data[r]
        desc = _pick_description(row)
        price = price_for_display(row)
        lines = ['Fila %d de la hoja \'%s\'' % (r + 2, hit['sheet'])]
        if price is not None:
            lines.insert(0, 'Precio tarifa: ' + money(price))
        changes = ['%s: %s' % (_short(c, 25), _short(row.get(c))) for c in diff_cols
                   if _norm_cell(row.get(c)) not in ('', 'NAN')]
        if changes:
            lines.append('Cambia: ' + '  ·  '.join(changes))
        option = {
            'provider': hit['provider_name'],
            'row': int(r),
            'title': _short(desc, 110) if desc else 'Opción %d' % n,
            'lines': lines,
        }
        extra = annotate(r, row) if annotate else None
        if extra:
            option.update({k: v for k, v in extra.items() if v})
        options.append(option)
    return {
        'success': False,
        'needs_choice': True,
        'kind': 'row',
        'reference': reference,
        'title': 'ESTA REFERENCIA SALE %d VECES EN LA TARIFA DE %s' % (len(distinct), hit['provider_name']),
        'subtitle': subtitle_for(options, 'CONCRETA CUÁL ES LA TUYA'),
        'candidates': options,
    }
