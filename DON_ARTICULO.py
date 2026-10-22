# -*- coding: utf-8 -*-
"""
DON ARTICULO
============
Wrapper GUI for ROBOT.py (Single Article Import).

- User only types the reference and clicks "Introducir".
- Reads provider configuration (provider number, discounts) from ayuda.xlsx
- Uses ROBOT.py internal classes to:
  1) search reference in DATABASE
  2) convert to ERP template
  3) fill %Dto1 / %Dto2 in columns L/M from the provider discounts table
  4) write importart.xlsx to the network folder
  5) automate Firefox import (always enabled)

NOTE: This file DOES NOT modify ROBOT.py. It imports and reuses its classes.
"""

import os
import re
import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

try:
    from PIL import Image, ImageTk, ImageSequence
    HAS_PIL = True
except Exception:
    HAS_PIL = False

try:
    import tariff_sql          # tarifas grandes leidas como SQL en vez de releer el Excel; si falta, todo sigue con el Excel
except Exception as _e_sql:
    tariff_sql = None
    print(f"[!] tariff_sql no disponible ({_e_sql}): las tarifas se leen del Excel como antes.")

# Consultas a servicios de terceros (buscador de imagenes, IA Gemini, buscador web): desactivadas por defecto.
# Activarlas (DONARTICULO_ONLINE_LOOKUPS=1) envia proveedor, referencia y descripcion a esos servicios (ver README, "Privacidad").
ONLINE_LOOKUPS = os.environ.get("DONARTICULO_ONLINE_LOOKUPS") == "1"

# -------------------------
# UI Theme (Blue, modern)
# -------------------------
AZUL_FONDO = "#071B33"
AZUL_CARD  = "#0B2A4A"
AZUL_BORDE = "#1D4E89"
AZUL_BOTON = "#2B8CFF"
TEXTO      = "#EAF6FF"
TEXTO_2    = "#A9C9FF"

# -------------------------
# Helpers
# -------------------------

def _normalize_provider(name: str) -> str:
    """Normalize provider names (upper, no spaces) and apply common aliases."""
    p = re.sub(r'\s+', '', str(name or '').strip().upper())
    return _PROVIDER_ALIASES.get(p, p)

# Nombres alternativos de un mismo proveedor (nombre normalizado -> nombre canonico). Ejemplo ficticio.
_PROVIDER_ALIASES = {
    'PROVEEDOR_002': 'PROVEEDOR_EJEMPLO',
}
# Proveedores cuyo catalogo es siempre ceramica (ver tambien ROBOT.FORCED_CERAMIC_PROVIDERS). Ejemplo ficticio.
_FORCED_CERAMIC_PROVIDERS = ['PROVEEDOR_001']
# Proveedores con varias categorias: su entrada en PROVS.xlsx es solo su categoria "principal". Ejemplo ficticio.
MULTI_CATEGORY_PROVIDERS = {'PROVEEDOR_003'}

def _fetch_google_image_bytes(query: str):
    if not ONLINE_LOOKUPS:
        return None
    import urllib.request
    import urllib.parse
    encoded_query = urllib.parse.quote(query)
    url = f"https://www.google.com/search?tbm=isch&q={encoded_query}"
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    try:
        response = urllib.request.urlopen(req, timeout=5)
        html = response.read().decode('utf-8', errors='ignore')
        matches = re.findall(r'(https://encrypted-tbn0\.gstatic\.com/images[^"]+)', html)
        if matches:
            img_url = matches[0]
            img_req = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
            img_response = urllib.request.urlopen(img_req, timeout=5)
            return img_response.read()
    except Exception:
        pass
    return None

def _analyze_image_with_gemini(image_bytes: bytes, description: str) -> str:
    if not ONLINE_LOOKUPS:
        return None
    try:
        from google import genai
    except ImportError:
        return None
        
    api_key = os.environ.get("GEMINI_API_KEY")
    key_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'gemini_key.txt')
    if not api_key and os.path.exists(key_path):
        with open(key_path, 'r', encoding='utf-8') as f:
            api_key = f.read().strip()
            
    if not api_key:
        return None
        
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
    Eres un experto en materiales de construcción, reformas y artículos de baño/cocina.
    Mira la imagen adjunta de este producto: '{description}'
    
    ¿A qué categoría pertenece? RESPONDE ÚNICAMENTE CON UNA DE ESTAS PALABRAS EXACTAS (en mayúsculas):
    CEMENTO
    CERÁMICA
    GRIFERÍA
    SANITARIO
    ACCESORIOS
    MUEBLES
    MAMPARAS
    COMPLEMENTOS
    REPUESTOS
    VARIOS
    
    Si ves un fregadero de cocina de acero o similar, es GRIFERÍA.
    Si ves una campana extractora o electrodoméstico, es VARIOS.
    Si ves un plato de ducha, examina si es de MAMPARAS o SANITARIO según corresponda al sistema. Pero en general plato de ducha a veces es SANITARIO.
    """
    try:
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(image_bytes))
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[img, prompt]
        )
        cat = response.text.strip().upper()
        valid_cats = ["CEMENTO", "CERÁMICA", "GRIFERÍA", "SANITARIO", "ACCESORIOS", "MUEBLES", "MAMPARAS", "COMPLEMENTOS", "REPUESTOS", "VARIOS"]
        for v in valid_cats:
            if v in cat:
                return v
    except Exception:
        pass
    return None

def _is_forced_ceramic(provider_name: str) -> bool:
    p = _normalize_provider(provider_name)
    return p in _FORCED_CERAMIC_PROVIDERS

def _safe_set(df, col: str, value):
    val = "" if value is None else value
    try:
        df.loc[0, col] = val
    except Exception:
        try:
            df[col] = df[col].astype(object)
            df.loc[0, col] = val
        except Exception:
            pass

def _to_float(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    s = s.replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

def _get_first_from_row(row: dict, names):
    for n in names:
        nn = _norm_cell(n)
        for kk, vv in (row or {}).items():
            if _norm_cell(kk) == nn:
                if vv is None:
                    continue
                sv = str(vv).strip()
                if not sv or sv.lower() == "nan":
                    continue
                return vv
    return None

def _get_first_number_from_row(row: dict, names):
    return _to_float(_get_first_from_row(row, names))

def _find_value_by_colpattern(row: dict, patterns):
    row = row or {}
    for kk, vv in row.items():
        if vv is None:
            continue
        sv = str(vv).strip()
        if not sv or sv.lower() == "nan":
            continue
        kn = _norm_cell(kk)
        for p in patterns:
            if p in kn:
                return vv
    return None

# Cabeceras habituales (variantes en castellano) de los datos logisticos de una tarifa de ceramica.
_HDR_UDX = ("UdxEnvase", "UDXENVASE", "UDXENVASES", "UD X ENVASE", "UD/ENVASE", "UD ENVASE", "UD CAJA", "UD/CAJA",
            "PIEZAS POR CAJA", "PIEZAS/CAJA", "PCS/CAJA", "PZ/CAJA", "PIEZAS CAJA", "PCS CAJA", "PZ CAJA")
_HDR_MIN = ("Unid.Minima", "Unid. Mínima", "UNID.MINIMA", "UNID. MINIMA", "UNID MINIMA", "UNIDAD MINIMA", "UNIDAD MÍNIMA",
            "M2/CAJA", "M2 CAJA", "M²/CAJA", "M2 POR CAJA")
_HDR_UPA = ("Ud.Palet", "UD.PALET", "UD PALET", "UD/PALET", "CAJAS/PALET", "CAJAS PALET", "UNIDADES PALET", "UD POR PALET",
            "BULTOS/PALET", "CAJAS X PALET", "CAJAS POR PALET")
_HDR_PP = ("Peso Palet", "PESO PALET", "KG/PALET", "PESO PALETIZADO", "PESO PALET.", "PESO POR PALET", "KG POR PALET",
           "KG X PALET")
_HDR_PESO = ("Peso", "PESO", "PESO CAJA", "PESO/CAJA", "KG/CAJA", "PESO KG", "KG CAJA", "KG POR CAJA", "KG X CAJA")

def _find_number_by_colpattern(row: dict, patterns):
    return _to_float(_find_value_by_colpattern(row, patterns))

def _norm_cell(x) -> str:
    # Codigos numericos leidos como float por pandas (cuando la columna de
    # referencia tiene huecos NaN, -> float64) llegan como
    # 123456789012.0; sin esto str() da '123456789012.0' y NUNCA casa con la
    # referencia tecleada '123456789012'. Solo afecta a valores float (los
    # codigos de texto de otros proveedores quedan intactos).
    if isinstance(x, float):
        if x != x:                 # NaN
            s = ""
        elif x.is_integer():
            s = str(int(x))
        else:
            s = str(x)
    else:
        s = "" if x is None else str(x)
    return re.sub(r"\s+", " ", s).strip().upper()

def _norm_header(h: str) -> str:
    h = "" if h is None else str(h)
    h = h.strip().upper()
    trans = str.maketrans("ÁÀÄÂÉÈËÊÍÌÏÎÓÒÖÔÚÙÜÛÑ", "AAAAEEEEIIIIOOOOUUUUN")
    h = h.translate(trans)
    h = re.sub(r"[\.:/\-_]+", " ", h)
    h = re.sub(r"\s+", " ", h).strip()
    return h

def _find_col_by_norm(df, *candidates):
    cand_norm = set()
    for c in candidates:
        if isinstance(c, (list, tuple, set)):
            for x in c:
                cand_norm.add(_norm_header(str(x)))
        else:
            cand_norm.add(_norm_header(str(c)))
            
    for c in df.columns:
        if _norm_header(str(c)) in cand_norm:
            return c
    return None

def _find_col_by_norm_multi(df, candidates):
    return _find_col_by_norm(df, candidates)

def _norm_colname(name: str) -> str:
    return _norm_cell(name).replace(".", "").replace("Á","A").replace("É","E").replace("Í","I").replace("Ó","O").replace("Ú","U")

def _safe_set_norm(df, wanted_cols, value):
    c = _find_col_by_norm(df, wanted_cols)
    if c is not None:
        _safe_set(df, c, value)

def _is_ean_col(colname: str) -> bool:
    c = _norm_cell(colname)
    return "EAN" in c

def _candidate_ref_columns(df):
    cols=[]
    for c in df.columns:
        cn=_norm_cell(c)
        if not cn or _is_ean_col(cn):
            continue
        # Quitar tildes para que cabeceras tipo "CÓDIGO TARIFA" o "ARTÍCULO" se detecten igual que sin tilde
        cn_ascii = _strip_accents(cn)
        if cn_ascii in ("REFERENCIA","REF","CODIGO","COD.","COD ARTIC","COD.ARTIC","CODARTIC","ARTICULO","ARTIC","COD.ARTICULO","COD. ARTICULO"):
            cols.insert(0,c)
        elif ("REFER" in cn_ascii or cn_ascii == "REF" or "COD" in cn_ascii) and "DESC" not in cn_ascii:
            cols.append(c)
    return cols

def _row_matches_reference_exact(row_data: dict, reference: str) -> bool:
    ref_norm = _norm_cell(reference)
    if not ref_norm:
        return False
        
    reference_clean = str(reference).upper().replace(' ', '').replace('-', '').replace('.', '').lstrip('0')
        
    for kk, vv in (row_data or {}).items():
        if vv is None:
            continue
        kn = _norm_cell(kk)
        if _is_ean_col(kn):
            continue
            
        val_clean = str(vv).upper().replace(' ', '').replace('-', '').replace('.', '').lstrip('0')
            
        if kn in ("REFERENCIA","REF","CÓDIGO","CODIGO","COD.ARTIC","COD ARTIC","CODARTIC","ARTICULO","ARTÍCULO","ARTIC","COD.ARTICULO","COD ARTICULO"):
            if _norm_cell(vv).lstrip('0') == ref_norm.lstrip('0') or val_clean == reference_clean:
                return True
        if ("REFER" in kn or kn == "REF" or "COD" in kn) and "DESC" not in kn:
            if _norm_cell(vv).lstrip('0') == ref_norm.lstrip('0') or val_clean == reference_clean:
                return True
                
        # If any other non-descriptive column perfectly matches the clean value, allow it
        # This handles cases where the column name is completely different but contains the exact reference
        if "DESC" not in kn and val_clean == reference_clean:
            return True
            
    return False

def _provider_name_for(root: str, filename: str) -> str:
    """Proveedor de una tarifa: la carpeta que la contiene (o el nombre del fichero si
    cuelga directamente de DATABASE)."""
    provider_name = os.path.splitext(filename)[0]
    parent = os.path.basename(root)
    if parent and parent.lower() not in ("", "database", "tarifas"):
        provider_name = parent
    return provider_name

def _iter_tariff_files(database_root: str):
    """(ruta, fichero, proveedor) de cada tarifa .xlsx de DATABASE en el orden de busqueda:
    carpetas por orden de os.walk y, dentro de cada una, primero las tarifas con un año de
    4 digitos en el nombre."""
    for root, _, files in os.walk(database_root):
        xlsx_files = sorted(
            [f for f in files if f.lower().endswith(".xlsx") and not f.startswith("~$")],
            key=lambda x: (0 if re.search(r'\d{4}', x) else 1, x),
        )
        for f in xlsx_files:
            yield os.path.join(root, f), f, _provider_name_for(root, f)

def _leer_excel_hojas(path):
    """Hojas de una tarifa leidas del Excel: {nombre de hoja: DataFrame}."""
    import pandas as pd
    return pd.read_excel(path, sheet_name=None)

def _read_tariff_sheets(path, logger=None):
    """Hojas de una tarifa de DATABASE listas para buscar.
    Las tarifas de tariff_sql.SQL_PROVIDERS salen de SQLite/memoria con el MISMO contenido que el Excel."""
    if tariff_sql is not None and tariff_sql.usa_sql(_provider_name_for(os.path.dirname(path), os.path.basename(path))):
        try:
            return tariff_sql.hojas(path, lambda: _leer_excel_hojas(path), _candidate_ref_columns, _norm_cell, logger)
        except tariff_sql.NoSql:
            pass                    # no se puede servir desde SQL: el Excel de siempre
    return _leer_excel_hojas(path)

_SQL_SALTAR = object()      # el fichero SQL no trae la referencia: se pasa al siguiente fichero

def _hit_desde_sql(path, fichero, provider_name, ref_norm, row_index, logger):
    """find_reference_hit para una tarifa servida desde SQL: consulta el indice de referencias (sin recorrer el DataFrame).
    Devuelve el dict de coincidencia, None (la fila pedida no es de esta referencia: la busqueda termina), _SQL_SALTAR
    (no esta en este fichero) o lanza tariff_sql.NoSql (usar el Excel)."""
    coincidencias = tariff_sql.buscar(path, ref_norm)
    if coincidencias is None:        # primera vez o tarifa cambiada: se construye el SQL (lanza NoSql si no se puede) y se vuelve a preguntar
        tariff_sql.hojas(path, lambda: _leer_excel_hojas(path), _candidate_ref_columns, _norm_cell, logger)
        coincidencias = tariff_sql.buscar(path, ref_norm)
        if coincidencias is None:
            raise tariff_sql.NoSql('sin indice SQL valido')
    if not coincidencias:
        return _SQL_SALTAR
    sheet_name, first_col, first_idx, rows = coincidencias[0]
    idx = first_idx if row_index is None else row_index
    if idx not in rows:
        return None                  # la fila pedida no es de esta referencia
    df = tariff_sql.hojas(path, lambda: _leer_excel_hojas(path), _candidate_ref_columns, _norm_cell, logger)[sheet_name]
    if idx not in df.index or any(i not in df.index for i in rows):      # el Excel cambio entre la consulta al indice y la lectura: se usa el Excel
        raise tariff_sql.NoSql('la tarifa cambio durante la busqueda')
    if logger:
        logger(f"✓ Coincidencia EXACTA en {provider_name} | {fichero} | hoja '{sheet_name}' | columna '{first_col}' (SQL)")
    return {
        'provider_name': provider_name, 'path': path, 'file': fichero, 'sheet': sheet_name,
        'df': df, 'row_data': df.loc[idx].to_dict(), 'row_index': idx,
        'rows': rows, 'matched_ref': ref_norm,
    }

def find_reference_hit(database_root: str, reference: str, logger=None, provider_filter=None, row_index=None):
    """Primera coincidencia EXACTA de la referencia en las tarifas de DATABASE (mismo orden de
    busqueda de siempre). Devuelve un dict con provider_name, path, file, sheet, df, row_data,
    row_index (fila elegida), rows (TODAS las filas de esa hoja donde casa la referencia) y
    matched_ref; None si no esta. provider_filter limita la busqueda a un proveedor; row_index
    fuerza esa fila de la hoja ganadora en vez de la primera."""
    import os
    import pandas as pd

    def _search(ref_norm):
        if not ref_norm:
            return None
        for root, _, files in os.walk(database_root):
            # Sort files so that tariffs with year (e.g. "2026") come first
            xlsx_files = sorted(
                [f for f in files if f.lower().endswith(".xlsx") and not f.startswith("~$")],
                key=lambda x: (0 if re.search(r'\d{4}', x) else 1, x),
            )
            for f in xlsx_files:
                provider_name = _provider_name_for(root, f)
                if provider_filter is not None and provider_name != provider_filter:
                    continue
                path = os.path.join(root, f)
                if tariff_sql is not None and tariff_sql.usa_sql(provider_name):
                    try:
                        hit_sql = _hit_desde_sql(path, f, provider_name, ref_norm, row_index, logger)
                    except tariff_sql.NoSql:
                        hit_sql = False                  # sin SQL valido para este fichero: el camino de siempre
                    except Exception as e:
                        if logger:
                            logger(f"Aviso: SQL de {f} fallo ({e}); se lee el Excel")
                        hit_sql = False
                    if hit_sql is _SQL_SALTAR:
                        continue
                    if hit_sql is not False:
                        return hit_sql                   # coincidencia, o None si la fila pedida no es de esta referencia
                try:
                    sheets = _read_tariff_sheets(path)
                except Exception:
                    continue

                for sheet_name, df in sheets.items():
                    if df is None or df.empty:
                        continue
                    cand_cols = _candidate_ref_columns(df)
                    if not cand_cols:
                        continue

                    first_col, first_idx, rows = None, None, set()
                    for c in cand_cols:
                        try:
                            s = df[c].map(_norm_cell)
                        except Exception:
                            continue
                        mask = (s == ref_norm)
                        if mask.any():
                            if first_col is None:
                                first_col, first_idx = c, mask.idxmax()
                            rows.update(mask[mask].index.tolist())
                    if first_col is None:
                        continue
                    rows = sorted(rows)
                    idx = first_idx if row_index is None else row_index
                    if idx not in rows:
                        return None     # la fila pedida no es de esta referencia
                    if logger:
                        logger(f"✓ Coincidencia EXACTA en {provider_name} | {f} | hoja '{sheet_name}' | columna '{first_col}'")
                    return {
                        'provider_name': provider_name, 'path': path, 'file': f, 'sheet': sheet_name,
                        'df': df, 'row_data': df.loc[idx].to_dict(), 'row_index': idx,
                        'rows': rows, 'matched_ref': ref_norm,
                    }
        return None

    ref_norm = _norm_cell(reference)
    hit = _search(ref_norm)
    if hit:
        return hit

    return None

def find_reference_anysheet(database_root: str, reference: str, logger=None, provider_filter=None):
    """(proveedor, ruta, fila, hoja_df) de la primera coincidencia exacta, o None. Ver
    find_reference_hit para los detalles (filas repetidas, hoja, fila concreta)."""
    hit = find_reference_hit(database_root, reference, logger=logger, provider_filter=provider_filter)
    if not hit:
        return None
    return hit['provider_name'], hit['path'], hit['row_data'], hit['df']

def _strip_accents(s: str) -> str:
    try:
        import unicodedata
        return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    except Exception:
        return s

def _norm_tipo(s: str) -> str:
    s = _strip_accents(str(s or '')).upper().strip()
    s = re.sub(r'\s+', ' ', s)
    return s

# -------------------------
# Overrides system (Phase 6)
# -------------------------
_OVERRIDES_DATA = []  # List of dicts: {proveedor, referencia, patron, categoria}

def _cell_text(value, default=''):
    """Texto de una celda en mayusculas; una celda vacia (NaN) vale `default`, no la cadena 'NAN'."""
    if value is None or (isinstance(value, float) and value != value):
        return default
    return str(value).strip().upper() or default

def load_overrides(excel_path: str):
    """Load overrides.xlsx if it exists. Each row can override category by ref or pattern."""
    global _OVERRIDES_DATA
    _OVERRIDES_DATA = []
    if not os.path.exists(excel_path):
        return
    try:
        import pandas as pd
        df = pd.read_excel(excel_path)
        for _, r in df.iterrows():
            prov = _cell_text(r.get('PROVEEDOR'), '*')
            ref = _cell_text(r.get('REFERENCIA'))
            patron = _cell_text(r.get('PATRON'))
            cat = _norm_tipo(_cell_text(r.get('CATEGORIA')))
            if not cat:
                continue
            _OVERRIDES_DATA.append({
                'proveedor': prov,
                'referencia': ref,
                'patron': patron,
                'categoria': cat,
            })
    except Exception:
        pass

def check_category_override(provider_name, reference, description):
    """Check if overrides.xlsx forces a specific category. Returns category string or None."""
    if not _OVERRIDES_DATA:
        return None
    prov_up = _normalize_provider(provider_name) if provider_name else ''
    ref_up = str(reference or '').strip().upper()
    desc_up = str(description or '').strip().upper()
    for ov in _OVERRIDES_DATA:
        if ov['proveedor'] != '*' and _normalize_provider(ov['proveedor']) != prov_up:
            continue
        if ov['referencia'] and ov['referencia'] != ref_up:
            continue
        if ov['patron']:
            import fnmatch
            search_text = ref_up + ' ' + desc_up
            if not fnmatch.fnmatch(search_text, ov['patron']):
                continue
        return ov['categoria']
    return None

def load_provs_map(excel_path: str):
    import pandas as pd
    df = pd.read_excel(excel_path)
    out = {}
    if not {'PROVEEDOR','TIPO'}.issubset(set(df.columns)):
        return out
    for _, r in df.iterrows():
        prov = _normalize_provider(r.get('PROVEEDOR'))
        tip = _norm_tipo(r.get('TIPO'))
        if prov:
            out[prov] = tip
    return out

def load_exceltodo_rules(excel_path: str):
    import pandas as pd
    df = pd.read_excel(excel_path)
    out = {}
    for _, r in df.iterrows():
        tip = _norm_tipo(r.get('TIPO'))
        prov = _normalize_provider(r.get('PROVEEDOR'))
        if not tip or not prov:
            continue
        try:
            dp = float(r.get('DESCUENTO PROFESINAL'))
            dpa = float(r.get('DESCUENTO PARTICULAR'))
        except Exception:
            continue
        out[(tip, prov)] = (dp, dpa)
    return out

def _pick_description(row_data) -> str:
    """Devuelve la descripción del producto leyendo la primera clave disponible
    (Descripcion Corta, Descripcion, Descripcion Larga, Denominacion, Nombre…),
    tolerante a mayúsculas/minúsculas, tildes y espacios.
    """
    if not row_data:
        return ''
    wanted = (
        'DESCRIPCION CORTA', 'DESCRIPCION LARGA', 'DESCRIPCION',
        'DENOMINACION', 'NOMBRE', 'DETALLE', 'TEXTO',
    )
    norm_map = {}
    for k, v in row_data.items():
        if v is None:
            continue
        ks = str(k).upper()
        for a, b in (('Á', 'A'), ('É', 'E'), ('Í', 'I'), ('Ó', 'O'), ('Ú', 'U'), ('Ñ', 'N')):
            ks = ks.replace(a, b)
        ks = re.sub(r'[^A-Z0-9]+', ' ', ks).strip()
        if ks and ks not in norm_map:
            norm_map[ks] = v
    for w in wanted:
        v = norm_map.get(w)
        if v is None:
            continue
        s = str(v).strip()
        if s and s.lower() != 'nan':
            return s
    return ''


_GRIFERIA_FORCE_WORDS = (
    'MEZCLADOR', 'MEZCLADORES',
    'MONOMANDO', 'MONOMANDOS',
    'BIMANDO', 'BIMANDOS',
    'GRIFO', 'GRIFOS', 'GRIFERIA',
    'TERMOSTATO', 'TERMOSTATICO', 'TERMOSTATICA',
)

_GRIFERIA_NEGATIVE_WORDS = (
    'REPUESTO', 'REPUESTOS', 'RECAMBIO', 'RECAMBIOS',
    'CARTUCHO', 'CARTUCHOS', 'AIREADOR', 'JUNTA',
)

def _is_unambiguous_griferia(text) -> bool:
    """Devuelve True si la descripción contiene un término inequívoco de grifería
    (mezclador, monomando, grifo, termostato…) y no es un repuesto/cartucho.
    Sirve como override duro: cualquier proveedor con esos términos -> GRIFERÍA.
    """
    if not text:
        return False
    t = str(text).upper()
    for a, b in (('Á', 'A'), ('É', 'E'), ('Í', 'I'), ('Ó', 'O'), ('Ú', 'U'), ('Ñ', 'N')):
        t = t.replace(a, b)
    padded = ' ' + re.sub(r'[^A-Z0-9]+', ' ', t).strip() + ' '
    if not any(' ' + w + ' ' in padded for w in _GRIFERIA_FORCE_WORDS):
        return False
    if any(' ' + w + ' ' in padded for w in _GRIFERIA_NEGATIVE_WORDS):
        return False
    return True


def _bucket_from_category(cat: str) -> str:
    c = _norm_tipo(cat)
    if c.startswith('CER'):
        return 'CERAMICA'
    if c.startswith('GRIF'):
        return 'GRIFERIA'
    if c.startswith('SANIT'):
        return 'SANITARIO'
    if c.startswith('ACCES'):
        return 'ACCESORIOS'
    if c.startswith('MUEB') or c.startswith('MOBIL'):
        return 'MUEBLES'
    if c.startswith('MAMP'):
        return 'MAMPARAS'
    if c.startswith('CEMEN'):
        return 'CEMENTO'
    if c.startswith('COMPL'):
        return 'COMPLEMENTOS'
    if c.startswith('REPU') or c.startswith('RECAM'):
        return 'REPUESTOS'
    if c.startswith('VAR'):
        return 'VARIOS'
    return 'RESTO'

def pick_discounts(provider_name: str, detected_category: str,
                   base_pro: float, base_part: float,
                   provs_map: dict, rules: dict):
    prov_key = _normalize_provider(provider_name)
    bucket = _bucket_from_category(detected_category)

    # For multi-category providers, skip the provs_map shortcut and go straight to exceltodo
    # Multi-category providers: their PROVS.xlsx entry is just their "main" category
    is_multi = prov_key in MULTI_CATEGORY_PROVIDERS

    main = provs_map.get(prov_key)
    if main and not is_multi:
        main_bucket = _bucket_from_category(main)
        if main_bucket == bucket:
            return base_pro, base_part, f"Proveedor (tipo principal {main_bucket})"

    for key in [
        (bucket, prov_key),
        (bucket, _normalize_provider('RESTO')),
        (bucket, _normalize_provider('TODOS')),
        ('RESTO', _normalize_provider('TODOS')),
    ]:
        if key in rules:
            dp, dpa = rules[key]
            return dp, dpa, f"exceltodo {key[0]}/{key[1]}"
    return base_pro, base_part, "Fallback proveedor"

def _is_converted_template(df) -> bool:
    if df is None or getattr(df, "columns", None) is None:
        return False
    cols = {str(c).strip().lower() for c in df.columns}
    template_cols = {
        "proveedor", "fabricante",
        "tarifa 1", "tarifa1", "tarifa 2", "tarifa2", "tarifa 3", "tarifa3",
        "%dto1", "%dto2",
        "precio compra", "precio base tarifa",
        "codigo y referencia", "código y referencia", "codigo", "código", "articulo", "artículo", "referencia", "ref"
    }
    hits = sum(1 for c in cols if c in template_cols)
    has_supplier_fields = ("proveedor" in cols and "fabricante" in cols)
    has_prices = any(c in cols for c in ("tarifa 1","tarifa1")) and any(c in cols for c in ("tarifa 2","tarifa2","tarifa 3","tarifa3"))
    return hits >= 6 and has_supplier_fields and has_prices

def _parse_descuentos(val):
    if val is None or (isinstance(val, float) and val != val):
        return '', '', None, None

    s = str(val).strip()
    if re.fullmatch(r'\d+(?:[.,]\d+)?\s*\+\s*\d+(?:[.,]\d+)?', s) and (',' in s or '.' in s):
        a, b = re.split(r'\s*\+\s*', s)
        a = float(a.replace(',', '.'))
        b = float(b.replace(',', '.'))
        if 1.5 <= a <= 5.0:
            return '', '', a, b

    if re.fullmatch(r'\d+\s*\+\s*\d+', s):
        a, b = re.split(r'\s*\+\s*', s)
        return a.strip(), b.strip(), None, None

    if re.fullmatch(r'\d+(?:[.,]\d+)?', s):
        return s.replace(',', '.').strip(), '', None, None

    return '', '', None, None

def load_provider_table(excel_path: str):
    import pandas as pd
    df = pd.read_excel(excel_path)
    required = {'PROVEEDOR', 'DESCUENTOS', 'NÚMERO PROVEEDOR', 'D.PROFESIONAL', 'D.PARTICULAR'}
    missing = required - set(df.columns)
    if missing:
        raise RuntimeError(f"Faltan columnas en ayuda.xlsx: {', '.join(sorted(missing))}")

    table = {}
    for _, r in df.iterrows():
        prov = _normalize_provider(r['PROVEEDOR'])
        dto1, dto2, cmult, cadd = _parse_descuentos(r['DESCUENTOS'])
        table[prov] = {
            'provider_number': str(int(r['NÚMERO PROVEEDOR'])) if str(r['NÚMERO PROVEEDOR']).strip() else '',
            'pro_discount': float(r['D.PROFESIONAL']),
            'private_discount': float(r['D.PARTICULAR']),
            'dto1': dto1,
            'dto2': dto2,
            'ceramic_mult': cmult,
            'ceramic_add': cadd,
        }
    return table

# -------------------------
# UI
# -------------------------

class AnimatedGif(ttk.Label):
    def __init__(self, master, gif_path, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.gif_path = gif_path
        self.frames = []
        self._job = None
        self._i = 0

        if HAS_PIL and os.path.exists(gif_path):
            im = Image.open(gif_path)
            for frame in ImageSequence.Iterator(im):
                fr = frame.convert("RGBA")
                self.frames.append(ImageTk.PhotoImage(fr))

    def start(self, delay=90):
        if not self.frames:
            return
        self.stop()
        self._i = 0
        def _step():
            self.configure(image=self.frames[self._i])
            self._i = (self._i + 1) % len(self.frames)
            self._job = self.after(delay, _step)
        _step()

    def stop(self):
        if self._job:
            try:
                self.after_cancel(self._job)
            except Exception:
                pass
            self._job = None


class DonArticuloApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DON ARTICULO")
        self.root.geometry("860x820")
        self.root.resizable(True, True)

        base_dir = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__)))
        self.base_dir = base_dir

        try:
            self.root.iconbitmap(os.path.join(base_dir, "don_articulo.ico"))
        except Exception:
            pass

        try:
            self.root.configure(bg=AZUL_FONDO)
        except Exception:
            pass

        self.style = ttk.Style(self.root)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.style.configure("TFrame", background=AZUL_FONDO)
        self.style.configure("Card.TFrame", background=AZUL_CARD)
        self.style.configure("TLabel", background=AZUL_FONDO, foreground=TEXTO)
        self.style.configure("Sub.TLabel", background=AZUL_FONDO, foreground=TEXTO_2)
        self.style.configure("TLabelframe", background=AZUL_FONDO, foreground=TEXTO)
        self.style.configure("TLabelframe.Label", background=AZUL_FONDO, foreground=TEXTO)
        self.style.configure("TEntry", fieldbackground="#0E335C", foreground=TEXTO)
        self.style.configure("TButton", background=AZUL_BOTON, foreground="white", padding=(14,10))
        self.style.map("TButton", background=[("active", "#1E6FD6")])

        self.provider_table = load_provider_table(os.path.join(base_dir, "ayuda.xlsx"))
        self.provs_map = load_provs_map(os.path.join(base_dir, "PROVS.xlsx")) if os.path.exists(os.path.join(base_dir, "PROVS.xlsx")) else {}
        self.exceltodo_rules = load_exceltodo_rules(os.path.join(base_dir, "exceltodo.xlsx")) if os.path.exists(os.path.join(base_dir, "exceltodo.xlsx")) else {}
        load_overrides(os.path.join(base_dir, "overrides.xlsx"))

        sys.path.insert(0, base_dir)
        try:
            import ROBOT as robot
        except Exception as e:
            raise RuntimeError("No se pudo importar ROBOT.py. Asegúrate de que ROBOT.py esté en la misma carpeta que DON_ARTICULO.py") from e
        self.robot = robot

        self.gif_path = os.path.join(base_dir, "en_proceso.gif")
        self.completed_img_path = os.path.join(base_dir, "completado.png")
        self.logo_path = os.path.join(base_dir, "toilet_logo.png")
        self.logo_img = None
        if HAS_PIL and os.path.exists(self.logo_path):
            try:
                img = Image.open(self.logo_path).convert("RGBA")
                img = img.resize((64, 64))
                self.logo_img = ImageTk.PhotoImage(img)
            except Exception:
                self.logo_img = None
        self.completed_img = None
        if HAS_PIL and os.path.exists(self.completed_img_path):
            self.completed_img = ImageTk.PhotoImage(Image.open(self.completed_img_path).convert("RGBA").resize((720, 400)))

        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self.root, padding=14, style="TFrame")
        frm.pack(fill="both", expand=True)

        top = ttk.Frame(frm, style="TFrame")
        top.pack(fill="x", pady=(0, 10))
        if getattr(self, "logo_img", None) is not None:
            ttk.Label(top, image=self.logo_img, style="TLabel").pack(side="left", padx=(0, 10))
        title = ttk.Label(top, text="DON ARTICULO", font=("Segoe UI", 24, "bold"), style="TLabel")
        title.pack(side="left")

        sub = ttk.Label(frm, text="Escribe la referencia y pulsa Introducir.  El resto lo hace el robot.",
                        font=("Segoe UI", 12), style="Sub.TLabel")
        sub.pack(pady=(0, 12), anchor="w")

        ref_box = ttk.LabelFrame(frm, text="Referencia", padding=10)
        ref_box.pack(fill="x")

        self.ref_var = tk.StringVar()
        entry = ttk.Entry(ref_box, textvariable=self.ref_var, font=("Segoe UI", 16))
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        entry.bind("<Return>", lambda e: self.on_start())
        entry.focus_set()

        self.btn = ttk.Button(ref_box, text="Introducir", command=self.on_start)
        self.btn.pack(side="right")

        status_box = ttk.LabelFrame(frm, text="Estado", padding=10)
        status_box.pack(fill="both", pady=14, expand=False)

        self.status_var = tk.StringVar(value="Listo.")
        self.status_lbl = ttk.Label(status_box, textvariable=self.status_var, font=("Segoe UI", 16, "bold"), style="TLabel")
        self.status_lbl.pack(anchor="w")

        self.anim = AnimatedGif(status_box, self.gif_path)
        self.anim.pack(pady=12)

        self.img_lbl = ttk.Label(status_box)
        self.img_lbl.pack(pady=12)

        log_box = ttk.LabelFrame(frm, text="Log", padding=10)
        log_box.pack(fill="both", expand=True)

        self.log_txt = tk.Text(log_box, height=18, wrap="word", bg="#06162B", fg=TEXTO, insertbackground=TEXTO, relief="flat")
        self.log_txt.pack(fill="both", expand=True)
        self.log_txt.configure(state="disabled")

        self._set_completed(False)

    def _log(self, msg: str):
        self.log_txt.configure(state="normal")
        self.log_txt.insert("end", msg + "\n")
        self.log_txt.see("end")
        self.log_txt.configure(state="disabled")

    def _set_busy(self, busy: bool):
        self.btn.configure(state=("disabled" if busy else "normal"))
        if busy:
            self._set_completed(False)
            try:
                self.img_lbl.pack_forget()
            except Exception:
                pass
            try:
                self.anim.pack(pady=12)
            except Exception:
                pass
            self.anim.start()
        else:
            self.anim.stop()

    def _set_completed(self, completed: bool):
        if completed and self.completed_img:
            try:
                self.anim.stop()
                self.anim.pack_forget()
            except Exception:
                pass
            self.img_lbl.configure(image=self.completed_img)
            self.img_lbl.image = self.completed_img
            try:
                self.img_lbl.pack(pady=8)
            except Exception:
                pass
        else:
            self.img_lbl.configure(image="")
            self.img_lbl.image = None

    def on_start(self):
        ref = self.ref_var.get().strip()
        if not ref:
            messagebox.showwarning("Falta referencia", "Introduce una referencia.")
            return

        self._set_busy(True)
        self.status_var.set("EN PROCESO…")
        self._log("========================================")
        self._log(f"Referencia: {ref}")

        t = threading.Thread(target=self._run_process, args=(ref,), daemon=True)
        t.start()

    def _run_process(self, reference: str):
        try:
            robot = self.robot

            self._log(f"Buscando referencia: {reference} (coincidencia exacta, todas las hojas)")
            found = find_reference_anysheet(robot.DEFAULT_DATABASE_PATH, reference, logger=self._log)

            if not found:
                scanner = robot.ProviderDatabaseScanner(robot.DEFAULT_DATABASE_PATH, logger=self._log)
                found = scanner.find_reference(reference)
                if found:
                    _p, _tf, _rd, _df = found
                    if not _row_matches_reference_exact(_rd, reference):
                        self._log("⚠ Coincidencia parcial ignorada (no coincide toda la celda).")
                        found = None

            if not found:
                raise RuntimeError("No se encontró la referencia en ningún proveedor.")

            provider_name, tariff_file, row_data, tariff_df = found
            self._log(f"Encontrado en proveedor: {provider_name}")
            prov_key = _normalize_provider(provider_name)

            if prov_key not in self.provider_table:
                raise RuntimeError(f"El proveedor '{provider_name}' no está en ayuda.xlsx (tabla de descuentos).")

            cfg = self.provider_table[prov_key]
            provider_number = cfg['provider_number']
            pro_discount = cfg['pro_discount']
            private_discount = cfg['private_discount']
            dto1 = cfg['dto1']
            dto2 = cfg['dto2']
            cmult = cfg['ceramic_mult']
            cadd = cfg['ceramic_add']

            self._log(f"Nº proveedor: {provider_number} | Dto profesional: {pro_discount}% | Dto particular: {private_discount}%")
            if dto1 or dto2:
                self._log(f"%Dto1/%Dto2 (L/M): {dto1}{' + ' + dto2 if dto2 else ''}")

            converter = robot.SingleRowConverter()
            user_inputs = {
                'provider_number': provider_number,
                'pro_discount': pro_discount,
                'private_discount': private_discount,
            }
            result_df, category, cat_info = converter.convert_row(row_data, tariff_df, provider_name, user_inputs)
            self._log(f"Categoría detectada: {category}")

            if _is_forced_ceramic(provider_name):
                if str(category).upper()[:3] != "CER":
                    self._log(f"Forzando categoría CERÁMICA por proveedor: {provider_name}")
                category = "CERAMICA"

            # Phase 6: Check overrides.xlsx
            ref_val = str(row_data.get('Referencia', row_data.get('referencia', reference)))
            desc_val = str(row_data.get('Descripcion', row_data.get('Descripción', '')))
            cat_override = check_category_override(provider_name, ref_val, desc_val)
            if cat_override:
                self._log(f"Override manual: {category} -> {cat_override}")
                category = cat_override
                cat_info = robot.CATEGORIES.get(category, robot.CATEGORIES.get('VARIOS', {}))

            # Phase 7: AI fallback for low confidence or VARIOS
            confidence = converter.classifier.get_confidence()

            # Phase 7b: Generic AI fallback
            if ONLINE_LOOKUPS and (str(category).upper() == "VARIOS" or confidence < 30):
                self._log(f"Confianza baja ({confidence}) o VARIOS - intentando IA...")
                try:
                    ai_desc = str(row_data.get('Descripcion', row_data.get('Descripción', reference)))
                    ai_query = f"{provider_name} {reference} {ai_desc}"
                    img_bytes = _fetch_google_image_bytes(ai_query)
                    if img_bytes:
                        ai_cat = _analyze_image_with_gemini(img_bytes, ai_desc)
                        if ai_cat and ai_cat != "VARIOS":
                            self._log(f"IA sugiere: {ai_cat}")
                            category = ai_cat
                            cat_info = robot.CATEGORIES.get(category, robot.CATEGORIES.get('VARIOS', {}))
                except Exception as e:
                    self._log(f"Error IA: {e}")

            # Override universal: términos inequívocos de grifería siempre fuerzan GRIFERÍA
            _griferia_desc = _pick_description(row_data)
            if _is_unambiguous_griferia(_griferia_desc):
                if str(category).upper().replace('Í', 'I') != 'GRIFERIA':
                    self._log(f"Override GRIFERÍA por término inequívoco en '{_griferia_desc[:60]}' (antes: {category})")
                    category = "GRIFERÍA"
                    cat_info = robot.CATEGORIES.get(category, robot.CATEGORIES.get('VARIOS', {}))

            adj_pro, adj_part, reason = pick_discounts(provider_name, category, pro_discount, private_discount, self.provs_map, self.exceltodo_rules)
            if adj_pro != pro_discount or adj_part != private_discount:
                self._log(f"Ajuste descuentos: {pro_discount}%/{private_discount}% -> {adj_pro}%/{adj_part}% ({reason})")
                pro_discount, private_discount = adj_pro, adj_part

                base_val = None
                for c in ("Tarifa 1", "Precio Base Tarifa", "Tarifa1"):
                    if c in result_df.columns:
                        try:
                            base_val = float(result_df.loc[0, c])
                            break
                        except Exception:
                            pass
                if base_val is not None:
                    if "Tarifa 2" in result_df.columns:
                        result_df.loc[0, "Tarifa 2"] = round(base_val * (1 - pro_discount/100.0), 6)
                    if "Tarifa 3" in result_df.columns:
                        result_df.loc[0, "Tarifa 3"] = round(base_val * (1 - private_discount/100.0), 6)
            else:
                self._log(f"Descuentos OK ({reason}): {pro_discount}%/{private_discount}%")

            # Extraer descuentos de la fila del proveedor si existen (H, I, J...)
            def _pick_dto(row: dict, *names):
                for n in names:
                    nn = _norm_cell(n)
                    for kk, vv in (row or {}).items():
                        if _norm_cell(kk) == nn and vv is not None and str(vv).strip():
                            v = str(vv).strip()
                            if v.lower() != 'nan':
                                return v
                return None

            dto1_row = _pick_dto(row_data, "%Dto1", "Dto1", "Descuento 1", "% Dto1", "%Dto 1", "Dto 1")
            dto2_row = _pick_dto(row_data, "%Dto2", "Dto2", "Descuento 2", "% Dto2", "%Dto 2", "Dto 2")
            dto3_row = _pick_dto(row_data, "%Dto3", "Dto3", "Descuento 3", "% Dto3", "%Dto 3", "Dto 3")
            dto4_row = _pick_dto(row_data, "%Dto4", "Dto4", "Descuento 4", "% Dto4", "%Dto 4", "Dto 4")

            if dto1_row: dto1 = dto1_row
            if dto2_row: dto2 = dto2_row
            dto3 = dto3_row if dto3_row else ''
            dto4 = dto4_row if dto4_row else ''

            if dto1:
                _safe_set(result_df, '%Dto1', dto1)
            if dto2:
                _safe_set(result_df, '%Dto2', dto2)
            if dto3:
                _safe_set(result_df, '%Dto3', dto3)
            if dto4:
                _safe_set(result_df, '%Dto4', dto4)

            # =========================================================
            # 1. LÓGICA DE CÁLCULO BASE Y FÓRMULAS (A vs B)
            # =========================================================
            # Detección global de descuento:
            has_discount = bool(dto1 or dto2)
            
            for kk, vv in (row_data or {}).items():
                kn = _norm_header(kk).replace(" ", "")
                if kn in ("%DTO1", "%DTO2", "%DTO3", "%DTO4", "DTO1", "DTO2", "DTO3", "DTO4", "DESCUENTO", "DTO", "%DTO"):
                    val = str(vv).strip()
                    if val and val.lower() != "nan" and val not in ("0", "0.0", "0,0"):
                        has_discount = True
                        break

            if str(category).upper().startswith("CER"):
                # Un perfil con price_as_is (precio tal cual, sin formula ceramica) fuerza Calculo Base B
                if robot.provider_profile(provider_name).get('price_as_is'):
                    has_discount = True

            col_calculo_base = _find_col_by_norm(result_df, ["Calculo Base", "Cálculo Base"])
            if col_calculo_base:
                default_cb = cat_info.get('calculo_base', 'A') if isinstance(cat_info, dict) else 'A'
                final_cb = "B" if has_discount else default_cb
                _safe_set(result_df, col_calculo_base, final_cb)

            def _clean_text(s: str) -> str:
                s = re.sub(r"\s+", " ", (s or "")).strip()
                s = re.sub(r"\bPRODUCTO\b", "", s, flags=re.I).strip()
                s = re.sub(r"\bREF:\s*[^\s]+", "", s, flags=re.I).strip()
                s = re.sub(r"\s+", " ", s).strip(" -_")
                return s

            def _is_numeric_code(s: str) -> bool:
                s = (s or "").strip()
                return bool(re.fullmatch(r"\d+(\.\d+)*", s))

            def _bad_token(tok: str, refv: str) -> bool:
                t = (tok or "").strip()
                if not t:
                    return True
                if t.lower() == "nan":
                    return True
                if refv and t.upper() == refv.upper():
                    return True
                if t.upper() in {"M2","M²","UD","PZ","PCS","KG","M","MM","CM"}:
                    return True
                if re.fullmatch(r"[A-Z]", t.upper()):
                    return True
                if t.upper() in {"S","N"}:
                    return True
                return False

            def _pick(row: dict, *names):
                for n in names:
                    nn = _norm_cell(n)
                    for kk, vv in (row or {}).items():
                        if _norm_cell(kk) == nn and vv is not None and str(vv).strip():
                            v = _clean_text(str(vv))
                            if v and not _is_numeric_code(v):
                                return v
                return None

            def _get_ref_from_result(df):
                for c in ("Código", "Codigo", "Codigo y referencia", "Código y referencia", "Articulo", "Artículo", "Referencia", "Ref.Prov.", "Ref Prov", "Ref"):
                    if c in df.columns:
                        v = _clean_text(str(df.loc[0, c]))
                        if v:
                            return v
                return None

            # Rescatar valores EXACTOS de la fila original del proveedor para blindar la lectura
            source_pc = None
            source_pbt = None
            source_t1 = None
            
            for kk, vv in (row_data or {}).items():
                kn = _norm_header(kk)
                if kn in ("PRECIO COMPRA", "PRECIO DE COMPRA"):
                    if source_pc is None: source_pc = _to_float(vv)
                elif kn == "PRECIO BASE TARIFA":
                    source_pbt = _to_float(vv)
                elif kn in ("TARIFA 1", "TARIFA1"):
                    source_t1 = _to_float(vv)

            col_pc = _find_col_by_norm(result_df, ["Precio Compra", "Precio de Compra"])
            col_q = _find_col_by_norm(result_df, ["Precio Base Tarifa"])
            col_u = _find_col_by_norm(result_df, ["Tarifa 1", "Tarifa1"])
            
            if source_pc is not None and col_pc:
                _safe_set(result_df, col_pc, source_pc)
                pc_val = source_pc
                self._log(f"Precio Compra ({pc_val}) forzado desde origen.")
            else:
                pc_val = _to_float(result_df.loc[0, col_pc]) if col_pc else 0.0

            is_full_template = (source_pc is not None and source_pbt is not None and source_t1 is not None)
            
            if is_full_template and not has_discount:
                base_calc = source_pbt
                t1_calc = source_t1
                self._log(f"Plantilla elaborada detectada. Tarifas rescatadas directas: Base={base_calc}, T1={t1_calc}")
            elif has_discount:
                # Si tiene descuento (Cálculo Base B), el precio base no sufre incremento logístico
                base_calc = pc_val
                t1_calc = base_calc
                self._log(f"Calculo Base B: Precio Compra ({pc_val}) aplicado directamente sin fórmula cerámica")
            elif str(category).upper().startswith("CER"):
                mult = cmult if cmult is not None else robot.CERAMIC_MULT
                add = cadd if cadd is not None else robot.CERAMIC_ADD
                base_calc = round((pc_val * mult) + add, 6)
                t1_calc = base_calc
                self._log(f"Fórmula cerámica clásica aplicada: ({pc_val} x {mult}) + {add} = {base_calc}")
            else:
                base_calc = pc_val
                t1_calc = base_calc
                self._log(f"Artículo No-Cerámico (Cat: {category}): Base Tarifa directa = {base_calc}")
            
            if col_q: _safe_set(result_df, col_q, base_calc)
            if col_u: _safe_set(result_df, col_u, t1_calc)
            
            # Descuentos de Tarifa 2 / Tarifa 3 (la categoria GRIFERIA y el perfil del proveedor los pueden cambiar)
            final_pro_discount = pro_discount
            final_private_discount = private_discount
            
            _profile = robot.provider_profile(provider_name)      # particularidades del proveedor (ROBOT.PROVIDER_PROFILES)
            if _profile.get('dto1') is not None:
                col_dto1 = _find_col_by_norm(result_df, ["%Dto1", "Dto1", "% Dto1"])
                if col_dto1:
                    _safe_set(result_df, col_dto1, _profile['dto1'])

            if hasattr(robot, "ProductClassifier"):
                use_web = bool(_profile.get('verify_with_web'))
                classifier = robot.ProductClassifier(enable_web_verification=use_web)
                description_for_cat = _pick(row_data, "DESCRIPCION", "DESCRIPCIÓN", "DESCRIPCION CORTA", "DESCRIPCIÓN CORTA", "DENOMINACION", "DENOMINACIÓN", "NOMBRE") or "PRODUCTO"
                cat = classifier.classify_product(description_for_cat, verify_with_web=use_web)

                # Override universal: términos inequívocos de grifería siempre fuerzan GRIFERÍA
                if _is_unambiguous_griferia(description_for_cat) and cat != "GRIFERÍA":
                    self._log(f"Override GRIFERÍA por término inequívoco en '{str(description_for_cat)[:60]}' (antes: {cat})")
                    cat = "GRIFERÍA"

                if cat == "VARIOS" and ONLINE_LOOKUPS:
                    self._log(f"Categoría VARIOS. Intentando clasificar por Inteligencia Artificial (Gemini) usando imagen web...")
                    try:
                        img_bytes = _fetch_google_image_bytes(description_for_cat)
                        if img_bytes:
                            cat_ia = _analyze_image_with_gemini(img_bytes, description_for_cat)
                            if cat_ia:
                                cat = cat_ia
                                self._log(f"Gemini clasificó el artículo como: {cat}")
                            else:
                                self._log("Gemini devolvió categoría nula o error (revisa si tienes GEMINI_API_KEY o gemini_key.txt). Manteniendo VARIOS.")
                        else:
                            self._log("No se pudo obtener imagen para la IA. Manteniendo VARIOS.")
                    except Exception as e:
                        self._log(f"Error en IA: {e}")
                
                if cat == "GRIFERÍA":
                    final_pro_discount = robot.GRIFERIA_PRO_DISCOUNT
                    final_private_discount = robot.GRIFERIA_PRIVATE_DISCOUNT

            # El perfil del proveedor manda sobre la regla por categoria
            if _profile.get('pro_discount') is not None:
                final_pro_discount = _profile['pro_discount']
            if _profile.get('private_discount') is not None:
                final_private_discount = _profile['private_discount']

            col_t2 = _find_col_by_norm(result_df, ["Tarifa 2", "Tarifa2"])
            if col_t2: _safe_set(result_df, col_t2, round(t1_calc * (1 - final_pro_discount / 100.0), 6))
            col_t3 = _find_col_by_norm(result_df, ["Tarifa 3", "Tarifa3"])
            if col_t3: _safe_set(result_df, col_t3, round(t1_calc * (1 - final_private_discount / 100.0), 6))

            ref_val = _get_ref_from_result(result_df) or _pick(row_data, "REFERENCIA","REF","CÓDIGO","CODIGO","ARTICULO","ARTÍCULO") or _clean_text(str(reference))

            base_from_short = _pick(row_data, "Descripcion Corta", "Descripción Corta", "DESCRIPCION CORTA", "DESCRIPCIÓN CORTA")
            if base_from_short:
                text = base_from_short
            else:
                desc_main = _pick(row_data, "DESCRIPCION", "DESCRIPCIÓN", "DESCRIPCION LARGA", "DESCRIPCIÓN LARGA", "DENOMINACION", "DENOMINACIÓN", "NOMBRE")
                serie = _pick(row_data, "SERIE","COLECCION","COLECCIÓN","GAMA","LINEA","LÍNEA","COLECION")
                modelo = _pick(row_data, "MODELO","MODEL","NOMBRE MODELO")
                color = _pick(row_data, "COLOR","ACABADO","TERMINACION","TERMINACIÓN")
                medida = _pick(row_data, "MEDIDA","FORMATO","DIMENSION","DIMENSIÓN","TAMAÑO","TAMANO")

                base_parts = []
                for p in (serie, desc_main, modelo, color, medida):
                    if p:
                        base_parts.append(p)

                if not base_parts:
                    for kk, vv in (row_data or {}).items():
                        if vv is None:
                            continue
                        if _is_ean_col(str(kk)):
                            continue
                        if _norm_cell(kk).startswith("UNNAMED"):
                            continue
                        s = _clean_text(str(vv))
                        if not s or _is_numeric_code(s):
                            continue
                        if s.upper() == ref_val.upper():
                            continue
                        if re.fullmatch(r"[\d\.,\-]+", s):
                            continue
                        base_parts.append(s)

                tokens=[]
                for part in base_parts:
                    part=_clean_text(part)
                    if not part:
                        continue
                    for tok in part.split(" "):
                        if _bad_token(tok, ref_val):
                            continue
                        tokens.append(tok)

                text = re.sub(r"\s+", " ", " ".join(tokens)).strip()

            def _strip_ref(s: str) -> str:
                s = s or ""
                s = re.sub(r"\bREF:\s*[^\s]+", "", s, flags=re.I)
                s = re.sub(r"\s+", " ", s).strip(" -_")
                return s

            if text:
                base = _strip_ref(text)
                desc_short = base.upper()
                desc_long = (base + f" REF: {ref_val}").strip().upper()
            else:
                desc_long = f"REF: {ref_val}".upper()
                desc_short = ""

            for c in ("Descripción","Descripcion","DESCRIPCION","DESCRIPCIÓN"):
                if c in result_df.columns:
                    _safe_set(result_df, c, desc_long)
                    break

            for c in ("Descripcion Corta","Descripción Corta","DESCRIPCION CORTA","DESCRIPCIÓN CORTA","Descripcion corta","Descripción corta","BA","Descripcion Corta"):
                if c in result_df.columns:
                    _safe_set(result_df, c, desc_short)
                    break

            is_cer = str(category).upper().startswith("CER")
            for col in ("Aplicar palet (S/N)", "Paletizado S/N"):
                if col in result_df.columns:
                    _safe_set(result_df, col, "S" if is_cer else "N")
            if not is_cer and "Ud.Palet" in result_df.columns:
                _safe_set(result_df, "Ud.Palet", "")
            if "Ud. Med. Compra" in result_df.columns and "UD. MEDIDA" in result_df.columns:
                _safe_set(result_df, "Ud. Med. Compra", str(result_df.loc[0, "UD. MEDIDA"]))


            if str(category).upper().startswith("CER"):
                for col, val in (
                    ("Sección", "1"), ("Seccion", "1"),
                    ("Familia", "1.1"),
                    ("Subfamilia", "1.1.1"),
                ):
                    if col in result_df.columns:
                        _safe_set(result_df, col, val)


            # =========================================================
            # 2. CORRECCIÓN DE UNIDAD DE MEDIDA PARA TODAS LAS CERÁMICAS
            # =========================================================
            if str(category).upper().startswith("CER"):
                um = None
                
                # Para saber si una "Caja" es en realidad una pieza especial (Ej: Rodapié, Peldaño...)
                desc_val = ""
                for k in ("Descripcion Corta", "Descripción Corta", "Descripcion", "Descripción"):
                    if k in result_df.columns:
                        desc_val = str(result_df.loc[0, k] or "").upper()
                        break
                piezas_kws = ["RODAPIE", "RODAPIÉ", "PELDAÑO", "PELDANO", "TACO", "TABICA", "ZOCALO", "ZÓCALO", "CENEFA", "DECORADO", "PIEZA", "MALLA", "MOSAICO", "LISTELO"]
                is_pieza_especial = any(kw in desc_val for kw in piezas_kws)

                for kk, vv in (row_data or {}).items():
                    kn = _norm_header(kk)
                    if "UD MEDIDA" in kn or kn in ("UM","U M","UNIDAD MEDIDA","UD VENTA","UNIDAD VENTA","UNIDAD"):
                        sv = str(vv).strip().upper() if vv is not None else ""
                        if sv and sv != "NAN":
                            if sv in ("M2", "M²", "MT2", "MT"):
                                um = "M2"
                            elif sv in ("PZ", "PZA", "PCS", "PIEZA", "PIEZAS", "PI", "UD", "UDS", "UN"):
                                um = "PZ"
                            elif sv in ("CA", "CJ", "CJA", "CAJAS", "CAJA"):
                                um = "PZ" if is_pieza_especial else "M2"
                            elif sv in ("M", "ML", "METRO", "METROS", "METRO LINEAL", "METROS LINEALES"):
                                um = "M"
                            else:
                                um = sv.replace("M²","M2").replace("PCS","PZ").replace("PIEZAS","PZ").replace("PIEZA","PZ").replace("ML","M")
                            break

                if not um:
                    desc_try = ""
                    for k in ("Descripcion Corta","Descripción Corta","Descripción","Descripcion"):
                        if k in result_df.columns:
                            desc_try = str(result_df.loc[0, k] or "")
                            break
                    d = desc_try.upper()
                    if " M2" in d or "M2" in d or "M²" in d:
                        um = "M2"
                    elif " PZ" in d or "PZ" in d or "PCS" in d or "PIEZA" in d or " PI " in d:
                        um = "PZ"
                    elif " ML" in d or "ML" in d or " M " in d or "METRO" in d:
                        um = "M"
                    else:
                        um = "M2"

                col_um = _find_col_by_norm_multi(result_df, ["Ud.Medida", "UD. MEDIDA", "UD MEDIDA"])
                if col_um:
                    _safe_set(result_df, col_um, um)

                col_umc = _find_col_by_norm_multi(result_df, ["Ud. Med. Compra", "UD. MED. COMPRA", "UD MED COMPRA", "UD MEDIDA COMPRA"])
                if col_umc:
                    _safe_set(result_df, col_umc, um)
            else:
                # =========================================================
                # FORZADO DE UNIDAD DE MEDIDA 'UD' PARA NO-CERÁMICAS
                # =========================================================
                um = "UD"
                col_um = _find_col_by_norm_multi(result_df, ["Ud.Medida", "UD. MEDIDA", "UD MEDIDA"])
                if col_um:
                    _safe_set(result_df, col_um, um)

                col_umc = _find_col_by_norm_multi(result_df, ["Ud. Med. Compra", "UD. MED. COMPRA", "UD MED COMPRA", "UD MEDIDA COMPRA"])
                if col_umc:
                    _safe_set(result_df, col_umc, um)


            # =========================================================
            # 3. ESCANEO Y VOLCADO LOGÍSTICO EXACTO (ESPECIAL PLANTILLAS ELABORADAS)
            # =========================================================
            if str(category).upper().startswith("CER"):
                
                # Leemos la Unidad de Medida que hemos decidido en el paso anterior
                um_col = _find_col_by_norm(result_df, ["Ud.Medida", "UD MEDIDA"])
                um_val = str(result_df.loc[0, um_col] or "").upper().strip() if um_col else ""

                # Extraemos los datos escaneando directamente el Excel del proveedor
                v_udx = _get_first_number_from_row(row_data, _HDR_UDX)
                v_min = _get_first_number_from_row(row_data, _HDR_MIN)
                v_upa = _get_first_number_from_row(row_data, _HDR_UPA)
                v_pp  = _get_first_number_from_row(row_data, _HDR_PP)
                v_peso = _get_first_number_from_row(row_data, _HDR_PESO)

                # Si el proveedor oculta el peso bajo nombres raros, buscamos cualquier cosa con PESO (menos palet, pz, m2, etc)
                if v_peso is None:
                    pat_peso = ["PESO/CAJA", "KG/CAJA", "PESO CAJA", "PESO KG", "PESO"]
                    exclusion_patterns = ["PALET", "PL", "PZ", "PIEZA", "M2"]
                    for kk, vv in (row_data or {}).items():
                        kn = _norm_cell(kk)
                        if any(p in kn for p in pat_peso) and not any(e in kn for e in exclusion_patterns):
                            val = _to_float(vv)
                            if val is not None:
                                v_peso = val
                                break

                # Si no dan peso de palet, lo multiplicamos (Peso Caja * Cajas Palet)
                if v_pp is None and v_peso is not None and v_upa is not None:
                    v_pp = round(v_peso * v_upa, 2)
                    
                # Si falta el peso caja pero tenemos peso de palet y unidades por palet
                if v_peso is None and v_pp is not None and v_upa is not None and v_upa > 0:
                    v_peso = round(v_pp / v_upa, 2)
                
                # --- REGLA CRÍTICA PARA PIEZAS (PZ) VS METROS (M2) ---
                if um_val in ("PZ", "PZA", "PCS", "PIEZA", "PIEZAS", "PI"):
                    if v_udx is not None:
                        v_min = v_udx
                    elif v_min is None:
                        v_min = 1

                # EL GOLPE FINAL: Forzamos los múltiplos a ser idénticos a la Unidad Mínima de venta (AD)
                # Ignoramos el "1" mentiroso que ponen algunos proveedores.
                v_mcom = v_min
                v_mven = v_min

                # --- ESCRIBIR TODO EN EL ARCHIVO IMPORTART (con los nombres de columna exactos del formato) ---
                if v_udx is not None:
                    _safe_set_norm(result_df, ["UdxEnvase", "Udx Envase", "UDXENVASE"], v_udx)
                    
                if v_min is not None:
                    _safe_set_norm(result_df, ["Unid. Mínima", "Unid.Minima", "Unid Minima"], v_min)
                    
                if v_upa is not None:
                    _safe_set_norm(result_df, ["Ud.Palet", "Ud Palet"], v_upa)
                    
                if v_peso is not None:
                    _safe_set_norm(result_df, ["Peso"], v_peso)
                    
                if v_pp is not None:
                    _safe_set_norm(result_df, ["Peso Palet"], v_pp)
                    
                if v_mcom is not None:
                    _safe_set_norm(result_df, ["Multiplo Compra", "Multiplo Com"], v_mcom)
                    
                if v_mven is not None:
                    _safe_set_norm(result_df, ["Multiplo Venta", "Multiplo Ven"], v_mven)

            else:
                # Para productos NO cerámicos -> forzar UdxEnvase=1, Unid.Minima=1
                _safe_set_norm(result_df, ["UdxEnvase", "Udx Envase", "UDXENVASE"], 1)
                _safe_set_norm(result_df, ["Unid. Mínima", "Unid.Minima", "Unid Minima"], 1)
                _safe_set_norm(result_df, ["Multiplo Compra", "Multiplo Com"], 1)
                _safe_set_norm(result_df, ["Multiplo Venta", "Multiplo Ven"], 1)

            # 6) Save file to network folder
            fm = robot.NetworkFileManager(robot.DEFAULT_NETWORK_PATH, filename='importart')
            file_path = fm.save_import_file(result_df)
            self._log(f"Archivo generado: {file_path}")

            # 7) Automate import in Firefox (SIEMPRE)
            if not robot.HAS_SELENIUM:
                raise RuntimeError("ROBOT.py no tiene Selenium instalado. Instala dependencias o ejecuta en el PC donde funciona el robot.")
            automator = robot.FirefoxERPAutomator(robot.DEFAULT_ERP_URL, logger=self._log)
            self._log("Iniciando automatización en Firefox…")
            automator.automate_import()

            # Done
            self.root.after(0, self._on_success)

        except Exception as e:
            tb = traceback.format_exc()
            self.root.after(0, self._on_error, str(e), tb)

    def _on_success(self):
        self._set_busy(False)
        self._set_completed(True)
        self.status_var.set("COMPLETADO ✅")

    def _on_error(self, msg: str, tb: str):
        self._set_busy(False)
        self._set_completed(False)
        self.status_var.set("ERROR ❌")
        self._log("ERROR:")
        self._log(msg)
        self._log(tb)
        messagebox.showerror("Error", msg)


def main():
    root = tk.Tk()
    try:
        # nicer Windows theme if available
        root.tk.call('source', 'azure.tcl')
        root.tk.call('set_theme', 'light')
    except Exception:
        pass

    app = DonArticuloApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()