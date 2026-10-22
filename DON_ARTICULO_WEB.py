# -*- coding: utf-8 -*-
"""
DON ARTICULO WEB - Servidor web para consulta de referencias.
Accede desde cualquier equipo de la red: http://<TU-IP>:5050
"""
import os, sys, json, traceback, threading, socket, logging, re
import hashlib, hmac, time
from datetime import datetime

# Silenciar warnings de urllib3 generados por cleanup de Selenium.
# Cuando geckodriver/Firefox ya esta cerrado y algun camino del codigo
# intenta una ultima llamada HTTP (DELETE /session/...), urllib3 emite
# 3 reintentos con "WinError 10061" que ensucian la consola sin aportar
# informacion accionable. La logica de negocio no usa urllib3 directamente.
logging.getLogger('urllib3.connectionpool').setLevel(logging.ERROR)
logging.getLogger('urllib3').setLevel(logging.ERROR)

# Lock para que solo haya una importacion a la vez
IMPORT_LOCK = threading.Lock()

# Force UTF-8 on Windows console
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
if PROJECT_DIR not in sys.path:
    sys.path.insert(0, PROJECT_DIR)

from flask import Flask, render_template, request, jsonify
from werkzeug.exceptions import RequestEntityTooLarge
import ROBOT
import ref_index
import row_choice
import historico
from DON_ARTICULO import (
    find_reference_anysheet, find_reference_hit, _normalize_provider, _is_forced_ceramic,
    _safe_set, _to_float, _get_first_number_from_row, _find_col_by_norm,
    _find_col_by_norm_multi, _safe_set_norm, _norm_cell, _norm_header,
    _row_matches_reference_exact, load_provs_map, load_exceltodo_rules,
    load_provider_table, pick_discounts,
    _fetch_google_image_bytes, _analyze_image_with_gemini,
    load_overrides, check_category_override,
    _is_unambiguous_griferia, _pick_description, ONLINE_LOOKUPS,
    _HDR_UDX, _HDR_MIN, _HDR_UPA, _HDR_PP, _HDR_PESO,
)

app = Flask(__name__, template_folder=os.path.join(PROJECT_DIR, 'web_templates'),
            static_folder=os.path.join(PROJECT_DIR, 'static'), static_url_path='/static')

AYUDA_PATH = os.path.join(PROJECT_DIR, "ayuda.xlsx")
PROVIDER_TABLE = {}
PROVS_MAP = {}
EXCELTODO_RULES = {}
try:
    import conjuntos
    _AVISOS = conjuntos.GestorAvisos(PROJECT_DIR)     # aviso "referencia individual de un conjunto" (ver conjuntos.py)
except Exception as _e_avisos:
    _AVISOS = None
    print("[!] Aviso de conjuntos no disponible (%s): se importa como antes, sin aviso." % _e_avisos)

def load_config():
    global PROVIDER_TABLE, PROVS_MAP, EXCELTODO_RULES
    PROVIDER_TABLE = load_provider_table(AYUDA_PATH)

    provs_path = os.path.join(PROJECT_DIR, "PROVS.xlsx")
    exceltodo_path = os.path.join(PROJECT_DIR, "exceltodo.xlsx")
    overrides_path = os.path.join(PROJECT_DIR, "overrides.xlsx")

    PROVS_MAP = load_provs_map(provs_path) if os.path.exists(provs_path) else {}
    EXCELTODO_RULES = load_exceltodo_rules(exceltodo_path) if os.path.exists(exceltodo_path) else {}
    load_overrides(overrides_path)

def _firma(descripcion, precio):
    """Clave de lo que se importa de una fila (descripcion + precio de compra): con el mismo codigo
    tecleado y el mismo proveedor, misma firma = mismo articulo."""
    try:
        return '%s|%.6f' % (' '.join(str(descripcion).split()).upper(), float(precio))
    except (TypeError, ValueError):
        return None

def _convert_preview(provider_name, df, row_data):
    """Lo que el programa importaria de una fila: {'codigo', 'firma'}; None si el proveedor no esta
    en ayuda.xlsx o la conversion falla."""
    cfg = PROVIDER_TABLE.get(_normalize_provider(provider_name))
    if not cfg:
        return None
    try:
        values = dict(row_data)
        user_inputs = {'provider_number': cfg['provider_number'], 'pro_discount': cfg['pro_discount'],
                       'private_discount': cfg['private_discount']}
        out, _cat, _info = ROBOT.SingleRowConverter().convert_row(values, df, provider_name, user_inputs)
        return {'codigo': str(out.loc[0, 'Codigo']).strip(),
                'firma': _firma(out.loc[0, 'Descripcion'], out.loc[0, 'Precio Compra'])}
    except Exception:
        return None

# Proveedores cuyo codigo, si YA esta en el sistema en OTRO proveedor, no se sufija solo (x.1): se avisa y la
# persona escribe otra referencia, que pasa a ser el Codigo, la Ref.Prov. y el "REF: x" de la Descripcion
# (este ultimo solo si la descripcion lo lleva). Para activarlo en otro proveedor basta anadir su nombre aqui.
CODE_ALERT_PROVIDERS = {'PROVEEDOR_EJEMPLO'}
# Proveedores que ya estan ENTEROS en el ERP con los valores buenos: DON ARTICULO no importa nada de ellos para no pisar
# lo que ya hay; al teclear una de sus referencias sale el aviso "REFERENCIA YA EN EL ERP".
ERP_LOCKED_PROVIDERS = {'PROVEEDOR_004'}
# Proveedores a los que NO se aplican ni la IA de respaldo ni el override universal de griferia, p. ej. porque su
# categoria la fija un lookup determinista por codigo (ROBOT.PROVIDER_PROFILES['...']['category_lookup']).
CLASSIFIER_EXCLUDED_PROVIDERS = set()
CODE_ALERT_TITLE = 'ARTÍCULO YA METIDO CON ESA REFERENCIA EN EL SISTEMA EN OTRO PROVEEDOR, ELIJA OTRA REFERENCIA'
CODE_MAX_LEN = 15          # el codigo mas largo de las tarifas en formato ERP de DATABASE
CODE_PATTERN = re.compile(r'[A-ZÑ0-9][A-ZÑ0-9._/+-]*')     # como los codigos reales del ERP (mayusculas, sin espacios);
                                                            # la 1a letra o numero tambien evita que Excel lo lea como formula

class _ComprobacionCodigoFallida(Exception):
    """No se pudo consultar ni el historico ni el indice de codigos del ERP: no se puede saber si el codigo esta libre."""

def _proveedores_con_codigo(code, provider_name, log=None, incluir_propio=False):
    """Proveedores en los que ese codigo YA esta en el sistema: lo que importo DON ARTICULO (historico) y las
    tarifas en formato ERP de DATABASE (articulos ya cargados en el ERP). Por defecto solo los DISTINTOS de
    provider_name; con incluir_propio=True tambien el propio. Si fallan las DOS fuentes no se puede saber y
    se lanza _ComprobacionCodigoFallida (no se importa a ciegas)."""
    found = []
    fallos = 0

    def add(name):
        if name and (incluir_propio or name != provider_name) and name not in found:
            found.append(name)

    try:
        for uso in historico.usos_codigo_erp(code):
            add(uso['proveedor'])
    except Exception as e:
        fallos += 1
        if log: log("Aviso: no se pudo consultar el historico de importaciones: %s" % e)
    try:
        for prov in ref_index.erp_providers_for(ROBOT.DEFAULT_DATABASE_PATH, code):
            add(prov)
    except Exception as e:
        fallos += 1
        if log: log("Aviso: no se pudo consultar el indice de codigos del ERP: %s" % e)
    if fallos == 2:
        raise _ComprobacionCodigoFallida("fallan el historico y el indice de codigos")
    return found

def _problema_referencia_nueva(new_code, provider_name, original, log=None):
    """None si la referencia elegida (ya en mayusculas) esta libre; si no, el motivo (texto para la persona)."""
    if not new_code:
        return 'ESCRIBA UNA REFERENCIA.'
    if any(ch.isspace() for ch in new_code):
        return 'LA REFERENCIA NO PUEDE LLEVAR ESPACIOS.'
    if len(new_code) > CODE_MAX_LEN:
        return 'LA REFERENCIA ES DEMASIADO LARGA (MÁXIMO %d CARACTERES).' % CODE_MAX_LEN
    if not CODE_PATTERN.fullmatch(new_code):
        return 'SOLO LETRAS, NÚMEROS Y . _ / + - (EMPEZANDO POR UNA LETRA O UN NÚMERO).'
    if new_code == str(original).strip().upper():
        return 'ES LA MISMA REFERENCIA QUE YA ESTÁ EN USO, ELIJA OTRA.'
    en_uso = _proveedores_con_codigo(new_code, provider_name, log, incluir_propio=True)
    try:
        for prov in ref_index.providers_for(ROBOT.DEFAULT_DATABASE_PATH, new_code):   # cualquier tarifa que la traiga
            if prov not in en_uso:
                en_uso.append(prov)
    except Exception as e:
        if log: log("Aviso: no se pudo consultar el indice de referencias: %s" % e)
    if en_uso:
        return 'ESA REFERENCIA TAMBIÉN ESTÁ EN USO (%s), ELIJA OTRA.' % ', '.join(en_uso)
    return None

def _aplicar_referencia(result_df, codigo_orig, nuevo):
    """La referencia elegida pasa a ser el Codigo y la Ref.Prov. y sustituye al 'REF: x' del final de la
    Descripcion, solo si la descripcion lo lleva (hay tarifas cuya descripcion no trae REF)."""
    _safe_set(result_df, 'Codigo', nuevo)
    _safe_set(result_df, 'Ref.Prov.', nuevo)
    sufijo = ' REF: %s' % codigo_orig
    descripcion = str(result_df.loc[0, 'Descripcion'])
    if descripcion.endswith(sufijo):
        _safe_set(result_df, 'Descripcion', descripcion[:-len(sufijo)] + ' REF: %s' % nuevo)

def _pedir_otra_referencia(reference, provider_name, codigo, en_uso, row=None, error=''):
    """Respuesta 'hay que elegir' del programa: alerta con un recuadro en blanco para la nueva referencia."""
    return {
        'success': False,
        'needs_choice': True,
        'kind': 'codigo',
        'reference': reference,
        'provider': provider_name,
        'row': row,
        'code': codigo,
        'en_uso': en_uso,
        'title': CODE_ALERT_TITLE,
        'subtitle': 'La referencia %s de %s ya está en: %s' % (codigo, provider_name, ', '.join(en_uso)) if en_uso else '',
        'error': error or '',
    }

def _codigo_destino(reference, provider_name, preview):
    """Codigo con el que entraria en el ERP: el que ya tenia si es el MISMO articulo; x.1, x.2...
    si ese codigo ya lo usa OTRO articulo (el ERP no admite dos iguales); si no, el suyo."""
    if not preview or not preview.get('codigo'):
        return None
    mismo = historico.mismo_articulo(reference, provider_name, preview.get('firma'))
    if mismo:
        return mismo['codigo_erp']
    if _normalize_provider(provider_name) in CODE_ALERT_PROVIDERS:
        return preview['codigo']          # no se sufija solo: si esta en uso, la persona elige otra referencia
    if historico.uso_codigo_erp(preview['codigo']):
        return historico.siguiente_codigo_libre(preview['codigo'])
    return preview['codigo']

def _articulo_de_fila(row):
    """Descripcion de las tarifas cuya columna se llama 'Articulo', que _pick_description no conoce."""
    for key, value in (row or {}).items():
        if _norm_cell(key).replace('Í', 'I') == 'ARTICULO' and value is not None:
            texto = ' '.join(str(value).split())
            if texto and texto.lower() != 'nan':
                return texto[:160]
    return ''

def _describe_providers(reference, providers):
    """Datos del banner cuando la referencia esta en varios proveedores: proveedor, descripcion,
    precio de tarifa y fichero de cada candidato. Lo ya importado (historico) sale en gris."""
    candidates = []
    for provider in providers:
        item = {'provider': provider, 'title': provider, 'lines': []}
        try:
            hit = find_reference_hit(ROBOT.DEFAULT_DATABASE_PATH, reference, provider_filter=provider)
            if hit:
                row = hit['row_data']
                desc = _pick_description(row)[:160] or _articulo_de_fila(row)
                if desc:
                    item['lines'].append(desc)
                meta = []
                price = row_choice.price_for_display(row)
                if price is not None:
                    meta.append('Precio tarifa: ' + row_choice.money(price))
                meta.append(hit['file'])
                item['lines'].append('  ·  '.join(meta))
                previas = historico.importaciones(reference, provider)
                if previas:
                    ultima = previas[-1]
                    item['badge'] = 'YA IMPORTADO el %s como %s' % (historico.fecha_corta(ultima['fecha']), ultima['codigo_erp'])
                    item['disabled'] = len(hit['rows']) == 1     # con varias filas, al entrar se ve cual falta
                else:
                    preview = _convert_preview(provider, hit['df'], row)
                    destino = _codigo_destino(reference, provider, preview)
                    if destino and destino != preview['codigo']:
                        item['note'] = 'SE IMPORTARÁ COMO %s (el %s ya está en el ERP)' % (destino, preview['codigo'])
        except Exception as e:
            item['lines'].append('(no se pudo leer la tarifa: %s)' % e)
        candidates.append(item)
    return candidates

def _provider_titles(options):
    """Titulo y subtitulo del banner de proveedores segun el historico."""
    hechas = [o for o in options if o.get('badge')]
    if not hechas:
        return 'HEMOS ENCONTRADO TU ARTÍCULO EN VARIOS PROVEEDORES', 'CONCRETA CUÁL ES EL TUYO'
    titulo = 'ESTE CÓDIGO YA SE IMPORTÓ: ' + '; '.join(
        '%s %s' % (o['provider'], o['badge'].replace('YA IMPORTADO ', '')) for o in hechas)
    return titulo, row_choice.subtitle_for(options, 'CONCRETA CUÁL ES EL TUYO')

def process_reference(reference, logger_cb=None, provider_choice=None, row_choice_idx=None, new_code=None, conjunto_ok=False):
    logs = []
    new_code = str(new_code or '').strip().upper()       # los codigos del ERP van en mayusculas
    def log(msg):
        msg_str = str(msg)
        logs.append(msg_str)
        if logger_cb:
            logger_cb(msg_str)
        try:
            print(f"  [LOG] {msg_str}")
        except UnicodeEncodeError:
            print(f"  [LOG] {msg_str.encode('ascii', 'replace').decode('ascii')}")

    def fail(error):
        # Fallo de negocio: queda en consola y en el log vivo (el front pinta "Error" en rojo)
        log("Error: %s" % error)
        return {'success': False, 'error': error, 'logs': logs}

    try:
        log("Buscando referencia: %s" % reference)
        # 1) Codigo en VARIOS proveedores: no se adivina, la persona elige (banner del front).
        #    Con un solo proveedor se sigue como siempre (sin preguntar).
        provider_filter = None
        try:
            candidates = ref_index.providers_for(ROBOT.DEFAULT_DATABASE_PATH, reference, logger=log)
        except Exception as e:
            candidates = []
            log("Aviso: no se pudo comprobar si la referencia esta en varios proveedores: %s" % e)
        if provider_choice:
            if provider_choice not in candidates:
                return fail("El proveedor '%s' no tiene la referencia '%s'." % (provider_choice, reference))
            provider_filter = provider_choice
            log("Proveedor elegido por el usuario: %s" % provider_choice)
        elif len(candidates) > 1:
            log("La referencia '%s' esta en %d proveedores (%s): hay que elegir cual es." % (reference, len(candidates), ', '.join(candidates)))
            options = _describe_providers(reference, candidates)
            title, subtitle = _provider_titles(options)
            return {
                'success': False,
                'needs_choice': True,
                'kind': 'provider',
                'reference': reference,
                'title': title,
                'subtitle': subtitle,
                'candidates': options,
                'logs': logs,
            }
        elif len(candidates) == 1:
            provider_filter = candidates[0]

        hit = find_reference_hit(ROBOT.DEFAULT_DATABASE_PATH, reference, logger=log,
                                 provider_filter=provider_filter, row_index=row_choice_idx)
        found = None
        if hit:
            found = (hit['provider_name'], hit['path'], hit['row_data'], hit['df'])
        elif row_choice_idx is not None:
            return fail("La fila elegida ya no corresponde a la referencia '%s' (¿ha cambiado la tarifa?)." % reference)

        if not found:
            scanner = ROBOT.ProviderDatabaseScanner(ROBOT.DEFAULT_DATABASE_PATH, logger=log)
            found = scanner.find_reference(reference)
            if found:
                _p, _tf, _rd, _df = found
                if not _row_matches_reference_exact(_rd, reference):
                    log("Coincidencia parcial ignorada.")
                    found = None

        if not found:
            return fail("No se encontro la referencia '%s' en DATABASE." % reference)

        provider_name, tariff_file, row_data, tariff_df = found
        log("Encontrado en proveedor: %s" % provider_name)
        prov_key = _normalize_provider(provider_name)

        # Proveedor ya entero en el ERP: no se convierte ni se importa nada; el front saca el aviso rojo.
        if prov_key in ERP_LOCKED_PROVIDERS:
            log("⚠ REFERENCIA YA EN EL ERP: %s es de %s, que ya esta entero en el ERP. No se importa nada." % (reference, provider_name))
            return {
                'success': False,
                'ya_en_erp_blocked': True,
                'reference': reference,
                'provider': provider_name,
                'error': 'REFERENCIA YA EN EL ERP: %s es de %s; no se importa nada.' % (reference, provider_name),
                'logs': logs,
            }

        # Referencia INDIVIDUAL que forma parte de un conjunto -> se avisa ANTES de importar.
        # ACEPTAR (conjunto_ok) = se importa solo esa, como siempre; NO = la persona elige el conjunto. Solo para los
        # proveedores y las referencias de las tablas de conjuntos (conjuntos.py); el resto no pasa por aqui.
        if not conjunto_ok and _AVISOS is not None:
            try:
                _aviso = _AVISOS.aviso(provider_name, reference)
            except Exception as e:                  # el aviso es una ayuda: si su tabla falla no se bloquea ninguna importacion
                _aviso = None
                log("Aviso: no se pudo comprobar si %s forma parte de un conjunto: %s" % (reference, e))
            if _aviso:
                log("⚠ %s: %s" % (reference, _aviso.mensaje))
                return {
                    'success': False,
                    'needs_choice': True,
                    'kind': 'conjunto',
                    'reference': reference,
                    'provider': provider_filter,
                    'row': row_choice_idx,
                    'title': _aviso.titulo,
                    'message': _aviso.mensaje,
                    'accept': _aviso.boton_aceptar,
                    'conjunto_btn': _aviso.boton_conjunto,
                    'conjuntos': list(_aviso.conjuntos),
                    'tambien_suelta': _aviso.tambien_suelta,
                    'logs': logs,
                }

        if prov_key not in PROVIDER_TABLE:
            return fail("Proveedor '%s' (clave '%s') no esta en ayuda.xlsx. Anadelo (PROVEEDOR, DESCUENTOS, NUMERO PROVEEDOR, D.PROFESIONAL, D.PARTICULAR) y REINICIA el servidor (INICIAR_WEB.bat): la tabla solo se carga al arrancar." % (provider_name, prov_key))

        cfg = PROVIDER_TABLE[prov_key]
        provider_number = cfg['provider_number']
        pro_discount = cfg['pro_discount']
        private_discount = cfg['private_discount']
        dto1 = cfg['dto1']
        dto2 = cfg['dto2']

        log("N proveedor: %s | Dto pro: %s%% | Dto part: %s%%" % (provider_number, pro_discount, private_discount))

        converter = ROBOT.SingleRowConverter()
        user_inputs = {
            'provider_number': provider_number,
            'pro_discount': pro_discount,
            'private_discount': private_discount,
        }
        # 2) Codigo repetido en la MISMA hoja del proveedor: si las filas importarian lo mismo NO se
        #    pregunta; si importarian algo distinto (formato, color, acabado, precio...) se pregunta cual
        #    es. Lo ya importado (historico) sale en gris. Hojas/ficheros distintos de la misma carpeta
        #    NO preguntan: gana la primera tarifa, como siempre.
        if hit is not None and row_choice_idx is None:
            _previews = {}

            def _preview(idx, row):
                if idx not in _previews:
                    _previews[idx] = _convert_preview(provider_name, tariff_df, row)
                return _previews[idx]

            def _signature(idx, row):
                p = _preview(idx, row)
                return p['firma'] if p else None

            def _annotate(idx, row):
                p = _preview(idx, row)
                if not p:
                    return None
                previa = historico.mismo_articulo(reference, provider_name, p['firma'])
                if previa:
                    return {'disabled': True, 'badge': 'YA IMPORTADO el %s como %s' % (historico.fecha_corta(previa['fecha']), previa['codigo_erp'])}
                destino = _codigo_destino(reference, provider_name, p)
                if destino and destino != p['codigo']:
                    return {'note': 'SE IMPORTARÁ COMO %s (el %s ya está en el ERP)' % (destino, p['codigo'])}
                return None

            choice = row_choice.row_choice_payload(hit, reference, signature=_signature, annotate=_annotate)
            if choice and choice.get('error'):
                return fail(choice['error'])
            if choice:
                log("La referencia '%s' sale %d veces distintas en la tarifa de %s: hay que elegir cual es." % (reference, len(choice['candidates']), hit['provider_name']))
                choice['logs'] = logs
                return choice
        result_df, category, cat_info = converter.convert_row(row_data, tariff_df, provider_name, user_inputs)
        firma = _firma(result_df.loc[0, 'Descripcion'], result_df.loc[0, 'Precio Compra'])
        log("Categoria detectada: %s" % category)

        if _is_forced_ceramic(provider_name):
            if str(category).upper()[:3] != "CER":
                log("Forzando categoria CERAMICA por proveedor: %s" % provider_name)
            category = "CERAMICA"

        # Phase 6: Check overrides.xlsx
        ref_val = str(row_data.get('Referencia', row_data.get('referencia', reference)))
        desc_for_override = str(row_data.get('Descripcion', row_data.get('Descripción', '')))
        cat_override = check_category_override(provider_name, ref_val, desc_for_override)
        if cat_override:
            log("Override manual: %s -> %s" % (category, cat_override))
            category = cat_override
            cat_info = ROBOT.CATEGORIES.get(category, ROBOT.CATEGORIES.get('VARIOS', {}))

        # Phase 7: Classification fallback
        confidence = converter.classifier.get_confidence()
        prov_upper = _normalize_provider(provider_name)

        # Phase 7b: Generic AI fallback (clasificacion dudosa o VARIOS), salvo proveedores excluidos
        if ONLINE_LOOKUPS and prov_upper not in CLASSIFIER_EXCLUDED_PROVIDERS and (str(category).upper() == "VARIOS" or confidence < 30):
            log("Confianza baja (%s) o VARIOS - intentando IA..." % confidence)
            try:
                desc = str(row_data.get('Descripcion', row_data.get('Descripcion', reference)))
                query = "%s %s %s" % (provider_name, reference, desc)
                img_bytes = _fetch_google_image_bytes(query)
                if img_bytes:
                    ai_cat = _analyze_image_with_gemini(img_bytes, desc)
                    if ai_cat and ai_cat != "VARIOS":
                        log("IA sugiere: %s" % ai_cat)
                        category = ai_cat
                        cat_info = ROBOT.CATEGORIES.get(category, ROBOT.CATEGORIES.get('VARIOS', {}))
            except Exception as e:
                log("Error IA: %s" % str(e))

        # Override universal: terminos inequivocos de griferia siempre fuerzan GRIFERIA (salvo proveedores excluidos)
        _griferia_desc = _pick_description(row_data)
        if prov_upper not in CLASSIFIER_EXCLUDED_PROVIDERS and _is_unambiguous_griferia(_griferia_desc):
            if str(category).upper().replace('Í', 'I') != 'GRIFERIA':
                log("Override GRIFERIA por termino inequivoco en '%s' (antes: %s)" % (_griferia_desc[:60], category))
                category = "GRIFERÍA"
                cat_info = ROBOT.CATEGORIES.get(category, ROBOT.CATEGORIES.get('VARIOS', {}))

        adj_pro, adj_part, reason = pick_discounts(
            provider_name, category, pro_discount, private_discount, PROVS_MAP, EXCELTODO_RULES)
        if adj_pro != pro_discount or adj_part != private_discount:
            log("Ajuste descuentos: %s%%/%s%% -> %s%%/%s%% (%s)" % (pro_discount, private_discount, adj_pro, adj_part, reason))
            pro_discount, private_discount = adj_pro, adj_part
            base_val = None
            for c in ("Tarifa 1", "Precio Base Tarifa"):
                if c in result_df.columns:
                    try: base_val = float(result_df.loc[0, c]); break
                    except: pass
            if base_val is not None:
                if "Tarifa 2" in result_df.columns:
                    result_df.loc[0, "Tarifa 2"] = round(base_val * (1 - pro_discount/100.0), 6)
                if "Tarifa 3" in result_df.columns:
                    result_df.loc[0, "Tarifa 3"] = round(base_val * (1 - private_discount/100.0), 6)

        # Dtos from row
        def _pick_dto(row, *names):
            for n in names:
                nn = _norm_cell(n)
                for kk, vv in (row or {}).items():
                    if _norm_cell(kk) == nn and vv is not None and str(vv).strip():
                        v = str(vv).strip()
                        if v.lower() != 'nan': return v
            return None

        dto1_row = _pick_dto(row_data, "%Dto1", "Dto1", "Descuento 1")
        dto2_row = _pick_dto(row_data, "%Dto2", "Dto2", "Descuento 2")
        dto3_row = _pick_dto(row_data, "%Dto3", "Dto3", "Descuento 3")
        dto4_row = _pick_dto(row_data, "%Dto4", "Dto4", "Descuento 4")
        if dto1_row: dto1 = dto1_row
        if dto2_row: dto2 = dto2_row
        dto3 = dto3_row or ''
        dto4 = dto4_row or ''
        # Un perfil con %Dto1 fijo (ROBOT.PROVIDER_PROFILES) manda sobre el de ayuda.xlsx y el de la fila
        _perfil = ROBOT.provider_profile(provider_name)
        if _perfil.get('dto1') is not None:
            dto1 = _perfil['dto1']
        if dto1: _safe_set(result_df, '%Dto1', dto1)
        if dto2: _safe_set(result_df, '%Dto2', dto2)
        if dto3: _safe_set(result_df, '%Dto3', dto3)
        if dto4: _safe_set(result_df, '%Dto4', dto4)

        # =========================================================
        # CORRECCIÓN DE UNIDAD DE MEDIDA PARA CERÁMICAS Y NO-CERÁMICAS
        # =========================================================
        if str(category).upper().startswith('CER'):
            um = None

            # Detectar si es pieza especial (rodapié, peldaño...) para mapear CAJA -> PZ
            desc_val = ""
            for k in ("Descripcion Corta", "Descripción Corta", "Descripcion", "Descripción"):
                col_d = _find_col_by_norm(result_df, [k])
                if col_d:
                    desc_val = str(result_df.loc[0, col_d] or "").upper()
                    break
            piezas_kws = ["RODAPIE", "RODAPIÉ", "PELDAÑO", "PELDANO", "TACO", "TABICA",
                          "ZOCALO", "ZÓCALO", "CENEFA", "DECORADO", "PIEZA", "MALLA",
                          "MOSAICO", "LISTELO"]
            is_pieza_especial = any(kw in desc_val for kw in piezas_kws)

            # Buscar columna de unidad de medida en los datos originales del proveedor
            for kk, vv in (row_data or {}).items():
                kn = _norm_header(kk)
                if "UD MEDIDA" in kn or kn in ("UM", "U M", "UNIDAD MEDIDA", "UD VENTA",
                                                "UNIDAD VENTA", "UNIDAD"):
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
                            um = sv.replace("M²", "M2").replace("PCS", "PZ").replace("PIEZAS", "PZ").replace("PIEZA", "PZ").replace("ML", "M")
                        break

            # Si no se encontró, intentar deducir de la descripción
            if not um:
                desc_try = ""
                for k in ("Descripcion Corta", "Descripción Corta", "Descripción", "Descripcion"):
                    col_d2 = _find_col_by_norm(result_df, [k])
                    if col_d2:
                        desc_try = str(result_df.loc[0, col_d2] or "")
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

            # Proveedores con ficha de logistica (ROBOT._CERAMIC_LOGISTICS_SPEC): la unidad la traduce SU ficha (p. ej. CA/PI/UN -> PZ).
            _um_ficha = ROBOT._ceramic_unit(row_data, provider_name)
            if _um_ficha:
                um = _um_ficha

            # Escribir en ambas columnas: Ud.Medida (D) y Ud.Med.Compra (AE)
            col_um = _find_col_by_norm_multi(result_df, ["Ud.Medida", "UD. MEDIDA", "UD MEDIDA"])
            if col_um:
                _safe_set(result_df, col_um, um)
            col_umc = _find_col_by_norm_multi(result_df, ["Ud. Med. Compra", "UD. MED. COMPRA", "UD MED COMPRA", "UD MEDIDA COMPRA"])
            if col_umc:
                _safe_set(result_df, col_umc, um)
            log("Ud. Medida cerámica: %s" % um)
        else:
            # No-cerámica: forzar UD en ambas columnas
            um = "UD"
            col_um = _find_col_by_norm_multi(result_df, ["Ud.Medida", "UD. MEDIDA", "UD MEDIDA"])
            if col_um:
                _safe_set(result_df, col_um, um)
            col_umc = _find_col_by_norm_multi(result_df, ["Ud. Med. Compra", "UD. MED. COMPRA", "UD MED COMPRA", "UD MEDIDA COMPRA"])
            if col_umc:
                _safe_set(result_df, col_umc, um)

        # =========================================================
        # LOGICA LOGISTICA (UdxEnvase, Unid.Minima, Peso, Peso Palet, etc)
        # =========================================================
        if str(category).upper().startswith('CER'):
            # Ceramica: leer datos logisticos del Excel del proveedor
            from DON_ARTICULO import _get_first_number_from_row as _gfn
            v_udx = _gfn(row_data, _HDR_UDX)
            v_min = _gfn(row_data, _HDR_MIN)
            v_upa = _gfn(row_data, _HDR_UPA)
            v_pp  = _gfn(row_data, _HDR_PP)
            v_peso = _gfn(row_data, _HDR_PESO)

            # Fallback: buscar cualquier columna con PESO (excluyendo palet/pz/m2)
            if v_peso is None:
                pat_peso = ['PESO/CAJA','KG/CAJA','PESO CAJA','PESO KG','PESO']
                exclusion_patterns = ['PALET','PL','PZ','PIEZA','M2']
                for kk, vv in (row_data or {}).items():
                    kn = _norm_cell(kk)
                    if any(p in kn for p in pat_peso) and not any(e in kn for e in exclusion_patterns):
                        val = _to_float(vv)
                        if val is not None:
                            v_peso = val
                            break

            # Calcular peso palet si falta: Peso Caja * Cajas Palet
            if v_pp is None and v_peso is not None and v_upa is not None:
                v_pp = round(v_peso * v_upa, 2)

            # Calcular peso caja si falta pero tenemos peso palet y cajas/palet
            if v_peso is None and v_pp is not None and v_upa is not None and v_upa > 0:
                v_peso = round(v_pp / v_upa, 2)

            um_col = _find_col_by_norm(result_df, ['Ud.Medida','UD MEDIDA'])
            um_val = str(result_df.loc[0, um_col] or '').upper().strip() if um_col else ''
            if um_val in ('PZ','PZA','PCS','PIEZA','PIEZAS','PI'):
                if v_udx is not None: v_min = v_udx
                elif v_min is None: v_min = 1
            v_mcom = v_min
            v_mven = v_min
            if v_udx is not None: _safe_set_norm(result_df, ['UdxEnvase','Udx Envase','UDXENVASE'], v_udx)
            if v_min is not None: _safe_set_norm(result_df, ['Unid. Mínima','Unid.Minima','Unid Minima'], v_min)
            if v_upa is not None: _safe_set_norm(result_df, ['Ud.Palet','Ud Palet'], v_upa)
            if v_peso is not None: _safe_set_norm(result_df, ['Peso'], v_peso)
            if v_pp is not None: _safe_set_norm(result_df, ['Peso Palet'], v_pp)
            if v_mcom is not None: _safe_set_norm(result_df, ['Multiplo Compra','Multiplo Com'], v_mcom)
            if v_mven is not None: _safe_set_norm(result_df, ['Multiplo Venta','Multiplo Ven'], v_mven)

            # Proveedores de ceramica con ficha propia: su tarifa trae caja/palet, kgs/caja y kgs/palet con nombres que el
            # lector generico no conoce, y su pieza especial puede venderse suelta (no por caja): logistica propia segun la unidad.
            if ROBOT._ceramic_logistics_spec(provider_name):
                # Resto de proveedores de ceramica con ficha propia (ROBOT._CERAMIC_LOGISTICS_SPEC): Peso por unidad de
                # venta, minimo = multiplo = unidades por caja. El Peso generico (kg de la caja) NUNCA se deja: sin dato,
                # en blanco. Lo demas solo se sobrescribe si hay dato (si no, queda lo que dejo el bloque generico).
                lg = ROBOT._ceramic_logistics(row_data, um, provider_name)
                if lg is not None:
                    _safe_set_norm(result_df, ['Peso'], lg['peso'])
                    for destino, campo in ((['UdxEnvase','Udx Envase','UDXENVASE'], 'udx_envase'),
                                           (['Unid. Mínima','Unid.Minima','Unid Minima'], 'unid_minima'),
                                           (['Multiplo Compra','Multiplo Com'], 'multiplo'),
                                           (['Multiplo Venta','Multiplo Ven'], 'multiplo'),
                                           (['Ud.Palet','Ud Palet'], 'ud_palet'),
                                           (['Peso Palet'], 'peso_palet')):
                        if lg[campo] is not None:
                            _safe_set_norm(result_df, destino, lg[campo])
        else:
            # No ceramica: forzar todo a 1
            _safe_set_norm(result_df, ['UdxEnvase','Udx Envase','UDXENVASE'], 1)
            _safe_set_norm(result_df, ['Unid. Mínima','Unid.Minima','Unid Minima'], 1)
            _safe_set_norm(result_df, ['Multiplo Compra','Multiplo Com'], 1)
            _safe_set_norm(result_df, ['Multiplo Venta','Multiplo Ven'], 1)

        # =========================================================
        # CONTROL DE PRECIO 0 / SIN PRECIO / "A CONSULTAR" / NO NUMERICO
        # Aplica a TODOS los proveedores de DATABASE. Cualquier valor no numerico de PRECIO
        # ("A CONSULTAR", simbolos raros, vacio) cae a 0 al convertir, por lo que
        # precio<=0 cubre todos los casos. Si NO esta confirmado, se DETIENE aqui
        # -antes de guardar el importart y antes de tocar el ERP- y se pide
        # confirmacion al usuario (popup rojo en el frontend).
        # =========================================================
        _precio_num = _to_float(result_df.loc[0, 'Precio Compra']) if 'Precio Compra' in result_df.columns else None
        zero_price = (_precio_num is None) or (_precio_num <= 0)
        if zero_price:
            log('PRECIO 0 / sin precio: IMPORTACION CANCELADA (no se ha tocado el ERP). Meter manualmente.')
            return {
                'success': False,
                'zero_price_blocked': True,
                'reference': reference,
                'provider': provider_name,
                'category': category,
                'message': 'AVISO: PRECIO DE "0€" O SIN ESPECIFICAR EN TARIFA, CONSULTAR CON FABRICANTE, METER MANUALMENTE',
                'logs': logs,
            }

        # =========================================================
        # HISTORICO: el ERP no admite dos articulos con el mismo codigo. Si este codigo ya entro en
        # el ERP con OTRO articulo, este entra como x.1 (o x.2...); si es el MISMO articulo, entra con
        # el codigo que ya tenia. La Ref.Prov. se queda con el codigo real del proveedor (salvo en los proveedores
        # con alerta de codigo, CODE_ALERT_PROVIDERS: ahi Codigo, Ref.Prov. y REF: llevan siempre el mismo codigo).
        # =========================================================
        codigo_erp = str(result_df.loc[0, 'Codigo']).strip()
        if prov_key in CODE_ALERT_PROVIDERS:
            # Si su codigo ya esta en el sistema en OTRO proveedor NO se sufija solo: se avisa y la persona escribe
            # otra referencia, que pasa a ser Codigo, Ref.Prov. y el "REF: x" de la Descripcion. Si lo usa otro
            # articulo del MISMO proveedor, sufijo x.1 como en todos los demas.
            codigo_orig = codigo_erp
            try:
                mismo = historico.mismo_articulo(reference, provider_name, firma)
            except Exception as e:
                mismo = None
                log("Aviso: no se pudo consultar el historico de importaciones: %s" % e)
            try:
                if mismo:
                    codigo_erp = mismo['codigo_erp']
                    log("Nota: este articulo ya se importo el %s como %s; se vuelve a importar con ese codigo." % (historico.fecha_corta(mismo['fecha']), codigo_erp))
                elif new_code:
                    problema = _problema_referencia_nueva(new_code, provider_name, codigo_orig, log)
                    if problema:
                        log("Referencia elegida '%s' rechazada: %s" % (new_code[:40], problema))
                        pregunta = _pedir_otra_referencia(reference, provider_name, codigo_orig,
                                                          _proveedores_con_codigo(codigo_orig, provider_name, log),
                                                          row_choice_idx, problema)
                        pregunta['logs'] = logs
                        return pregunta
                    codigo_erp = new_code
                    log("Referencia elegida por la persona: %s (en vez de %s)." % (codigo_erp, codigo_orig))
                else:
                    en_uso = _proveedores_con_codigo(codigo_orig, provider_name, log)
                    if en_uso:
                        log("El codigo %s ya esta en el sistema en otro proveedor (%s): hay que elegir otra referencia." % (codigo_orig, ', '.join(en_uso)))
                        pregunta = _pedir_otra_referencia(reference, provider_name, codigo_orig, en_uso, row_choice_idx)
                        pregunta['logs'] = logs
                        return pregunta
                    try:
                        previo = historico.uso_codigo_erp(codigo_erp)
                    except Exception as e:
                        previo = None
                        log("Aviso: no se pudo consultar el historico de importaciones: %s" % e)
                    if previo:
                        codigo_erp = historico.siguiente_codigo_libre(codigo_erp)
                        log("El codigo %s ya esta en el ERP (%s, importado el %s): este articulo entra como %s." % (codigo_orig, previo['proveedor'], historico.fecha_corta(previo['fecha']), codigo_erp))
            except _ComprobacionCodigoFallida as e:
                return fail("No se pudo comprobar si el codigo %s ya esta en el sistema (%s). No se ha importado nada." % (codigo_orig, e))
            if codigo_erp != codigo_orig:
                _aplicar_referencia(result_df, codigo_orig, codigo_erp)
        else:
            try:
                mismo = historico.mismo_articulo(reference, provider_name, firma)
                if mismo:
                    if mismo['codigo_erp'] != codigo_erp:
                        codigo_erp = mismo['codigo_erp']
                        _safe_set(result_df, 'Codigo', codigo_erp)
                    log("Nota: este articulo ya se importo el %s como %s; se vuelve a importar con ese codigo." % (historico.fecha_corta(mismo['fecha']), codigo_erp))
                else:
                    previo = historico.uso_codigo_erp(codigo_erp)
                    if previo:
                        nuevo = historico.siguiente_codigo_libre(codigo_erp)
                        log("El codigo %s ya esta en el ERP (%s, importado el %s): este articulo entra como %s." % (codigo_erp, previo['proveedor'], historico.fecha_corta(previo['fecha']), nuevo))
                        codigo_erp = nuevo
                        _safe_set(result_df, 'Codigo', codigo_erp)
            except Exception as e:
                log("Aviso: no se pudo consultar el historico de importaciones: %s" % e)

        # =========================================================
        # GUARDAR importart.xlsx EN CARPETA DE RED
        # =========================================================
        log('Guardando importart.xlsx...')
        fm = ROBOT.NetworkFileManager(ROBOT.DEFAULT_NETWORK_PATH, filename='importart')
        file_path = fm.save_import_file(result_df)
        log('Archivo generado: %s' % file_path)

        # =========================================================
        # AUTOMATIZAR FIREFOX (importar al ERP)
        # =========================================================
        if ROBOT.HAS_SELENIUM:
            log('Iniciando automatizacion Firefox...')
            automator = ROBOT.FirefoxERPAutomator(ROBOT.DEFAULT_ERP_URL, logger=log)
            try:
                automator.automate_import()
            except Exception as e:
                return fail("importart.xlsx generado (%s) pero la importacion al ERP FALLO: %s. El articulo NO esta importado." % (file_path, e))
            log('Importacion completada en el ERP.')
            try:
                historico.registrar(reference, codigo_erp, provider_name, provider_number,
                                    os.path.basename(str(tariff_file)), hit['sheet'] if hit else None,
                                    int(hit['row_index']) + 2 if hit else None,
                                    str(result_df.loc[0, 'Descripcion']), _to_float(result_df.loc[0, 'Precio Compra']), firma)
                log('Guardado en el historico: %s -> %s (%s).' % (reference, codigo_erp, provider_name))
            except Exception as e:
                log('Aviso: no se pudo guardar en el historico: %s' % e)
        else:
            return fail("Selenium no disponible: importart.xlsx generado (%s) pero NO importado al ERP." % file_path)

        def _sv(col):
            try:
                v = result_df.loc[0, col]
                if v is None or str(v) == 'nan' or str(v) == '': return ''
                return v
            except: return ''

        result = {
            'success': True,
            'provider': provider_name,
            'provider_number': provider_number,
            'reference': reference,
            'codigo_erp': codigo_erp,
            'description': str(_sv('Descripcion')),
            'category': category,
            'unit': str(_sv('Ud.Medida')),
            'precio_compra': str(_sv('Precio Compra')),
            'precio_base': str(_sv('Precio Base Tarifa')),
            'tarifa_1': str(_sv('Tarifa 1')),
            'tarifa_2': str(_sv('Tarifa 2')),
            'tarifa_3': str(_sv('Tarifa 3')),
            'dto1': str(_sv('%Dto1')), 'dto2': str(dto2 or ''),
            'dto3': str(dto3 or ''), 'dto4': str(dto4 or ''),
            'familia': str(_sv('Familia')),
            'subfamilia': str(_sv('Subfamilia')),
            'seccion': str(_sv('Sección')),
            'ean': str(_sv('Ean')),
            'calculo_base': str(_sv('Calculo Base')),
            'imported': True,
            'logs': logs,
            'timestamp': datetime.now().strftime('%H:%M:%S')
        }
        log('OK - Importado: %s | PVP: %s' % (category, result['tarifa_1']))
        return result
    except Exception as e:
        log(traceback.format_exc())
        return fail("Error inesperado: %s" % e)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/procesar', methods=['POST'])
def procesar():
    try:
        data = request.get_json()
        ref = (data.get('reference', '') or '').strip()
        if not ref:
            return jsonify({'success': False, 'error': 'Falta la referencia.', 'logs': []})
        # Solo una importacion a la vez (Firefox no puede tener dos sesiones)
        acquired = IMPORT_LOCK.acquire(timeout=120)
        if not acquired:
            return jsonify({'success': False, 'error': 'Hay otra importacion en curso. Espera un momento.', 'logs': []})
        try:
            _row = str(data.get('row', '')).strip()
            result = process_reference(ref, provider_choice=(data.get('provider') or '').strip() or None,
                                       row_choice_idx=int(_row) if _row.isdigit() else None,
                                       new_code=str(data.get('new_code') or '').strip() or None,
                                       conjunto_ok=(str(data.get('conjunto') or '').strip() == 'ok'))
        finally:
            IMPORT_LOCK.release()
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print('ERROR EN /procesar: ' + tb)
        result = {'success': False, 'error': str(e), 'logs': [tb]}
    return jsonify(result)

@app.route('/procesar_stream')
def procesar_stream():
    from flask import Response
    import queue
    ref = request.args.get('ref', '').strip()
    provider = request.args.get('provider', '').strip()
    row_arg = request.args.get('row', '').strip()
    row_idx = int(row_arg) if row_arg.isdigit() else None
    new_code = request.args.get('newcode', '').strip()
    conjunto_ok = request.args.get('conjunto', '').strip() == 'ok'     # la persona acepto introducir solo la pieza individual
    if not ref:
        return jsonify({'success': False, 'error': 'Falta la referencia.'})

    def generate():
        q = queue.Queue()
        def log_cb(msg):
            q.put({'type': 'log', 'msg': msg})
        
        def worker():
            acquired = IMPORT_LOCK.acquire(timeout=120)
            if not acquired:
                q.put({'type': 'error', 'msg': 'Hay otra importacion en curso. Espera un momento.'})
                q.put(None)
                return
            try:
                result = process_reference(ref, logger_cb=log_cb, provider_choice=provider or None, row_choice_idx=row_idx,
                                           new_code=new_code or None, conjunto_ok=conjunto_ok)
                if result.get('needs_choice'):
                    result.setdefault('conjunto_ok', conjunto_ok)       # lo aceptado se conserva si hay que elegir proveedor/fila/referencia despues
                    q.put({'type': 'choose', 'payload': result})
                elif result.get('success') or result.get('zero_price_blocked') or result.get('ya_en_erp_blocked'):
                    q.put({'type': 'result', 'payload': result})
                else:
                    q.put({'type': 'error', 'msg': result.get('error') or 'Fallo sin detalle (ver el log)'})
            except Exception as e:
                q.put({'type': 'error', 'msg': str(e)})
            finally:
                IMPORT_LOCK.release()
                q.put(None)

        threading.Thread(target=worker, daemon=True).start()

        while True:
            item = q.get()
            if item is None:
                break
            # Formato SSE
            yield f"data: {json.dumps(item)}\n\n"

    return Response(generate(), mimetype='text/event-stream')

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'providers': len(PROVIDER_TABLE)})

# ======================================================================
# REINICIAR REFERENCIA (pestaña inferior de la web).
# Olvida lo que el programa tiene apuntado de una referencia (historico.py) para que vuelva a entrar con su codigo
# original: sirve cuando el articulo se borro del ERP pero el programa sigue creyendo que ya esta metido. El ERP no se
# toca y lo borrado se archiva en la tabla 'reinicios'. La contraseña se comprueba SIEMPRE EN EL SERVIDOR, en cada
# peticion; no esta en claro en el codigo: se guarda como hash PBKDF2-SHA256 y se lee de variables de entorno
# (DONARTICULO_RESET_SALT y DONARTICULO_RESET_HASH, en hexadecimal). Por defecto la contraseña es la de ejemplo "y".
# Tras REINICIO_MAX_FALLOS fallos seguidos desde un mismo equipo se bloquea REINICIO_BLOQUEO_S segundos. Es un freno
# contra pulsaciones por error, no una proteccion fuerte: usa una contraseña larga y no expongas el servidor a internet.
# ======================================================================
REINICIO_SAL = bytes.fromhex(os.environ.get('DONARTICULO_RESET_SALT', '78'))      # por defecto, un placeholder (b'x')
REINICIO_VUELTAS = 100000
_REINICIO_HASH = (os.environ.get('DONARTICULO_RESET_HASH')
                  or hashlib.pbkdf2_hmac('sha256', b'y', REINICIO_SAL, REINICIO_VUELTAS).hex())
REINICIO_MAX_FALLOS = 5
REINICIO_BLOQUEO_S = 60
REFERENCIA_MAX = 40
REINICIO_MAX_CUERPO = 4096        # bytes: las peticiones de esta pestaña son diminutas; mas es un abuso
_REINICIO_FALLOS = {}            # ip -> (fallos seguidos, bloqueado hasta)
_REINICIO_LOCK = threading.Lock()


def _reinicio_clave_ok(clave):
    if not isinstance(clave, str) or not clave or len(clave) > 64:
        return False
    try:
        octetos = clave.encode('utf-8')
    except UnicodeEncodeError:           # sustituto Unicode suelto ("\ud800"): no es una contraseña, y no debe dar un 500
        return False
    intento = hashlib.pbkdf2_hmac('sha256', octetos, REINICIO_SAL, REINICIO_VUELTAS).hex()
    return hmac.compare_digest(intento, _REINICIO_HASH)


def _reinicio_json(datos, estado=200):
    resp = jsonify(datos)
    resp.status_code = estado
    resp.headers['Cache-Control'] = 'no-store'
    return resp


def _reinicio_autorizar():
    """(datos, None) si la contraseña es correcta; (None, respuesta) con 401 (mala) o 429 (bloqueado por intentos)."""
    if (request.content_length or 0) > REINICIO_MAX_CUERPO:
        return None, _reinicio_json({'ok': False, 'error': 'Petición demasiado grande.'}, 413)
    request.max_content_length = REINICIO_MAX_CUERPO        # tambien limita los cuerpos sin Content-Length (chunked)
    try:
        datos = request.get_json(silent=True)
    except RequestEntityTooLarge:
        return None, _reinicio_json({'ok': False, 'error': 'Petición demasiado grande.'}, 413)
    datos = datos if isinstance(datos, dict) else {}
    ip = request.remote_addr or '?'
    ahora = time.time()
    with _REINICIO_LOCK:
        fallos, hasta = _REINICIO_FALLOS.get(ip, (0, 0.0))
        if hasta > ahora:
            espera = int(hasta - ahora) + 1
            return None, _reinicio_json({'ok': False, 'error': 'Demasiados intentos. Espera %d segundos.' % espera, 'espera': espera}, 429)
        # El intento se CUENTA antes de comprobar la clave y en el mismo cerrojo: con varias conexiones a la vez no se
        # cuela ninguna (contar despues dejaba pasar cientos de intentos mientras el PBKDF2 tardaba).
        fallos += 1
        _REINICIO_FALLOS[ip] = (0, ahora + REINICIO_BLOQUEO_S) if fallos >= REINICIO_MAX_FALLOS else (fallos, 0.0)
    if not _reinicio_clave_ok(datos.get('password')):
        return None, _reinicio_json({'ok': False, 'error': 'Contraseña incorrecta.'}, 401)
    with _REINICIO_LOCK:
        _REINICIO_FALLOS.pop(ip, None)
    return datos, None


def _reinicio_referencia(datos):
    """(referencia normalizada, None) o (None, motivo del rechazo)."""
    ref = datos.get('reference')
    if not isinstance(ref, str) or not ref.strip():
        return None, 'Escribe una referencia.'
    if any(ord(c) < 32 or ord(c) == 127 for c in ref):
        return None, 'La referencia lleva caracteres no válidos.'
    try:
        ref.encode('utf-8')
    except UnicodeEncodeError:           # sustituto Unicode suelto: no se puede guardar ni buscar
        return None, 'La referencia lleva caracteres no válidos.'
    ref = historico._norm(ref)
    if len(ref) > REFERENCIA_MAX:
        return None, 'La referencia es demasiado larga (máximo %d caracteres).' % REFERENCIA_MAX
    return ref, None


def _reinicio_esperadas(datos):
    """(n, None), (None, None) si no viene, o (None, motivo). 'esperadas' = cuantas importaciones vio la persona al buscar."""
    if 'esperadas' not in datos:
        return None, None
    n = datos.get('esperadas')
    if isinstance(n, bool) or not isinstance(n, int) or n < 0 or n > 10000:
        return None, 'El dato «esperadas» no es válido.'
    return n, None


def _reinicio_fila(f):
    return {'fecha': historico.fecha_corta(f['fecha']), 'proveedor': f['proveedor'],
            'codigo_tecleado': f['codigo_tecleado'], 'codigo_erp': f['codigo_erp'], 'descripcion': f['descripcion'] or ''}


@app.route('/reiniciar/entrar', methods=['POST'])
def reiniciar_entrar():
    datos, error = _reinicio_autorizar()
    if error:
        return error
    return _reinicio_json({'ok': True})


@app.route('/reiniciar/buscar', methods=['POST'])
def reiniciar_buscar():
    datos, error = _reinicio_autorizar()
    if error:
        return error
    ref, problema = _reinicio_referencia(datos)
    if problema:
        return _reinicio_json({'ok': False, 'error': problema}, 400)
    try:
        filas = historico.coincidencias(ref)
    except Exception as e:
        print('ERROR leyendo el historico (reiniciar/buscar): %s' % e)
        return _reinicio_json({'ok': False, 'error': 'No se pudo leer el historial del programa.'}, 500)
    return _reinicio_json({'ok': True, 'referencia': ref, 'total': len(filas), 'filas': [_reinicio_fila(f) for f in filas]})


@app.route('/reiniciar/confirmar', methods=['POST'])
def reiniciar_confirmar():
    datos, error = _reinicio_autorizar()
    if error:
        return error
    ref, problema = _reinicio_referencia(datos)
    if problema:
        return _reinicio_json({'ok': False, 'error': problema}, 400)
    esperadas, problema = _reinicio_esperadas(datos)
    if problema:
        return _reinicio_json({'ok': False, 'error': problema}, 400)
    # No se toca el historial mientras hay una importacion en marcha (la importacion lo apunta al terminar).
    if not IMPORT_LOCK.acquire(timeout=3):
        return _reinicio_json({'ok': False, 'error': 'Hay una importación en curso. Espera a que termine y vuelve a intentarlo.'}, 409)
    try:
        filas = historico.reiniciar(ref, esperadas)
    except historico.HistorialCambiado as e:
        return _reinicio_json({'ok': False, 'cambiado': True, 'error': 'Lo apuntado de esa referencia ha cambiado desde que lo miraste (ahora hay %d). Vuelve a buscar.' % e.hay}, 409)
    except Exception as e:
        print('ERROR reiniciando la referencia %s: %s' % (ref, e))
        return _reinicio_json({'ok': False, 'error': 'No se pudo reiniciar la referencia (el historial no se ha modificado).'}, 500)
    finally:
        IMPORT_LOCK.release()
    print('REINICIAR REFERENCIA: %s -> %d fila(s) del historial olvidadas (archivadas en la tabla reinicios).' % (ref, len(filas)))
    return _reinicio_json({'ok': True, 'referencia': ref, 'borradas': len(filas), 'filas': [_reinicio_fila(f) for f in filas]})


def _warm_ref_index():
    """Construye el indice referencia->proveedores en segundo plano al arrancar, para que la
    primera busqueda no espere (con tarifas grandes puede tardar)."""
    try:
        print("[*] Indexando DATABASE en segundo plano (referencias repetidas entre proveedores)...")
        ref_index.providers_for(ROBOT.DEFAULT_DATABASE_PATH, 'ARRANQUE')
        total, secs = ref_index.stats()
        print("[*] Indice listo: %d referencias en %s s." % (total, secs))
    except Exception as e:
        print("[!] No se pudo construir el indice de referencias: %s" % e)

def get_local_ip():
    """Get the local IP address of the machine."""
    try:
        # Create a dummy socket to detect the local IP
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        # Doesn't need to be reachable (TEST-NET-1: UDP connect no envia nada)
        s.connect(("192.0.2.1", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "0.0.0.0"

def _create_desktop_shortcuts():
    """Crea la carpeta WEB en el escritorio con un .html y un .url que apuntan al servidor por nombre de equipo."""
    try:
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        web_folder = os.path.join(desktop_path, "WEB")
        os.makedirs(web_folder, exist_ok=True)

        hostname = socket.gethostname()
        shortcut_html = os.path.join(web_folder, "Don_Articulo.html")
        shortcut_url = os.path.join(web_folder, "Don_Articulo.url")

        # HTML que redirige al hostname del servidor
        html_content = f'<!DOCTYPE html><html><head><meta charset="utf-8"><meta http-equiv="refresh" content="0; url=http://{hostname}:5050" /><title>Don Articulo Web</title></head><body style="font-family: Arial; text-align: center; margin-top: 50px;"><h2>Conectando al servidor...</h2><p>Destino: http://{hostname}:5050</p></body></html>'

        # Acceso directo nativo de Windows (.url)
        url_content = f'[InternetShortcut]\nURL=http://{hostname}:5050/\n'

        with open(shortcut_html, "w", encoding="utf-8") as f:
            f.write(html_content)
        with open(shortcut_url, "w", encoding="utf-8") as f:
            f.write(url_content)

        print(f"\n[*] CARPETA CREADA: {web_folder}")
    except Exception as e:
        print(f"[!] Error creando carpeta WEB en el escritorio: {e}")


if __name__ == '__main__':
    load_config()
    threading.Thread(target=_warm_ref_index, daemon=True).start()
    local_ip = get_local_ip()
    print("DON ARTICULO WEB - Servidor iniciado")
    print(f"Accede desde: http://{local_ip}:5050")
    print(f"O desde: http://localhost:5050")
    # Opcional: carpeta "WEB" en el escritorio con accesos directos al servidor por nombre de equipo
    if os.environ.get("DONARTICULO_DESKTOP_SHORTCUTS") == "1":
        _create_desktop_shortcuts()

    # ======================================================================
    # SERVIDOR WSGI DE PRODUCCION (Waitress)
    # ----------------------------------------------------------------------
    # El servidor de desarrollo de Flask/Werkzeug NO esta pensado para uso
    # sostenido: acumula sockets en TIME_WAIT, filtra descriptores y se queda
    # sin hilos cuando los navegadores mantienen Keep-Alive ocioso. Tras
    # varias horas deja de aceptar conexiones aunque el proceso siga vivo.
    #
    # Waitress es un WSGI puro Python, nativo Windows, que sirve la MISMA
    # app Flask sin cambiar rutas ni logica. Esta en requirements.txt; si no
    # esta instalado se cae al app.run() para no dejar al usuario sin servidor.
    # ======================================================================
    try:
        import waitress  # noqa: F401
        _use_waitress = True
    except ImportError:
        print("[!] Waitress no esta instalado (pip install waitress).")
        _use_waitress = False

    if _use_waitress:
        from waitress import serve
        print("[*] Motor: Waitress (WSGI de produccion). Estable para uso continuo.")
        # threads: workers concurrentes. connection_limit: maximo de conexiones
        # aceptadas simultaneamente. channel_timeout: cierra clientes ociosos
        # para que nadie monopolice un worker indefinidamente.
        serve(
            app,
            host='0.0.0.0',
            port=5050,
            threads=8,
            connection_limit=200,
            channel_timeout=120,
            cleanup_interval=30,
            ident='DON_ARTICULO_WEB'
        )
    else:
        print("[!] Cayendo al servidor de desarrollo de Flask (puede colgarse tras horas de uso).")
        app.run(host='0.0.0.0', port=5050, debug=False, use_reloader=False, threaded=True)
