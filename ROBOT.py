#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Single Article Import
=====================

Searches the provider tariffs (one Excel file per provider in the DATABASE folder) for a single
article reference, converts that row to the ERP import format, writes importart.xlsx and
automates the import into the ERP with Selenium/Firefox.

Features:
- Searches all provider tariff files in the DATABASE folder
- Weighted keyword category classifier
- Generates a single-row import file
- Copies the file to the import folder
- Automates the Firefox ERP import with LOGIN support
"""

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import re
import os
import sys
import unicodedata
import random
import hashlib
import time
import json
from datetime import datetime, date
from pathlib import Path

# Try to import optional libraries
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    from selenium import webdriver
    from selenium.webdriver.firefox.service import Service as FirefoxService
    from selenium.webdriver.firefox.options import Options as FirefoxOptions
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.action_chains import ActionChains
    HAS_SELENIUM = True
except ImportError:
    HAS_SELENIUM = False

try:
    import pytz
    HAS_PYTZ = True
except ImportError:
    HAS_PYTZ = False

# ============================================================================
# CONSTANTS AND CONFIGURATION
# ============================================================================

VERSION = "5.0"  # OPTIMIZADO: Velocidad mejorada + Cierre automático Firefox

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Rutas y URL por defecto; todas se pueden cambiar con variables de entorno (ver README, "Configuracion").
DEFAULT_DATABASE_PATH = os.environ.get("DONARTICULO_DATABASE", os.path.join(BASE_DIR, "DATABASE"))
DEFAULT_NETWORK_PATH = os.environ.get("DONARTICULO_IMPORT_DIR", os.path.join(BASE_DIR, "import"))
DEFAULT_ERP_URL = os.environ.get("ERP_URL", "http://erp.empresa-x.example/")

# Consultas a servicios de terceros (buscador web para verificar una categoria dudosa): desactivadas por defecto.
# Activarlas (DONARTICULO_ONLINE_LOOKUPS=1) envia la descripcion del articulo a un buscador (ver README, "Privacidad").
ONLINE_LOOKUPS = os.environ.get("DONARTICULO_ONLINE_LOOKUPS") == "1"

# Zona horaria con la que se decide "es la primera ejecucion de hoy" (identificador estandar de pytz/IANA)
LOCAL_TIMEZONE = os.environ.get("DONARTICULO_TIMEZONE", "Europe/Madrid")

# Logo paths
LOGO_PATHS = [
    os.path.join(BASE_DIR, 'logo.png'),
    os.path.join(BASE_DIR, 'logo.webp'),
]

# ERP Template columns
ERP_COLUMNS = [
    'Codigo', 'Descripcion', 'Familia', 'Ud.Medida', 'Subfamilia', 'Sección',
    'Fabricante', 'Ean', 'Proveedor', 'Ref.Prov.', 'Precio Compra', '%Dto1',
    '%Dto2', '%Dto3', '%Dto4', '%Rappel', 'Precio Base Tarifa', '%Margen',
    'Calculo Base', 'Incr.Tarifa', 'Tarifa 1', 'Tarifa 2', 'Tarifa 3',
    'Tarifa 4', 'Tarifa 5', 'Tarifa 6', 'Tarifa 7', 'Tarifa 8', 'Tarifa 9',
    'Tarifa 10', 'Ud.Med.Compra', 'Factor Conver', 'UdxEnvase', 'Unid.Minima',
    'Peso', 'Volumen', 'Ctrl.Stock S/N', 'Lote S/N', '%IVA', 'Ud.Min.Compra',
    'Aplicar palet (S/N)', 'Apli.P.Coste S/N', 'Aplicar Dto/Rec S/N',
    'Recargo/Dto (R/D)', '% Rec/Dto', 'Paletizado S/N', 'Ud.Palet',
    'Grupo Precios', 'Agencia Trans', 'Factor Agt', 'Clave libre',
    'Stock Minimo', 'Descripcion Corta', 'Multiplo Com', 'Multiplo Ven',
    'Largo', 'Ancho', 'Profundo', 'Kit (S/N)', 'Tar.Man(S/N)', 'Peso Palet'
]

# ============================================================================
# CATEGORY CONFIGURATION
# ============================================================================

CATEGORIES = {
    'CERÁMICA': {
        'unit': 'M2',
        'familia': '1.1',
        'subfamilia': '1.1.1',
        'seccion': '1',
        'calculo_base': 'A',
        'use_ceramic_formula': True
    },
    'GRIFERÍA': {
        'unit': 'UD',
        'familia': '2.1',
        'subfamilia': '2.1.1',
        'seccion': '2',
        'calculo_base': 'B',
        'use_ceramic_formula': False
    },
    'SANITARIO': {
        'unit': 'UD',
        'familia': '3.1',
        'subfamilia': '3.1.1',
        'seccion': '3',
        'calculo_base': 'B',
        'use_ceramic_formula': False
    },
    'ACCESORIOS': {
        'unit': 'UD',
        'familia': '4.1',
        'subfamilia': '4.1.1',
        'seccion': '4',
        'calculo_base': 'B',
        'use_ceramic_formula': False
    },
    'MUEBLES': {
        'unit': 'UD',
        'familia': '5.1',
        'subfamilia': '5.1.1',
        'seccion': '5',
        'calculo_base': 'B',
        'use_ceramic_formula': False
    },
    'MAMPARAS': {
        'unit': 'UD',
        'familia': '6.1',
        'subfamilia': '6.1.1',
        'seccion': '6',
        'calculo_base': 'B',
        'use_ceramic_formula': False
    },
    'CEMENTO': {
        'unit': 'KG',
        'familia': '7.1',
        'subfamilia': '7.1.1',
        'seccion': '7',
        'calculo_base': 'B',
        'use_ceramic_formula': False
    },
    'COMPLEMENTOS': {
        'unit': 'UD',
        'familia': '8.1',
        'subfamilia': '8.1.1',
        'seccion': '8',
        'calculo_base': 'B',
        'use_ceramic_formula': False
    },
    'VARIOS': {
        'unit': 'UD',
        'familia': '9.1',
        'subfamilia': '9.1.1',
        'seccion': '9',
        'calculo_base': 'B',
        'use_ceramic_formula': False
    },
    'REPUESTOS': {
        'unit': 'UD',
        'familia': '10.1',
        'subfamilia': '10.1.1',
        'seccion': '10',
        'calculo_base': 'B',
        'use_ceramic_formula': False
    }
}

# v6.1 EXPANDED Weighted keyword dictionary for better classification
KEYWORD_DICT = {
    'CERÁMICA': {
        'primary': [
            # Core ceramic terms
            'ceramica', 'cerámica', 'azulejo', 'baldosa', 'porcelanico', 'porcelánico',
            'gres', 'pavimento', 'revestimiento', 'tile', 'floor', 'wall', 'mosaico',
            'tesela', 'formato', 'rectificado', 'mate', 'pulido', 'lapado', 'antideslizante',
            # Surface types
            'satinado', 'brillante', 'rustico', 'rústico', 'natural', 'esmaltado',
            # Material types
            'pasta blanca', 'pasta roja', 'pasta gris', 'monococcion', 'monoporosa',
            # Installation terms
            'rodapie', 'rodapié', 'zocalo', 'zócalo', 'cenefa', 'listelo', 'listello',
            'peldano', 'peldaño', 'escalera', 'esquina', 'angulo', 'ángulo',
            # Dimensions patterns
            'formato', 'calibrado', 'biselado', 'modular',
            # Decorative types
            'decor', 'decorado', 'decoración', 'relieve', 'textura', 'veteado',
            # Specific applications
            'fachada', 'exterior', 'interior', 'cocina', 'salon', 'salón',
            'terraza', 'piscina', 'baño',
            # Technical terms
            'r9', 'r10', 'r11', 'r12', 'r13', 'c1', 'c2', 'pei', 'absorcion',
            # Style terms  
            'hidraulico', 'hidráulico', 'vintage', 'metro', 'hexagonal', 'subway'
        ],
        'secondary': [
            'm2', 'm²', 'metro', 'suelo', 'pared',
            'ceramicas', 'cerámicas',
            # Colors removed: 'blanco','negro','gris','beige','marron','crema' are too generic
            # Material imitations
            'piedra', 'marmol', 'mármol', 'granito', 'pizarra', 'travertino',
            'madera', 'wood', 'cemento', 'hormigon', 'hormigón', 'terrazo',
            # Series/Collection
            'serie', 'coleccion', 'colección', 'modelo', 'linea', 'línea'
        ],
        'multi_word': [
            'pavimento porcelanico', 'revestimiento ceramico', 'gres porcelanico',
            'baldosa ceramica', 'azulejo ceramico', 'porcelanico rectificado', 
            'mosaico ceramico', 'pasta blanca', 'pasta roja', 'suelo porcelanico',
            'azulejo cocina', 'azulejo baño', 'baldosa exterior', 'baldosa interior',
            'gres extruido', 'porcelanico natural', 'porcelanico pulido',
            'revestimiento pared', 'pavimento suelo', 'ceramica decorativa'
        ],
        'exclusions': [
            'grifo', 'inodoro', 'mampara', 'sanitario', 'mueble baño', 'cisterna',
            # v3.0: CEMENT PRODUCT TYPES - CRITICAL EXCLUSIONS
            'mortero', 'adhesivo', 'cola', 'cemento', 'lechada', 'rejunte', 'borada'
        ]
    },
    'GRIFERÍA': {
        'primary': [
            # Core faucet terms
            'grifo', 'griferia', 'grifería', 'mezclador', 'monomando', 'bimando',
            'caño', 'faucet', 'tap', 'termostato', 'termostatico', 'termostático',
            # Shower components
            'teleducha', 'rociador', 'alcachofa', 'maneral', 'ducha mano',
            'flexo', 'flexo ducha', 'barra ducha', 'columna ducha', 'set ducha',
            # Specific faucet types (removed 'lavabo','fregadero','cocina','bidé','bide','bidet','bañera' - too ambiguous alone)
            'empotrado', 'mural', 'caño alto', 'caño giratorio', 'extraible',
            # Technical terms
            'cartucho', 'aireador', 'regulador', 'inversor', 'desviador',
            # Connection terms
            'latiguillo', 'racor', 'conexion', 'conexión', 'rosca', 'escudo', 'floron'
        ],
        'secondary': [
            # Finishes
            'cromo', 'cromado', 'acero', 'inox', 'inoxidable', 'negro', 'negro mate',
            'dorado', 'oro', 'bronce', 'niquel', 'níquel', 'cobre',
            # Types
            'monoblock', 'temporizado', 'electronico', 'electrónico', 'sensor'
        ],
        'multi_word': [
            'grifo monomando', 'grifo bimando', 'mezclador lavabo', 'grifo cocina',
            'columna ducha', 'set ducha', 'barra ducha', 'grifo fregadero',
            'grifo bañera', 'grifo bidet', 'grifo lavabo', 'grifo empotrado',
            'termostato ducha', 'termostato bañera', 'kit ducha', 'conjunto ducha',
            'grifo caño alto', 'grifo extraible', 'ducha empotrada', 'rociador techo',
            'toma de agua', 'cuerpo empotrado', 'parte empotrada', 'set empotrado',
            'cuerpo de empotrar', 'griferia bañera', 'monomando empotrado',
            'griferia lavabo', 'griferia bide', 'griferia ducha', 'cuerpo oculto'
        ],
        'exclusions': ['ceramica', 'cerámica', 'azulejo', 'baldosa', 'mampara', 'porcelanico']
    },
    'SANITARIO': {
        'primary': [
            # Core sanitary terms (removed standalone 'lavabo' - too ambiguous, kept in multi_word)
            'inodoro', 'sanitario', 'wc', 'water', 'bide', 'bidé', 'bidet',
            'urinario', 'taza', 'toilet', 'basin', 'pedestal', 'semipedestal',
            # Toilet types
            'cisterna', 'tanque', 'fluxor', 'suspendido', 'adosado', 'compacto',
            'rimless', 'sin brida', 'salida dual', 'salida horizontal', 'salida vertical',
            # Basin types
            'encimera', 'sobre encimera', 'bajo encimera', 'encastrar', 'mural',
            'esquina', 'angular', 'ovalado', 'redondo', 'rectangular', 'cuadrado',
            # Accessories
            'asiento', 'tapa', 'tapa wc', 'caida amortiguada', 'soft close',
            'mecanismo', 'descarga', 'pulsador', 'flotador', 'llenado',
            # English Equivalents
            'urinal', 'cistern', 'washbasin', 'flush', 'flush valve', 'tank',
            'bowl', 'above counter', 'wall hung', 'back to wall', 'rimless'
        ],
        'secondary': [
            # Materials
            'porcelana', 'vitrificado', 'vitreous', 'china', 'ceramica',
            # Colors
            'blanco', 'pergamon', 'hueso', 'negro', 'gris'
        ],
        'multi_word': [
            'taza inodoro', 'lavabo suspendido', 'inodoro compacto', 'pack inodoro',
            'lavabo pedestal', 'lavabo encimera', 'bidet suspendido', 'wc suspendido',
            'inodoro rimless', 'inodoro sin brida', 'tapa wc', 'asiento inodoro',
            'lavabo sobre encimera', 'lavabo bajo encimera', 'lavabo mural',
            'conjunto sanitario', 'pack sanitario', 'juego sanitario',
            'above counter bowl', 'above counter washbasin', 'wall hung wc',
            'back to wall wc', 'back to wall bidet', 'wall hung bidet'
        ],
        'exclusions': ['grifo', 'griferia', 'mampara', 'azulejo', 'baldosa', 'mueble']
    },
    'ACCESORIOS': {
        'primary': [
            # Bathroom accessories (removed 'vaso' - too generic)
            'toallero', 'jabonera', 'portarrollos', 'percha', 'colgador',
            'escobillero', 'dosificador', 'dispensador', 'accesorio',
            'aro', 'argolla', 'gancho', 'repisa', 'bandeja', 'barra',
            # Specific items
            'portacepillos', 'portavasos', 'porta', 'soporte',
            'secamanos', 'secador manos', 'papelera', 'cubo',
            # Mirrors and lighting
            'espejo baño', 'luz baño', 'aplique',
            # Safety items
            'asidero', 'barra seguridad', 'barra apoyo',
            # English Equivalents
            'dispenser', 'towel', 'towel bar', 'soap dispenser', 'paper towel',
            'grab bar', 'hand dryer', 'waste receptacle', 'hook', 'tissue',
            'napkin', 'waste', 'receptacle', 'brush holder', 'holder', 'tray'
        ],
        'secondary': [
            # Materials/Finishes
            'cromado', 'cromo', 'acero', 'inox', 'inoxidable',
            'negro', 'blanco', 'mate', 'brillo',
            'cristal', 'vidrio', 'laton', 'latón', 'zamak',
            'stainless steel', 'matte black',
            # General
            'baño', 'aseo', 'cuarto baño', 'washroom', 'bathroom'
        ],
        'multi_word': [
            'toallero baño', 'portarrollos papel', 'accesorio baño', 
            'escobillero baño', 'jabonera pared', 'vaso cepillos',
            'dosificador jabon', 'percha pared', 'toallero barra',
            'toallero anilla', 'porta toallas', 'porta cepillos',
            'barra toalla', 'set accesorios', 'conjunto accesorios',
            'soap dispenser', 'paper towel dispenser', 'grab bar',
            'waste receptacle', 'hand dryer', 'tray for accessories',
            'roll holder', 'towel holder', 'cup holder'
        ],
        'exclusions': ['grifo', 'inodoro', 'mampara', 'mueble', 'ceramica', 'azulejo', 'baldosa', 'faucet', 'toilet']
    },
    'MUEBLES': {
        'primary': [
            # Furniture types
            'mueble', 'armario', 'espejo', 'columna', 'cajon', 'cajón',
            'estanteria', 'estantería', 'mobiliario', 'encimera', 'auxiliar',
            # Specific furniture
            'bajo', 'alto', 'camerino', 'tocador', 'vanity', 'modulo', 'módulo',
            'aplique', 'foco', 'iluminacion', 'iluminación',
            # Construction
            'puerta', 'tirador', 'bisagra', 'pata', 'balda',
            # English Equivalents
            'cabinet', 'medicine cabinet', 'mirror', 'drawer', 'shelf',
            'shelving', 'cupboard', 'vanity unit', 'wardrobe', 'furniture'
        ],
        'secondary': [
            # Materials
            'madera', 'melamina', 'lacado', 'dm', 'mdf', 'aglomerado',
            'roble', 'wenge', 'nogal', 'fresno', 'blanco brillo',
            'wood', 'oak', 'walnut', 'veneer',
            # General
            'baño', 'lavabo', 'aseo', 'tocador', 'bathroom', 'washroom'
        ],
        'multi_word': [
            'mueble baño', 'mueble lavabo', 'espejo baño', 'columna baño',
            'mueble auxiliar', 'espejo led', 'espejo luz', 'mueble suspendido',
            'mueble suelo', 'conjunto mueble', 'mueble con lavabo',
            'medicine cabinet', 'bathroom cabinet', 'vanity unit',
            'wall cabinet', 'base cabinet', 'mirror cabinet'
        ],
        'exclusions': ['grifo', 'ceramica', 'inodoro', 'azulejo', 'baldosa', 'mampara', 'faucet', 'toilet', 'tile']
    },
    'MAMPARAS': {
        'primary': [
            # Screen types
            'mampara', 'cancel', 'pantalla', 'cabina', 'panel fijo', 'hoja',
            # Door types
            'frontal', 'angular', 'semicircular', 'cuadrante', 'pentagono', 'pentágono',
            'corredera', 'abatible', 'plegable', 'pivotante', 'batiente',
            # Shower bases
            'plato', 'plato ducha', 'receptor', 'shower tray',
            # Technical terms
            'perfil', 'rodamiento', 'junta', 'goma', 'sello', 'guia', 'guía'
        ],
        'secondary': [
            # Materials
            'cristal', 'vidrio', 'templado', 'serigrafiado', 'transparente', 'opaco',
            'acrílico', 'acrilico', 'antical',
            # Finishes
            'cromo', 'negro', 'blanco', 'plata',
            # General
            'ducha', 'bañera', 'baño', 'aseo'
        ],
        'multi_word': [
            'mampara ducha', 'mampara bañera', 'plato ducha', 'cabina ducha',
            'frontal ducha', 'angular ducha', 'mampara corredera', 'mampara abatible',
            'mampara plegable', 'panel fijo', 'cristal templado', 'vidrio templado',
            'plato resina', 'plato acrilico', 'plato porcelanico'
        ],
        'exclusions': ['grifo', 'ceramica', 'inodoro', 'azulejo', 'baldosa', 'mueble']
    },
    'CEMENTO': {
        'primary': [
            # Core materials
            'cemento', 'adhesivo', 'mortero', 'cola', 'junta', 'rejunte', 'rejuntado',
            'sellador', 'masilla', 'pegamento', 'lechada', 'impermeabilizante',
            # Types
            'imprimacion', 'imprimación', 'primer', 'nivelador', 'autonivelante',
            'yeso', 'escayola', 'enlucido', 'enfoscado',
            # Waterproofing
            'impermeabilizacion', 'lamina', 'manta', 'tela asfaltica',
            # Sealants
            'silicona', 'poliuretano', 'ms polimer', 'acrilica', 'acrílica',
            # Product types
            'cola porcelanico', 'cola azulejo', 'cola ceramica', 'cola baldosa',
            'mortero cola', 'cemento cola', 'adhesivo ceramico', 'adhesivo azulejo',
            'borada', 'rejunte', 'junta ceramica', 'junta azulejo',
            'c1', 'c2', 'c2te', 'c2e', 'c2s1', 'c2s2',  # Cement classification codes
        ],
        'secondary': [
            # Units
            'kg', 'kilo', 'saco', 'bolsa', 'bote', 'cartucho', 'cubo',
            # Colors
            'blanco', 'gris', 'negro', 'beige', 'marron'
        ],
        'multi_word': [
            'adhesivo ceramico', 'mortero cola', 'junta ceramica', 'cemento cola',
            'adhesivo flexible', 'adhesivo porcelanico', 'lechada epoxy',
            'junta epoxi', 'sellador silicona', 'impermeabilizante liquido',
            'mortero nivelador', 'pasta niveladora', 'masilla reparadora',
            'cola para ceramica', 'cola para porcelanico', 'cola para azulejos',
            'mortero para ceramica', 'pegamento ceramico', 'pegamento azulejo'
        ],
        'exclusions': []
    },
    'COMPLEMENTOS': {
        'primary': [
            # Plumbing fittings
            'sifon', 'sifón', 'valvula', 'válvula', 'desague', 'desagüe',
            'tubo', 'flexible', 'racor', 'codo', 'enlace', 'manguito',
            'latiguillo', 'llave', 'llave paso', 'llave escuadra',
            # Trims and profiles
            'perfil', 'cantonera', 'remate', 'junquillo', 'esquinero', 'canto',
            'rodapie', 'rodapié', 'transicion', 'transición', 'prolongador',
            'cobertura', 'embellecedor',
            # Small parts
            'tornillo', 'taco', 'fijacion', 'fijación', 'anclaje',
            'soporte', 'escuadra', 'angulo', 'ángulo'
        ],
        'secondary': [
            # Materials
            'fontaneria', 'fontanería', 'pvc', 'metal', 'laton', 'latón',
            'aluminio', 'acero', 'inox', 'plastico', 'plástico',
            # General
            'instalacion', 'instalación', 'conexion', 'conexión'
        ],
        'multi_word': [
            'sifon lavabo', 'valvula descarga', 'desague ducha', 'llave paso',
            'llave escuadra', 'tubo conexion', 'flexible conexion',
            'perfil aluminio', 'cantonera aluminio', 'perfil transicion',
            'junquillo pvc', 'rodapie pvc', 'taco expansion',
            'partes externas', 'parte externa', 'parte exterior'
        ],
        'exclusions': ['grifo completo', 'monomando', 'bimando', 'mezclador']
    },
    'REPUESTOS': {
        'primary': [
            # Core terms
            'repuesto', 'recambio', 'pieza', 'mecanismo', 'cartucho',
            'spare', 'replacement', 'kit reparacion', 'reparacion', 'reparación',
            # Specific parts
            'junta', 'oring', 'o-ring', 'membrana', 'diafragma',
            'reten', 'retén', 'sello', 'vastago', 'vástago',
            'maneta', 'palanca', 'pomo', 'inversor', 'desviador',
            # Toilet parts
            'flotador', 'pulsador', 'tirador descarga', 'mecanismo cisterna',
            'boya', 'llenado', 'descarga',
            # Griferia parts
            'escudo', 'floron', 'florón', 'aireador'
        ],
        'secondary': [
            # Types
            'sustitucion', 'sustitución', 'cambio', 'original', 'compatible',
            'universal', 'especifico', 'específico',
            # Context words
            'para', 'modelo', 'referencia'
        ],
        'multi_word': [
            'repuesto grifo', 'cartucho grifo', 'mecanismo cisterna', 
            'kit reparacion', 'pieza repuesto', 'junta recambio',
            'cartucho monomando', 'mecanismo descarga', 'flotador cisterna',
            'pulsador cisterna', 'repuesto ducha', 'recambio original'
        ],
        'exclusions': []
    },
    'VARIOS': {
        'primary': [],
        'secondary': [],
        'multi_word': [],
        'exclusions': []
    }
}

CATEGORY_PRIORITY = [
    'CERÁMICA', 'GRIFERÍA', 'SANITARIO', 'ACCESORIOS', 'MUEBLES',
    'MAMPARAS', 'CEMENTO', 'COMPLEMENTOS', 'REPUESTOS', 'VARIOS'
]

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def clean_text_for_erp(text):
    """Clean text for ERP compatibility and force uppercase."""
    if pd.isna(text) or text is None:
        return ''
    
    text = str(text).strip()
    if not text:
        return ''
    
    text = unicodedata.normalize('NFC', text)
    
    replacements = {
        '\u2013': '-', '\u2014': '-', '\u2026': '...',
        '\u201c': '"', '\u201d': '"', '\u2018': "'", '\u2019': "'",
        '\x91': "'", '\x92': "'", '\x93': '"', '\x94': '"',
        '\r\n': ' ', '\n': ' ', '\r': ' ', '\t': ' '
    }

    for old, new in replacements.items():
        text = text.replace(old, new)
        
    # Remove control characters except normal printable ones
    text = ''.join(c for c in text if unicodedata.category(c)[0] != 'C' or c == ' ')
    
    # Remove multiple spaces
    text = ' '.join(text.split())
    
    return text.upper().upper()


def _dedup_repeated_words(text):
    """Quita palabras duplicadas (case-insensitive, ignorando acentos y plurales s/es)
    preservando el orden. Usado por los perfiles con 'dedup_words'."""
    if not text:
        return ''
    seen = set()
    out = []
    for w in str(text).split():
        norm = normalize_for_matching(w)
        if norm.endswith('es') and len(norm) > 3:
            stem = norm[:-2]
        elif norm.endswith('s') and len(norm) > 2:
            stem = norm[:-1]
        else:
            stem = norm
        if stem in seen:
            continue
        seen.add(stem)
        out.append(w)
    return ' '.join(out)


# ============================================================================
# LOOKUPS DE CATEGORIA POR CODIGO (patron generico)
# ----------------------------------------------------------------------------
# Para algunos proveedores la clasificacion por palabras clave no es fiable y la categoria se fija a mano en un
# Excel (CODIGO | CATEGORIA) que vive JUNTO AL PROGRAMA y no en DATABASE, para que la busqueda de referencias no
# lo confunda con una tarifa. Se carga una sola vez (cache). Un perfil de proveedor lo activa con la clave
# 'category_lookup' (ver PROVIDER_PROFILES).
# ============================================================================
_CATEGORY_LOOKUPS = {}
_CATEGORY_CANON = {'GRIFERIA': 'GRIFERÍA', 'CERAMICA': 'CERÁMICA'}


def _norm_code(code):
    return str(code).strip().upper().replace(' ', '')


def load_category_lookup(filename, key_col='CODIGO', cat_col='CATEGORIA'):
    """{codigo normalizado: categoria} de un Excel junto al programa; {} si no existe o no se puede leer."""
    if filename in _CATEGORY_LOOKUPS:
        return _CATEGORY_LOOKUPS[filename]
    lookup = {}
    path = os.path.join(BASE_DIR, filename)
    try:
        if os.path.exists(path):
            valid = set(CATEGORIES.keys())
            df = pd.read_excel(path, dtype=str)
            for _, r in df.iterrows():
                cod = _norm_code(r.get(key_col, '') or '')
                cat = str(r.get(cat_col, '') or '').strip().upper()
                cat = _CATEGORY_CANON.get(cat, cat)
                if cod and cat in valid:
                    lookup[cod] = cat
    except Exception as e:
        print(f"Aviso: no se pudo leer {filename}: {e}")
    _CATEGORY_LOOKUPS[filename] = lookup
    return lookup


def category_for_code(filename, codigo, default=None):
    """Categoria fijada para un codigo en el Excel `filename`, o `default` si el codigo no esta."""
    return load_category_lookup(filename).get(_norm_code(codigo), default)


# ============================================================================
# LOGISTICA DE CERAMICA POR PROVEEDOR
# ----------------------------------------------------------------------------
# No es un script general: cada proveedor trae sus columnas con otro nombre, asi que cada uno tiene su FICHA en
# _CERAMIC_LOGISTICS_SPEC con los nombres EXACTOS de SUS columnas (o None si no las trae). Reglas:
#   Peso        = kgs/caja / (m2/caja si se vende en M2; pz/caja si se vende en PZ)  -> kg de UNA unidad de venta
#   Unid.Minima = Multiplo = m2/caja (M2) o pz/caja (PZ): la caja es el minimo a pedir
#   UdxEnvase   = pz/caja;  Ud.Palet = cajas/palet;  Peso Palet = kg/palet de la tarifa (o kgs/caja x cajas/palet)
# Sin el dato, el campo queda en blanco: no se inventa nada ni se cae al peso de la caja. Los proveedores que no
# estan aqui siguen con la logica generica.
# UNIDADES: la ficha traduce la unidad de SU tarifa a la del ERP ('unidades'). Una unidad "caja" puede tratarse como
# pieza salvo en el PRECIO: si la tarifa trae el de la CAJA ('precio_por_caja') se divide entre las piezas por caja
# (_ceramic_price_divisor).
# OJO: la columna de kg NO significa lo mismo en todos los proveedores (aunque se llame "kg/caja"):
#   'kg_caja'   = kg de la CAJA entera  -> hay que dividir
#   'kg_unidad' = YA es el kg de UNA unidad de venta (kg/m2 o kg/pieza) -> se deja tal cual
# En esta copia solo queda una ficha de ejemplo (proveedor ficticio).
# ============================================================================
_CERAMIC_LOGISTICS_SPEC = {
    'PROVEEDOR_001': {
        'pz_caja': 'PIEZAS/CAJA', 'm2_caja': 'M2/CAJA', 'kg_caja': 'KG/CAJA',
        'cajas_palet': 'CAJAS/PALET', 'kg_palet': 'KG/PALET',
        'unidad_col': 'UNIDAD', 'unidades': {'M2': 'M2', 'PZ': 'PZ', 'PI': 'PZ', 'UN': 'PZ', 'CA': 'PZ'},
        'precio_por_caja': ['CA'],
    },
}


def _header_key(col):
    """Cabecera en minusculas, sin tildes y con los espacios/saltos de linea colapsados."""
    return re.sub(r'\s+', ' ', normalize_for_matching(unicodedata.normalize('NFC', str(col)))).strip()


def _whole_number(valor):
    """17.0 -> 17 (cajas, piezas); el resto de numeros se dejan como estan."""
    return int(valor) if valor is not None and float(valor).is_integer() else valor


def _ceramic_logistics_spec(provider_name):
    """Ficha de columnas del proveedor, o None si ese proveedor no tiene logistica propia (sigue la generica)."""
    return _CERAMIC_LOGISTICS_SPEC.get(str(provider_name or '').strip().upper())


def _ceramic_num(row_data, key):
    """Numero > 0 de la celda (admite texto con coma decimal), o None si esta vacia, a 0 o no es numerica."""
    if key is None:
        return None
    valor = row_data.get(key)
    try:
        n = float(valor)
    except (TypeError, ValueError):
        try:
            n = float(str(valor).strip().replace(',', '.'))
        except ValueError:
            return None
    return n if n == n and n > 0 else None


def _ceramic_unit(row_data, provider_name):
    """Unidad del ERP (M2 / PZ) que corresponde a la unidad de la tarifa segun la ficha del proveedor, o None si el
    proveedor no tiene ficha o la unidad no esta en ella (entonces se deja la traduccion generica de siempre)."""
    spec = _ceramic_logistics_spec(provider_name)
    if not spec or not spec.get('unidad_col'):
        return None
    claves = {_header_key(k): k for k in row_data}
    clave = claves.get(_header_key(spec['unidad_col']))
    valor = str(row_data.get(clave, '')).strip().upper() if clave else ''
    return (spec.get('unidades') or {}).get(valor)


def _ceramic_price_divisor(row_data, provider_name):
    """Piezas por caja entre las que hay que dividir el PRECIO de la tarifa cuando ese precio es el de la CAJA (ficha:
    'precio_por_caja' = unidades de la tarifa afectadas), o None si el precio ya es por unidad de venta o falta el dato."""
    spec = _ceramic_logistics_spec(provider_name)
    if not spec or not spec.get('precio_por_caja') or not spec.get('unidad_col'):
        return None
    claves = {_header_key(k): k for k in row_data}
    unidad = str(row_data.get(claves.get(_header_key(spec['unidad_col'])), '')).strip().upper()
    if unidad not in spec['precio_por_caja']:
        return None
    return _ceramic_num(row_data, claves.get(_header_key(spec['pz_caja'])))


def _ceramic_logistics(row_data, unidad, provider_name):
    """Logistica de una fila de ceramica segun la ficha del proveedor y la unidad de venta (M2/PZ), o None si el proveedor no
    tiene ficha o la unidad no es M2/PZ (entonces sigue la logica generica). Los campos sin dato salen None."""
    spec = _ceramic_logistics_spec(provider_name)
    if spec is None or unidad not in ('M2', 'PZ'):
        return None
    claves = {_header_key(k): k for k in row_data}
    num = lambda campo: _ceramic_num(row_data, claves.get(_header_key(spec[campo]))) if spec.get(campo) else None
    pz_caja, m2_caja, kg_caja, kg_unidad = num('pz_caja'), num('m2_caja'), num('kg_caja'), num('kg_unidad')
    cajas_palet, kg_palet = num('cajas_palet'), num('kg_palet')
    unidades_caja = m2_caja if unidad == 'M2' else pz_caja          # unidades de venta que caben en una caja
    if kg_unidad is not None and kg_caja is None and unidades_caja:
        kg_caja = kg_unidad * unidades_caja                          # kg de la caja deducidos (solo para el palet si falta)
    if cajas_palet is None and unidades_caja:                       # la tarifa no trae cajas/palet pero si m2 o piezas por palet
        por_palet = num('m2_palet') if unidad == 'M2' else num('pz_palet')
        if por_palet:
            cajas_palet = round(por_palet / unidades_caja)
    if kg_palet is None and kg_caja and cajas_palet:
        kg_palet = round(kg_caja * cajas_palet, 2)
    return {
        'udx_envase': _whole_number(pz_caja),
        'unid_minima': _whole_number(unidades_caja),
        'multiplo': _whole_number(unidades_caja),
        'ud_palet': _whole_number(cajas_palet),
        'peso': kg_unidad if kg_unidad is not None else (round(kg_caja / unidades_caja, 6) if kg_caja and unidades_caja else None),
        'peso_palet': kg_palet,
    }


# ============================================================================
# PERFILES POR PROVEEDOR
# ----------------------------------------------------------------------------
# Las particularidades de un proveedor se declaran aqui, no con ramas `if proveedor == ...` dentro del convertidor.
# Clave: nombre del proveedor (carpeta de DATABASE) en mayusculas y sin espacios. Claves admitidas:
#   column_patterns         {'code': [regex], 'name': [regex]}: columnas de codigo / nombre cuando el patron generico
#                           se equivoca (p. ej. 'articulo' casa antes con 'Cod. Articulo')
#   strip_parentheses       quita las anotaciones "(...)" del nombre (marca blanca, notas de embalaje o corte...)
#   price_decimal_comma     el precio viene como texto europeo ("69,64" o "1.108,94")
#   price_factor            factor sobre el precio de la tarifa (la tarifa no es el precio de compra real)
#   verify_with_web         verifica la clasificacion dudosa con una busqueda en internet
#   force_category          categoria fija (anula al clasificador)
#   category_lookup         Excel CODIGO|CATEGORIA junto al programa que fija la categoria por codigo
#   dedup_words             quita las palabras repetidas de las descripciones
#   price_as_is             Precio Base Tarifa = Precio Compra (sin formula ceramica)
#   pro_discount / private_discount   % de Tarifa 2 / Tarifa 3 sobre Tarifa 1
#   dto1                    %Dto1 fijo
#   calculo_base            fuerza el Calculo Base ('A' o 'B')
#   provider_number         nº de proveedor en Fabricante (G) y Proveedor (I)
#   supplier_column_number  nº distinto solo para la columna Proveedor (I)
# Los perfiles reales de cada proveedor se han omitido de esta copia: solo queda uno ficticio de ejemplo.
# ============================================================================
PROVIDER_PROFILES = {
    'PROVEEDOR_EJEMPLO': {
        'column_patterns': {'code': ['^cod\\.? ?articulo$'], 'name': ['^articulo$']},
        'strip_parentheses': True,
        'price_as_is': True,
        'pro_discount': 28.0,
        'private_discount': 17.0,
        'dto1': 38.0,
        'calculo_base': 'B',
        'provider_number': 1000,
    },
}

# Proveedores cuyo catalogo es siempre ceramica: el clasificador por palabras clave se salta (coincidencia por subcadena).
FORCED_CERAMIC_PROVIDERS = ('PROVEEDOR_001',)

# Politica de precios de EJEMPLO (valores ficticios: pon la tuya).
# Formula ceramica por defecto: Precio Base Tarifa = (Precio Compra x CERAMIC_MULT) + CERAMIC_ADD
CERAMIC_MULT = 1.8
CERAMIC_ADD = 4.5
# Descuentos (%) de la categoria GRIFERIA en Tarifa 2 (profesional) y Tarifa 3 (particular)
GRIFERIA_PRO_DISCOUNT = 21.0
GRIFERIA_PRIVATE_DISCOUNT = 9.0
# Descuentos (%) por defecto de la ventana de escritorio
DEFAULT_PRO_DISCOUNT = 27
DEFAULT_PRIVATE_DISCOUNT = 11


def provider_profile(provider_name):
    """Perfil del proveedor (dict, vacio si no tiene)."""
    return PROVIDER_PROFILES.get(re.sub(r'\s+', '', str(provider_name or '').upper()), {})


def normalize_for_matching(text):
    """Normalize text for keyword matching."""
    if not text:
        return ''
    text = str(text).lower().strip()
    replacements = {'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u', 'ü': 'u', 'ñ': 'n'}
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def generate_ean13(base_code):
    """Generate a valid EAN-13 code, stable for a given base_code (the same reference always gets the same EAN)."""
    seed = int(hashlib.sha256(str(base_code).encode('utf-8')).hexdigest(), 16)
    rng = random.Random(seed)
    digits = [8, 4] + [rng.randint(0, 9) for _ in range(10)]
    odd_sum = sum(digits[::2])
    even_sum = sum(digits[1::2])
    check = (10 - ((odd_sum + even_sum * 3) % 10)) % 10
    return ''.join(map(str, digits)) + str(check)


def get_local_date():
    """Get today's date in LOCAL_TIMEZONE. With an unknown timezone (or without pytz) it is the machine's local date."""
    if HAS_PYTZ:
        try:
            return datetime.now(pytz.timezone(LOCAL_TIMEZONE)).date()
        except pytz.UnknownTimeZoneError:
            print(f"Aviso: zona horaria desconocida '{LOCAL_TIMEZONE}' (DONARTICULO_TIMEZONE); se usa la fecha local del equipo.")
    return date.today()


# ============================================================================
# PROVIDER CATEGORY COLUMN TRANSLATION
# ============================================================================
# Maps provider-specific category values to our standard categories.
# Used when a provider's Excel has a "Categoria"/"Familia"/"Tipo" column.

PROVIDER_CATEGORY_TRANSLATIONS = {
    # Common Spanish terms
    'grifo': 'GRIFERÍA', 'grifos': 'GRIFERÍA', 'griferia': 'GRIFERÍA',
    'fregadero': 'GRIFERÍA', 'fregaderos': 'GRIFERÍA',
    'sanitario': 'SANITARIO', 'sanitarios': 'SANITARIO',
    'inodoro': 'SANITARIO', 'inodoros': 'SANITARIO',
    'lavabo': 'SANITARIO', 'lavabos': 'SANITARIO',
    'bide': 'SANITARIO', 'bidet': 'SANITARIO', 'bidets': 'SANITARIO',
    'mueble': 'MUEBLES', 'muebles': 'MUEBLES', 'mobiliario': 'MUEBLES',
    'mampara': 'MAMPARAS', 'mamparas': 'MAMPARAS',
    'plato ducha': 'MAMPARAS', 'platos ducha': 'MAMPARAS',
    'accesorio': 'ACCESORIOS', 'accesorios': 'ACCESORIOS',
    'complemento': 'COMPLEMENTOS', 'complementos': 'COMPLEMENTOS',
    'ceramica': 'CERÁMICA', 'azulejo': 'CERÁMICA', 'pavimento': 'CERÁMICA',
    'cemento': 'CEMENTO', 'adhesivo': 'CEMENTO', 'mortero': 'CEMENTO',
    'repuesto': 'REPUESTOS', 'repuestos': 'REPUESTOS', 'recambio': 'REPUESTOS',
    # English terms
    'faucet': 'GRIFERÍA', 'faucets': 'GRIFERÍA', 'tap': 'GRIFERÍA', 'taps': 'GRIFERÍA',
    'sink': 'GRIFERÍA', 'sinks': 'GRIFERÍA',
    'toilet': 'SANITARIO', 'basin': 'SANITARIO', 'washbasin': 'SANITARIO',
    'furniture': 'MUEBLES', 'cabinet': 'MUEBLES', 'vanity': 'MUEBLES',
    'shower enclosure': 'MAMPARAS', 'shower screen': 'MAMPARAS',
    'accessory': 'ACCESORIOS', 'accessories': 'ACCESORIOS',
    'dispenser': 'ACCESORIOS', 'towel': 'ACCESORIOS', 'grab bar': 'ACCESORIOS',
    'waste': 'ACCESORIOS', 'dryer': 'ACCESORIOS',
}

def _translate_provider_category(raw_value):
    """Translate a provider's category column value to our standard category."""
    if not raw_value:
        return None
    val = str(raw_value).strip().lower()
    val = normalize_for_matching(val)
    if not val or val == 'nan':
        return None
    # Direct match
    if val in PROVIDER_CATEGORY_TRANSLATIONS:
        return PROVIDER_CATEGORY_TRANSLATIONS[val]
    # Partial match
    for key, cat in PROVIDER_CATEGORY_TRANSLATIONS.items():
        if key in val:
            return cat
    return None


# ============================================================================
# v6.0 PRODUCT CLASSIFIER
# ============================================================================

class ProductClassifier:
    """v3.1 Robust classifier with CONTEXTUAL PRIORITY, weighted keywords and web verification."""
    
    PRIMARY_WEIGHT = 20
    SECONDARY_WEIGHT = 5
    MULTI_WORD_WEIGHT = 60
    EXCLUSION_PENALTY = -50
    BRAND_WEIGHT = 15  # v3.1: Brands are a hint, not an absolute truth
    
    # v7.0: MULTI-CATEGORY BRAND MAPPING
    # Single-category brands get a definitive bonus.
    # Multi-category brands get a bonus to ALL their categories; keywords decide the winner.
    # En esta copia el mapa va con marcas ficticias; el real tenia muchos fabricantes.
    BRAND_CATEGORIES = {
        # Marca de una sola categoria: bonus definitivo
        'marca ejemplo cemento': ['CEMENTO'],
        'marca ejemplo ceramica': ['CERÁMICA'],
        # Marca multi-categoria: bonus a TODAS sus categorias; las palabras clave deciden
        'marca ejemplo multi': ['SANITARIO', 'GRIFERÍA', 'MAMPARAS'],
    }

    # v2.5: Web verification hints for each category
    WEB_VERIFICATION_HINTS = {
        'CEMENTO': ['mortero', 'adhesivo', 'cola', 'cemento', 'pegamento', 'construccion', 'colocacion', 'junta', 'rejunte'],
        'CERÁMICA': ['ceramica', 'azulejo', 'baldosa', 'porcelanico', 'revestimiento', 'pavimento', 'suelo', 'pared'],
        'GRIFERÍA': ['grifo', 'griferia', 'monomando', 'ducha', 'agua', 'lavabo', 'cocina'],
        'SANITARIO': ['sanitario', 'inodoro', 'lavabo', 'wc', 'baño', 'porcelana'],
        'ACCESORIOS': ['accesorio', 'toallero', 'jabonera', 'portarrollos', 'percha'],
        'MUEBLES': ['mueble', 'armario', 'espejo', 'mobiliario', 'cajon'],
        'MAMPARAS': ['mampara', 'ducha', 'cristal', 'plato ducha', 'cabina'],
        'COMPLEMENTOS': ['sifon', 'valvula', 'desague', 'tubo', 'perfil', 'fontaneria'],
        'REPUESTOS': ['repuesto', 'recambio', 'pieza', 'cartucho', 'mecanismo'],
    }
    
    def __init__(self, enable_web_verification=True):
        self.detected_category = None
        self.category_scores = {}
        self.enable_web_verification = enable_web_verification
    
    def classify_product(self, text, verify_with_web=False, provider_category=None):
        """Classify a single product based on its text.

        v7.0: Multi-category brand support. Provider category column takes priority.

        Args:
            text: Product description text
            verify_with_web: If True, verify classification with web search
            provider_category: Category from provider's own Excel column (highest priority)

        Returns:
            Category string
        """
        if not text:
            return 'VARIOS'

        # Phase 2: If provider's Excel has a category column, use it directly
        if provider_category:
            translated = _translate_provider_category(provider_category)
            if translated:
                self.category_scores = {cat: 0 for cat in CATEGORIES}
                self.category_scores[translated] = 200
                return translated

        normalized_text = normalize_for_matching(text)

        # Calculate base scores from text keywords
        scores = self._calculate_scores(normalized_text)

        # v7.0: Apply brand bonus using multi-category system
        brand_categories = self._check_brand_categories(normalized_text)
        if brand_categories:
            if len(brand_categories) == 1:
                # Single-category brand: definitive bonus
                scores[brand_categories[0]] += self.BRAND_WEIGHT + 50
            else:
                # Multi-category brand: bonus to ALL possible categories, keywords decide
                for cat in brand_categories:
                    scores[cat] += 50
                    

        self.category_scores = scores

        initial_category = self._get_best_category(scores)

        # v2.5: Optional web verification for uncertain classifications
        if verify_with_web and self.enable_web_verification:
            max_score = max(scores.values()) if scores else 0
            if max_score < 20:
                verified_category = self._verify_with_web_search(text, initial_category)
                if verified_category and verified_category != initial_category:
                    return verified_category

        return initial_category

    def get_confidence(self):
        """Return the confidence score (max score) of the last classification."""
        if not self.category_scores:
            return 0
        return max(self.category_scores.values())
    
    def _check_brand_categories(self, normalized_text):
        """
        v7.0: Check for known brand names. Returns list of possible categories, or None.
        """
        sorted_brands = sorted(self.BRAND_CATEGORIES.keys(), key=len, reverse=True)

        for brand in sorted_brands:
            if brand in normalized_text:
                return self.BRAND_CATEGORIES[brand]

        return None
    
    def _verify_with_web_search(self, description, initial_category):
        """
        Verify classification by searching product on the internet using raw HTML.
        Returns verified category or initial if verification fails.
        """
        if not ONLINE_LOOKUPS:
            return initial_category
        try:
            import urllib.request
            import urllib.parse
            import re
            
            # Busqueda de la descripcion usando duckduckgo HTML
            query = f"{description} categoria"
            encoded_query = urllib.parse.quote(query)
            
            url = f"https://html.duckduckgo.com/html/?q={encoded_query}"
            
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'})
            response = urllib.request.urlopen(req, timeout=8)
            html = response.read().decode('utf-8')
            
            search_text = re.sub(r'<[^>]+>', ' ', html)
            search_text = ' '.join(search_text.split()).lower()
            
            if not search_text.strip():
                return initial_category
            
            # Check which category has the most hints in the search result
            best_match = None
            best_count = 0
            
            for category, hints in self.WEB_VERIFICATION_HINTS.items():
                count = sum(1 for hint in hints if hint in search_text)
                if count > best_count:
                    best_count = count
                    best_match = category
            
            if best_match and best_count >= 1: # Reduced requirement as html scrape is dense but sometimes brief
                return best_match
            
            return initial_category
            
        except Exception:
            # If web search fails, return initial classification
            return initial_category
    
    def _calculate_scores(self, text):
        """Calculate weighted scores for each category."""
        scores = {cat: 0 for cat in CATEGORIES}
        
        for category, keywords in KEYWORD_DICT.items():
            score = 0
            
            # Multi-word phrases (highest weight)
            for phrase in keywords.get('multi_word', []):
                if normalize_for_matching(phrase) in text:
                    score += self.MULTI_WORD_WEIGHT
            
            # Primary keywords
            for keyword in keywords.get('primary', []):
                if normalize_for_matching(keyword) in text:
                    score += self.PRIMARY_WEIGHT
            
            # Secondary keywords
            for keyword in keywords.get('secondary', []):
                if normalize_for_matching(keyword) in text:
                    score += self.SECONDARY_WEIGHT
            
            # Exclusion penalties
            for exclusion in keywords.get('exclusions', []):
                if normalize_for_matching(exclusion) in text:
                    score += self.EXCLUSION_PENALTY
            
            scores[category] = score
        
        return scores
    
    def _get_best_category(self, scores):
        """Get the best category using scores and tie-break rules."""
        if not scores:
            return 'VARIOS'
        
        max_score = max(scores.values())
        
        if max_score <= 0:
            return 'VARIOS'
        
        tied_categories = [cat for cat, score in scores.items() if score == max_score]
        
        if len(tied_categories) == 1:
            return tied_categories[0]
        
        # Tie-break using priority
        for cat in CATEGORY_PRIORITY:
            if cat in tied_categories:
                return cat
        
        return tied_categories[0]
    
    def get_category_info(self, category):
        """Get category configuration. NEVER returns blank Section/Family/Subfamily."""
        info = CATEGORIES.get(category)
        if not info or not info.get('seccion'):
            info = CATEGORIES['VARIOS']
        return info


# ============================================================================
# PROVIDER DATABASE SCANNER
# ============================================================================

class ProviderDatabaseScanner:
    """Searches all provider tariff files for article references."""
    
    def __init__(self, database_path, logger=None):
        self.database_path = database_path
        self.logger = logger
    
    def log(self, message):
        """Log message if logger available."""
        if self.logger:
            self.logger(message)
        else:
            print(message)
    
    def find_reference(self, reference):
        """
        Search all provider tariff files for the reference.
        Returns: (provider_name, file_path, row_data, df) or None if not found
        """
        if not os.path.exists(self.database_path):
            raise Exception(f"Carpeta DATABASE no accesible: {self.database_path}")
        
        providers = self.get_all_providers()
        
        if not providers:
            raise Exception(f"No se encontraron carpetas de proveedores en: {self.database_path}")
        
        self.log(f"Buscando en {len(providers)} proveedores...")
        
        for provider in providers:
            tariff_files = self.get_tariff_file(provider)
            if tariff_files:
                self.log(f"  Buscando en {provider}...")
                for tariff_file in tariff_files:
                    result = self.search_in_tariff(tariff_file, reference)
                    if result:
                        row_data, df = result
                        self.log(f"  ✓ ¡Encontrado en {provider}! ({os.path.basename(tariff_file)})")
                        return (provider, tariff_file, row_data, df)
        
        return None
    
    def get_all_providers(self):
        """Get list of all provider folders."""
        try:
            return [d for d in os.listdir(self.database_path)
                    if os.path.isdir(os.path.join(self.database_path, d))
                    and not d.startswith('.')]
        except Exception as e:
            raise Exception(f"Error leyendo carpeta DATABASE: {str(e)}")
    
    def get_tariff_file(self, provider):
        """Get tariff file paths for a provider, sorted by year priority (newest first)."""
        provider_folder = os.path.join(self.database_path, provider)
        
        extensions = ['.xlsx', '.xls', '.ods']
        
        found = []
        
        # Collect all matching files
        try:
            for file in os.listdir(provider_folder):
                file_lower = file.lower()
                if file.startswith("~$"):
                    continue
                if not any(file_lower.endswith(ext) for ext in extensions):
                    continue
                if ('tarifa' in file_lower or provider.lower() in file_lower):
                    found.append(os.path.join(provider_folder, file))
        except:
            pass
        
        if not found:
            return None
        
        # Sort: files with year numbers (e.g. 2026) come FIRST
        found.sort(key=lambda x: (0 if re.search(r'\d{4}', os.path.basename(x)) else 1, x))
        
        return found
    
    def search_in_tariff(self, file_path, reference):
        """Search for reference in tariff file."""
        try:
            # Read file
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.ods':
                df = pd.read_excel(file_path, engine='odf')
            else:
                df = pd.read_excel(file_path)
            
            # Find header row
            df = self._find_header_row(df)
            
            if df.empty:
                return None
            
            reference_upper = str(reference).upper().strip()
            reference_clean = reference_upper.replace(' ', '').replace('-', '').replace('.', '')
            
            # Search in all columns
            for col in df.columns:
                try:
                    # Convert column to string for searching
                    col_values = df[col].astype(str).str.upper().str.strip()
                    
                    # Exact match
                    matches = df[col_values == reference_upper]
                    if not matches.empty:
                        return (matches.iloc[0].to_dict(), df)
                    
                    # Clean match (without spaces/dashes and leading zeros)
                    col_clean = col_values.str.replace(' ', '', regex=False).str.replace('-', '', regex=False).str.replace('.', '', regex=False).str.lstrip('0')
                    matches = df[col_clean == reference_clean.lstrip('0')]
                    if not matches.empty:
                        return (matches.iloc[0].to_dict(), df)
                    
                except Exception:
                    continue
            
            return None
            
        except Exception as e:
            self.log(f"    Error leyendo {file_path}: {str(e)}")
            return None
    
    def _find_header_row(self, df):
        """Find the actual header row in the data."""
        header_keywords = ['codigo', 'referencia', 'descripcion', 'precio', 'articulo',
                         'producto', 'ean', 'ref', 'tarifa', 'pvp', 'descripción',
                         'nombre', 'serie', 'coleccion', 'acabado']
        
        best_row = -1
        best_score = 0
        
        for i in range(min(15, len(df))):
            row = df.iloc[i]
            if row.notna().sum() < 3:
                continue
            
            score = 0
            for v in row:
                if pd.notna(v):
                    v_str = str(v).lower()
                    for kw in header_keywords:
                        if kw in v_str:
                            score += 1
                            break
            
            if score > best_score:
                best_score = score
                best_row = i
        
        if best_row >= 0 and best_score >= 2:
            new_df = df.iloc[best_row+1:].copy()
            new_df.columns = df.iloc[best_row].values
            new_df = new_df.dropna(how='all').reset_index(drop=True)
            return new_df
        
        return df


# ============================================================================
# SINGLE ROW CONVERTER
# ============================================================================

class SingleRowConverter:
    """Converts a single row to ERP import format using v6.0 logic."""
    
    def __init__(self):
        self.classifier = ProductClassifier()
    
    def convert_row(self, row_data, df, provider_name, user_inputs):
        """
        Convert a single row to import template format.

        Args:
            row_data: Dictionary of row data from tariff
            df: Full DataFrame (for column detection)
            provider_name: Name of provider
            user_inputs: Dict with provider_number, pro_discount, private_discount

        Returns:
            DataFrame with single row in template format
        """
        # Particularidades del proveedor (ver PROVIDER_PROFILES); {} si no tiene
        profile = provider_profile(provider_name)

        # Detect columns - more comprehensive patterns
        code_col = self._find_column(df, ['referencia', 'ref\\.?\\s*prov', 'ref(?!.*ean)', '^codigo$', 'cod(?!.*ean)', 'sku'])
        name_col = self._find_column(df, ['descripci.n', 'descripcion', 'nombre', 'producto', 'articulo', 'artículo', 'detalle'])
        price_col = self._find_column(df, ['precio.*compra', 'precio.*neto', 'precio.*base',
                                           'p\\.?v\\.?p', '^precio$', 'precio', 'tarifa', 'coste', '^neto$'])
        ean_col = self._find_column(df, ['codigo.*ean', 'ean.*13', 'ean', 'gtin', 'barcode', 'cod.*barras'])
        collection_col = self._find_column(df, ['coleccion', 'colección', 'serie', 'linea', 'gama'])
        # Phase 2: Detect provider category column separately (familia/categoria/tipo)
        provider_cat_col = self._find_column(df, ['categoria', 'categoría', 'familia', 'tipo.*producto', 'grupo.*producto'])
        finish_col = self._find_column(df, ['acabado', 'terminacion', 'finish', 'tono', 'color(?!.*ean)'])
        unit_col = self._find_column(df, ['^u\\.?m\\.?$', 'ud.*medida', 'unidad', 'medida'])
        weight_col = self._find_column(df, ['peso', 'kg'])

        # El perfil del proveedor puede fijar las columnas de codigo / nombre cuando el patron generico se equivoca
        column_patterns = profile.get('column_patterns', {})
        if column_patterns.get('code'):
            code_col = self._find_column(df, column_patterns['code']) or code_col
        if column_patterns.get('name'):
            name_col = self._find_column(df, column_patterns['name']) or name_col

        # Extract values
        codigo = self._get_value(row_data, code_col) or self._get_first_value(row_data)
        # Normaliza codigos numericos leidos como float por pandas (columna con huecos
        # -> float64, p.ej. '123456789012.0'): los deja como '123456789012'
        # para que A (Codigo), J (Ref.Prov.) y "REF: ..." no arrastren el '.0'.
        if isinstance(codigo, float):
            codigo = str(int(codigo)) if codigo.is_integer() else str(codigo)

        nombre = self._get_value(row_data, name_col) or 'PRODUCTO'
        collection = self._get_value(row_data, collection_col) or ''
        finish = self._get_value(row_data, finish_col) or ''

        # Anotaciones entre parentesis del nombre (marca blanca de cliente, notas de embalaje o corte...):
        # no interesan en el ERP
        if profile.get('strip_parentheses'):
            nombre = re.sub(r'\([^)]*\)', '', str(nombre))
            nombre = re.sub(r'\s+', ' ', nombre).strip()

        # Phase 2: Extract provider category from their Excel column
        provider_category = self._get_value(row_data, provider_cat_col) if provider_cat_col else None

        # Get price
        precio_compra = 0.0
        if price_col:
            try:
                precio_compra = float(self._get_value(row_data, price_col) or 0)
            except (ValueError, TypeError):
                pass

        # Precios en formato europeo con coma decimal y, en alguna fila, punto de millar (ej. "69,64" o
        # "1.108,94"): el float() nativo no admite la coma y el bloque de arriba los deja silenciosamente
        # en 0.0 (ValueError capturado).
        if profile.get('price_decimal_comma') and price_col:
            _raw_price = self._get_value(row_data, price_col)
            if _raw_price is not None:
                _s = str(_raw_price).strip()
                if ',' in _s:
                    _s = _s.replace('.', '').replace(',', '.')
                try:
                    precio_compra = float(_s)
                except (ValueError, TypeError):
                    precio_compra = 0.0

        # El "Precio" de la tarifa no es el Precio Compra real: se multiplica por el factor del perfil. A partir
        # de ahi se aplica igual la formula generica de ceramica.
        if profile.get('price_factor'):
            precio_compra = round(precio_compra * profile['price_factor'], 6)

        # Si el precio de la tarifa es el de la CAJA y el ERP lo lleva por PIEZA (PZ): precio / piezas por caja.
        _piezas_caja = _ceramic_price_divisor(row_data, provider_name)
        if _piezas_caja and precio_compra:
            precio_compra = round(precio_compra / _piezas_caja, 6)

        # Classify product using v7.0 classifier with provider category support
        safe_provider = re.sub(r'\bCERAMICAS?\b', '', str(provider_name).upper()).strip()
        description_text = f"{safe_provider} {collection} {nombre} {finish}".strip()

        category = self.classifier.classify_product(
            description_text, verify_with_web=bool(profile.get('verify_with_web')),
            provider_category=provider_category)

        # Proveedores de ceramica "pura": se salta el clasificador por palabras clave
        if provider_name and any(p in provider_name.upper() for p in FORCED_CERAMIC_PROVIDERS):
            category = "CERÁMICA"

        # Categoria fijada por el perfil: Excel por codigo (verdad absoluta) o categoria fija
        if profile.get('category_lookup'):
            category = category_for_code(profile['category_lookup'], codigo, default=category)
        if profile.get('force_category'):
            category = profile['force_category']

        cat_info = self.classifier.get_category_info(category)

        # Build descriptions
        nombre_clean = clean_text_for_erp(nombre)
        collection_clean = clean_text_for_erp(collection)
        finish_clean = clean_text_for_erp(finish)

        desc_parts = []
        if collection_clean:
            desc_parts.append(collection_clean)
        desc_parts.append(nombre_clean if nombre_clean else 'PRODUCTO')
        if finish_clean:
            desc_parts.append(finish_clean)
        descripcion = ' '.join(desc_parts) + f" REF: {codigo}"

        desc_corta_parts = []
        if collection_clean:
            desc_corta_parts.append(collection_clean)
        desc_corta_parts.append(nombre_clean if nombre_clean else 'PRODUCTO')
        if finish_clean:
            desc_corta_parts.append(finish_clean)
        desc_corta = ' '.join(desc_corta_parts)

        # Quitar palabras repetidas (sing/plural, sin acentos) en Descripcion y Descripcion Corta: el nombre de
        # serie a veces YA esta dentro de la propia descripcion del proveedor y a veces no.
        if profile.get('dedup_words'):
            descripcion = _dedup_repeated_words(' '.join(desc_parts)) + f" REF: {codigo}"
            desc_corta = _dedup_repeated_words(' '.join(desc_corta_parts))

        # Check for discount to set Calculo Base
        has_discount = bool(profile.get('price_as_is'))

        for k, v in (row_data or {}).items():
            k_str = str(k).strip().upper().replace(" ", "")
            if k_str in ("%DTO1", "%DTO2", "%DTO3", "%DTO4", "DTO1", "DTO2", "DTO3", "DTO4", "DESCUENTO", "DTO", "%DTO", "CALCULOBASE"):
                val = str(v).strip()
                if val and val.lower() != "nan" and val not in ("0", "0.0", "0,0"):
                    if k_str == "CALCULOBASE" and val != "B":
                        continue
                    has_discount = True
                    break

        cat_calculo_base_default = cat_info.get('calculo_base', 'A')
        calculo_base = 'B' if has_discount else cat_calculo_base_default

        # Calculate prices
        if cat_info.get('use_ceramic_formula', False) and not has_discount:
            precio_base_tarifa = (precio_compra * CERAMIC_MULT) + CERAMIC_ADD
        else:
            precio_base_tarifa = precio_compra

        # Descuentos de Tarifa 2 (profesional) y Tarifa 3 (particular)
        final_pro_discount = user_inputs['pro_discount']
        final_private_discount = user_inputs['private_discount']

        if category == "GRIFERÍA":
            final_pro_discount = GRIFERIA_PRO_DISCOUNT
            final_private_discount = GRIFERIA_PRIVATE_DISCOUNT

        # El perfil del proveedor manda sobre la regla por categoria
        if profile.get('pro_discount') is not None:
            final_pro_discount = profile['pro_discount']
        if profile.get('private_discount') is not None:
            final_private_discount = profile['private_discount']

        tarifa_1 = precio_base_tarifa
        tarifa_2 = round(precio_base_tarifa * (1 - final_pro_discount / 100.0), 6)
        tarifa_3 = round(precio_base_tarifa * (1 - final_private_discount / 100.0), 6)

        # Get EAN
        ean = ''
        if ean_col:
            ean_val = self._get_value(row_data, ean_col) or ''
            ean_str = str(ean_val).replace('.0', '').strip()
            if re.match(r'^\d{13}$', ean_str):
                ean = ean_str
        if not ean:
            ean = generate_ean13(codigo)

        # Get unit
        ud_medida = cat_info['unit']
        if unit_col:
            unit_val = str(self._get_value(row_data, unit_col) or '').upper().strip()
            if unit_val in ['M2', 'M²', 'UD', 'UN', 'ML', 'KG', 'L']:
                ud_medida = 'M2' if unit_val in ['M2', 'M²'] else unit_val

        # Override to ensure non-ceramic products always use 'UD' instead of 'KG', 'L', etc.
        if not str(category).upper().startswith("CER"):
            ud_medida = "UD"

        # Get weight
        peso = ''
        if weight_col:
            try:
                peso = float(self._get_value(row_data, weight_col) or '')
            except (ValueError, TypeError):
                pass

        calculo_base = cat_info['calculo_base']

        # Determine %Dto1 value, Calculo Base and provider number (perfil del proveedor)
        dto1_val = ''
        if profile.get('dto1') is not None:
            dto1_val = profile['dto1']
        if profile.get('calculo_base'):
            calculo_base = profile['calculo_base']
        if profile.get('provider_number') is not None:
            user_inputs['provider_number'] = profile['provider_number']  # G (Fabricante) e I (Proveedor)

        # Columna I (Proveedor): por defecto = columna G (Fabricante). Un perfil puede dar de alta el
        # proveedor con otro numero SOLO en la columna I, dejando la G como esta.
        proveedor_col_i = user_inputs['provider_number']
        if profile.get('supplier_column_number') is not None:
            proveedor_col_i = profile['supplier_column_number']

        # Build ERP row
        erp_row = [
            codigo,                          # A: Codigo
            descripcion,                     # B: Descripcion
            cat_info['familia'],             # C: Familia
            ud_medida,                       # D: Ud.Medida
            cat_info['subfamilia'],          # E: Subfamilia
            cat_info['seccion'],             # F: Sección
            user_inputs['provider_number'],  # G: Fabricante
            ean,                             # H: Ean
            proveedor_col_i,                 # I: Proveedor
            codigo,                          # J: Ref.Prov.
            precio_compra,                   # K: Precio Compra
            dto1_val,                        # L: %Dto1
            '',                              # M: %Dto2
            '',                              # N: %Dto3
            '',                              # O: %Dto4
            '',                              # P: %Rappel
            precio_base_tarifa,              # Q: Precio Base Tarifa
            '',                              # R: %Margen
            calculo_base,                    # S: Calculo Base
            '',                              # T: Incr.Tarifa
            tarifa_1,                        # U: Tarifa 1
            tarifa_2,                        # V: Tarifa 2
            tarifa_3,                        # W: Tarifa 3
            '',                              # X: Tarifa 4
            '',                              # Y: Tarifa 5
            '',                              # Z: Tarifa 6
            '',                              # AA: Tarifa 7
            '',                              # AB: Tarifa 8
            '',                              # AC: Tarifa 9
            '',                              # AD: Tarifa 10
            ud_medida,                       # AE: Ud.Med.Compra
            1 if calculo_base == 'A' else '', # AF: Factor Conver
            '' if calculo_base == 'A' else 1,  # AG: UdxEnvase (1 para no-cerámica; ceramica en B: sin dato, en blanco)
            '' if calculo_base == 'A' else 1,  # AH: Unid.Minima (1 para no-cerámica)
            peso,                            # AI: Peso
            '',                              # AJ: Volumen
            'N',                             # AK: Ctrl.Stock S/N
            '',                              # AL: Lote S/N
            21,                              # AM: %IVA
            '',                              # AN: Ud.Min.Compra
            'S' if calculo_base == 'A' else 'N',  # AO: Aplicar palet (ceramica en B sigue paletizada)
            'N',                             # AP: Apli.P.Coste S/N
            'N',                             # AQ: Aplicar Dto/Rec S/N
            'N',                             # AR: Recargo/Dto (R/D)
            '',                              # AS: % Rec/Dto
            'S' if calculo_base == 'A' else 'N',  # AT: Paletizado S/N
            '',                              # AU: Ud.Palet
            '',                              # AV: Grupo Precios
            '',                              # AW: Agencia Trans
            '',                              # AX: Factor Agt
            '',                              # AY: Clave libre
            '',                              # AZ: Stock Minimo
            desc_corta,                      # BA: Descripcion Corta
            '',                              # BB: Multiplo Com
            '',                              # BC: Multiplo Ven
            '',                              # BD: Largo
            '',                              # BE: Ancho
            '',                              # BF: Profundo
            'N',                             # BG: Kit (S/N)
            '',                              # BH: Tar.Man(S/N)
            '',                              # BI: Peso Palet
        ]

        # Create DataFrame
        result_df = pd.DataFrame([erp_row], columns=ERP_COLUMNS)

        return result_df, category, cat_info

    def _find_column(self, df, patterns):
        """Find column matching patterns."""
        for pattern in patterns:
            for col in df.columns:
                col_lower = str(col).lower().strip()
                # Handle accented characters
                col_normalized = normalize_for_matching(col_lower)
                try:
                    if re.search(pattern, col_lower) or re.search(pattern, col_normalized):
                        return col
                except:
                    if pattern in col_lower or pattern in col_normalized:
                        return col
        return None
    
    def _get_value(self, row_data, col):
        """Get value from row data for a column."""
        if not col:
            return None
        
        # Try exact match
        if col in row_data:
            val = row_data[col]
            if pd.notna(val):
                return val
        
        # Try case-insensitive match
        col_lower = str(col).lower()
        for key, val in row_data.items():
            if str(key).lower() == col_lower and pd.notna(val):
                return val
        
        return None
    
    def _get_first_value(self, row_data):
        """Get first non-null value from row."""
        for key, val in row_data.items():
            if pd.notna(val) and str(val).strip():
                return str(val).strip()
        return 'UNKNOWN'


# ============================================================================
# NETWORK FILE MANAGER
# ============================================================================

class NetworkFileManager:
    """Manages saving import files to network folder."""
    
    def __init__(self, network_path, filename='importart'):
        self.network_path = network_path
        self.filename = filename
    
    def save_import_file(self, df):
        """Save DataFrame to network folder as 'importart'."""
        try:
            # Check if network path exists
            if not os.path.exists(self.network_path):
                # Try to create it (for local testing)
                try:
                    os.makedirs(self.network_path, exist_ok=True)
                except:
                    raise Exception(f"Carpeta de red no accesible: {self.network_path}")
            
            # Full file path
            file_path = os.path.join(self.network_path, f"{self.filename}.xlsx")
            
            # Save Excel file
            df.to_excel(file_path, index=False, engine='openpyxl')
            
            return file_path
            
        except Exception as e:
            raise Exception(f"Error guardando archivo en red: {str(e)}")


# ============================================================================
# FIREFOX ERP AUTOMATOR v2.1 - Robust Element Finding
# ============================================================================

class FirefoxERPAutomator:
    """Automates ERP import process in Firefox with robust element finding."""
    
    # Login credentials (placeholders: se leen de variables de entorno)
    ERP_USERNAME = os.environ.get("ERP_USER", "x")
    ERP_PASSWORD = os.environ.get("ERP_PASSWORD", "y")
    
    def __init__(self, erp_url, logger=None):
        self.erp_url = erp_url
        self.logger = logger
        self.last_run_file = os.path.join(os.path.dirname(__file__), 'last_run_date.txt')
        self.screenshot_dir = os.path.dirname(__file__)
    
    def log(self, message):
        """Log message."""
        if self.logger:
            self.logger(message)
        else:
            print(message)
    

    # ----------------------------
    # Download detection helpers
    # ----------------------------
    def _get_download_dir(self):
        """Return the Downloads folder used by the ERP downloads."""
        userprofile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        candidates = [
            os.path.join(userprofile, "Downloads"),
            os.path.join(userprofile, "Descargas"),
        ]
        for c in candidates:
            try:
                if c and os.path.isdir(c):
                    return c
            except Exception:
                pass
        return userprofile

    def _snapshot_downloads(self, download_dir):
        """Snapshot current files in download dir: returns dict(name->mtime) and max_mtime."""
        mt = {}
        max_m = 0.0
        try:
            for name in os.listdir(download_dir):
                p = os.path.join(download_dir, name)
                try:
                    if os.path.isfile(p):
                        m = os.path.getmtime(p)
                        mt[name] = m
                        if m > max_m:
                            max_m = m
                except Exception:
                    continue
        except Exception:
            pass
        return mt, max_m

    def _wait_for_new_download(self, before_map, before_max_mtime, timeout=60):
        """Wait for a new/updated file to appear in Downloads.

        Nota: en Firefox el fichero puede aparecer primero como *.part o con extensiones
        inesperadas. Para dar el proceso por *éxito*, aquí consideramos descarga *cualquier*
        archivo nuevo o modificado cuyo mtime sea posterior al snapshot previo.

        Devuelve el nombre del archivo detectado, o None si no se detecta nada dentro del timeout.
        """
        download_dir = self._get_download_dir()
        end = time.time() + timeout

        # Aceptar cualquier extensión (incluye .part / .tmp / etc.).
        # Solo excluimos algunos archivos de ruido que pueden crearse en Downloads.
        ignore_prefixes = ("debug_", "~$")

        while time.time() < end:
            try:
                for name in os.listdir(download_dir):
                    p = os.path.join(download_dir, name)
                    if not os.path.isfile(p):
                        continue
                    try:
                        m = os.path.getmtime(p)
                    except Exception:
                        continue

                    if name.startswith(ignore_prefixes):
                        continue

                    if name not in before_map or m > max(before_map.get(name, 0.0), before_max_mtime):
                        return name
            except Exception:
                pass
            time.sleep(0.4)

        return None

    def is_first_run_today(self):
        """Check if this is the first run today (LOCAL_TIMEZONE)."""
        try:
            today = get_local_date()
            if os.path.exists(self.last_run_file):
                with open(self.last_run_file, 'r') as f:
                    last_run = f.read().strip()
                if last_run == str(today):
                    return False
        except:
            pass
        
        return True
    
    def update_last_run_date(self):
        """Update last run date to today."""
        try:
            today = get_local_date()
            with open(self.last_run_file, 'w') as f:
                f.write(str(today))
        except:
            pass
    
    def find_element_robust(self, driver, element_name, strategies, timeout=5, clickable=True):
        """
        v3.0: Try multiple strategies to find an element with REDUCED TIMEOUT.
        
        Args:
            driver: Selenium WebDriver
            element_name: Name of element for logging
            strategies: List of (By, selector) tuples to try
            timeout: Max time to wait per strategy (v3.0: reduced from 15s to 5s)
            clickable: If True, wait for element_to_be_clickable; else presence_of_element_located
        
        Returns:
            WebElement if found, None otherwise
        """
        condition = EC.element_to_be_clickable if clickable else EC.presence_of_element_located
        
        for by, selector in strategies:
            try:
                element = WebDriverWait(driver, timeout).until(
                    condition((by, selector))
                )
                self.log(f"  ✓ Encontrado '{element_name}' usando {by}='{selector}'")
                return element
            except Exception:
                # v3.0: Reduced logging for failed selectors (less noise)
                continue
        
        self.log(f"  ✗ FALLO: No se pudo encontrar '{element_name}' con ninguna estrategia")
        return None
    

    def _popup_import_visible(self, driver):
        """Return True only if the *import confirmation* popup seems visible.

        Nota: evitamos falsos positivos (en este ERP siempre hay texto 'Aceptar/Cancelar' en la barra).
        Solo consideramos popup si aparecen textos característicos del diálogo de importación.
        """
        needles = ['importacion articulos','importación articulos','usted va a importar','se recomienda realizar una copia']
        js = r"""
            function isVisible(el){
                if(!el) return false;
                const r = el.getBoundingClientRect();
                const s = window.getComputedStyle(el);
                return r.width>2 && r.height>2 && s.visibility!=='hidden' && s.display!=='none' && (parseFloat(s.opacity||'1')>0.05);
            }
            const needles = arguments[0] || [];
            const nodes = document.querySelectorAll('div, span, td, th, p, label');
            for (const n of nodes){
                const t = (n.innerText||'').toLowerCase();
                if(!t) continue;
                if(needles.some(k => t.includes(k))){
                    // buscar un contenedor visible cercano
                    let cur = n;
                    for (let i=0;i<10 && cur;i++){
                        if(isVisible(cur)) return true;
                        cur = cur.parentElement;
                    }
                    return true;
                }
            }
            return false;
        """
        try:
            return bool(driver.execute_script(js, needles))
        except Exception:
            return False

    def click_aceptar_import_popup(self, driver, timeout=20):
        """
        Click robusto del botón **Aceptar** del popup de importación.
    
        Problema típico en ERP basados en dhtmlx: el botón puede estar dentro de iframes, ser un div clicable,
        o estar cubierto por overlays. Aquí NO damos el click por bueno hasta verificar que el popup se cerró.
        """
        self.log("Buscando popup de importación visible...")
    
        # 1) Espera a que el popup exista (en cualquier frame posible lo detectaremos después)
        end_time = time.time() + timeout
        while time.time() < end_time:
            try:
                if self._popup_import_visible(driver):
                    break
            except Exception:
                pass
            time.sleep(0.25)
    
        if not self._popup_import_visible(driver):
            self.log("  ⚠ No se detectó popup visible (puede que no aparezca en este flujo)")
            return False
    
        def _verify_popup_gone(max_wait=6):
            t0 = time.time()
            while time.time() - t0 < max_wait:
                try:
                    if not self._popup_import_visible(driver):
                        return True
                except Exception:
                    pass
                time.sleep(0.25)
            return False
    
        def _try_click_candidates_in_context():
            """Intenta encontrar/clicar Aceptar en el DOM actual (frame actual)."""
            # Estrategia 1: XPath flexible (button/input/a/div/span) con contains y normalize-space
            xpaths = [
                # botón/inputs típicos
                "//button[normalize-space(.)='Aceptar' or contains(translate(normalize-space(.),'ACEPTAR','aceptar'),'aceptar')]" ,
                "//input[(translate(@value,'ACEPTAR','aceptar')='aceptar') or contains(translate(@value,'ACEPTAR','aceptar'),'aceptar')]" ,
                # dhtmlx a veces usa div/span
                "//*[self::div or self::span or self::a][contains(translate(normalize-space(.),'ACEPTAR','aceptar'),'aceptar')]" ,
            ]
    
            candidates = []
            for xp in xpaths:
                try:
                    found = driver.find_elements(By.XPATH, xp)
                    if found:
                        candidates.extend(found)
                except Exception:
                    continue
    
            # Estrategia 2: JS - localizar el contenedor del popup por el título y buscar dentro
            try:
                el = driver.execute_script(r'''
                    function isVisible(el){
                        if(!el) return false;
                        const r = el.getBoundingClientRect();
                        const s = window.getComputedStyle(el);
                        return r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden' && (parseFloat(s.opacity||'1')>0.05);
                    }
                    function norm(t){ return (t||'').replace(/\s+/g,' ').trim().toLowerCase(); }
    
                    // intenta localizar el popup por el título
                    const titleNeedle = 'importacion articulos';
                    let popupRoot = null;
    
                    // candidatos a título (muchas veces es un td/div en header)
                    const titleEls = Array.from(document.querySelectorAll('div,span,td,b,strong'));
                    for(const te of titleEls){
                        const t = norm(te.innerText);
                        if(t === titleNeedle || t.includes(titleNeedle)){
                            // sube a un contenedor "modal-like"
                            let p = te;
                            for(let i=0;i<8 && p;i++){
                                const cls = (p.className||'').toString();
                                if(/dhx|dhtmlx|win|modal|window|dialog/i.test(cls)){
                                    popupRoot = p;
                                    break;
                                }
                                p = p.parentElement;
                            }
                            if(!popupRoot) popupRoot = te.parentElement;
                            break;
                        }
                    }
                    if(!popupRoot) popupRoot = document.body;
    
                    const elems = Array.from(popupRoot.querySelectorAll('button,input[type=button],input[type=submit],a,div,span'));
                    const good = [];
                    for(const e of elems){
                        const t = norm((e.value||'') + ' ' + (e.innerText||''));
                        if(t.includes('aceptar') && isVisible(e)){
                            // preferimos el elemento que esté realmente arriba en el punto central
                            const r = e.getBoundingClientRect();
                            const x = r.left + r.width/2;
                            const y = r.top + r.height/2;
                            const top = document.elementFromPoint(x,y);
                            if(top && (top === e || e.contains(top) || top.contains(e))) good.push(e);
                        }
                    }
                    // devuelve el último (suele ser el más “arriba”)
                    return good.length ? good[good.length-1] : null;
                ''')
                if el:
                    candidates.append(el)
            except Exception:
                pass

            # Estrategia 3 (CRÍTICA): en muchos dhtmlx el texto "Aceptar" está en un nodo de texto
            # (p.ej. dentro de una celda), pero el elemento clicable REAL no contiene ese texto.
            # En ese caso: buscamos el *texto* "Aceptar", sacamos su rectángulo (Range.getBoundingClientRect),
            # y clicamos el elemento REAL que haya en ese punto con elementFromPoint + MouseEvents.
            try:
                ok = driver.execute_script(r'''
                    (function(){
                      function norm(s){
                        return (s||'').toString().replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim().toLowerCase();
                      }
                      function isVisibleRect(r){
                        return r && r.width>3 && r.height>3 && r.bottom>0 && r.right>0 && r.top < window.innerHeight && r.left < window.innerWidth;
                      }
                      function fireAt(x,y){
                        const el = document.elementFromPoint(x,y);
                        if(!el) return {clicked:false, why:'no_elementFromPoint'};
                        const evOpts = {bubbles:true, cancelable:true, clientX:x, clientY:y};
                        el.dispatchEvent(new MouseEvent('mousemove', evOpts));
                        el.dispatchEvent(new MouseEvent('mouseover', evOpts));
                        el.dispatchEvent(new MouseEvent('mousedown', evOpts));
                        el.dispatchEvent(new MouseEvent('mouseup', evOpts));
                        el.dispatchEvent(new MouseEvent('click', evOpts));
                        return {clicked:true, tag: el.tagName, cls: (el.className||'').toString().slice(0,120)};
                      }

                      // localizar nodos de texto que contengan "Aceptar"
                      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null);
                      const hits = [];
                      while(walker.nextNode()){
                        const n = walker.currentNode;
                        const t = norm(n.nodeValue);
                        if(!t) continue;
                        if(t === 'aceptar' || t.includes('aceptar')){
                          try{
                            const r = document.createRange();
                            r.selectNodeContents(n);
                            const rect = r.getBoundingClientRect();
                            if(isVisibleRect(rect)){
                              hits.push({rect, txt: t});
                            }
                          }catch(e){}
                        }
                      }
                      if(!hits.length) return {clicked:false, why:'no_text_hits'};

                      // Elegimos el "Aceptar" más abajo (normalmente botones abajo)
                      hits.sort((a,b)=> (b.rect.top - a.rect.top) || (b.rect.left - a.rect.left));
                      const target = hits[0].rect;
                      const x = target.left + target.width/2;
                      const y = target.top + target.height/2;
                      return fireAt(x,y);
                    })();
                ''')
                if ok and ok.get('clicked'):
                    return True
            except Exception:
                pass
    
            # Deduplicar (por id de web element)
            uniq = []
            seen = set()
            for c in candidates:
                try:
                    k = getattr(c, 'id', None) or str(c)
                except Exception:
                    k = str(c)
                if k in seen:
                    continue
                seen.add(k)
                uniq.append(c)
    
            if not uniq:
                return False
    
            # Intentos de click: normal → ActionChains → JS MouseEvent
            for attempt, el in enumerate(uniq[:8], start=1):
                try:
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", el)
                    except Exception:
                        pass
                    time.sleep(0.15)
    
                    clicked = False
                    # 1) click directo
                    try:
                        el.click()
                        clicked = True
                    except Exception:
                        pass
    
                    # 2) ActionChains (mueve+click)
                    if not clicked:
                        try:
                            ActionChains(driver).move_to_element(el).pause(0.05).click(el).perform()
                            clicked = True
                        except Exception:
                            pass
    
                    # 3) JS MouseEvent “hard click”
                    if not clicked:
                        try:
                            driver.execute_script(r'''
                                const el = arguments[0];
                                const r = el.getBoundingClientRect();
                                const x = r.left + r.width/2;
                                const y = r.top + r.height/2;
                                function fire(target, type){
                                    target.dispatchEvent(new MouseEvent(type,{bubbles:true,cancelable:true,view:window,clientX:x,clientY:y}));
                                }
                                fire(el,'mouseover'); fire(el,'mousemove'); fire(el,'mousedown'); fire(el,'mouseup'); fire(el,'click');
                            ''', el)
                            clicked = True
                        except Exception:
                            pass
    
                    if clicked and _verify_popup_gone():
                        return True
                except Exception:
                    continue
    
            # Último recurso dentro del contexto actual: click por coordenadas usando elementFromPoint
            try:
                ok = driver.execute_script(r'''
                    function norm(t){ return (t||'').replace(/\s+/g,' ').trim().toLowerCase(); }
                    function isVisible(el){
                        if(!el) return false;
                        const r = el.getBoundingClientRect();
                        const s = window.getComputedStyle(el);
                        return r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden';
                    }
                    // busca cualquier nodo que muestre “Aceptar”
                    const nodes = Array.from(document.querySelectorAll('button,input[type=button],input[type=submit],a,div,span'));
                    let target = null;
                    for(const n of nodes){
                        const t = norm((n.value||'') + ' ' + (n.innerText||''));
                        if(t === 'aceptar' || t.endsWith(' aceptar') || t.includes('aceptar')){
                            if(isVisible(n)){ target = n; }
                        }
                    }
                    if(!target) return false;
                    const r = target.getBoundingClientRect();
                    const x = r.left + r.width/2;
                    const y = r.top + r.height/2;
                    const top = document.elementFromPoint(x,y);
                    if(!top) return false;
                    // dispara eventos en el elemento realmente superior
                    function fire(type){
                        top.dispatchEvent(new MouseEvent(type,{bubbles:true,cancelable:true,view:window,clientX:x,clientY:y}));
                    }
                    fire('mouseover'); fire('mousemove'); fire('mousedown'); fire('mouseup'); fire('click');
                    return true;
                ''')
                if ok and _verify_popup_gone():
                    return True
            except Exception:
                pass
    
            return False
    
        # 2) Probar en default_content y en TODOS los iframes anidados (profundidad limitada)
        def _walk_frames_and_try(depth=0, max_depth=4):
            if _try_click_candidates_in_context():
                return True
            if depth >= max_depth:
                return False
            try:
                frames = driver.find_elements(By.TAG_NAME, "iframe")
            except Exception:
                frames = []
            for i in range(min(len(frames), 25)):
                try:
                    driver.switch_to.frame(frames[i])
                    if _walk_frames_and_try(depth+1, max_depth):
                        return True
                except Exception:
                    pass
                finally:
                    try:
                        driver.switch_to.parent_frame()
                    except Exception:
                        driver.switch_to.default_content()
            return False
    
        driver.switch_to.default_content()
        if _walk_frames_and_try():
            driver.switch_to.default_content()
            return True
    
        driver.switch_to.default_content()
    
        # 3) Último último recurso: teclado
        self.log("  ⚠ Fallback: intentando ENTER sobre el popup...")
        try:
            ActionChains(driver).send_keys(Keys.ENTER).perform()
        except Exception:
            pass
        if _verify_popup_gone(max_wait=4):
            return True
    
        return False

    # --------------------------------------------------------------------
    # Generic "Aceptar" popup closer (used for 2nd result popup)
    # --------------------------------------------------------------------
    def _any_visible_aceptar(self, driver):
        """True if there is any visible element that looks like an 'Aceptar' control."""
        try:
            return bool(driver.execute_script(r"""
                function isVisible(el){
                    if(!el) return false;
                    const r = el.getBoundingClientRect();
                    const s = window.getComputedStyle(el);
                    return r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden' && (parseFloat(s.opacity||'1')>0.05);
                }
                function txt(el){
                    return ((el.value||'') + ' ' + (el.innerText||'') + ' ' + (el.textContent||'')).replace(/\s+/g,' ').trim();
                }
                const sel = 'button, input[type="button"], input[type="submit"], a, div, span';
                const els = Array.from(document.querySelectorAll(sel));
                for(const el of els){
                    const t = txt(el);
                    if(!t) continue;
                    const tl = t.toLowerCase();
                    if(tl.includes('aceptar')){
                        if(isVisible(el)) return true;
                    }
                }
                return false;
            """))
        except Exception:
            return False

    def _click_one_visible_aceptar(self, driver):
        """Clicks the top-most visible 'Aceptar' using coordinate-based MouseEvents. Returns True if a click was dispatched."""
        try:
            return bool(driver.execute_script(r"""
                function isVisible(el){
                    if(!el) return false;
                    const r = el.getBoundingClientRect();
                    const s = window.getComputedStyle(el);
                    return r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden' && (parseFloat(s.opacity||'1')>0.05);
                }
                function txt(el){
                    return ((el.value||'') + ' ' + (el.innerText||'') + ' ' + (el.textContent||'')).replace(/\s+/g,' ').trim();
                }
                function getZ(el){
                    let cur = el;
                    let best = -1;
                    for(let i=0;i<10 && cur;i++){
                        const z = window.getComputedStyle(cur).zIndex;
                        const zi = (z==='auto') ? -1 : parseInt(z,10);
                        if(!isNaN(zi)) best = Math.max(best, zi);
                        cur = cur.parentElement;
                    }
                    return best;
                }
                const sel = 'button, input[type="button"], input[type="submit"], a, div, span';
                const els = Array.from(document.querySelectorAll(sel))
                    .filter(el => {
                        const t = txt(el);
                        const tl = t.toLowerCase();
                        if(!tl.includes('aceptar')) return false;
                        if(tl.includes('cancelar')) return false;
                        return isVisible(el);
                    })
                    .sort((a,b) => getZ(b) - getZ(a));
                if(!els.length) return false;
                const el = els[0];
                const r = el.getBoundingClientRect();
                const x = r.left + r.width/2;
                const y = r.top + r.height/2;
                const target = document.elementFromPoint(x, y) || el;
                ['mousemove','mouseover','mousedown','mouseup','click'].forEach(type => {
                    const evt = new MouseEvent(type, {bubbles:true, cancelable:true, view:window, clientX:x, clientY:y});
                    target.dispatchEvent(evt);
                });
                // also call native click as fallback
                try{ el.click(); }catch(e){}
                return true;
            """))
        except Exception:
            return False

    def close_all_aceptar_popups(self, driver, total_timeout=15):
        """Close any 'Aceptar' popups that appear (e.g., the result popup after import). Verified by disappearance."""
        start_time = time.time()
        end_time = start_time + total_timeout
        appear_timeout = min(8, total_timeout)  # segundos máximos esperando a que aparezca el popup
        closed_any = False

        # Try also inside iframes (some dhtmlx layouts nest frames)
        def all_frames():
            frames = driver.find_elements(By.TAG_NAME, 'iframe')
            return frames

        while time.time() < end_time:
            # Espera activa a que aparezca el popup de resultado.
            # OJO: en este ERP el popup puede tardar varios segundos en renderizarse,
            # y si salimos aquí demasiado pronto, el log dirá "no se detectó" aunque sí esté.
            driver.switch_to.default_content()
            if not self._any_visible_aceptar(driver):
                # Si aún no hemos cerrado nada, damos una "ventana" para que aparezca.
                if not closed_any and time.time() < (start_time + appear_timeout):
                    time.sleep(0.25)
                    continue
                return closed_any

            clicked = False

            # 1) Try in main document
            if self._click_one_visible_aceptar(driver):
                clicked = True

            # 2) Try in each iframe if still visible
            if not clicked or self._any_visible_aceptar(driver):
                for fr in all_frames():
                    try:
                        driver.switch_to.default_content()
                        driver.switch_to.frame(fr)
                        if self._any_visible_aceptar(driver):
                            if self._click_one_visible_aceptar(driver):
                                clicked = True
                    except Exception:
                        continue

            driver.switch_to.default_content()
            if clicked:
                closed_any = True
                # give UI time to close
                t0 = time.time()
                while time.time() - t0 < 4:
                    try:
                        driver.switch_to.default_content()
                        if not self._any_visible_aceptar(driver):
                            return True
                    except Exception:
                        pass
                    time.sleep(0.25)

            # If still there, try keyboard fallback
            try:
                ActionChains(driver).send_keys(Keys.ENTER).perform()
            except Exception:
                pass
            time.sleep(0.4)

        return closed_any




    
    # --------------------------------------------------------------------
    # Deep frame walker utilities (for popups inside nested iframes)
    # --------------------------------------------------------------------
    def _iter_frame_paths(self, driver, max_depth=6, max_frames_per_level=40):
        """
        Returns a list of iframe index-paths to reach nested frames.
        Path [] means default_content.
        """
        paths = [[]]

        def walk(cur_path, depth):
            if depth >= max_depth:
                return
            try:
                driver.switch_to.default_content()
                for idx in cur_path:
                    frs = driver.find_elements(By.TAG_NAME, "iframe")
                    if idx >= len(frs):
                        return
                    driver.switch_to.frame(frs[idx])
                frs = driver.find_elements(By.TAG_NAME, "iframe")
            except Exception:
                return

            for i in range(min(len(frs), max_frames_per_level)):
                p = cur_path + [i]
                paths.append(p)
                walk(p, depth + 1)

        walk([], 0)
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        return paths

    def _switch_to_frame_path(self, driver, path):
        """Switch to default_content then follow iframe indexes in path. Returns True on success."""
        try:
            driver.switch_to.default_content()
            for idx in path:
                frs = driver.find_elements(By.TAG_NAME, "iframe")
                if idx >= len(frs):
                    return False
                driver.switch_to.frame(frs[idx])
            return True
        except Exception:
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
            return False

    def _debug_acceptar_snapshot(self, driver, label="ACEPTAR"):
        """
        Compact debug snapshot: where we are, top visible dialogs, and aceptar-like candidates.
        Keep it short but very informative.
        """
        try:
            info = driver.execute_script(r"""
                function norm(s){
                    return (s||"").toString().replace(/\u00a0/g," ").replace(/\s+/g," ").trim();
                }
                function isVisible(el){
                    if(!el) return false;
                    const r = el.getBoundingClientRect();
                    const s = window.getComputedStyle(el);
                    if(r.width < 5 || r.height < 5) return false;
                    if(s.display === 'none' || s.visibility === 'hidden') return false;
                    if(parseFloat(s.opacity||'1') < 0.05) return false;
                    return true;
                }
                // dialog-like containers (best-effort)
                const dlgSel = '[role="dialog"], .dhx_popup, .dhx_win, .dhtmlx_window, .modal, .popup, .window, .dialog';
                const dialogs = Array.from(document.querySelectorAll(dlgSel)).filter(isVisible).slice(0,3).map(d=>{
                    const t = norm(d.innerText).slice(0,140);
                    const r = d.getBoundingClientRect();
                    return {text:t, rect:[Math.round(r.left),Math.round(r.top),Math.round(r.width),Math.round(r.height)]};
                });

                const candSel = 'button,input[type="button"],input[type="submit"],[role="button"],a,div,span';
                const cands = [];
                for(const el of Array.from(document.querySelectorAll(candSel))){
                    if(!isVisible(el)) continue;
                    const t = norm((el.value||"") + " " + (el.innerText||"") + " " + (el.textContent||""));
                    if(!t) continue;
                    const tl = t.toLowerCase();
                    if(tl.includes("aceptar") || tl === "ok" || tl.includes("insertados") || tl.includes("actualizados")){
                        const r = el.getBoundingClientRect();
                        const cx = r.left + r.width/2, cy = r.top + r.height/2;
                        const top = document.elementFromPoint(cx, cy);
                        cands.push({
                            text: t.slice(0,80),
                            tag: (el.tagName||"").toLowerCase(),
                            cls: (el.className||"").toString().slice(0,60),
                            rect: [Math.round(r.left),Math.round(r.top),Math.round(r.width),Math.round(r.height)],
                            topTag: top ? (top.tagName||"").toLowerCase() : null,
                            topCls: top ? (top.className||"").toString().slice(0,60) : null,
                            topMatch: top ? (top===el || el.contains(top) || top.contains(el)) : false
                        });
                        if(cands.length>=6) break;
                    }
                }
                return {dialogs, cands};
            """)
            self.log(f"  [{label}] Snapshot: dialogs={len(info.get('dialogs',[]))} cands={len(info.get('cands',[]))}")
            for i,d in enumerate(info.get("dialogs",[]), start=1):
                self.log(f"  [{label}] Dialog#{i} rect={d['rect']} text='{d['text']}'")
            for i,c in enumerate(info.get("cands",[]), start=1):
                self.log(f"  [{label}] Cand#{i} <{c['tag']}> rect={c['rect']} text='{c['text']}' top=<{c['topTag']}> match={c['topMatch']}")
        except Exception:
            pass

    # --------------------------------------------------------------------
    # Strict closer for the RESULT popup (2nd 'Aceptar' after import)
    # --------------------------------------------------------------------
    def close_result_popup_strict(self, driver, timeout=25):
        """
        Cierra el popup de RESULTADO tras la importación:
        "Articulos insertados: X  Actualizados: Y" + botones Aceptar/Cancelar.
        - Recorre TODOS los iframes anidados (profundo).
        - NO depende de que el botón sea <button>.
        - Verifica cierre: desaparece el texto de resultado en todos los frames.
        """
        RESULT_HINTS = ["articulos insertados", "artículos insertados", "actualizados"]
        ACCEPT_HINTS = ["aceptar", "ok", "accept"]

        def result_visible_anywhere():
            paths = self._iter_frame_paths(driver, max_depth=6)
            for p in paths:
                if not self._switch_to_frame_path(driver, p):
                    continue
                try:
                    txt = driver.execute_script("return document.body ? (document.body.innerText||'') : ''")
                    low = (txt or "").lower()
                    if any(h in low for h in RESULT_HINTS):
                        return True
                except Exception:
                    pass
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
            return False

        # wait for the result popup text to appear
        t0 = time.time()
        self.log("  [POPUP2] Esperando popup de resultado (texto insertados/actualizados)...")
        while time.time() - t0 < timeout:
            if result_visible_anywhere():
                break
            time.sleep(0.25)

        if not result_visible_anywhere():
            # Sometimes the ERP doesn't show it, so we treat as not needed
            self.log("  [POPUP2] ℹ No apareció el texto del popup de resultado (puede no salir en este caso)")
            return True

        # try to click the best accept-like candidate within the same frame where the text is visible
        JS_CLICK_RESULT_ACCEPT = r"""
        (function(){
          function norm(s){ return (s||"").toString().replace(/\u00a0/g," ").replace(/\s+/g," ").trim().toLowerCase(); }
          function isVisible(el){
            if(!el) return false;
            const r = el.getBoundingClientRect();
            const st = window.getComputedStyle(el);
            if(r.width < 5 || r.height < 5) return false;
            if(st.display==='none' || st.visibility==='hidden') return false;
            if(parseFloat(st.opacity||'1') < 0.05) return false;
            const cx = r.left + r.width/2, cy = r.top + r.height/2;
            const top = document.elementFromPoint(cx, cy);
            return !!(top && (top===el || el.contains(top) || top.contains(el)));
          }

          const resultHints = arguments[0];
          const acceptHints = arguments[1];

          // Find an element that contains result text
          const all = Array.from(document.querySelectorAll("body *"));
          let anchor = null;
          for(const el of all){
            const t = norm(el.innerText);
            if(!t) continue;
            if(resultHints.some(h => t.includes(h))){
              anchor = el; break;
            }
          }
          if(!anchor) return {found:false, clicked:false, why:"no_result_anchor"};

          // Walk up to a likely dialog container
          let root = anchor;
          for(let i=0;i<10 && root && root!==document.body;i++){
            const cls = (root.className||"").toString();
            const role = norm(root.getAttribute("role"));
            if(/dhx|dhtmlx|win|modal|window|dialog|popup/i.test(cls) || role==="dialog"){
              break;
            }
            root = root.parentElement || root;
          }

          // --- Estrategia A (CRÍTICA): click por coordenadas del *texto* "Aceptar" ---
          // En algunos dhtmlx, el nodo clicable no contiene texto, pero el texto sí existe en un text-node.
          try{
            const tw = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
            const hits = [];
            while(tw.nextNode()){
              const n = tw.currentNode;
              const tt = norm(n.nodeValue);
              if(!tt) continue;
              if(tt === 'aceptar' || tt.includes('aceptar') || tt === 'ok'){
                const rg = document.createRange();
                rg.selectNodeContents(n);
                const rr = rg.getBoundingClientRect();
                if(rr && rr.width>3 && rr.height>3){
                  hits.push(rr);
                }
              }
            }
            if(hits.length){
              hits.sort((a,b)=> (b.top - a.top) || (b.left - a.left));
              const r0 = hits[0];
              const x0 = r0.left + r0.width/2;
              const y0 = r0.top + r0.height/2;
              const el0 = document.elementFromPoint(x0,y0);
              if(el0){
                const ev = {bubbles:true,cancelable:true,view:window,clientX:x0,clientY:y0};
                el0.dispatchEvent(new MouseEvent('mousemove', ev));
                el0.dispatchEvent(new MouseEvent('mouseover', ev));
                el0.dispatchEvent(new MouseEvent('mousedown', ev));
                el0.dispatchEvent(new MouseEvent('mouseup', ev));
                el0.dispatchEvent(new MouseEvent('click', ev));
                try{ el0.click(); }catch(e){}
                return {found:true, clicked:true, picked:'textnode@'+(el0.tagName||''), bestS:999};
              }
            }
          }catch(e){}

          function scoreBtn(el){
            const t = norm((el.value||"") + " " + (el.innerText||"") + " " + (el.textContent||""));
            let s = 0;
            if(acceptHints.some(h => t===h || t.includes(h))) s += 100;
            // prefer smallish things (buttons)
            const r = el.getBoundingClientRect();
            s += Math.max(0, 30 - (r.width*r.height)/4000);
            // prefer bottom-right inside dialog
            const rr = root.getBoundingClientRect();
            const cx = r.left + r.width/2, cy = r.top + r.height/2;
            s += Math.max(0, 40 - Math.hypot((rr.right-10)-cx, (rr.bottom-10)-cy)/20);
            return s;
          }

          const sel = "button,input[type='button'],input[type='submit'],[role='button'],a,div,span";
          const candidates = Array.from(root.querySelectorAll(sel)).filter(isVisible);
          if(!candidates.length) return {found:true, clicked:false, why:"no_candidates"};

          let best = null, bestS=-1;
          for(const el of candidates){
            const sc = scoreBtn(el);
            if(sc > bestS){ bestS=sc; best=el; }
          }
          if(!best) return {found:true, clicked:false, why:"no_best"};

          const r = best.getBoundingClientRect();
          const x = r.left + r.width/2, y = r.top + r.height/2;
          const top = document.elementFromPoint(x,y) || best;

          function fire(type){
            top.dispatchEvent(new MouseEvent(type,{bubbles:true,cancelable:true,view:window,clientX:x,clientY:y}));
          }
          try{ best.scrollIntoView({block:'center', inline:'center'}); }catch(e){}
          try{ best.focus(); }catch(e){}
          fire("mousemove"); fire("mouseover"); fire("mousedown"); fire("mouseup"); fire("click");
          try{ best.click(); }catch(e){}
          return {found:true, clicked:true, picked: norm((best.innerText||best.value||"")).slice(0,60), bestS:bestS};
        })();
        """

        paths = self._iter_frame_paths(driver, max_depth=6)

        for attempt in range(1, 9):
            any_click = False
            for p in paths:
                if not self._switch_to_frame_path(driver, p):
                    continue
                try:
                    # quick debug snapshot sometimes (only first attempts)
                    if attempt in (1, 3):
                        self._debug_acceptar_snapshot(driver, label=f"POPUP2@{p}")
                    res = driver.execute_script(JS_CLICK_RESULT_ACCEPT, RESULT_HINTS, ACCEPT_HINTS)
                    if res and res.get("clicked"):
                        any_click = True
                        self.log(f"  [POPUP2] ✓ Click dispatch (intento {attempt}) en frame {p} -> '{res.get('picked','')}' score={res.get('bestS','')}")
                except Exception:
                    pass

            try:
                driver.switch_to.default_content()
            except Exception:
                pass

            # keyboard fallback (sometimes the dialog captures ENTER)
            if not any_click:
                try:
                    ActionChains(driver).send_keys(Keys.ENTER).perform()
                    self.log(f"  [POPUP2] ⚠ Fallback ENTER (intento {attempt})")
                except Exception:
                    pass

            time.sleep(0.6)

            if not result_visible_anywhere():
                self.log("  [POPUP2] ✓ Popup resultado cerrado (verificado por desaparición de texto)")
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
                return True

        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        return False


    def save_debug_screenshot(self, driver, name="error"):
        """Save a screenshot for debugging."""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = os.path.join(self.screenshot_dir, f"debug_{name}_{timestamp}.png")
            driver.save_screenshot(screenshot_path)
            self.log(f"  📸 Captura de pantalla guardada: {screenshot_path}")
            return screenshot_path
        except Exception as e:
            self.log(f"  ⚠ No se pudo guardar captura: {str(e)}")
            return None
    
    def login_to_erp(self, driver):
        """
        Login to ERP system with robust element finding.
        
        Args:
            driver: Selenium WebDriver instance
        
        Returns:
            True if login successful, False otherwise
        """
        try:
            self.log("--- PASO 1: LOGIN ---")
            self.log("Esperando página de login...")
            time.sleep(3)  # Longer initial wait
            
            # Find username field - multiple strategies
            username_strategies = [
                (By.ID, "username"),
                (By.NAME, "username"),
                (By.NAME, "_username"),
                (By.ID, "_username"),
                (By.CSS_SELECTOR, "input[name='username']"),
                (By.CSS_SELECTOR, "input[name='_username']"),
                (By.CSS_SELECTOR, "input[type='text']"),
                (By.XPATH, "//input[@type='text']"),
                (By.XPATH, "//input[contains(@placeholder, 'usuario')]"),
                (By.XPATH, "//input[contains(@placeholder, 'Usuario')]"),
                (By.XPATH, "//input[contains(@name, 'user')]"),
            ]
            
            username_field = self.find_element_robust(
                driver, "campo usuario", username_strategies, timeout=10, clickable=False
            )
            
            if not username_field:
                self.save_debug_screenshot(driver, "login_no_username")
                raise Exception("No se pudo encontrar el campo de usuario")
            
            username_field.clear()
            username_field.send_keys(self.ERP_USERNAME)
            self.log(f"  ✓ Usuario introducido: {self.ERP_USERNAME}")
            
            # Find password field - multiple strategies
            password_strategies = [
                (By.ID, "password"),
                (By.NAME, "password"),
                (By.NAME, "_password"),
                (By.ID, "_password"),
                (By.CSS_SELECTOR, "input[type='password']"),
                (By.CSS_SELECTOR, "input[name='password']"),
                (By.CSS_SELECTOR, "input[name='_password']"),
                (By.XPATH, "//input[@type='password']"),
            ]
            
            password_field = self.find_element_robust(
                driver, "campo contraseña", password_strategies, timeout=10, clickable=False
            )
            
            if not password_field:
                self.save_debug_screenshot(driver, "login_no_password")
                raise Exception("No se pudo encontrar el campo de contraseña")
            
            password_field.clear()
            password_field.send_keys(self.ERP_PASSWORD)
            self.log("  ✓ Contraseña introducida: ****")
            
            # Find Accept/Login button - multiple strategies
            accept_strategies = [
                (By.XPATH, "//button[contains(text(), 'Aceptar')]"),
                (By.XPATH, "//button[contains(text(), 'Accept')]"),
                (By.XPATH, "//button[contains(text(), 'Entrar')]"),
                (By.XPATH, "//button[contains(text(), 'Login')]"),
                (By.XPATH, "//button[contains(text(), 'Iniciar')]"),
                (By.XPATH, "//input[@type='submit']"),
                (By.XPATH, "//button[@type='submit']"),
                (By.CSS_SELECTOR, "button[type='submit']"),
                (By.CSS_SELECTOR, "input[type='submit']"),
                (By.XPATH, "//button[contains(@class, 'login')]"),
                (By.XPATH, "//button[contains(@class, 'btn-primary')]"),
                (By.XPATH, "//button[contains(@class, 'btn')]"),
                (By.CSS_SELECTOR, ".btn-primary"),
                (By.CSS_SELECTOR, ".login-button"),
            ]
            
            accept_button = self.find_element_robust(
                driver, "botón Aceptar/Login", accept_strategies, timeout=10
            )
            
            if not accept_button:
                self.save_debug_screenshot(driver, "login_no_button")
                raise Exception("No se pudo encontrar el botón de Aceptar/Login")
            
            accept_button.click()
            self.log("  ✓ Click en botón Aceptar")
            
            # CRITICAL: Wait longer for main page to load after login
            self.log("Esperando que cargue la página principal (5 segundos)...")
            time.sleep(5)
            
            # v3.0: Verify login was successful - WORKING selectors FIRST (no accents)
            main_page_indicators = [
                # v3.0: WITHOUT accents - these WORK, try first
                (By.XPATH, "//*[contains(text(), 'Articulos')]"),
                (By.XPATH, "//button[contains(text(), 'Articulos')]"),
                (By.XPATH, "//a[contains(text(), 'Articulos')]"),
                (By.XPATH, "//*[contains(text(), 'Proceso')]"),
                (By.XPATH, "//button[contains(text(), 'Proceso')]"),
                (By.XPATH, "//a[contains(text(), 'Proceso')]"),
                # Generic menu indicators
                (By.XPATH, "//*[contains(@class, 'menu')]"),
                (By.XPATH, "//*[contains(@class, 'navbar')]"),
                (By.XPATH, "//*[contains(@class, 'sidebar')]"),
                (By.XPATH, "//*[contains(@id, 'menu')]"),
                # WITH accents - backup only
                (By.XPATH, "//*[contains(text(), 'Artículos')]"),
                (By.XPATH, "//button[contains(text(), 'Artículos')]"),
                (By.XPATH, "//a[contains(text(), 'Artículos')]"),
            ]
            
            main_page_element = self.find_element_robust(
                driver, "indicador página principal", main_page_indicators, timeout=15, clickable=False
            )
            
            if main_page_element:
                self.log("✓ Login exitoso - página principal cargada")
                return True
            else:
                # Check if still on login page
                try:
                    driver.find_element(By.XPATH, "//input[@type='password']")
                    self.save_debug_screenshot(driver, "login_failed")
                    raise Exception("Login fallido - aún en página de login")
                except:
                    # No password field = probably logged in
                    self.log("✓ Login completado (verificación parcial)")
                    return True
        
        except Exception as e:
            self.log(f"✗ Error en login: {str(e)}")
            raise Exception(f"Error de login: {str(e)}")
    
    def detect_page_state(self, driver):
        """
        v2.5: Detect the current state of the ERP page.
        
        Returns:
            'login' - On login page, needs full flow
            'main' - On main page, needs Artículos then rest
            'articles' - On articles page, can go directly to Proceso
            'unknown' - Unknown state
        """
        self.log("--- DETECTANDO ESTADO DE LA PÁGINA ---")
        
        # Check if on login page (password field visible)
        login_indicators = [
            (By.ID, "password"),
            (By.NAME, "password"),
            (By.NAME, "_password"),
            (By.CSS_SELECTOR, "input[type='password']"),
            (By.XPATH, "//input[@type='password']"),
        ]
        
        for by, selector in login_indicators:
            try:
                driver.find_element(by, selector)
                self.log("  → Estado detectado: PÁGINA DE LOGIN")
                return 'login'
            except:
                continue
        
        # Check if on articles page (Proceso menu visible)
        proceso_indicators = [
            (By.XPATH, "//*[normalize-space(text())='Proceso']"),
            (By.XPATH, "//a[contains(text(), 'Proceso')]"),
            (By.XPATH, "//span[contains(text(), 'Proceso')]"),
            (By.PARTIAL_LINK_TEXT, "Proceso"),
            (By.XPATH, "//*[contains(@class, 'dropdown')]//*[contains(text(), 'Proceso')]"),
        ]
        
        for by, selector in proceso_indicators:
            try:
                driver.find_element(by, selector)
                self.log("  → Estado detectado: PÁGINA DE ARTÍCULOS (Proceso visible)")
                return 'articles'
            except:
                continue
        
        # Check if on main page (Artículos icon visible but not Proceso)
        articulos_indicators = [
            (By.XPATH, "//*[normalize-space(text())='Artículos']"),
            (By.XPATH, "//*[normalize-space(text())='Articulos']"),
            (By.XPATH, "//div[contains(text(), 'Artículos')]"),
            (By.XPATH, "//span[contains(text(), 'Artículos')]"),
        ]
        
        for by, selector in articulos_indicators:
            try:
                driver.find_element(by, selector)
                self.log("  → Estado detectado: PÁGINA PRINCIPAL (Artículos visible)")
                return 'main'
            except:
                continue
        
        self.log("  → Estado detectado: DESCONOCIDO")
        return 'unknown'
    
    def do_articulos_doubleclick(self, driver):
        """Execute double-click on Artículos icon."""
        self.log("")
        self.log("--- PASO: DOBLE-CLICK EN ICONO ARTÍCULOS ---")
        
        # v3.0: OPTIMIZED - Working selectors FIRST (without accents)
        articulos_strategies = [
            # v3.0: WITHOUT accents - these WORK, try FIRST
            (By.XPATH, "//*[normalize-space(text())='Articulos']"),
            (By.XPATH, "//*[contains(text(), 'Articulos')]"),
            (By.XPATH, "//div[contains(text(), 'Articulos')]"),
            (By.XPATH, "//span[contains(text(), 'Articulos')]"),
            (By.XPATH, "//a[contains(., 'Articulos')]"),
            (By.XPATH, "//*[contains(normalize-space(.), 'Articulos') and not(descendant::*[contains(normalize-space(.), 'Articulos')])]"),
            # Generic selectors
            (By.XPATH, "//td[contains(., 'Articulos')]"),
            (By.XPATH, "//label[contains(., 'Articulos')]"),
            (By.XPATH, "//*[@title='Articulos']"),
            (By.XPATH, "//*[@alt='Articulos']"),
            # WITH accents - backup only
            (By.XPATH, "//*[normalize-space(text())='Artículos']"),
            (By.XPATH, "//*[contains(text(), 'Artículos')]"),
            (By.XPATH, "//div[contains(text(), 'Artículos')]"),
            (By.XPATH, "//span[contains(text(), 'Artículos')]"),
            (By.XPATH, "//a[contains(., 'Artículos')]"),
            (By.XPATH, "//*[@title='Artículos']"),
        ]
        
        # v3.0: Reduced timeout from 20 to 10
        articulos_element = self.find_element_robust(
            driver, "icono Artículos", articulos_strategies, timeout=10
        )
        
        if not articulos_element:
            self.save_debug_screenshot(driver, "no_articulos")
            raise Exception("No se pudo encontrar el icono de Artículos")
        
        # TRUE FAST DOUBLE-CLICK using ActionChains
        self.log("⚠ IMPORTANTE: Requiere DOBLE-CLICK RÁPIDO (< 0.15s entre clicks)")
        
        try:
            actions = ActionChains(driver)
            actions.double_click(articulos_element).perform()
            self.log("  ✓ Doble-click ejecutado con ActionChains")
        except Exception as e:
            self.log(f"  ⚠ Error en ActionChains: {str(e)}")
            articulos_element.click()
            time.sleep(0.05)
            articulos_element.click()
            self.log("  ✓ Doble-click manual ejecutado (50ms)")
        
        # Wait and verify
        time.sleep(1)
        
        # Check if screen changed
        proceso_check = [
            (By.XPATH, "//a[contains(text(), 'Proceso')]"),
            (By.XPATH, "//*[contains(text(), 'Proceso')]"),
            (By.PARTIAL_LINK_TEXT, "Proceso"),
        ]
        
        for by, selector in proceso_check:
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((by, selector))
                )
                self.log("  ✓ Pantalla cambió correctamente - módulo Artículos abierto")
                return True
            except:
                continue
        
        # Retry if needed
        self.log("  ⚠ Reintentando doble-click...")
        articulos_element = self.find_element_robust(
            driver, "icono Artículos", articulos_strategies, timeout=5
        )
        if articulos_element:
            try:
                actions = ActionChains(driver)
                actions.double_click(articulos_element).perform()
            except:
                articulos_element.click()
                time.sleep(0.05)
                articulos_element.click()
            time.sleep(1)
        
        return True
    
    def do_proceso_and_import(self, driver):
        """Execute Proceso → Importar artículos → Aceptar buttons (all single clicks)."""
        
        # ===== PASO: CLICK EN MENÚ PROCESO =====
        self.log("")
        self.log("--- PASO: CLICK EN MENÚ PROCESO ---")
        
        # v3.3: ULTRA-FAST - Solo las 2 estrategias que sabemos que funcionan
        proceso_element = None
        start_time = time.time()
        
        try:
            # Estrategia que funciona: normalize-space
            proceso_element = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.XPATH, "//*[normalize-space(text())='Proceso']"))
            )
            elapsed = time.time() - start_time
            self.log(f"  ✓ Encontrado 'Proceso' ({elapsed:.1f}s)")
        except:
            try:
                # Fallback rápido
                proceso_element = WebDriverWait(driver, 1).until(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Proceso')]"))
                )
                elapsed = time.time() - start_time
                self.log(f"  ✓ Encontrado 'Proceso' con fallback ({elapsed:.1f}s)")
            except:
                self.save_debug_screenshot(driver, "no_proceso")
                raise Exception("No se pudo encontrar el menú Proceso")
        
        # Click con offset para la flecha (lo que sabemos que funciona)
        try:
            actions = ActionChains(driver)
            actions.move_to_element(proceso_element)
            actions.move_by_offset(25, 0)  # Mover a la derecha para la flecha
            actions.click()
            actions.perform()
            self.log("  ✓ Click ejecutado (+25px offset)")
        except Exception as e:
            # Fallback: click normal
            proceso_element.click()
            self.log("  ✓ Click normal ejecutado")
        
        # Wait for dropdown
        self.log("Esperando menú desplegable (2 segundos)...")
        time.sleep(2)
        
        # v3.0: Verify dropdown opened, retry with different strategy if not
        importar_check = driver.find_elements(By.XPATH, "//*[contains(text(), 'Importar')]")
        if not importar_check:
            self.log("  ⚠ Dropdown no abierto, reintentando con click directo...")
            try:
                proceso_element.click()
                time.sleep(1)
            except:
                pass
        
        # ===== PASO: CLICK EN 'IMPORTAR ARTÍCULOS' =====
        self.log("")
        self.log("--- PASO: CLICK EN 'IMPORTAR ARTÍCULOS' ---")
        
        # v3.3: FAST - Solo las estrategias que funcionan
        start_time = time.time()
        importar_element = None
        
        try:
            # La que funciona (sin acentos)
            importar_element = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Importar articulos')]"))
            )
            elapsed = time.time() - start_time
            self.log(f"  ✓ Encontrado 'Importar articulos' ({elapsed:.1f}s)")
        except:
            try:
                # Fallback
                importar_element = WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Importar')]"))
                )
                elapsed = time.time() - start_time
                self.log(f"  ✓ Encontrado con fallback ({elapsed:.1f}s)")
            except:
                self.save_debug_screenshot(driver, "no_importar")
                raise Exception("No se pudo encontrar 'Importar artículos'")
        
        importar_element.click()
        self.log("  ✓ Click ejecutado")
        
        # Esperar 1 segundo para que aparezca el diálogo (optimizado)
        self.log("Esperando diálogo...")
        time.sleep(1)
        
        # ===== BÚSQUEDA Y CLIC EN BOTÓN ACEPTAR =====
        self.log("")
        self.log("="*60)
        self.log("BUSCANDO BOTÓN ACEPTAR:")
        self.log("="*60)
        # v3.7: Click REAL con verificación (solo loguea éxito si el popup desaparece)
        self.log("Estrategia: Click verificado en popup de importación...")
        # Snapshot de descargas ANTES de clicar Aceptar (para detectar el archivo que genera el ERP)
        _dl_dir = self._get_download_dir()
        _dl_map, _dl_max = self._snapshot_downloads(_dl_dir)
        verified = self.click_aceptar_import_popup(driver, timeout=18)
        if verified:
            self.log("✓ Primer Aceptar completado (VERIFICADO: popup cerrado)")
        else:
            self.log("✗ NO SE PUDO CERRAR EL POPUP DE IMPORTACIÓN (ABORTANDO)")
            self._debug_acceptar_snapshot(driver, label="POPUP1_FAIL")
            self.save_debug_screenshot(driver, "popup_aceptar_no_click")
            return False
        self.log("")

        # Si tras el primer Aceptar se inicia/crea una descarga, damos el proceso por COMPLETADO
        # y cerramos Firefox para evitar quedarnos atascados en el segundo popup.
        # Esperamos más tiempo porque a veces Firefox tarda en crear el archivo en Descargas.
        new_file = self._wait_for_new_download(_dl_map, _dl_max, timeout=60)
        if new_file:
            self.log(f"✓ Descarga detectada tras el primer Aceptar: {new_file}")
            self.log("→ Cerrando Firefox (driver.quit) y finalizando el flujo aquí (sin segundo popup)...")
            try:
                driver.quit()
            except Exception:
                pass
            return True

        # Esperar procesamiento (optimizado)
        time.sleep(2)


        # ===== SEGUNDO POPUP (RESULTADO IMPORTACIÓN) =====
        # Popup de resumen: "Articulos insertados: X  Actualizados: Y" + botones Aceptar/Cancelar.
        # IMPORTANTÍSIMO: aquí el botón puede NO contener 'Aceptar' en el DOM (a veces es un DIV, o está en otro iframe).
        # Por eso lo cerramos detectándolo por el TEXTO del popup y clicando el mejor candidato (aceptar/ok) con eventos reales.
        self.log("")
        self.log("--- CERRANDO POPUP DE RESULTADO (ACEPTAR) ---")

        ok2 = self.close_result_popup_strict(driver, timeout=25)
        if ok2:
            self.log("  ✓ Popup de resultado cerrado (VERIFICADO)")
        else:
            self.log("✗ NO SE PUDO CERRAR EL POPUP DE RESULTADO (ABORTANDO)")
            self.save_debug_screenshot(driver, "popup_resultado_no_cierra")
            return False

        time.sleep(0.5)
        
        # Cerrar Firefox inmediatamente tras completar
        self.log("→ Cerrando Firefox...")
        try:
            driver.quit()
            self.log("  ✓ Firefox cerrado")
        except Exception as e:
            self.log(f"  ⚠ Error cerrando: {e}")
        
        return True

    def automate_import(self):
        """Automate the ERP import process in Firefox with robust element finding.
        
        v2.5: INTELLIGENT STATE DETECTION
        - Detects if on login page → full flow (login → Artículos → Proceso → ...)
        - Detects if on main page → partial flow (Artículos → Proceso → ...)
        - Detects if on articles page → direct flow (Proceso → ...)
        
        v2.5: IMPROVED CLICKS
        - DOUBLE-CLICK only for Artículos (required by ERP)
        - SINGLE-CLICK for all other steps (Proceso, Importar, Aceptar)
        
        v2.5: EXPANDED PROCESO SELECTORS
        - Many more XPath patterns to find "Proceso" menu
        """
        if not HAS_SELENIUM:
            raise Exception("Selenium no instalado. Ejecute: pip install selenium")
        
        driver = None
        
        try:
            self.log("=" * 60)
            self.log("INICIANDO AUTOMATIZACIÓN FIREFOX v3.4")
            self.log("=" * 60)
            
            # Configure Firefox
            self.log("Iniciando Firefox...")
            options = FirefoxOptions()
            # Forzar descargas a carpeta conocida (evita prompts y permite detectar 'éxito' por descarga)
            download_dir = self._get_download_dir()
            try:
                options.set_preference('browser.download.folderList', 2)
                options.set_preference('browser.download.dir', download_dir)
                options.set_preference('browser.download.useDownloadDir', True)
                options.set_preference('browser.download.manager.showWhenStarting', False)
                options.set_preference('browser.download.alwaysOpenPanel', False)
                options.set_preference('browser.download.manager.focusWhenStarting', False)
                options.set_preference('browser.download.manager.alertOnEXEOpen', False)
                options.set_preference('browser.download.manager.closeWhenDone', True)
                options.set_preference('pdfjs.disabled', True)
                options.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv,application/pdf,application/zip,application/octet-stream')
            except Exception as e:
                self.log(f"  ⚠ No se pudieron aplicar preferencias de descarga: {e}")
            self.log(f"  ✓ Descargas forzadas a: {download_dir}")
            driver = webdriver.Firefox(options=options)
            driver.maximize_window()
            driver.implicitly_wait(3)
            
            # Navigate to ERP
            self.log(f"Navegando a: {self.erp_url}")
            driver.get(self.erp_url)
            time.sleep(3)
            
            # v2.5: DETECT PAGE STATE
            page_state = self.detect_page_state(driver)
            
            if page_state == 'login':
                # Full flow: Login → Artículos → Proceso → ...
                self.log("")
                self.log(">>> FLUJO COMPLETO: LOGIN → ARTÍCULOS → PROCESO → IMPORTAR")
                self.login_to_erp(driver)
                self.do_articulos_doubleclick(driver)
                time.sleep(2)
                ok = self.do_proceso_and_import(driver)
                if ok is False:
                    raise Exception("La importación no pudo finalizar: no se cerraron todos los popups.")
                
            elif page_state == 'main':
                # Partial flow: Artículos → Proceso → ...
                self.log("")
                self.log(">>> FLUJO PARCIAL: ARTÍCULOS → PROCESO → IMPORTAR")
                self.log("(Ya logueado, en página principal)")
                self.do_articulos_doubleclick(driver)
                time.sleep(2)
                ok = self.do_proceso_and_import(driver)
                if ok is False:
                    raise Exception("La importación no pudo finalizar: no se cerraron todos los popups.")
                
            elif page_state == 'articles':
                # Direct flow: Proceso → ...
                self.log("")
                self.log(">>> FLUJO DIRECTO: PROCESO → IMPORTAR")
                self.log("(Ya en página de artículos, saltando login y Artículos)")
                ok = self.do_proceso_and_import(driver)
                if ok is False:
                    raise Exception("La importación no pudo finalizar: no se cerraron todos los popups.")
                
            else:
                # Unknown state - try full flow
                self.log("")
                self.log(">>> ESTADO DESCONOCIDO - Intentando flujo completo")
                self.save_debug_screenshot(driver, "unknown_state")
                self.login_to_erp(driver)
                self.do_articulos_doubleclick(driver)
                time.sleep(2)
                ok = self.do_proceso_and_import(driver)
                if ok is False:
                    raise Exception("La importación no pudo finalizar: no se cerraron todos los popups.")
            
            # Update last run date
            self.update_last_run_date()
            
            self.log("")
            self.log("=" * 60)
            self.log("✓✓✓ IMPORTACIÓN AUTOMÁTICA COMPLETADA v3.4 ✓✓✓")
            self.log("=" * 60)
            
        except Exception as e:
            self.log("")
            self.log("=" * 60)
            self.log("✗✗✗ ERROR EN AUTOMATIZACIÓN ✗✗✗")
            self.log("=" * 60)
            self.log(f"Error: {str(e)}")
            
            if driver:
                self.save_debug_screenshot(driver, "error")
            
            raise Exception(f"Error en automatización Firefox: {str(e)}")
        
        finally:
            # Solo llamamos a quit() si geckodriver sigue vivo. Si
            # do_proceso_and_import() ya cerro Firefox, el proceso ha muerto y
            # un quit() redundante intentaria hablar via HTTP con un puerto
            # local que ya nadie escucha, generando 3 warnings de urllib3
            # ("WinError 10061") en la consola. Comportamiento exterior
            # identico: Firefox queda cerrado igual.
            if driver:
                try:
                    geckoproc = getattr(getattr(driver, 'service', None), 'process', None)
                    still_alive = (geckoproc is None) or (geckoproc.poll() is None)
                    if still_alive:
                        driver.quit()
                except Exception:
                    pass


# ============================================================================
# CONFIGURATION MANAGER
# ============================================================================

class ConfigManager:
    """Manages application configuration."""
    
    def __init__(self):
        self.config_file = os.path.join(os.path.dirname(__file__), 'config.json')
        self.config = self.load_config()
    
    def load_config(self):
        """Load configuration from file."""
        default_config = {
            'database_path': DEFAULT_DATABASE_PATH,
            'network_path': DEFAULT_NETWORK_PATH,
            'erp_url': DEFAULT_ERP_URL,
            'default_provider': '1000',
            'default_pro_discount': DEFAULT_PRO_DISCOUNT,
            'default_private_discount': DEFAULT_PRIVATE_DISCOUNT,
            'automate_firefox': True
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    saved_config = json.load(f)
                    default_config.update(saved_config)
        except:
            pass
        
        return default_config
    
    def save_config(self):
        """Save configuration to file."""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=2)
        except:
            pass
    
    def get(self, key, default=None):
        """Get configuration value."""
        return self.config.get(key, default)
    
    def set(self, key, value):
        """Set configuration value."""
        self.config[key] = value
        self.save_config()


# ============================================================================
# MAIN GUI APPLICATION
# ============================================================================

class SingleArticleImportApp:
    """Main GUI application for single article import."""
    
    def __init__(self, root):
        self.root = root
        self.root.title(f"Importación Individual de Artículos v{VERSION}")
        self.root.geometry("700x700")
        self.root.resizable(True, True)
        
        self.config = ConfigManager()
        self.logo_image = None
        
        self._load_logo()
        self._create_widgets()
    
    def _load_logo(self):
        """Load and prepare logo image."""
        if not HAS_PIL:
            return
        
        try:
            for path in LOGO_PATHS:
                if os.path.exists(path):
                    img = Image.open(path)
                    if img.mode != 'RGBA':
                        img = img.convert('RGBA')
                    
                    max_width = 150
                    ratio = max_width / img.width
                    new_size = (max_width, int(img.height * ratio))
                    img = img.resize(new_size, Image.Resampling.LANCZOS)
                    
                    self.logo_image = ImageTk.PhotoImage(img)
                    break
        except Exception as e:
            print(f"Could not load logo: {e}")
    
    def _create_widgets(self):
        """Create GUI widgets."""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Logo
        if self.logo_image:
            logo_label = ttk.Label(main_frame, image=self.logo_image)
            logo_label.image = self.logo_image
            logo_label.pack(pady=5)
        
        # Title
        title = ttk.Label(main_frame, text=f"Importación Individual de Artículos v{VERSION}",
                         font=('Helvetica', 14, 'bold'))
        title.pack(pady=5)
        
        subtitle = ttk.Label(main_frame,
            text="Busca referencia en BASE DE DATOS → Convierte → Importa al ERP",
            font=('Helvetica', 9, 'italic'))
        subtitle.pack(pady=2)
        
        # Reference input frame
        ref_frame = ttk.LabelFrame(main_frame, text="Referencia del Artículo", padding="10")
        ref_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(ref_frame, text="Referencia:", font=('Helvetica', 11)).pack(side=tk.LEFT, padx=5)
        self.reference_var = tk.StringVar()
        ref_entry = ttk.Entry(ref_frame, textvariable=self.reference_var, width=40, font=('Helvetica', 11))
        ref_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ref_entry.bind('<Return>', lambda e: self._search_reference())
        
        search_btn = ttk.Button(ref_frame, text="🔍 Buscar", command=self._search_reference)
        search_btn.pack(side=tk.RIGHT, padx=5)
        
        # Parameters frame
        params_frame = ttk.LabelFrame(main_frame, text="Parámetros de Conversión", padding="10")
        params_frame.pack(fill=tk.X, pady=5)
        
        # Row 1: Provider code
        ttk.Label(params_frame, text="Nº Proveedor:").grid(row=0, column=0, sticky='w', pady=3)
        self.provider_var = tk.StringVar(value=str(self.config.get('default_provider', '1000')))
        ttk.Entry(params_frame, textvariable=self.provider_var, width=15).grid(row=0, column=1, sticky='w', pady=3)
        
        # Row 2: Professional discount
        ttk.Label(params_frame, text="Dto. Profesional (%):").grid(row=1, column=0, sticky='w', pady=3)
        self.pro_discount_var = tk.StringVar(value=str(self.config.get('default_pro_discount', DEFAULT_PRO_DISCOUNT)))
        ttk.Entry(params_frame, textvariable=self.pro_discount_var, width=15).grid(row=1, column=1, sticky='w', pady=3)
        ttk.Label(params_frame, text="ej: 27 = 27% dto", font=('Helvetica', 8)).grid(row=1, column=2, sticky='w', padx=5)
        
        # Row 3: Private discount
        ttk.Label(params_frame, text="Dto. Particular (%):").grid(row=2, column=0, sticky='w', pady=3)
        self.private_discount_var = tk.StringVar(value=str(self.config.get('default_private_discount', DEFAULT_PRIVATE_DISCOUNT)))
        ttk.Entry(params_frame, textvariable=self.private_discount_var, width=15).grid(row=2, column=1, sticky='w', pady=3)
        ttk.Label(params_frame, text="ej: 11 = 11% dto", font=('Helvetica', 8)).grid(row=2, column=2, sticky='w', padx=5)
        
        # Row 4: Firefox automation checkbox
        self.automate_var = tk.BooleanVar(value=self.config.get('automate_firefox', True))
        auto_check = ttk.Checkbutton(params_frame, text="Automatizar importación en Firefox",
                                     variable=self.automate_var)
        auto_check.grid(row=3, column=0, columnspan=3, sticky='w', pady=5)
        
        # Main button
        self.import_btn = ttk.Button(main_frame, text="📥 IMPORTAR ARTÍCULO",
                                     command=self._process_import,
                                     style='Accent.TButton')
        self.import_btn.pack(pady=15)
        
        # Status label
        self.status_var = tk.StringVar(value="Introduzca una referencia para buscar")
        status_label = ttk.Label(main_frame, textvariable=self.status_var,
                                font=('Helvetica', 10, 'bold'))
        status_label.pack()
        
        # Log area
        log_frame = ttk.LabelFrame(main_frame, text="Registro de Operaciones", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, wrap=tk.WORD,
                                                  font=('Consolas', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Settings button
        settings_btn = ttk.Button(main_frame, text="⚙ Configuración", command=self._show_settings)
        settings_btn.pack(side=tk.RIGHT, pady=5)
    
    def log(self, message):
        """Add message to log area."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def update_status(self, message):
        """Update status label."""
        self.status_var.set(message)
        self.root.update()
    
    def _search_reference(self):
        """Search for reference (just validates input)."""
        reference = self.reference_var.get().strip()
        if not reference:
            messagebox.showwarning("Atención", "Por favor, introduzca una referencia")
            return
        
        self.update_status(f"Referencia preparada: {reference}")
        self.log(f"Referencia introducida: {reference}")
    
    def _process_import(self):
        """Main import process."""
        # Get inputs
        reference = self.reference_var.get().strip()
        if not reference:
            messagebox.showwarning("Atención", "Por favor, introduzca una referencia")
            return
        
        try:
            provider_number = self.provider_var.get().strip()
            pro_discount = float(self.pro_discount_var.get().strip())
            private_discount = float(self.private_discount_var.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Los descuentos deben ser números válidos")
            return
        
        # Disable button during process
        self.import_btn.config(state='disabled')
        self.log_text.delete(1.0, tk.END)
        
        try:
            # STEP 1: Search for reference
            self.update_status("Buscando referencia en base de datos...")
            self.log(f"Buscando referencia: {reference}")
            
            database_path = self.config.get('database_path')
            scanner = ProviderDatabaseScanner(database_path, logger=self.log)
            result = scanner.find_reference(reference)
            
            if not result:
                raise Exception(f"Referencia '{reference}' no encontrada en ningún proveedor")
            
            provider_name, tariff_file, row_data, df = result
            self.log(f"✓ Encontrado en: {provider_name}")
            self.log(f"  Archivo: {os.path.basename(tariff_file)}")
            
            # STEP 2: Convert row
            self.update_status("Convirtiendo datos...")
            self.log("Convirtiendo datos al formato ERP...")
            
            converter = SingleRowConverter()
            user_inputs = {
                'provider_number': provider_number,
                'pro_discount': pro_discount,
                'private_discount': private_discount
            }
            output_df, category, cat_info = converter.convert_row(row_data, df, provider_name, user_inputs)
            
            self.log(f"✓ Categoría detectada: {category}")
            self.log(f"  Sección: {cat_info['seccion']}, Familia: {cat_info['familia']}, Subfamilia: {cat_info['subfamilia']}")
            
            # STEP 3: Save to network folder
            self.update_status("Guardando archivo en carpeta de red...")
            self.log("Guardando archivo de importación...")
            
            network_path = self.config.get('network_path')
            file_manager = NetworkFileManager(network_path)
            saved_path = file_manager.save_import_file(output_df)
            
            self.log(f"✓ Archivo guardado: {saved_path}")
            
            # STEP 4: Firefox automation (optional)
            if self.automate_var.get():
                self.update_status("Automatizando importación en Firefox...")
                self.log("Iniciando automatización Firefox...")
                
                erp_url = self.config.get('erp_url')
                automator = FirefoxERPAutomator(erp_url, logger=self.log)
                
                if HAS_SELENIUM:
                    try:
                        automator.automate_import()
                        self.log("✓ Importación automática completada")
                    except Exception as e:
                        self.log(f"⚠ Error en automatización: {str(e)}")
                        self.log("  Puede importar manualmente desde el ERP")
                else:
                    self.log("⚠ Selenium no instalado - automatización no disponible")
                    self.log("  Por favor, importe manualmente desde el ERP")
            else:
                self.log("ℹ Automatización Firefox desactivada")
                self.log(f"  Archivo listo para importar manualmente: {saved_path}")
            
            # Success
            self.update_status("✓ ¡Importación completada exitosamente!")
            self.log("")
            self.log("═" * 50)
            self.log("¡ARTÍCULO IMPORTADO CORRECTAMENTE!")
            self.log("═" * 50)
            
            # -----------------------------------------------------------------
            # AUTO-CIERRE INMEDIATO
            #
            # El cierre con `after()` no es fiable porque parte del flujo corre en
            # un hilo/worker (y el event-loop puede no ejecutar los callbacks).
            # Para garantizar el cierre siempre:
            #   1) cerramos ventana (si existe)
            #   2) forzamos salida del proceso
            #
            # Para importar otra referencia hay que abrir el programa de nuevo.
            # -----------------------------------------------------------------
            try:
                self.update_status("Cerrando programa...")
            except Exception:
                pass

            try:
                try:
                    self.root.destroy()
                except Exception:
                    pass
            finally:
                os._exit(0)

            return
            
        except Exception as e:
            self.update_status(f"Error: {str(e)}")
            self.log(f"✗ ERROR: {str(e)}")
            messagebox.showerror("Error", str(e))
        
        finally:
            self.import_btn.config(state='normal')
    
    def _show_settings(self):
        """Show settings dialog."""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Configuración")
        settings_window.geometry("500x300")
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        frame = ttk.Frame(settings_window, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Database path
        ttk.Label(frame, text="Ruta DATABASE:").grid(row=0, column=0, sticky='w', pady=5)
        db_var = tk.StringVar(value=self.config.get('database_path'))
        ttk.Entry(frame, textvariable=db_var, width=50).grid(row=0, column=1, pady=5)
        
        # Network path
        ttk.Label(frame, text="Ruta red (import):").grid(row=1, column=0, sticky='w', pady=5)
        net_var = tk.StringVar(value=self.config.get('network_path'))
        ttk.Entry(frame, textvariable=net_var, width=50).grid(row=1, column=1, pady=5)
        
        # ERP URL
        ttk.Label(frame, text="URL ERP:").grid(row=2, column=0, sticky='w', pady=5)
        erp_var = tk.StringVar(value=self.config.get('erp_url'))
        ttk.Entry(frame, textvariable=erp_var, width=50).grid(row=2, column=1, pady=5)
        
        # Default provider
        ttk.Label(frame, text="Proveedor por defecto:").grid(row=3, column=0, sticky='w', pady=5)
        prov_var = tk.StringVar(value=self.config.get('default_provider'))
        ttk.Entry(frame, textvariable=prov_var, width=20).grid(row=3, column=1, sticky='w', pady=5)
        
        def save_settings():
            self.config.set('database_path', db_var.get())
            self.config.set('network_path', net_var.get())
            self.config.set('erp_url', erp_var.get())
            self.config.set('default_provider', prov_var.get())
            self.config.save_config()
            messagebox.showinfo("Guardado", "Configuración guardada correctamente")
            settings_window.destroy()
        
        ttk.Button(frame, text="Guardar", command=save_settings).grid(row=5, column=1, sticky='e', pady=20)
        ttk.Button(frame, text="Cancelar", command=settings_window.destroy).grid(row=5, column=0, sticky='w', pady=20)


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """Main entry point."""
    root = tk.Tk()
    
    # Try to set a nicer theme
    try:
        root.tk.call('source', 'azure.tcl')
        root.tk.call('set_theme', 'light')
    except:
        pass
    
    app = SingleArticleImportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()