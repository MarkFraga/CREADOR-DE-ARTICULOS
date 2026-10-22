# DON ARTICULO

Herramienta interna de alta de artículos en un ERP para una empresa del sector cerámica y baño.

La persona escribe **una referencia**; el programa busca en las tarifas de los proveedores (un Excel por
proveedor), convierte esa fila al formato de importación del ERP (código, descripción, familia, precios, tarifas,
descuentos, logística…), genera el fichero `importart.xlsx` y **un robot Selenium/Firefox lo importa en el ERP**.
Hay una versión web (Flask, para toda la red local) y una de escritorio (Tkinter).

> Esta es una copia **anonimizada y publicable** del programa. Todos los nombres de empresas, proveedores,
> personas, URL, usuarios, contraseñas y valores comerciales (márgenes, descuentos, nº de proveedor) son ficticios
> (`PROVEEDOR_EJEMPLO`, `PROVEEDOR_001`, `empresa-x.example`, `x`/`y`…). Los datos (tarifas, históricos, capturas,
> ejecutables) **no** forman parte del repositorio.

## Qué hace

1. **Busca** la referencia (coincidencia exacta, todas las hojas de todas las tarifas de `DATABASE`).
   - Si está en **varios proveedores**, un banner pide elegir cuál (lo ya importado sale en gris).
   - Si sale **varias veces en la misma hoja** con datos distintos (color, formato, precio…), pide elegir la fila;
     si todas importarían lo mismo, no pregunta.
   - Si es una pieza suelta de un **conjunto**, avisa antes de importar (`conjuntos.py`).
2. **Convierte** la fila al formato del ERP (`ROBOT.py`, `SingleRowConverter`):
   clasificación en sección/familia/subfamilia por palabras clave ponderadas, fórmula cerámica, tarifas 1-3,
   EAN (si falta se genera uno válido y estable), unidades y logística por caja/palet.
3. **Controla** los casos peligrosos: precio 0 o ilegible (se cancela sin tocar el ERP), códigos ya usados en el
   ERP (el nuevo entra como `x.1`, `x.2`…), proveedores que ya están enteros en el ERP, referencias a elegir.
4. **Importa** con Selenium: login, menú *Artículos › Proceso › Importar artículos*, confirmación de los
   popups (cada uno verificado), y deja constancia en un histórico SQLite.
5. Una pestaña **REINICIAR REFERENCIA** (protegida por contraseña) permite olvidar lo apuntado de una referencia
   cuando el artículo se borró del ERP; lo borrado se archiva, no se pierde.

## Estructura

| Fichero | Para qué sirve |
|---|---|
| `ROBOT.py` | Motor: constantes del ERP, clasificador, buscador de tarifas, convertidor fila→ERP, perfiles por proveedor, robot Selenium y ventana original |
| `DON_ARTICULO.py` | Utilidades compartidas (búsqueda exacta, normalización, descuentos, overrides) y la ventana de escritorio |
| `DON_ARTICULO_WEB.py` | Servidor Flask/Waitress: API con *Server-Sent Events*, flujo completo de importación, pestaña de reinicio |
| `web_templates/index.html` | Interfaz web (una sola página, sin librerías de JS ni recursos de terceros) |
| `row_choice.py` | Banner de elección cuando una referencia sale varias veces en una hoja |
| `ref_index.py` | Índice en memoria referencia → proveedores (se reconstruye solo si cambia una tarifa) |
| `historico.py` | Histórico SQLite de lo importado, para no duplicar códigos y poder reiniciar una referencia |
| `tariff_sql.py` | Tarifas enormes servidas desde SQLite en vez de releer el Excel (verifica que el resultado es idéntico) |
| `conjuntos.py` | Aviso «esta referencia es parte de un conjunto» a partir de una tabla Excel |
| `examples/make_example_config.py` | Genera una configuración y una carpeta `DATABASE` de ejemplo con datos inventados |
| `tests/` | Pruebas con `unittest` (también las ejecuta `pytest`) |
| `INICIAR_WEB.bat`, `INICIAR_SERVIDOR_WEB.bat` | Lanzadores del servidor web |
| `compile_exe.py`, `build_windows.bat` | Empaquetado con PyInstaller (incluye solo los ficheros de datos e imágenes que existan) |

### Qué se ha omitido de esta copia

- Los **adaptadores por proveedor** (reglas de columnas, precios, descuentos, categorías por código, logística…):
  en la versión interna había varios. Aquí se conserva el **mecanismo**: `ROBOT.PROVIDER_PROFILES` (perfil
  declarativo por proveedor), `ROBOT._CERAMIC_LOGISTICS_SPEC` (ficha de columnas de logística),
  `ROBOT.category_for_code` (categoría fijada por un Excel `CODIGO|CATEGORIA`) y proveedores ficticios de ejemplo.
- Las listas de marcas del clasificador (`ProductClassifier.BRAND_CATEGORIES`) llevan marcas ficticias.
- La política de precios de `ROBOT.py` (`CERAMIC_MULT`, `CERAMIC_ADD`, descuentos por defecto) son **valores de
  ejemplo**: pon los tuyos.
- Datos, tarifas, `.xlsx`, bases de datos, imágenes, iconos y ejecutables.

## Instalación

Requiere Python 3.10+, Firefox y geckodriver (solo para el robot de importación).

```bash
pip install -r requirements.txt
```

## Configuración

Todo lo sensible o dependiente del entorno se lee de **variables de entorno**; los valores por defecto son
marcadores de posición:

| Variable | Por defecto | Qué es |
|---|---|---|
| `ERP_URL` | `http://erp.empresa-x.example/` | Página de entrada del ERP |
| `ERP_USER` | `x` | Usuario del ERP |
| `ERP_PASSWORD` | `y` | Contraseña del ERP |
| `DONARTICULO_DATABASE` | `./DATABASE` | Carpeta con una subcarpeta por proveedor y sus tarifas `.xlsx` |
| `DONARTICULO_IMPORT_DIR` | `./import` | Carpeta donde se escribe `importart.xlsx` (en la empresa, una carpeta compartida) |
| `DONARTICULO_RESET_SALT` / `DONARTICULO_RESET_HASH` | sal `x` / hash de la contraseña `y` | Contraseña de la pestaña REINICIAR (PBKDF2-SHA256, hexadecimal) |
| `DONARTICULO_TIMEZONE` | `Europe/Madrid` | Zona horaria (IANA) para decidir «primera ejecución de hoy» |
| `DONARTICULO_DESKTOP_SHORTCUTS` | desactivado | `1` crea en el escritorio una carpeta `WEB` con accesos directos al servidor |
| `DONARTICULO_ONLINE_LOOKUPS` | desactivado | `1` activa las consultas a servicios de terceros (ver *Privacidad*) |
| `GEMINI_API_KEY` | — | Clasificación por imagen (IA). **Exige además `DONARTICULO_ONLINE_LOOKUPS=1`**; instala `pip install google-genai` |

Para fijar tu propia contraseña de reinicio:

```bash
python -c "import hashlib,os;s=os.urandom(16);print('SALT',s.hex());print('HASH',hashlib.pbkdf2_hmac('sha256',b'tu-clave',s,100000).hex())"
```

### Ficheros de datos (junto al programa)

| Fichero | Columnas |
|---|---|
| `ayuda.xlsx` | `PROVEEDOR`, `DESCUENTOS` (`37+13` = %Dto1+%Dto2; `41` = %Dto1; `1,8+4,5` = fórmula cerámica ×1,8 + 4,5), `NÚMERO PROVEEDOR`, `D.PROFESIONAL`, `D.PARTICULAR` |
| `PROVS.xlsx` | `PROVEEDOR`, `TIPO` (categoría principal) |
| `exceltodo.xlsx` | `TIPO`, `PROVEEDOR`, `DESCUENTO PROFESINAL`, `DESCUENTO PARTICULAR` |
| `overrides.xlsx` | `PROVEEDOR`, `REFERENCIA`, `PATRON`, `CATEGORIA` (categoría forzada; las celdas vacías valen «cualquiera») |

Para probar sin datos reales:

```bash
python examples/make_example_config.py        # crea los cuatro Excel y DATABASE/ con datos ficticios
set DONARTICULO_DATABASE=%CD%\DATABASE        # (Windows)  export DONARTICULO_DATABASE=$PWD/DATABASE  (Linux/macOS)
python DON_ARTICULO_WEB.py                    # http://localhost:5050
```

### Imágenes e iconos

No se publican con el código (`.gitignore` los excluye) y el programa funciona sin ellos:

- **Web:** `static/en_proceso.gif` y `static/completado.png`.
- **Ventana de escritorio** (y `compile_exe.py`): `en_proceso.gif`, `completado.png`, `toilet_logo.png`, `logo.png` y
  `don_articulo.ico`, junto a `DON_ARTICULO.py`.

## Ejecución

```bash
python DON_ARTICULO_WEB.py     # versión web (puerto 5050)
python DON_ARTICULO.py         # versión de escritorio
```

El robot navega por la interfaz web de un ERP concreto (menús *Artículos*, *Proceso*, *Importar artículos*,
popups *Aceptar*); si tu ERP usa otros textos, adapta los selectores en `FirefoxERPAutomator`.

## Privacidad

Por defecto el programa **no contacta con ningún servicio de terceros**: la web usa las fuentes del sistema y no
carga recursos externos. Con `DONARTICULO_ONLINE_LOOKUPS=1` se activan, solo para artículos que el clasificador no
consigue clasificar, estas consultas:

- una búsqueda en un buscador web (se envía la descripción del artículo, o proveedor + referencia + descripción);
- una búsqueda de imágenes y, si hay `GEMINI_API_KEY`, el envío de esa imagen y la descripción a un servicio de IA.

Actívalo solo si las tarifas de tus proveedores no son confidenciales.

## Tests

```bash
python -m unittest discover -s tests -v
```

Cubren la normalización, el análisis de descuentos, la búsqueda de referencias, el convertidor (con y sin
perfil de proveedor), la logística cerámica, el histórico, el banner de elección, los conjuntos, la copia SQL
de las tarifas, la limpieza de texto, el EAN estable, el carácter opt-in de las consultas externas y el flujo web
completo con el ERP y Firefox sustituidos por un doble.

## Seguridad

- El servidor escucha en todas las interfaces (`0.0.0.0:5050`) y **no tiene autenticación**, salvo la pestaña de
  reinicio; está pensado para una red local de confianza. No lo expongas a Internet.
- La carpeta estática es `static/` (no la raíz del proyecto), para no servir el código ni los Excel.
- Los datos de las tarifas se pintan en la web con `textContent`, nunca como HTML.
- Ninguna credencial está en el código: ERP y reinicio salen del entorno.

## Licencia

Sin licencia definida todavía.
