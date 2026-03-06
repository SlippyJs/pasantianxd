# -*- coding: utf-8 -*-
"""
extractor_tablas_pdf.py
=======================
Extrae filas de tablas (ítems, cantidades, precios) desde archivos PDF.

Flujo principal:
  1. Intenta extracción digital con pdfplumber (tablas con líneas o texto).
  2. Si no obtiene resultados, usa OCR con Tesseract + pdf2image.

Dependencias externas requeridas:
  pip install pdfplumber pytesseract pdf2image Pillow
  - Tesseract-OCR instalado en el sistema (ver sección RUTAS)
  - Poppler instalado en el sistema        (ver sección RUTAS)
"""

# =============================================================================
# LIBRERÍAS ESTÁNDAR
# =============================================================================
import os
import re
from typing import Any, Dict, List, Optional

# =============================================================================
# LIBRERÍAS DE TERCEROS
# =============================================================================
import pdfplumber

try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image
    _TESSERACT_DISPONIBLE = True
except ImportError:
    _TESSERACT_DISPONIBLE = False

# =============================================================================
# RUTAS EXTERNAS (Tesseract y Poppler)
# Edita estas rutas si cambias la instalación.
# =============================================================================

TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_BIN   = r"C:\Users\DETPC\OneDrive\Desktop\tesseract\poppler-24.02.0\Library\bin"

if _TESSERACT_DISPONIBLE:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE

# =============================================================================
# CONFIGURACIÓN OCR
# Ajusta estos valores sin necesidad de tocar la lógica del extractor.
# =============================================================================

TESSERACT_LANG = "spa"   # Idioma. Cambia a "eng" para PDFs en inglés.
TESSERACT_DPI  = 300     # Resolución de imagen para OCR. 300 es el estándar.

# =============================================================================
# CONFIGURACIÓN DE COLUMNAS
# Define qué columnas se buscan y cuáles son obligatorias para aceptar una tabla.
# Agrega o quita palabras clave aquí sin tocar otra lógica.
# =============================================================================

# Nombres internos de las columnas (no cambiar, se usan en todo el código)
COL_ITEM           = "item"
COL_CANTIDAD       = "cantidad"
COL_DESCRIPCION    = "descripcion"
COL_VALOR_UNITARIO = "valor_unitario"
COL_VALOR_TOTAL    = "valor_total"

# Palabras clave para detectar cada columna en los encabezados del PDF
PALABRAS_CLAVE_COLUMNAS: Dict[str, List[str]] = {
    COL_ITEM: [
        "item", "número", "nro", "n°", "#", "ítem", "no.", "cod",
        "codigo", "código", "referencia", "ref", "id", "sku",
        "lote", "linea", "n.", "num.", "seq", "secuencia",
    ],
    COL_CANTIDAD: [
        "cantidad", "cant", "qty", "q", "cant.", "unidad", "und",
        "cantidad solicitada", "cant sol", "cantidadsol",
    ],
    COL_DESCRIPCION: [
        "descripción", "descripcion", "producto", "desc", "concepto",
        "detalle", "servicio", "artículo", "articulo", "nombre producto",
        "nombre del producto", "denominación", "denominacion",
        "presentación", "presentacion",
    ],
    COL_VALOR_UNITARIO: [
        "valor unitario", "precio unitario", "v.u", "vu", "precio",
        "valor u", "p.u", "unit", "unitario", "valor unit", "punit",
        "precio unit", "valor x unidad", "valor unidad",
    ],
    COL_VALOR_TOTAL: [
        "valor total", "total", "importe", "monto", "subtotal",
        "valor linea", "valor líquido", "valor liquido",
        "subtotal linea", "importe total", "total linea",
    ],
}

# Columnas que DEBEN estar presentes para considerar una tabla válida.
# Filtro estricto: si falta cualquiera de estas, la tabla se descarta.
COLUMNAS_REQUERIDAS: List[str] = [
    COL_DESCRIPCION,
    COL_CANTIDAD,
    COL_VALOR_UNITARIO,
    COL_VALOR_TOTAL,
]

# Unidades textuales que se aceptan en la columna "cantidad" (además de números)
UNIDADES_TEXTUALES_VALIDAS = {"unidad", "und", "varios", "x", "gl", "kit", "m2", "m3"}


# =============================================================================
# UTILIDADES DE TEXTO
# Funciones puras de bajo nivel. No dependen de ninguna otra sección.
# =============================================================================

def normalizar(texto: str) -> str:
    """Convierte a minúsculas y colapsa espacios múltiples."""
    if not texto:
        return ""
    return re.sub(r'\s+', ' ', str(texto).lower().strip())


def es_numero(texto: str) -> bool:
    """Devuelve True si el texto representa un número válido (acepta comas y puntos)."""
    if not texto:
        return False
    limpio = str(texto).replace(',', '.').replace(' ', '').replace('\xa0', '')
    try:
        float(limpio)
        return True
    except ValueError:
        return False


def celda_segura(valor: Any) -> str:
    """Convierte cualquier valor de celda a string limpio; nunca devuelve None."""
    return "" if valor is None else str(valor).strip()


def obtener_celda(fila: List[str], indice: int) -> str:
    """Accede a una celda por índice de forma segura; devuelve '' si está fuera de rango."""
    if indice < 0 or indice >= len(fila):
        return ""
    return celda_segura(fila[indice])


# =============================================================================
# DETECCIÓN DE COLUMNAS
# Identifica a qué índice corresponde cada columna según el encabezado.
# =============================================================================

def detectar_columnas(filas: List[List[str]]) -> Dict[str, int]:
    """
    Recorre las filas dadas buscando palabras clave de PALABRAS_CLAVE_COLUMNAS.
    Devuelve {nombre_columna: índice_en_la_fila}.
    Un índice no puede asignarse a dos columnas distintas.
    """
    indices: Dict[str, int] = {}
    indices_usados: set = set()

    for fila in filas:
        for col_idx, celda in enumerate(fila):
            texto = normalizar(celda)
            if not texto or col_idx in indices_usados:
                continue
            for nombre_col, palabras in PALABRAS_CLAVE_COLUMNAS.items():
                if nombre_col in indices:
                    continue
                if any(normalizar(p) in texto for p in palabras):
                    indices[nombre_col] = col_idx
                    indices_usados.add(col_idx)
                    break

    return indices


def tabla_tiene_columnas_requeridas(indices: Dict[str, int]) -> bool:
    """
    Devuelve True solo si TODAS las columnas de COLUMNAS_REQUERIDAS
    fueron encontradas en los índices detectados.
    """
    return all(indices.get(col) is not None for col in COLUMNAS_REQUERIDAS)


def buscar_encabezado(tabla: List[List[str]], max_filas: int = 10) -> Optional[Dict[str, int]]:
    """
    Revisa las primeras `max_filas` de la tabla buscando un encabezado válido.
    Devuelve el dict de índices si lo encuentra, None si no.
    """
    for idx in range(min(max_filas, len(tabla))):
        filas_candidatas = [tabla[idx]]
        if idx + 1 < len(tabla):
            filas_candidatas.append(tabla[idx + 1])
        indices = detectar_columnas(filas_candidatas)
        if tabla_tiene_columnas_requeridas(indices):
            return indices
    return None


# =============================================================================
# VALIDACIÓN Y EXTRACCIÓN DE FILAS DE DATOS
# =============================================================================

def extraer_fila(fila: List[str], indices: Dict[str, int]) -> Dict[str, str]:
    """Mapea una fila cruda a un dict con los nombres de columna definidos."""
    return {
        "numero_item":    obtener_celda(fila, indices.get(COL_ITEM, -1)),
        "cantidad":       obtener_celda(fila, indices.get(COL_CANTIDAD, -1)),
        "descripcion":    obtener_celda(fila, indices.get(COL_DESCRIPCION, -1)),
        "valor_unitario": obtener_celda(fila, indices.get(COL_VALOR_UNITARIO, -1)),
        "valor_total":    obtener_celda(fila, indices.get(COL_VALOR_TOTAL, -1)),
    }


def fila_es_valida(datos: Dict[str, str]) -> bool:
    """
    Filtro estricto sobre una fila extraída. Todos los criterios deben cumplirse:
      - descripcion: al menos 3 caracteres.
      - cantidad:    número válido o unidad textual conocida.
      - valor_unitario y valor_total: deben ser numéricos.
    """
    desc  = datos.get("descripcion", "")
    cant  = datos.get("cantidad", "")
    v_uni = datos.get("valor_unitario", "")
    v_tot = datos.get("valor_total", "")

    if len(desc) < 3:
        return False
    if not (es_numero(cant) or normalizar(cant) in UNIDADES_TEXTUALES_VALIDAS):
        return False
    if not es_numero(v_uni):
        return False
    if not es_numero(v_tot):
        return False

    return True


def procesar_tabla(
    tabla: List[List[str]],
    numero_pagina: int,
    offset_linea: int = 0,
) -> List[Dict[str, Any]]:
    """
    Procesa una tabla completa (la primera fila debe ser el encabezado).
    Devuelve solo las filas que pasan fila_es_valida.
    """
    resultado = []

    if not tabla or len(tabla) < 2:
        return resultado

    indices = detectar_columnas([tabla[0]])
    if not tabla_tiene_columnas_requeridas(indices):
        return resultado

    for idx, fila in enumerate(tabla[1:], start=1):
        if len(fila) < 2 or all(not celda_segura(c) for c in fila):
            continue

        datos = extraer_fila(fila, indices)
        if not fila_es_valida(datos):
            continue

        datos["numero_pagina"] = numero_pagina
        datos["numero_linea"]  = offset_linea + idx + 1
        resultado.append(datos)

    return resultado


# =============================================================================
# CONCATENACIÓN DE TABLAS FRAGMENTADAS ENTRE PÁGINAS
# =============================================================================

def _fila_parece_datos(fila: List[str]) -> bool:
    """Heurística: True si la fila contiene un dígito o texto suficientemente largo."""
    texto = " ".join(celda_segura(c) for c in fila)
    return any(ch.isdigit() for ch in texto) or len(texto) > 20


def concatenar_tablas_fragmentadas(
    tablas_meta: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """
    Une tablas consecutivas que no tienen encabezado propio (continuación de página).
    Cada elemento de tablas_meta debe tener: 'page', 'tabla', 'enc_idx', 'indices'.
    """
    if not tablas_meta:
        return []

    resultado = []
    i = 0
    while i < len(tablas_meta):
        base = tablas_meta[i].copy()
        base["tabla"] = list(base["tabla"])
        j = i + 1

        while j < len(tablas_meta):
            siguiente = tablas_meta[j]
            if siguiente.get("enc_idx") is not None:
                break
            if siguiente["tabla"] and _fila_parece_datos(siguiente["tabla"][0]):
                base["tabla"].extend(siguiente["tabla"])
                j += 1
            else:
                break

        resultado.append(base)
        i = j

    return resultado


# =============================================================================
# PIPELINE COMPARTIDO (digital y OCR usan el mismo procesamiento final)
# =============================================================================

def _procesar_tablas_meta(tablas_meta: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Aplica concatenación de fragmentos y extracción de filas válidas
    sobre la lista de metadatos de tablas.
    """
    unidas = concatenar_tablas_fragmentadas(tablas_meta)
    todas  = []

    for meta in unidas:
        indices = meta.get("indices")
        if not indices or not tabla_tiene_columnas_requeridas(indices):
            continue
        filas = procesar_tabla(meta["tabla"], meta["page"])
        todas.extend(filas)

    return todas


# =============================================================================
# CLASE PRINCIPAL
# =============================================================================

class ExtractorTablasPDF:
    """
    Extrae filas de tablas desde un PDF.

    Uso básico:
        extractor = ExtractorTablasPDF("mi_documento.pdf")
        filas = extractor.extraer()
        for fila in filas:
            print(fila)
    """

    def __init__(self, ruta_pdf: str):
        if not os.path.isfile(ruta_pdf):
            raise FileNotFoundError(f"PDF no encontrado: {ruta_pdf}")
        self.ruta_pdf = ruta_pdf

    # -------------------------------------------------------------------------
    # Punto de entrada público
    # -------------------------------------------------------------------------

    def extraer(self) -> List[Dict[str, Any]]:
        """
        Intenta extracción digital primero; si no obtiene resultados, usa OCR.
        Devuelve lista de dicts con los ítems encontrados.
        """
        filas = self._extraer_digital()
        if not filas:
            filas = self._extraer_ocr()
        return filas

    # -------------------------------------------------------------------------
    # Extracción digital con pdfplumber
    # -------------------------------------------------------------------------

    def _extraer_digital(self) -> List[Dict[str, Any]]:
        """Extrae tablas usando las líneas o el texto embebido del PDF (sin OCR)."""
        tablas_meta = []

        try:
            with pdfplumber.open(self.ruta_pdf) as doc:
                for num_pag, pagina in enumerate(doc.pages, start=1):
                    for tabla in self._obtener_tablas_pdfplumber(pagina):
                        indices = buscar_encabezado(tabla)
                        tablas_meta.append({
                            "page":    num_pag,
                            "tabla":   tabla,
                            "enc_idx": 0 if indices else None,
                            "indices": indices,
                        })
        except Exception as e:
            print(f"[ExtractorDigital] Error: {e}")
            return []

        return _procesar_tablas_meta(tablas_meta)

    def _obtener_tablas_pdfplumber(self, pagina) -> List[List[List[str]]]:
        """
        Intenta dos estrategias en cascada:
          1. Líneas físicas del PDF.
          2. Alineación de texto cuando no hay líneas visibles.
        """
        tablas = pagina.extract_tables({
            "vertical_strategy":   "lines",
            "horizontal_strategy": "lines",
        })
        if tablas:
            return tablas

        tablas = pagina.extract_tables({
            "vertical_strategy":   "text",
            "horizontal_strategy": "text",
            "snap_tolerance":      5,
        })
        return tablas or []

    # -------------------------------------------------------------------------
    # Extracción OCR con Tesseract
    # -------------------------------------------------------------------------

    def _extraer_ocr(self) -> List[Dict[str, Any]]:
        """
        Convierte cada página a imagen con pdf2image y aplica Tesseract OCR.
        Luego reconstruye tablas desde el texto reconocido.
        """
        if not _TESSERACT_DISPONIBLE:
            print("[OCR] Instala: pip install pytesseract pdf2image Pillow")
            return []
        if not os.path.isfile(TESSERACT_EXE):
            print(f"[OCR] Tesseract no encontrado en: {TESSERACT_EXE}")
            return []

        try:
            imagenes = convert_from_path(
                self.ruta_pdf,
                dpi=TESSERACT_DPI,
                poppler_path=POPPLER_BIN,
            )
        except Exception as e:
            print(f"[OCR] Error convirtiendo PDF a imágenes: {e}")
            return []

        tablas_meta = []
        for num_pag, imagen in enumerate(imagenes, start=1):
            texto = self._ocr_imagen(imagen)
            if not texto:
                continue
            tabla = self._reconstruir_tabla_desde_texto(texto)
            if not tabla:
                continue
            indices = buscar_encabezado(tabla)
            tablas_meta.append({
                "page":    num_pag,
                "tabla":   tabla,
                "enc_idx": 0 if indices else None,
                "indices": indices,
            })

        return _procesar_tablas_meta(tablas_meta)

    def _ocr_imagen(self, imagen) -> str:
        """Aplica Tesseract OCR a una imagen PIL y devuelve el texto resultante."""
        try:
            config = f"--oem 3 --psm 6 -l {TESSERACT_LANG}"
            return pytesseract.image_to_string(imagen, config=config)
        except Exception as e:
            print(f"[OCR] Error en Tesseract: {e}")
            return ""

    # -------------------------------------------------------------------------
    # Reconstrucción de tabla desde texto plano (salida OCR)
    # -------------------------------------------------------------------------

    @staticmethod
    def _reconstruir_tabla_desde_texto(texto: str) -> List[List[str]]:
        """
        Convierte texto plano (salida OCR) en estructura de tabla.
        Filtra líneas muy cortas y divide cada línea en celdas.
        """
        lineas = [l.rstrip() for l in texto.splitlines() if len(l.strip()) > 10]
        tabla  = []
        for linea in lineas:
            celdas = ExtractorTablasPDF._dividir_linea(linea)
            celdas = [c for c in celdas if c]
            if celdas:
                tabla.append(celdas)
        return tabla

    @staticmethod
    def _dividir_linea(linea: str) -> List[str]:
        """
        Divide una línea en columnas usando tres estrategias en cascada:
          1. Dos o más espacios consecutivos.
          2. Tabulaciones.
          3. Números al final de la línea (cantidades/precios).
        Si ninguna aplica, devuelve la línea como una sola celda.
        """
        # Estrategia 1: múltiples espacios
        partes = re.split(r'\s{2,}', linea.strip())
        if len(partes) > 1:
            return [p.strip() for p in partes]

        # Estrategia 2: tabulaciones
        partes = linea.split('\t')
        if len(partes) > 1:
            return [p.strip() for p in partes]

        # Estrategia 3: números agrupados al final
        match = re.search(r'(\s[\d\.,]+\s*){1,3}$', linea)
        if match:
            prefijo = linea[:match.start()].strip()
            sufijo  = linea[match.start():].strip()
            return [prefijo] + sufijo.split()

        return [linea.strip()]


# =============================================================================
# PUNTO DE ENTRADA PARA PRUEBAS RÁPIDAS
# Ejecuta: python extractor_tablas_pdf.py mi_archivo.pdf
# =============================================================================

if __name__ == "__main__":
    import sys
    import json

    if len(sys.argv) < 2:
        print("Uso: python extractor_tablas_pdf.py <ruta_al_pdf>")
        sys.exit(1)

    extractor = ExtractorTablasPDF(sys.argv[1])
    filas     = extractor.extraer()

    print(f"\nTotal de ítems extraídos: {len(filas)}\n")
    print(json.dumps(filas, ensure_ascii=False, indent=2))