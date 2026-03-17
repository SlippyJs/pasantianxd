from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import os
import time
import re

import numpy as np
import pandas as pd
import pdfplumber
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import cv2


# ===================================================================
# 1) CONFIGURACIÓN (AJUSTABLE)
# ===================================================================

@dataclass
class OCRConfig:
    dpi: int = 200
    idiomas: str = "spa+eng"
    conf_minima: int = 35
    psm: int = 6
    thread_count_pdf: int = 4

@dataclass
class PathsConfig:
    tesseract_cmd: Optional[str] = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    poppler_bin: Optional[str]   = r"C:\Users\DETPC\OneDrive\Desktop\tesseract\poppler-24.02.0\Library\bin"

@dataclass
class HeuristicsConfig:
    plumber_table_cfg: Dict[str, object] = None
    tol_cols_suave: int = 1
    tol_cols_fuerte: int = 2
    # ENDURECIDO: validación relajada desactivada por defecto
    # Activar solo si se necesita recuperar tablas partidas con datos incompletos
    validacion_relajada_en_continuacion: bool = False
    fusion_lineas_colgantes: bool = True
    max_filas_perfil: int = 8

@dataclass
class BatchConfig:
    max_workers: int = field(default_factory=lambda: max(2, os.cpu_count() or 2))
    omitir_si_existe: bool = True

_DEFAULT_PLUMBER_CFG = {
    "vertical_strategy":    "lines",
    "horizontal_strategy":  "lines",
    "snap_tolerance":       5,
    "join_tolerance":       8,
    "edge_min_length":      18,
    "min_words_vertical":   1,
    "min_words_horizontal": 1,
    "text_tolerance":       2.0,
    "intersection_tolerance": 2,
}

COLUMNAS_EXCEL = ["Pdf", "Pagina", "Fila", "Numero", "Producto", "Cantidad", "Valor_Unitario", "Valor_Total"]

SINONIMOS_COLUMNAS: Dict[str, set] = {
    "Numero":         {"no", "nº", "n°", "no.", "num", "numero", "nro", "ítem", "item", "codigo", "cod"},
    "Producto":       {"descripcion", "descripción", "detalle", "producto", "nombre generico", "nombre genérico"},
    "Cantidad":       {"cantidad", "cant"},
    "Valor_Unitario": {"valor u", "valor u.", "valor unitario", "precio unitario", "pu", "unitario", "precio u."},
    "Valor_Total":    {"valor total", "total", "importe", "importe total", "precio total"},
}

PALABRAS_CLAVE_COLUMNAS = {
    "Numero": [
        "item", "numero", "número", "nro", "n°", "#", "no.", "cod",
        "codigo", "código", "referencia", "ref", "id", "sku",
        "lote", "linea", "n.", "num.", "seq"
    ],
    "Cantidad": [
        "cantidad", "cant", "qty", "q", "cant.", "unidad", "und"
    ],
    "Producto": [
        "descripcion", "descripción", "producto", "detalle",
        "concepto", "articulo", "artículo", "servicio"
    ],
    "Valor_Unitario": [
        "valor unitario", "precio unitario", "v.u", "vu",
        "precio", "valor u", "p.u", "unitario"
    ],
    "Valor_Total": [
        "valor total", "total", "importe", "monto",
        "subtotal", "importe total", "total linea"
    ]
}

# Los 5 campos obligatorios — todos deben estar presentes sin excepción
CAMPOS_REQUERIDOS = {"Numero", "Producto", "Cantidad", "Valor_Unitario", "Valor_Total"}

PALABRAS_OMITIR: set = {
    "subtotal", "sub total", "total", "gran total", "total general", "valor total",
    "importe total", "resumen", "observaciones", "formas de pago", "condiciones"
}


# ===================================================================
# 2) UTILIDADES DE TEXTO / MATRIZ
# ===================================================================

def a_cadena(valor) -> str:
    return "" if valor is None else str(valor)

def normalizar_espacios(texto: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[\r\n]+", " ", a_cadena(texto))).strip()

def canonico(texto: str) -> str:
    t = normalizar_espacios(texto).lower()
    return (t.replace("á", "a").replace("é", "e").replace("í", "i")
             .replace("ó", "o").replace("ú", "u")
             .replace(".", "").replace(":", "").replace(";", ""))

def solo_numeros(texto: str) -> str:
    t = a_cadena(texto).replace("USD", " ").replace("US$", " ").replace("$", " ")
    return re.sub(r"[^\d.,-]", "", t).strip()

def contiene_letras(texto: str) -> bool:
    return bool(re.search(r"[A-Za-zÁÉÍÓÚáéíóú]", a_cadena(texto)))

def parece_dinero(texto: str) -> bool:
    return bool(re.search(r"\d+[.,]\d{2,3}", a_cadena(texto)))

def normalizar_matriz(matriz: List[List[str]]) -> List[List[str]]:
    M = [[normalizar_espacios(c) for c in fila] for fila in matriz if any(a_cadena(x) for x in fila)]
    if not M:
        return []
    n = max(len(f) for f in M)
    return [f + [""] * (n - len(f)) for f in M]

def es_header_fila(fila: List[str]) -> bool:
    t = canonico(" ".join(map(a_cadena, fila)))
    claves = [
        "no", "nº", "n°", "numero",
        "nombre generico", "descripcion", "descripción", "detalle",
        "cantidad", "valor unitario", "precio unitario", "valor total", "precio total",
    ]
    return sum(1 for k in claves if k in t) >= 2

def omitir_fila(fila: List[str]) -> bool:
    t = canonico(" ".join(map(a_cadena, fila)))
    return any(p in t for p in PALABRAS_OMITIR)

def get_cell(fila: List[str], idx: Optional[int]) -> str:
    return fila[idx] if (idx is not None and 0 <= idx < len(fila)) else ""


# ===================================================================
# 3) MAPEADO DE COLUMNAS
# ===================================================================

def mapear_por_header(header: List[str]) -> Dict[str, int]:
    m: Dict[str, int] = {}
    for i, txt in enumerate(header):
        c = canonico(txt)
        if not c:
            continue
        for campo, vocab in SINONIMOS_COLUMNAS.items():
            if any(v in c for v in vocab):
                m[campo] = i
    return m

def buscar_encabezado_en_tabla(tabla: List[List[str]], max_filas: int = 8) -> Optional[Dict[str, int]]:
    """
    Busca la fila de encabezado dentro de las primeras filas de la tabla.
    Exige los 5 campos exactos — sin excepción.
    """
    limite = min(max_filas, len(tabla))
    for i in range(limite):
        filas_test = [tabla[i]]
        if i + 1 < len(tabla):
            filas_test.append(tabla[i + 1])
        indices = detectar_columnas_por_keywords(filas_test)
        if CAMPOS_REQUERIDOS.issubset(indices.keys()):
            return indices
    return None

def detectar_columnas_por_keywords(filas: List[List[str]]) -> Dict[str, int]:
    """Detecta columnas revisando palabras clave en una o dos filas de encabezado."""
    indices: Dict[str, int] = {}
    usados: set = set()
    for fila in filas:
        for idx, celda in enumerate(fila):
            texto = canonico(celda)
            if not texto or idx in usados:
                continue
            for campo, palabras in PALABRAS_CLAVE_COLUMNAS.items():
                if campo in indices:
                    continue
                for p in palabras:
                    if re.search(rf"\b{re.escape(canonico(p))}\b", texto):
                        indices[campo] = idx
                        usados.add(idx)
                        break
    return indices

def perfiles_columnas(matriz: List[List[str]], fila_ini: int, max_filas: int) -> Tuple[List[int], List[int], List[int]]:
    if not matriz:
        return [], [], []
    filas, cols = len(matriz), len(matriz[0])
    r1   = min(filas, fila_ini + max_filas)
    nums = [0] * cols
    decs = [0] * cols
    txts = [0] * cols
    for r in range(fila_ini, r1):
        for c in range(cols):
            v = matriz[r][c]
            if re.search(r"\d", a_cadena(v)): nums[c] += 1
            if parece_dinero(v):              decs[c] += 1
            if contiene_letras(v):            txts[c] += 1
    return nums, decs, txts


# ===================================================================
# 4) VALIDACIÓN Y PREPROCESO DE FILAS
# ===================================================================

@dataclass
class EstadoMapeo:
    """
    Mantiene el mapeo de columnas entre páginas para tablas que continúan.
    Solo se actualiza cuando se extrae una tabla que pasó el filtro estricto.
    """
    m_prev: Optional[Dict[str, int]] = None
    ncol_prev: Optional[int] = None
    en_continuacion: bool = False

def fila_valida(fila: List[str], mapeo: Dict[str, int], relajada: bool = False) -> bool:
    if not fila or all(not x for x in fila):
        return False
    if omitir_fila(fila) or es_header_fila(fila):
        return False

    desc = get_cell(fila, mapeo.get("Producto"))
    cant = get_cell(fila, mapeo.get("Cantidad"))
    unit = get_cell(fila, mapeo.get("Valor_Unitario"))
    tot  = get_cell(fila, mapeo.get("Valor_Total"))

    if len(desc.strip()) < 3:
        return False

    cant_ok = bool(re.fullmatch(r"\d+([.,]\d+)?", solo_numeros(cant)))
    unit_ok = parece_dinero(unit)
    tot_ok  = parece_dinero(tot)

    # ENDURECIDO: modo relajado exige al menos 2 de 3 campos numéricos
    # (antes bastaba con 1, lo que dejaba pasar filas casi vacías)
    if relajada:
        return sum([cant_ok, unit_ok, tot_ok]) >= 2

    # Modo estricto: los 3 campos numéricos deben estar presentes
    return cant_ok and unit_ok and tot_ok

def fusion_lineas_partidas(M: List[List[str]], m: Dict[str, int]) -> List[List[str]]:
    """
    Fusiona pares de filas donde la primera tiene descripción pero le faltan
    números, y la segunda tiene los números pero descripción vacía o mínima.
    Extiende la fila si hace falta para evitar IndexError.
    """
    if not M or not m or "Producto" not in m:
        return M

    idx_p = m.get("Producto")
    idx_c = m.get("Cantidad")
    idx_u = m.get("Valor_Unitario")
    idx_t = m.get("Valor_Total")

    indices_usados = [i for i in [idx_p, idx_c, idx_u, idx_t] if i is not None]
    ancho_minimo   = (max(indices_usados) + 1) if indices_usados else 0

    def extender(fila: List[str]) -> List[str]:
        if len(fila) < ancho_minimo:
            return fila + [""] * (ancho_minimo - len(fila))
        return fila[:]

    out, r = [], 0
    while r < len(M):
        fila  = extender(M[r])
        desc  = get_cell(fila, idx_p)
        cant  = get_cell(fila, idx_c)
        unit  = get_cell(fila, idx_u)
        tot   = get_cell(fila, idx_t)

        pocos_numeros = sum(
            bool(x) for x in [solo_numeros(cant), solo_numeros(unit), solo_numeros(tot)]
        ) <= 1

        if desc and len(desc) >= 6 and pocos_numeros and (r + 1 < len(M)):
            fila2  = extender(M[r + 1])
            desc2  = get_cell(fila2, idx_p)
            cant2  = get_cell(fila2, idx_c)
            unit2  = get_cell(fila2, idx_u)
            tot2   = get_cell(fila2, idx_t)
            senales_fila2 = sum(
                bool(x) for x in [solo_numeros(cant2), solo_numeros(unit2), solo_numeros(tot2)]
            )

            if senales_fila2 >= 1 and (not desc2 or len(desc2) <= 3 or desc2.lower() in {"-", "—", "_"}):
                nueva = fila[:]
                if idx_c is not None: nueva[idx_c] = cant  if solo_numeros(cant)  else cant2
                if idx_u is not None: nueva[idx_u] = unit  if parece_dinero(unit) else unit2
                if idx_t is not None: nueva[idx_t] = tot   if parece_dinero(tot)  else tot2
                out.append(nueva)
                r += 2
                continue

        out.append(fila)
        r += 1

    return out


# ===================================================================
# 5) PIPELINE DE EMISIÓN
# ===================================================================

def construir_headers(M: List[List[str]]) -> Tuple[List[str], List[str], int]:
    ncol_real = max(len(f) for f in M)
    if len(M) == 1:
        h1 = [M[0][c] if c < len(M[0]) else "" for c in range(ncol_real)]
        h2 = h1[:]
    else:
        h1 = [M[0][c] if c < len(M[0]) else "" for c in range(ncol_real)]
        h2 = [normalizar_espacios(
                  f"{M[0][c] if c < len(M[0]) else ''} {M[1][c] if c < len(M[1]) else ''}"
              ) for c in range(ncol_real)]
    return h1, h2, ncol_real


# ================================================================
# FILTRO ESTRICTO — única puerta de entrada para cualquier tabla
# ================================================================
def tabla_contiene_campos_requeridos(matriz: List[List[str]]) -> bool:
    """
    Filtro estricto: devuelve True SOLO cuando se cumplen TODAS estas condiciones:

    1. La tabla tiene al menos 1 fila de encabezado + 1 fila de datos.
    2. Se detectan los 5 campos exactos en el encabezado:
       Numero, Producto, Cantidad, Valor_Unitario, Valor_Total.
    3. Cada campo apunta a un índice de columna distinto.
    4. Al menos 1 fila de datos debajo del encabezado tiene:
       - Descripción con 4+ caracteres, Y
       - Los 3 campos numéricos presentes (Cantidad, Valor_Unitario, Valor_Total).

    Tablas con más de 5 columnas son aceptadas si contienen los 5 campos;
    las columnas extra se ignoran en la extracción.
    Tablas de un solo producto son aceptadas si esa fila pasa el punto 4.
    """
    if not matriz:
        return False

    M = normalizar_matriz(matriz)
    if not M or len(M) < 2:
        return False

    limite = min(6, len(M))

    for i in range(limite):
        filas_test = [M[i]]
        if i + 1 < len(M):
            filas_test.append(M[i + 1])

        indices = detectar_columnas_por_keywords(filas_test)

        # Condición 2: los 5 campos deben estar todos presentes
        if not CAMPOS_REQUERIDOS.issubset(indices.keys()):
            continue

        # Condición 3: cada campo debe apuntar a una columna distinta
        if len(set(indices.values())) < len(indices):
            continue

        # Condición 4: al menos 1 fila de datos válida debajo del encabezado
        revisar_hasta = min(i + 8, len(M))
        for r in range(i + 1, revisar_hasta):
            fila = M[r]
            if omitir_fila(fila) or es_header_fila(fila):
                continue

            desc = get_cell(fila, indices.get("Producto"))
            cant = get_cell(fila, indices.get("Cantidad"))
            unit = get_cell(fila, indices.get("Valor_Unitario"))
            tot  = get_cell(fila, indices.get("Valor_Total"))

            desc_ok = len(desc.strip()) >= 4
            # ENDURECIDO: se exigen los 3 campos numéricos (antes bastaba 1)
            cant_ok = bool(re.search(r"\d", cant))
            unit_ok = parece_dinero(unit)
            tot_ok  = parece_dinero(tot)

            if desc_ok and cant_ok and unit_ok and tot_ok:
                return True

    return False


def emitir_registros(M: List[List[str]],
                     nombre_pdf: str,
                     pagina: int,
                     mapeo: Dict[str, int],
                     fila_ini: int,
                     relajada: bool) -> List[Dict]:
    """
    Emite solo los 5 campos requeridos aunque el mapeo contenga columnas extra.
    """
    CAMPOS_SALIDA = ("Numero", "Producto", "Cantidad", "Valor_Unitario", "Valor_Total")

    out: List[Dict] = []
    k = 0

    for r in range(fila_ini, len(M)):
        fila = M[r]
        if all(x == "" for x in fila):
            continue
        if omitir_fila(fila) or es_header_fila(fila):
            continue
        if not fila_valida(fila, mapeo, relajada=relajada):
            continue

        reg = {"Pdf": nombre_pdf, "Pagina": pagina, "Fila": None}
        for campo in CAMPOS_SALIDA:
            reg[campo] = get_cell(fila, mapeo.get(campo)) if campo in mapeo else ""

        if not reg["Producto"]:
            continue

        k += 1
        reg["Fila"] = k
        out.append(reg)

    return out


def normalizar_y_emitir(matriz_cruda: List[List[str]],
                        nombre_pdf: str,
                        pagina: int,
                        estado: EstadoMapeo,
                        heur: HeuristicsConfig) -> List[Dict]:
    """
    Normaliza la matriz y emite registros.

    ENDURECIDO:
    - Se eliminó el fallback heurístico (mapear_por_heuristica).
      Solo se procesa si se encuentra un encabezado real con los 5 campos.
    - La continuación solo se acepta si el número de columnas es idéntico
      al de la tabla anterior (tolerancia 0, antes era ±2).
    - Sin encabezado detectado y sin continuación estricta → descartado.
    """
    M = normalizar_matriz(matriz_cruda)
    if not M:
        return []

    h1, h2, ncol_real = construir_headers(M)
    en_cont  = False
    fila_ini = 1

    # Paso 1: intentar detectar encabezado real en las primeras filas
    m = buscar_encabezado_en_tabla(M)

    if m is None:
        # Paso 2: sin encabezado → solo aceptar si es continuación EXACTA
        # ENDURECIDO: ncol debe coincidir exactamente (antes tolerancia ±2)
        es_continuacion_exacta = (
            estado.m_prev is not None
            and estado.ncol_prev is not None
            and ncol_real == estado.ncol_prev  # tolerancia 0
        )
        if not es_continuacion_exacta:
            return []  # sin encabezado y sin continuación → descartar

        m = estado.m_prev.copy()
        fila_ini = 0
        en_cont  = True
    else:
        # Encabezado encontrado: verificar que tenga los 5 campos
        if not CAMPOS_REQUERIDOS.issubset(m.keys()):
            return []  # encabezado incompleto → descartar

        # Si hay encabezado nuevo, el mapeo previo ya no aplica
        fila_ini = 1

    # Verificación final: el mapeo debe tener los 5 campos
    if not CAMPOS_REQUERIDOS.issubset(m.keys()):
        return []

    if heur.fusion_lineas_colgantes and "Producto" in m:
        M = fusion_lineas_partidas(M, m)

    filas = emitir_registros(
        M, nombre_pdf, pagina, m, fila_ini,
        relajada=(heur.validacion_relajada_en_continuacion and en_cont)
    )

    if filas:
        estado.m_prev          = m.copy()
        estado.ncol_prev       = ncol_real
        estado.en_continuacion = en_cont

    return filas


# ===================================================================
# 6) LECTOR PDFPLUMBER (RUTA A)
# ===================================================================

class LectorPlumber:
    def __init__(self, estado: EstadoMapeo, heur: HeuristicsConfig,
                 plumber_cfg: Optional[Dict[str, object]] = None):
        self.estado = estado
        self.heur   = heur
        self.cfg    = plumber_cfg or _DEFAULT_PLUMBER_CFG

    def extraer(self, ruta_pdf: str) -> Tuple[List[Dict], Set[int]]:
        nombre     = Path(ruta_pdf).name
        resultados: List[Dict] = []
        faltantes:  Set[int]   = set()

        with pdfplumber.open(ruta_pdf) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                tablas = page.extract_tables(self.cfg) or []
                hubo   = False

                for tbl in tablas:
                    try:
                        # ENDURECIDO: solo pasan tablas con los 5 campos validados
                        # La continuación solo se acepta con columnas exactas
                        es_tabla_valida = tabla_contiene_campos_requeridos(tbl)
                        es_continuacion = (
                            not es_tabla_valida
                            and self.estado.m_prev is not None
                            and self.estado.ncol_prev is not None
                            and tbl
                            and len(tbl[0]) == self.estado.ncol_prev  # exacto, no ±1
                        )
                        if not es_tabla_valida and not es_continuacion:
                            continue

                        filas = normalizar_y_emitir(tbl, nombre, i, self.estado, self.heur)
                        if filas:
                            resultados.extend(filas)
                            hubo = True
                    except Exception:
                        pass

                if not hubo:
                    faltantes.add(i)

        return resultados, faltantes


# ===================================================================
# 7) LECTOR OCR (RUTA B)
# ===================================================================

class LectorOCR:
    def __init__(self, estado: EstadoMapeo, ocr: OCRConfig,
                 paths: PathsConfig, heur: HeuristicsConfig):
        self.estado = estado
        self.ocr    = ocr
        self.paths  = paths
        self.heur   = heur
        if paths.tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = paths.tesseract_cmd

    # ------------------------------------------------------------------
    # Detección del grid de líneas dentro de un ROI
    # ------------------------------------------------------------------
    def _grid_roi(self, roi_bgr) -> Tuple[List[int], List[int]]:
        """Detecta líneas h/v de la tabla y devuelve coordenadas de filas y columnas."""
        gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)
        _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        H, W      = bw.shape
        kernel_h  = cv2.getStructuringElement(cv2.MORPH_RECT, (W // 20, 1))
        kernel_v  = cv2.getStructuringElement(cv2.MORPH_RECT, (1, H // 20))
        lineas_h  = cv2.morphologyEx(bw, cv2.MORPH_OPEN, kernel_h, iterations=2)
        lineas_v  = cv2.morphologyEx(bw, cv2.MORPH_OPEN, kernel_v, iterations=2)
        intersecs = cv2.bitwise_and(lineas_h, lineas_v)

        ys, xs = np.where(intersecs > 0)
        if len(xs) < 3 or len(ys) < 3:
            return [], []

        def agrupar(vals, tol=8) -> List[int]:
            vals   = sorted(map(int, vals))
            grupos = [[vals[0]]]
            for v in vals[1:]:
                if abs(v - grupos[-1][-1]) <= tol:
                    grupos[-1].append(v)
                else:
                    grupos.append([v])
            return [int(sum(g) / len(g)) for g in grupos]

        rows = sorted({0} | set(agrupar(ys, 8)) | {H - 1})
        cols = sorted({0} | set(agrupar(xs, 8)) | {W - 1})

        if len(rows) < 3 or len(cols) < 3:
            return [], []
        return rows, cols

    # ------------------------------------------------------------------
    # Detección de regiones de tabla (ROIs) en la página
    # ------------------------------------------------------------------
    def _detectar_rois(self, bgr) -> List[Tuple[int, int, int, int]]:
        """Encuentra rectángulos con estructura real de tabla (líneas cruzadas)."""
        gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        H, W     = bw.shape
        kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (W // 25, 1))
        kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, H // 25))
        lineas_h = cv2.morphologyEx(bw, cv2.MORPH_OPEN, kernel_h, iterations=2)
        lineas_v = cv2.morphologyEx(bw, cv2.MORPH_OPEN, kernel_v, iterations=2)
        grid     = cv2.add(lineas_h, lineas_v)

        cnts, _ = cv2.findContours(grid, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        rois = []
        for c in cnts:
            x, y, w, h = cv2.boundingRect(c)
            if w * h < H * W * 0.004 or w < W * 0.18 or h < H * 0.05:
                continue
            rois.append((x, y, x + w, y + h))

        return sorted(rois, key=lambda r: (r[1], r[0]))

    # ------------------------------------------------------------------
    # OCR por tabla completa — una sola llamada Tesseract por tabla
    # ------------------------------------------------------------------
    def _construir_matriz_ocr(self, roi_pil: Image.Image,
                              rows: List[int], cols: List[int]) -> List[List[str]]:
        """
        Ejecuta OCR UNA SOLA VEZ sobre la imagen completa de la tabla y
        asigna cada palabra a su celda usando el grid (rows × cols).
        """
        nfil = len(rows) - 1
        ncol = len(cols) - 1
        M    = [[""] * ncol for _ in range(nfil)]

        config_tess = f"-l {self.ocr.idiomas} --oem 3 --psm {self.ocr.psm}"
        df = pytesseract.image_to_data(
            roi_pil, config=config_tess,
            output_type=pytesseract.Output.DATAFRAME
        )

        if df is None or df.empty:
            return M

        # Columna 'text' puede contener NaN (float) → convertir a str primero
        df["text"] = df["text"].astype(str)
        df = df[df["text"].str.strip().replace("nan", "") != ""]
        df["conf"] = pd.to_numeric(df["conf"], errors="coerce")
        df = df[df["conf"] >= self.ocr.conf_minima]

        if df.empty:
            return M

        for _, w in df.iterrows():
            cx = int(w["left"]  + w["width"]  / 2)
            cy = int(w["top"]   + w["height"] / 2)

            fila_idx = next((r for r in range(nfil) if rows[r] <= cy < rows[r + 1]), None)
            col_idx  = next((c for c in range(ncol) if cols[c] <= cx < cols[c + 1]), None)

            if fila_idx is not None and col_idx is not None:
                tok = str(w["text"]).strip()
                if tok and tok != "nan":
                    sep = " " if M[fila_idx][col_idx] else ""
                    M[fila_idx][col_idx] += sep + tok

        return [[normalizar_espacios(celda) for celda in fila] for fila in M]

    # ------------------------------------------------------------------
    # Convertir un ROI a matriz de texto (sin tocar EstadoMapeo)
    # ------------------------------------------------------------------
    def _roi_a_matriz(self, img_pil: Image.Image,
                      roi: Tuple[int, int, int, int]) -> Optional[List[List[str]]]:
        """
        Convierte un ROI en matriz de texto.
        Devuelve None si el ROI no tiene grid válido.
        No modifica EstadoMapeo — eso lo hace _procesar_pagina() en orden.
        """
        x0, y0, x1, y1 = roi
        roi_pil = img_pil.crop((x0, y0, x1, y1))
        roi_bgr = cv2.cvtColor(np.array(roi_pil), cv2.COLOR_RGB2BGR)

        rows, cols = self._grid_roi(roi_bgr)
        if not rows or not cols:
            return None

        return self._construir_matriz_ocr(roi_pil, rows, cols)

    # ------------------------------------------------------------------
    # Procesar una página: ROIs en paralelo → emitir en orden
    # ------------------------------------------------------------------
    def _procesar_pagina(self, img_pil: Image.Image, nombre: str, pag: int,
                         workers: int) -> List[Dict]:
        """
        1. Detecta ROIs en la página.
        2. Convierte ROIs a matrices en paralelo (OCR es la parte lenta).
        3. Filtra y emite registros secuencialmente para respetar EstadoMapeo.
        """
        bgr  = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)
        rois = self._detectar_rois(bgr)

        if not rois:
            return []

        # Paso 1: OCR en paralelo para todos los ROIs de la página
        matrices: List[Optional[List[List[str]]]] = [None] * len(rois)

        with ThreadPoolExecutor(max_workers=workers) as pool:
            futuros = {
                pool.submit(self._roi_a_matriz, img_pil, roi): idx
                for idx, roi in enumerate(rois)
            }
            for futuro in as_completed(futuros):
                idx = futuros[futuro]
                try:
                    matrices[idx] = futuro.result()
                except Exception:
                    matrices[idx] = None

        # Paso 2: filtrar y emitir en orden (EstadoMapeo requiere secuencial)
        out: List[Dict] = []
        for M in matrices:
            if M is None:
                continue

            # ENDURECIDO: mismo filtro estricto que pdfplumber
            es_tabla_valida = tabla_contiene_campos_requeridos(M)
            es_continuacion = (
                not es_tabla_valida
                and self.estado.m_prev is not None
                and self.estado.ncol_prev is not None
                and M and len(M[0]) == self.estado.ncol_prev  # exacto, no tolerancia
            )
            if not es_tabla_valida and not es_continuacion:
                continue

            filas = normalizar_y_emitir(M, nombre, pag, self.estado, self.heur)
            if filas:
                out.extend(filas)

        return out

    # ------------------------------------------------------------------
    # Orquestar extracción OCR sobre las páginas indicadas
    # ------------------------------------------------------------------
    def extraer(self, ruta_pdf: str, paginas: Optional[Set[int]] = None) -> List[Dict]:
        nombre = Path(ruta_pdf).name

        imgs = convert_from_path(
            ruta_pdf,
            dpi=self.ocr.dpi,
            poppler_path=self.paths.poppler_bin,
            thread_count=self.ocr.thread_count_pdf,
        )

        out: List[Dict] = []
        roi_workers = max(2, (os.cpu_count() or 2) // 2)

        for i, img in enumerate(imgs, start=1):
            if paginas and i not in paginas:
                continue
            filas = self._procesar_pagina(img, nombre, i, workers=roi_workers)
            if filas:
                out.extend(filas)

        return out


# ===================================================================
# 8) EXPORTACIÓN A EXCEL
# ===================================================================

class ExtractorProductos:
    """
    Uso individual:
        extr  = ExtractorProductos(paths=PathsConfig(), ocr=OCRConfig())
        filas = extr.extraer_productos("archivo.pdf")
        ruta  = extr.exportar_excel(filas, "salida/")

    Uso en lote (paralelo):
        extr = ExtractorProductos()
        extr.extraer_y_exportar_lote(lista_de_pdfs, carpeta_salida)
    """

    def __init__(self,
                 paths: Optional[PathsConfig]     = None,
                 ocr:   Optional[OCRConfig]        = None,
                 heur:  Optional[HeuristicsConfig] = None,
                 batch: Optional[BatchConfig]      = None):
        self.paths = paths or PathsConfig()
        self.ocr   = ocr   or OCRConfig()
        self.heur  = heur  or HeuristicsConfig(plumber_table_cfg=_DEFAULT_PLUMBER_CFG.copy())
        self.batch = batch or BatchConfig()

        if self.heur.plumber_table_cfg is None:
            self.heur.plumber_table_cfg = _DEFAULT_PLUMBER_CFG.copy()
        if self.paths.tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = self.paths.tesseract_cmd

    # ------------------------------------------------------------------
    # Procesar un solo PDF
    # ------------------------------------------------------------------
    def extraer_productos(self, ruta_pdf: str) -> List[Dict]:
        """Extrae productos de un PDF usando pdfplumber + OCR como fallback."""
        estado         = EstadoMapeo()
        lector_plumber = LectorPlumber(estado, self.heur, self.heur.plumber_table_cfg)
        lector_ocr     = LectorOCR(estado, self.ocr, self.paths, self.heur)

        res_plumber, faltantes = lector_plumber.extraer(ruta_pdf)

        if not res_plumber and not faltantes:
            datos = lector_ocr.extraer(ruta_pdf, paginas=None)
        else:
            ocr_extra = lector_ocr.extraer(ruta_pdf, paginas=faltantes) if faltantes else []
            datos     = res_plumber + ocr_extra

        datos = sorted(datos, key=lambda r: (r["Pagina"], r.get("Fila") or 0))
        k_by_page: Dict[int, int] = {}
        for d in datos:
            p = d["Pagina"]
            k_by_page[p] = k_by_page.get(p, 0) + 1
            d["Fila"]    = k_by_page[p]

        return datos

    # ------------------------------------------------------------------
    # Exportar lista de registros a Excel
    # ------------------------------------------------------------------
    def exportar_excel(self, filas: List[Dict], carpeta_destino: str,
                       nombre_archivo: str = "productos_extraidos.xlsx") -> str:
        df = pd.DataFrame(filas)
        for c in COLUMNAS_EXCEL:
            if c not in df.columns:
                df[c] = ""
        df = df[COLUMNAS_EXCEL]

        ruta_base = Path(carpeta_destino) / nombre_archivo

        for intento in range(5):
            try:
                ruta_base.parent.mkdir(parents=True, exist_ok=True)
                with pd.ExcelWriter(str(ruta_base), engine="openpyxl") as w:
                    df.to_excel(w, index=False, sheet_name="PRODUCTOS")
                return str(ruta_base)
            except PermissionError:
                print(f"[ADVERTENCIA] Archivo bloqueado. Reintentando ({intento+1}/5)...")
                time.sleep(1.2)

        ruta_alt = Path(carpeta_destino) / f"productos_extraidos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(str(ruta_alt), engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="PRODUCTOS")
        print("[OK] Guardado en archivo alternativo:", ruta_alt)
        return str(ruta_alt)

    # ------------------------------------------------------------------
    # Procesar un PDF y exportar su Excel (usado internamente en el lote)
    # ------------------------------------------------------------------
    def _procesar_un_pdf(self, ruta_pdf: str, carpeta_destino: str) -> Tuple[str, int, str]:
        """
        Extrae y exporta un PDF individual.
        Devuelve (nombre_pdf, cantidad_items, ruta_excel_o_msg).
        Códigos especiales: -1 = omitido, -2 = error.
        """
        nombre = Path(ruta_pdf).name

        ruta_excel_esperada = Path(carpeta_destino) / f"{Path(ruta_pdf).stem}.xlsx"
        if self.batch.omitir_si_existe and ruta_excel_esperada.exists():
            return nombre, -1, str(ruta_excel_esperada)

        try:
            filas = self.extraer_productos(ruta_pdf)
            if not filas:
                return nombre, 0, "sin_datos"

            ruta = self.exportar_excel(
                filas, carpeta_destino,
                nombre_archivo=f"{Path(ruta_pdf).stem}.xlsx"
            )
            return nombre, len(filas), ruta

        except Exception as e:
            return nombre, -2, str(e)

    # ------------------------------------------------------------------
    # Procesar múltiples PDFs en paralelo
    # ------------------------------------------------------------------
    def extraer_y_exportar_lote(self,
                                rutas_pdf: List[str],
                                carpeta_destino: str) -> List[Dict]:
        """
        Procesa múltiples PDFs en paralelo con ThreadPoolExecutor.
        Cada PDF corre en su propio hilo con su propio EstadoMapeo.

        Devuelve un resumen por PDF:
        [{"pdf": nombre, "items": N, "estado": "ok|omitido|sin_datos|error", "detalle": ...}]
        """
        Path(carpeta_destino).mkdir(parents=True, exist_ok=True)
        total      = len(rutas_pdf)
        resumen    = []
        procesados = 0

        print(f"[LOTE] Procesando {total} PDFs con {self.batch.max_workers} workers...")

        with ThreadPoolExecutor(max_workers=self.batch.max_workers) as pool:
            futuros = {
                pool.submit(self._procesar_un_pdf, ruta, carpeta_destino): ruta
                for ruta in rutas_pdf
            }

            for futuro in as_completed(futuros):
                nombre, n_items, detalle = futuro.result()
                procesados += 1

                if   n_items == -1: estado_str = "omitido"
                elif n_items == -2: estado_str = "error"
                elif n_items ==  0: estado_str = "sin_datos"
                else:               estado_str = "ok"

                resumen.append({
                    "pdf":     nombre,
                    "items":   max(n_items, 0),
                    "estado":  estado_str,
                    "detalle": detalle,
                })

                simbolo = {"ok": "✓", "omitido": "↷", "sin_datos": "○", "error": "✗"}[estado_str]
                print(f"  [{procesados}/{total}] {simbolo} {nombre}  ({estado_str})")

        ok      = sum(1 for r in resumen if r["estado"] == "ok")
        errores = sum(1 for r in resumen if r["estado"] == "error")
        print(f"\n[LOTE] Finalizado — OK: {ok} | Sin datos: {total-ok-errores} | Errores: {errores}")

        return resumen