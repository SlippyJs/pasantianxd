from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set
from datetime import datetime
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
    dpi: int = 260
    idiomas: str = "spa+eng"
    conf_minima: int = 35     # Confianza mínima por palabra OCR
    psm: int = 6              # Page Segmentation Mode (6: bloque uniforme de texto)

@dataclass
class PathsConfig:
    tesseract_cmd: Optional[str] = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    poppler_bin: Optional[str] = r"C:\Users\DETPC\OneDrive\Desktop\tesseract\poppler-24.02.0\Library\bin"

@dataclass
class HeuristicsConfig:
    plumber_table_cfg: Dict[str, object] = None        # configuración de pdfplumber
    tol_cols_suave: int = 1                            # tolerancia leve para alinear columnas
    tol_cols_fuerte: int = 2                           # tolerancia fuerte para alinear columnas
    validacion_relajada_en_continuacion: bool = True   # validación más permisiva si es continuación
    fusion_lineas_colgantes: bool = True               # combinar filas partida (texto + números abajo)
    max_filas_perfil: int = 8                          # filas a muestrear para perfiles

# Defaults para pdfplumber
_DEFAULT_PLUMBER_CFG = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance": 5,
    "join_tolerance": 8,
    "edge_min_length": 18,
    "min_words_vertical": 1,
    "min_words_horizontal": 1,
    "text_tolerance": 2.0,
    "intersection_tolerance": 2,
}

# Columnas de salida
COLUMNAS_EXCEL = ["Pdf", "Pagina", "Fila", "Numero", "Producto", "Cantidad", "Valor_Unitario", "Valor_Total"]

# Sinónimos de encabezados para mapear columnas
SINONIMOS_COLUMNAS: Dict[str, set] = {
    "Numero": {"no", "nº", "n°", "no.", "num", "numero", "nro", "ítem", "item", "codigo", "cod"},
    "Producto": {"descripcion", "descripción", "detalle", "producto", "nombre generico", "nombre genérico"},
    "Cantidad": {"cantidad", "cant"},
    "Valor_Unitario": {"valor u", "valor u.", "valor unitario", "precio unitario", "pu", "unitario", "precio u."},
    "Valor_Total": {"valor total", "total", "importe", "importe total", "precio total"},
}

# Palabras clave extendidas para detectar encabezados en cualquier formato
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

# Campos mínimos exigidos para procesar una tabla (filtro seguro)
CAMPOS_REQUERIDOS = {"Numero", "Producto", "Cantidad", "Valor_Unitario", "Valor_Total"}

# Palabras que NO deben entrar como ítems
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
def buscar_encabezado_en_tabla(tabla: List[List[str]], max_filas: int = 8) -> Optional[Dict[str,int]]:
    """
    Busca encabezado dentro de las primeras filas de la tabla
    """

    limite = min(max_filas, len(tabla))

    for i in range(limite):

        filas_test = [tabla[i]]

        if i + 1 < len(tabla):
            filas_test.append(tabla[i+1])

        indices = detectar_columnas_por_keywords(filas_test)

        if CAMPOS_REQUERIDOS.issubset(indices.keys()):
            return indices

    return None

def detectar_columnas_por_keywords(filas: List[List[str]]) -> Dict[str, int]:
    """
    Detecta columnas revisando palabras clave en múltiples filas
    """
    indices = {}
    usados = set()

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
    r1 = min(filas, fila_ini + max_filas)
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

def mapear_por_heuristica(matriz: List[List[str]], fila_ini: int, max_filas: int) -> Dict[str, int]:
    if not matriz:
        return {}
    cols = len(matriz[0])
    nums, decs, txts = perfiles_columnas(matriz, fila_ini, max_filas)
    col_total = max(range(cols), key=lambda c: (decs[c], c))
    cand_unit = [c for c in range(cols) if c < col_total] or [0]
    col_unit = max(cand_unit, key=lambda c: (decs[c], nums[c]))
    cand_cant = [c for c in range(cols) if c < col_unit] or [0]
    col_cant = max(cand_cant, key=lambda c: (nums[c], -decs[c]))
    usados = {col_total, col_unit, col_cant}
    cand_prod = [c for c in range(cols) if c not in usados] or [0]
    col_prod = max(cand_prod, key=lambda c: txts[c])
    m = {"Producto": col_prod, "Valor_Total": col_total, "Valor_Unitario": col_unit, "Cantidad": col_cant}
    if cols:
        m["Numero"] = 0 if 0 not in (usados | {col_prod}) else None
    return {k: v for k, v in m.items() if v is not None}

def alinear_con_prev(m_actual: Dict[str, int],
                     m_prev: Optional[Dict[str, int]],
                     matriz: List[List[str]],fila_ini: int,
                     tol_cols: int,
                     max_filas: int) -> Dict[str, int]:
    if not m_prev or not matriz:
        return m_actual

    cols = len(matriz[0])
    nums, decs, txts = perfiles_columnas(matriz, fila_ini, max_filas)

    def mejor_col(tipo: str, prefer: Optional[int]) -> Optional[int]:
        if cols == 0: return None

        def score(c: int) -> float:
            base = 0.0
            if   tipo == "Valor_Total":    base = 3.0 * decs[c] + 1.0 * nums[c]
            elif tipo == "Valor_Unitario": base = 3.0 * decs[c] + 0.8 * nums[c]
            elif tipo == "Cantidad":       base = 2.0 * nums[c] - 0.5 * decs[c]
            elif tipo == "Producto":       base = 2.5 * txts[c]
            elif tipo == "Numero":         base = 0.5 * nums[c] + 0.2 * txts[c]
            if prefer is not None:
                base -= 0.9 * min(abs(c - prefer), tol_cols + 1)
            return base

        candidatos = list(range(cols))
        if prefer is not None:
            lo, hi = max(0, prefer - tol_cols), min(cols - 1, prefer + tol_cols)
            vecinos = list(range(lo, hi + 1))
            if vecinos: candidatos = vecinos

        usados = {v for k, v in m_actual.items()
                  if k in {"Valor_Total", "Valor_Unitario", "Cantidad", "Producto"} and isinstance(v, int)}

        for c in sorted(candidatos, key=lambda x: score(x), reverse=True):
            if tipo in {"Valor_Total", "Valor_Unitario", "Cantidad", "Producto"} and c in usados and m_actual.get(tipo) != c:
                continue
            return c
        return None

    for k in ["Valor_Total", "Valor_Unitario", "Cantidad", "Producto", "Numero"]:
        prefer = m_prev.get(k) if m_prev else None
        if prefer is None: continue
        mejor = mejor_col(k, prefer)
        if mejor is not None: m_actual[k] = mejor

    return m_actual


# ===================================================================
# 4) VALIDACIÓN Y PREPROCESO DE FILAS
# ===================================================================

@dataclass
class EstadoMapeo:
    """Mantiene mapeo de columnas entre páginas consecutivas (tablas partidas)."""
    m_prev: Optional[Dict[str, int]] = None
    ncol_prev: Optional[int] = None
    en_continuacion: bool = False  # flag para permitir validación relajada

def fila_valida(fila: List[str], mapeo: Dict[str, int], relajada: bool = False) -> bool:

    if not fila or all(not x for x in fila):
        return False

    if omitir_fila(fila) or es_header_fila(fila):
        return False

    get = lambda idx: (fila[idx] if (idx is not None and 0 <= idx < len(fila)) else "")

    desc = get(mapeo.get("Producto"))
    cant = get(mapeo.get("Cantidad"))
    unit = get(mapeo.get("Valor_Unitario"))
    tot  = get(mapeo.get("Valor_Total"))

    if len(desc.strip()) < 3:
        return False

    cant_ok = bool(re.fullmatch(r"\d+([.,]\d+)?", solo_numeros(cant)))
    unit_ok = parece_dinero(unit)
    tot_ok = parece_dinero(tot)

    if relajada:
        return sum([cant_ok, unit_ok, tot_ok]) >= 1

    return cant_ok and unit_ok and tot_ok

def fusion_lineas_partidas(M: List[List[str]], m: Dict[str, int]) -> List[List[str]]:
    """
    Funde filas del estilo:
      Fila r  : [Producto = "texto largo", Cantidad/Unit/Total casi vacíos]
      Fila r+1: [Producto vacío o mínimo, pero con números en Cant/Unit/Total]
    """
    if not M or not m or "Producto" not in m: return M
    idx_p, idx_c, idx_u, idx_t = m.get("Producto"), m.get("Cantidad"), m.get("Valor_Unitario"), m.get("Valor_Total")

    out, r = [], 0
    while r < len(M):
        fila = M[r]
        desc = get_cell(fila, idx_p)
        cant = get_cell(fila, idx_c)
        unit = get_cell(fila, idx_u)
        tot  = get_cell(fila, idx_t)

        pocos = sum(bool(x) for x in [solo_numeros(cant), solo_numeros(unit), solo_numeros(tot)]) <= 1
        if desc and len(desc) >= 6 and pocos and (r + 1 < len(M)):
            fila2 = M[r + 1]
            desc2 = get_cell(fila2, idx_p)
            cant2, unit2, tot2 = get_cell(fila2, idx_c), get_cell(fila2, idx_u), get_cell(fila2, idx_t)
            senales2 = sum(bool(x) for x in [solo_numeros(cant2), solo_numeros(unit2), solo_numeros(tot2)])

            if senales2 >= 1 and (not desc2 or len(desc2) <= 3 or desc2.lower() in {"-", "—", "_"}):
                nueva = fila[:]  # copia
                if idx_c is not None: nueva[idx_c] = cant if solo_numeros(cant) else cant2
                if idx_u is not None: nueva[idx_u] = unit if parece_dinero(unit) else unit2
                if idx_t is not None: nueva[idx_t] = tot if parece_dinero(tot) else tot2
                out.append(nueva); r += 2; continue

        out.append(fila); r += 1

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
        h2 = [normalizar_espacios(f"{M[0][c] if c < len(M[0]) else ''} {M[1][c] if c < len(M[1]) else ''}") for c in range(ncol_real)]
    return h1, h2, ncol_real

def tabla_contiene_campos_requeridos(matriz: List[List[str]]) -> bool:
    """
    Filtro MUY ESTRICTO:
    - Detecta encabezado real
    - Exige al menos 4 de los 5 campos obligatorios
    - Valida que las columnas tengan datos coherentes debajo
    """

    if not matriz:
        return False

    M = normalizar_matriz(matriz)
    if not M or len(M) < 3:
        return False

    limite = min(6, len(M))

    for i in range(limite):

        filas_test = [M[i]]

        if i + 1 < len(M):
            filas_test.append(M[i + 1])

        indices = detectar_columnas_por_keywords(filas_test)

        # -------------------------------
        # 1️⃣ exigir al menos 4 columnas
        # -------------------------------
        if len(indices) < 4:
            continue

        # -------------------------------
        # columnas deben ser distintas
        # -------------------------------
        if len(set(indices.values())) < len(indices):
            continue

        # -------------------------------
        # validar datos reales debajo
        # -------------------------------
        filas_validas = 0
        revisar = min(i + 8, len(M))

        for r in range(i + 1, revisar):

            fila = M[r]

            desc = get_cell(fila, indices.get("Producto"))
            cant = get_cell(fila, indices.get("Cantidad"))
            unit = get_cell(fila, indices.get("Valor_Unitario"))
            tot  = get_cell(fila, indices.get("Valor_Total"))

            desc_ok = len(desc.strip()) >= 4
            cant_ok = bool(re.search(r"\d", cant))
            unit_ok = parece_dinero(unit)
            tot_ok  = parece_dinero(tot)

            if desc_ok and (cant_ok or unit_ok or tot_ok):
                filas_validas += 1

        # exigir mínimo 2 filas válidas
        if filas_validas >= 2:
            return True

    return False

def emitir_registros(M: List[List[str]],
                     nombre_pdf: str,
                     pagina: int,
                     mapeo: Dict[str, int],
                     fila_ini: int,
                     relajada: bool) -> List[Dict]:
    out: List[Dict] = []
    k = 0
    for r in range(fila_ini, len(M)):
        fila = M[r]
        if all(x == "" for x in fila): continue
        if omitir_fila(fila) or es_header_fila(fila): continue
        if not fila_valida(fila, mapeo, relajada=relajada): continue

        reg = {
            "Pdf": nombre_pdf, "Pagina": pagina, "Fila": None,
            "Numero": get_cell(fila, mapeo.get("Numero", 0)),
            "Producto": get_cell(fila, mapeo.get("Producto")),
            "Cantidad": get_cell(fila, mapeo.get("Cantidad")) if "Cantidad" in mapeo else "",
            "Valor_Unitario": get_cell(fila, mapeo.get("Valor_Unitario")) if "Valor_Unitario" in mapeo else "",
            "Valor_Total": get_cell(fila, mapeo.get("Valor_Total")) if "Valor_Total" in mapeo else "",
        }
        if not reg["Producto"]: continue
        k += 1; reg["Fila"] = k; out.append(reg)
    return out

def normalizar_y_emitir(matriz_cruda: List[List[str]],
                        nombre_pdf: str,
                        pagina: int,
                        estado: EstadoMapeo,
                        heur: HeuristicsConfig) -> List[Dict]:
    M = normalizar_matriz(matriz_cruda)
    if not M: return []

    h1, h2, ncol_real = construir_headers(M)
    m = mapear_por_header(h1); fila_ini = 1
    m2 = mapear_por_header(h2)
    if len(m2) > len(m): m = m2; fila_ini = 2

    # Fallback heurístico si faltan campos claves
    if "Producto" not in m or "Valor_Total" not in m:
        m = mapear_por_heuristica(M, fila_ini, heur.max_filas_perfil)

    # Detección de continuación
    en_cont = False
    
    if estado.m_prev:
        encabezado_detectado = buscar_encabezado_en_tabla(M)
        if encabezado_detectado is None:
            #continuacion de tabla anterior
            m = estado.m_prev.copy()
            fila_ini = 0
            en_cont = True
        else:
            m = encabezado_detectado

    # Usar mapeo previo si aún falta Producto/Total y cantidad de columnas compatible
    if ("Producto" not in m or "Valor_Total" not in m):
        if estado.m_prev and estado.ncol_prev and (abs(ncol_real - estado.ncol_prev) <= 2):
            m = estado.m_prev.copy(); fila_ini = 0; en_cont = True
        else:
            return []

    # Fusión de líneas colgantes
    if heur.fusion_lineas_colgantes and "Producto" in m:
        M = fusion_lineas_partidas(M, m)

    filas = emitir_registros(M, nombre_pdf, pagina, m, fila_ini,
                             relajada=(heur.validacion_relajada_en_continuacion and en_cont))

    if filas:
        estado.m_prev, estado.ncol_prev, estado.en_continuacion = m.copy(), ncol_real, en_cont

    return filas


# ===================================================================
# 6) LECTOR PDFPLUMBER (RUTA A)
# ===================================================================

class LectorPlumber:
    def __init__(self, estado: EstadoMapeo, heur: HeuristicsConfig, plumber_cfg: Optional[Dict[str, object]] = None):
        self.estado = estado
        self.heur = heur
        self.cfg = plumber_cfg or _DEFAULT_PLUMBER_CFG

    def extraer(self, ruta_pdf: str) -> Tuple[List[Dict], Set[int]]:
        nombre = Path(ruta_pdf).name
        resultados: List[Dict] = []
        faltantes: Set[int] = set()

        with pdfplumber.open(ruta_pdf) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                tablas = page.extract_tables(self.cfg) or []
                hubo = False
                for tbl in tablas:
                    try:
                        # --------------------------------------------------------
                        # FILTRO SEGURO:
                        # 1) Acepta tablas con encabezado completo (5 campos)
                        # 2) Acepta posibles CONTINUACIONES (si ya existe m_prev)
                        # 3) Descarta ruido sin romper multipágina
                        # --------------------------------------------------------
                        permitir = False
                        if tabla_contiene_campos_requeridos(tbl):
                            permitir = True
                        elif (
                            self.estado.m_prev is not None
                            and self.estado.ncol_prev is not None
                            and abs(len(tbl[0]) - self.estado.ncol_prev) <= 1
                        ):
                            permitir = True

                        if not permitir:
                            continue

                        # Lógica original de extracción (no modificada)
                        filas = normalizar_y_emitir(tbl, nombre, i, self.estado, self.heur)
                        if filas:
                            resultados.extend(filas); hubo = True
                    except Exception:
                        # Silencioso por robustez en lotes
                        pass
                if not hubo:
                    faltantes.add(i)

        return resultados, faltantes


# ===================================================================
# 7) LECTOR OCR (RUTA B)
# ===================================================================

class LectorOCR:
    def __init__(self, estado: EstadoMapeo, ocr: OCRConfig, paths: PathsConfig, heur: HeuristicsConfig):
        self.estado = estado
        self.ocr = ocr
        self.paths = paths
        self.heur = heur
        if paths.tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = paths.tesseract_cmd

    # ---- OCR por celda ----
    def _ocr_cell(self, img_pil: Image.Image) -> str:
        cfg = f"-l {self.ocr.idiomas} --oem 3 --psm {self.ocr.psm}"
        df = pytesseract.image_to_data(img_pil, config=cfg, output_type=pytesseract.Output.DATAFRAME)
        if df is None or df.empty: return ""
        df = df[df.text.notnull()]
        if df.empty: return ""
        df["conf"] = pd.to_numeric(df["conf"], errors="coerce")
        df = df[df["conf"] >= self.ocr.conf_minima]
        if df.empty: return ""
        t = normalizar_espacios(" ".join(str(x) for x in df["text"].tolist()))
        return re.sub(r"(?<=\d) +(?!,)", "", t)  # une dígitos separados por espacios (no antes de comas)

    # ---- Detección de grilla en ROI ----
    def _grid_roi(self, roi_bgr):
        gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)
        _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        h = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (bw.shape[1]//20, 1)), 2)
        v = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, bw.shape[0]//20)), 2)
        inter = cv2.bitwise_and(h, v)
        ys, xs = np.where(inter > 0)
        if len(xs) < 3 or len(ys) < 3: return [], []

        def agrupar(vals, tol=8):
            vals = sorted(map(int, vals))
            grupos = [[vals[0]]]
            for vv in vals[1:]:
                if abs(vv - grupos[-1][-1]) <= tol: grupos[-1].append(vv)
                else: grupos.append([vv])
            return [int(sum(u)/len(u)) for u in grupos]

        rows = sorted(set([0] + agrupar(ys, 8) + [roi_bgr.shape[0]-1]))
        cols = sorted(set([0] + agrupar(xs, 8) + [roi_bgr.shape[1]-1]))
        if len(rows) < 3 or len(cols) < 3: return [], []
        return rows, cols

    # ---- Detección de ROIs de tabla ----
    def _detectar_rois(self, bgr):
        gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        H, W = bw.shape
        h = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (W//25, 1)), 2)
        v = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, H//25)), 2)
        grid = cv2.add(h, v)
        cnts, _ = cv2.findContours(grid, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        rois = []
        for c in cnts:
            x, y, w, h = cv2.boundingRect(c)
            if w * h < H * W * 0.004 or w < W * 0.18 or h < H * 0.05:  # filtros mínimos
                continue
            rois.append((x, y, x + w, y + h))
        return sorted(rois, key=lambda r: (r[1], r[0]))

    # ---- OCR por ROI con grilla -> emitir items ----
    def _ocr_tabla_en_roi(self, img_pil, roi, nombre, pag):
        x0, y0, x1, y1 = roi
        roi_pil = img_pil.crop((x0, y0, x1, y1))
        roi_bgr = cv2.cvtColor(np.array(roi_pil), cv2.COLOR_RGB2BGR)
        rows, cols = self._grid_roi(roi_bgr)
        if not rows or not cols: return []

        nfil, ncol = len(rows) - 1, len(cols) - 1
        M = [["" for _ in range(ncol)] for _ in range(nfil)]
        for r in range(nfil):
            ya, yb = y0 + rows[r], y0 + rows[r+1]
            for c in range(ncol):
                xa, xb = x0 + cols[c], x0 + cols[c+1]
                px = max(2, (xb - xa)//50); py = max(2, (yb - ya)//40)
                xa2, xb2, ya2, yb2 = xa + px, xb - px, ya + py, yb - py
                if xb2 <= xa2 or yb2 <= ya2: continue
                M[r][c] = self._ocr_cell(img_pil.crop((xa2, ya2, xb2, yb2)))

        # --------------------------------------------------------
        # FILTRO SEGURO (idéntico a pdfplumber)
        # - Acepta encabezado completo
        # - Acepta CONTINUACIÓN (si hay m_prev/ncol_prev)
        # - Descarta el resto sin romper multipágina
        # --------------------------------------------------------
        permitir = False
        if tabla_contiene_campos_requeridos(M):
            permitir = True
        elif self.estado.m_prev is not None and self.estado.ncol_prev is not None:
            permitir = True  # continuación válida

        if not permitir:
            return []

        # Flujo original
        return normalizar_y_emitir(M, nombre, pag, self.estado, self.heur)

    # ---- Orquestar OCR ----
    def extraer(self, ruta_pdf: str, paginas: Optional[Set[int]] = None) -> List[Dict]:
        nombre = Path(ruta_pdf).name
        imgs = convert_from_path(ruta_pdf, dpi=self.ocr.dpi, poppler_path=self.paths.poppler_bin)
        out: List[Dict] = []
        for i, img in enumerate(imgs, start=1):
            if paginas and i not in paginas: continue
            bgr = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
            for roi in self._detectar_rois(bgr):
                filas = self._ocr_tabla_en_roi(img, roi, nombre, i)
                if filas: out.extend(filas)
        return out


# ===================================================================
# 8) EXPORTACIÓN A EXCEL
# ===================================================================

class ExtractorProductos:
    """
        extr = ExtractorProductos(paths=PathsConfig(), ocr=OCRConfig())
        filas = extr.extraer_productos("archivo.pdf")
        ruta = extr.exportar_excel(filas, "salida/")
    """
    def __init__(self,
                 paths: Optional[PathsConfig] = None,
                 ocr: Optional[OCRConfig] = None,
                 heur: Optional[HeuristicsConfig] = None):
        self.paths = paths or PathsConfig()
        self.ocr = ocr or OCRConfig()
        self.heur = heur or HeuristicsConfig(plumber_table_cfg=_DEFAULT_PLUMBER_CFG.copy())
        if self.heur.plumber_table_cfg is None:
            self.heur.plumber_table_cfg = _DEFAULT_PLUMBER_CFG.copy()
        if self.paths.tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = self.paths.tesseract_cmd  # Configurar Tesseract (opcional)

    def extraer_productos(self, ruta_pdf: str) -> List[Dict]:
        estado = EstadoMapeo()
        lector_plumber = LectorPlumber(estado, self.heur, self.heur.plumber_table_cfg)
        lector_ocr = LectorOCR(estado, self.ocr, self.paths, self.heur)

        res_plumber, faltantes = lector_plumber.extraer(ruta_pdf)
        if not res_plumber and not faltantes:
            datos = lector_ocr.extraer(ruta_pdf, paginas=None)
        else:
            datos = res_plumber + (lector_ocr.extraer(ruta_pdf, paginas=faltantes) if faltantes else [])

        # Orden y reenumeración por página
        datos = sorted(datos, key=lambda r: (r["Pagina"], r.get("Fila") or 0))
        k_by_page: Dict[int, int] = {}
        for d in datos:
            p = d["Pagina"]
            k_by_page[p] = k_by_page.get(p, 0) + 1
            d["Fila"] = k_by_page[p]
        return datos

    def exportar_excel(self, filas: List[Dict], carpeta_destino: str) -> str:
        df = pd.DataFrame(filas)
        for c in COLUMNAS_EXCEL:
            if c not in df.columns: df[c] = ""
        df = df[COLUMNAS_EXCEL]

        base = Path(carpeta_destino) / "productos_extraidos.xlsx"
        ruta = str(base)

        for intento in range(5):
            try:
                base.parent.mkdir(parents=True, exist_ok=True)
                with pd.ExcelWriter(ruta, engine="openpyxl") as w:
                    df.to_excel(w, index=False, sheet_name="PRODUCTOS")
                return ruta
            except PermissionError:
                print(f"[ADVERTENCIA] Archivo bloqueado. Reintentando ({intento+1}/5)...")
                time.sleep(1.2)

        alternativo = Path(carpeta_destino) / f"productos_extraidos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(str(alternativo), engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="PRODUCTOS")
        print("[OK] Guardado en archivo alternativo:", alternativo)
        return str(alternativo)