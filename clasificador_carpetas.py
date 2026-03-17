"""
clasificador_carpetas.py — Clasificación + extracción secuencial inmediata

Diseño:
- Clasifica cada carpeta (una por una)
- Si una carpeta es APTA (debe_extraerse=True), extrae sus PDFs inmediatamente
- No espera a clasificar todas para extraer
- No ejecuta extracción en paralelo (evita saturar CPU)
- Código claro, limpio y fácil de modificar

Mejoras:
① Si se detecta frase gatillo pero NINGÚN token de categoría coincide
  → se clasifica como MEDICAMENTOS por defecto y se extrae igual
② Pase 2: si no hay gatillo, busca tokens directamente para evitar
  SIN_DETECCION cuando "medicamentos" aparece sin frase de gatillo
③ El detalle incluye el PDF exacto y la página donde se detectó la frase
④ El resumen global no incluye la columna Debe_Extraerse
"""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
import pdfplumber
import unicodedata


# =============================================================================
# Logging limpio
# =============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfplumber").setLevel(logging.ERROR)


# =============================================================================
# Categorías — cámbielas aquí si necesita renombrarlas
# =============================================================================
CATEGORIA_MEDICAMENTOS = "MEDICAMENTOS"
CATEGORIA_DISPOSITIVOS = "DISPOSITIVOS"
CATEGORIA_INSUMOS      = "INSUMOS"
CATEGORIA_SIN          = "SIN_DETECCION"

# Categoría por defecto cuando se detecta un gatillo pero ningún token coincide.
CATEGORIA_FALLBACK = CATEGORIA_MEDICAMENTOS


# =============================================================================
# Tokens de categoría
#
# PRIORIDAD DE EVALUACIÓN: Medicamentos → Dispositivos → Insumos
# Agregue o quite términos según la realidad de sus documentos.
# Todos los tokens se comparan SIN acentos y en minúsculas.
# =============================================================================

TOKENS_MEDICAMENTOS: Tuple[str, ...] = (
    "medicamento", "medicamentos",
    "farmaco", "farmacos",
    "farmaceutico", "farmaceuticos",
    "farmacia", "farmacias",
    "ampolla", "ampollas",
    "tableta", "tabletas",
    "capsula", "capsulas",
    "jarabe", "jarabes",
    "solucion inyectable", "soluciones inyectables",
    "vacuna", "vacunas",
    "medicacion", "medicaciones",
    "principio activo", "principios activos",
    "especialidad farmaceutica",
)

TOKENS_DISPOSITIVOS: Tuple[str, ...] = (
    "dispositivo", "dispositivos",
    "dispositivo medico", "dispositivos medicos",
    "equipo medico", "equipos medicos",
    "equipo biomedico", "equipos biomedicos",
    "instrumento medico", "instrumentos medicos",
    "aparato medico", "aparatos medicos",
    "tecnologia medica", "tecnologias medicas",
    "material biomedico",
)

TOKENS_INSUMOS: Tuple[str, ...] = (
    "insumo", "insumos",
    "insumo medico", "insumos medicos",
    "material medico", "materiales medicos",
    "descartable", "descartables",
    "consumible", "consumibles",
    "suministro medico", "suministros medicos",
    "material quirurgico", "materiales quirurgicos",
    "material hospitalario",
)

# Frases que "activan" el análisis de categoría en esa línea.
GATILLOS: Tuple[str, ...] = (
    "adquisicion de",
    "contratacion de",
    "compra de",
    "suministro de",
    "provision de",
    "abastecimiento de",
)


# =============================================================================
# Columnas del Excel Detalle (en este orden exacto)
# =============================================================================
COLUMNAS_DETALLE = (
    "carpeta",
    "clasificacion",
    "frase_detectada",
    "motivo",
    "pdf_origen",
    "pagina_origen",
    "ruta",
    "debe_extraerse",
)

# Alias de tipo para mayor claridad
Detalle = Dict[str, object]


# =============================================================================
# Utilidades
# =============================================================================

def normalizar(texto: str) -> str:
    """Quita acentos y convierte a minúsculas para comparaciones robustas."""
    if not isinstance(texto, str):
        return ""
    sin_acentos = "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )
    return sin_acentos.lower()


def extraer_lineas(pagina) -> List[str]:
    """
    Extrae líneas de texto de una página pdfplumber.
    Si extract_text() no devuelve nada, intenta extract_words() como fallback.
    """
    txt    = pagina.extract_text() or ""
    lineas = [l for l in txt.splitlines() if l.strip()]
    if lineas:
        return lineas

    # Fallback: reconstruir línea a partir de palabras sueltas
    try:
        palabras = pagina.extract_words() or []
        linea    = " ".join(w.get("text", "") for w in palabras if w.get("text"))
        if linea.strip():
            return [linea]
    except Exception:
        pass

    return []


# =============================================================================
# Clasificador
# =============================================================================

class ClasificadorCarpetas:
    """
    Clasifica carpetas analizando frases en sus PDFs.

    Lógica de clasificación:
    Pase 1 — busca líneas con gatillo + token de categoría.
    Pase 2 — si el pase 1 no encontró nada, busca tokens directamente
             sin requerir gatillo. Evita SIN_DETECCION cuando aparece
             "medicamentos" o similar sin frase de gatillo delante.

    Solo llega a SIN_DETECCION si no hay ni gatillo ni token en ningún PDF.
    """

    def __init__(
        self,
        paginas_a_leer: int = 3,
        tokens_meds: Tuple[str, ...] = TOKENS_MEDICAMENTOS,
        tokens_disp: Tuple[str, ...] = TOKENS_DISPOSITIVOS,
        tokens_insu: Tuple[str, ...] = TOKENS_INSUMOS,
        gatillos: Tuple[str, ...] = GATILLOS,
        categoria_fallback: str = CATEGORIA_FALLBACK,
    ):
        self.paginas_a_leer     = max(1, int(paginas_a_leer))
        self.tokens_meds        = tokens_meds
        self.tokens_disp        = tokens_disp
        self.tokens_insu        = tokens_insu
        self.gatillos           = gatillos
        self.categoria_fallback = categoria_fallback

    # ------------------------------------------------------------------
    # Métodos internos
    # ------------------------------------------------------------------

    def _linea_tiene_gatillo(self, linea_normalizada: str) -> bool:
        """Devuelve True si la línea contiene alguna frase gatillo."""
        return any(g in linea_normalizada for g in self.gatillos)

    def _categoria_de_linea(self, linea_normalizada: str) -> Tuple[Optional[str], Optional[str]]:
        """
        Determina la categoría según los tokens presentes en la línea.
        Devuelve (categoria, token_encontrado) o (None, None) si no hay match.
        El orden de evaluación define la prioridad entre categorías:
        MEDICAMENTOS → DISPOSITIVOS → INSUMOS
        """
        grupos = [
            (self.tokens_meds, CATEGORIA_MEDICAMENTOS),
            (self.tokens_disp, CATEGORIA_DISPOSITIVOS),
            (self.tokens_insu, CATEGORIA_INSUMOS),
        ]
        for tokens, categoria in grupos:
            for tok in tokens:
                if tok in linea_normalizada:
                    return categoria, tok

        return None, None

    def _leer_lineas_pdf(self, pdf_path: Path) -> List[Tuple[int, str]]:
        """
        Lee las primeras `paginas_a_leer` páginas de un PDF.
        Devuelve lista de (num_pagina_base1, linea).
        """
        lineas: List[Tuple[int, str]] = []
        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                paginas_a_revisar = min(self.paginas_a_leer, len(pdf.pages))
                for num_pagina in range(paginas_a_revisar):
                    for linea in extraer_lineas(pdf.pages[num_pagina]):
                        lineas.append((num_pagina + 1, linea))
        except Exception as e:
            logging.error(f"[ERROR] No se pudo leer '{pdf_path.name}': {e}")
        return lineas

    # ------------------------------------------------------------------
    # Método público principal
    # ------------------------------------------------------------------

    def clasificar_carpeta_por_frase(self, ruta_carpeta: Path) -> Detalle:
        """
        Clasifica una carpeta revisando sus PDFs línea por línea.

        Devuelve un diccionario con:
        - carpeta, ruta, clasificacion, frase_detectada, motivo
        - pdf_origen: nombre del PDF donde se encontró la frase
        - pagina_origen: número de página (base 1)
        - debe_extraerse: True si la carpeta debe procesarse
        """
        ruta_carpeta = Path(ruta_carpeta)
        nombre       = ruta_carpeta.name

        pdfs = sorted(
            (p for p in ruta_carpeta.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"),
            key=lambda p: p.name,
        )

        # ------------------------------------------------------------------
        # Pase 1: gatillo + token (resultado más preciso)
        # ------------------------------------------------------------------
        for pdf_path in pdfs:
            for num_pagina, linea in self._leer_lineas_pdf(pdf_path):
                linea_norm = normalizar(linea)

                if not self._linea_tiene_gatillo(linea_norm):
                    continue

                # Gatillo encontrado → buscar categoría
                categoria, token = self._categoria_de_linea(linea_norm)

                if categoria is None:
                    # Gatillo sin token → usar fallback para no perder la carpeta
                    categoria = self.categoria_fallback
                    token     = "[fallback desde gatillo sin token]"

                return {
                    "carpeta":         nombre,
                    "ruta":            str(ruta_carpeta),
                    "clasificacion":   categoria,
                    "frase_detectada": linea.strip(),
                    "motivo":          token or "",
                    "pdf_origen":      pdf_path.name,
                    "pagina_origen":   num_pagina,
                    "debe_extraerse":  True,
                }

        # ------------------------------------------------------------------
        # Pase 2: sin gatillo, buscar tokens directamente.
        # Evita SIN_DETECCION cuando "medicamentos" aparece sin la frase
        # exacta de un gatillo ("adquisición de", "compra de", etc.).
        # ------------------------------------------------------------------
        for pdf_path in pdfs:
            for num_pagina, linea in self._leer_lineas_pdf(pdf_path):
                linea_norm = normalizar(linea)
                categoria, token = self._categoria_de_linea(linea_norm)

                if categoria is None:
                    continue  # línea sin tokens de interés

                return {
                    "carpeta":         nombre,
                    "ruta":            str(ruta_carpeta),
                    "clasificacion":   categoria,
                    "frase_detectada": linea.strip(),
                    "motivo":          f"[token directo] {token}",
                    "pdf_origen":      pdf_path.name,
                    "pagina_origen":   num_pagina,
                    "debe_extraerse":  True,
                }

        # Ningún gatillo ni token encontrado en ningún PDF
        return {
            "carpeta":         nombre,
            "ruta":            str(ruta_carpeta),
            "clasificacion":   CATEGORIA_SIN,
            "frase_detectada": "",
            "motivo":          "",
            "pdf_origen":      "",
            "pagina_origen":   "",
            "debe_extraerse":  False,
        }


# =============================================================================
# Generación del Excel de resumen
# =============================================================================

def _dataframe_resumen(detalles: Iterable[Detalle]) -> pd.DataFrame:
    """
    Hoja Resumen: vista resumida para revisión rápida.
    Columnas: Carpeta, Clasificación, Frase_Detectada, PDF_Origen, Página_Origen.
    La columna Debe_Extraerse se omite intencionalmente del resumen.
    """
    filas = [
        {
            "Carpeta":         d.get("carpeta", ""),
            "Clasificación":   d.get("clasificacion", ""),
            "Frase_Detectada": d.get("frase_detectada", ""),
            "PDF_Origen":      d.get("pdf_origen", ""),
            "Página_Origen":   d.get("pagina_origen", ""),
        }
        for d in detalles
    ]
    df = pd.DataFrame(filas)
    return df.sort_values(by=["Clasificación", "Carpeta"])


def _dataframe_detalle(detalles: Iterable[Detalle]) -> pd.DataFrame:
    """Hoja Detalle: todas las columnas para análisis completo."""
    df = pd.DataFrame(list(detalles))

    # Garantizar que todas las columnas existan aunque estén vacías
    for col in COLUMNAS_DETALLE:
        if col not in df.columns:
            df[col] = ""

    return df[list(COLUMNAS_DETALLE)].sort_values(by=["clasificacion", "carpeta"])


def generar_resumen_excel_desde_detalles(detalles: List[Detalle], ruta_excel: Path | str) -> None:
    """
    Exporta el Excel global con dos hojas:
    - Resumen: Carpeta, Clasificación, Frase_Detectada, PDF_Origen, Página_Origen
    - Detalle: todas las columnas incluyendo ruta y debe_extraerse
    """
    ruta_excel = Path(ruta_excel)
    ruta_excel.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        _dataframe_resumen(detalles).to_excel(writer, sheet_name="Resumen", index=False)
        _dataframe_detalle(detalles).to_excel(writer, sheet_name="Detalle", index=False)

    logging.info(f"✔ Resumen global generado: {ruta_excel}")