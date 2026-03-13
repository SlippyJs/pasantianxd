"""
main.py — Orquestador principal

Flujo por cada subcarpeta:
1. Clasificar → lee PDFs y detecta frase gatillo + categoría
2. Si la categoría está en CATEGORIAS_A_EXTRAER → extraer inmediatamente
3. Al finalizar todas las carpetas → generar Excel resumen global

Configuración centralizada al inicio del archivo.
"""

from __future__ import annotations

import logging
import sys
import traceback
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

from clasificador_carpetas import (
    ClasificadorCarpetas,
    generar_resumen_excel_desde_detalles,
    CATEGORIA_SIN,
)
from product_extractor import ExtractorProductos, PathsConfig, OCRConfig, HeuristicsConfig
from registro_sanitario_extractor import RegistroSanitarioExtractor
from product_position_locator import ProductPositionLocator
from zip_utils import extraer_pdfs_de_zip, limpiar_zipcache


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
# CONFIGURACIÓN GLOBAL
# Cambie estos valores según su entorno. No toque el resto del código.
# =============================================================================

# Categorías que activan la extracción de productos
CATEGORIAS_A_EXTRAER = {"MEDICAMENTOS"}

# Si es True, las carpetas SIN_DETECCION también se extraen
EXTRAER_SIN_DETECCION = True

# Cuántas páginas leer por PDF durante la clasificación (más páginas = más lento pero más seguro)
PAGINAS_A_LEER = 3

# Rutas de herramientas externas
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_BIN    = r"C:\Users\DETPC\OneDrive\Desktop\tesseract\poppler-24.02.0\Library\bin"


# =============================================================================
# Utilidades
# =============================================================================

def solicitar_ruta_base_interactiva() -> Path:
    """Solicita una ruta válida al usuario por consola."""
    while True:
        ruta_ingresada = input("\nIngrese la ruta base con carpetas a procesar: ").strip('"').strip()
        ruta = Path(ruta_ingresada)
        if ruta.exists() and ruta.is_dir():
            return ruta
        logging.error("❌ La ruta ingresada no existe o no es carpeta.")


def convertir_a_entero(valor, por_defecto: int = 0) -> int:
    """Convierte un valor a entero de forma segura."""
    try:
        if valor is None or valor == "":
            return por_defecto
        return int(float(valor))
    except Exception:
        return por_defecto


def _debe_extraer(clasificacion: str) -> bool:
    """
    Decide si una carpeta clasificada debe ser extraída.
    Centraliza la lógica de decisión en un solo lugar.
    """
    if clasificacion in CATEGORIAS_A_EXTRAER:
        return True
    if clasificacion == CATEGORIA_SIN and EXTRAER_SIN_DETECCION:
        return True
    return False


# =============================================================================
# Creación de extractores
# Cambie los parámetros aquí si necesita ajustar OCR u heurísticas.
# =============================================================================

def crear_extractor_productos() -> ExtractorProductos:
    """Instancia el extractor de productos con la configuración global."""
    return ExtractorProductos(
        paths=PathsConfig(
            tesseract_cmd=TESSERACT_PATH,
            poppler_bin=POPPLER_BIN,
        ),
        ocr=OCRConfig(
            dpi=260,
            idiomas="spa+eng",
            conf_minima=35,
            psm=6,
        ),
        heur=HeuristicsConfig(
            tol_cols_suave=1,
            tol_cols_fuerte=2,
            validacion_relajada_en_continuacion=True,
            fusion_lineas_colgantes=True,
            plumber_table_cfg=None,
        ),
    )


# =============================================================================
# Extracción de una carpeta
# =============================================================================

def _recopilar_pdfs_de_carpeta(carpeta: Path) -> List[tuple[str, str]]:
    """
    Recopila todos los PDFs de una carpeta (sueltos y dentro de ZIPs).
    Devuelve lista de (ruta_pdf, origen) donde origen es "directo" o el nombre del ZIP.
    """
    lista: List[tuple[str, str]] = []

    # PDFs sueltos
    for p in carpeta.iterdir():
        if p.is_file() and p.suffix.lower() == ".pdf":
            lista.append((str(p), "directo"))

    # PDFs dentro de ZIPs
    zipcache = carpeta / "__zipcache"
    limpiar_zipcache(zipcache)

    for p in carpeta.iterdir():
        if p.is_file() and p.suffix.lower() == ".zip":
            destino = zipcache / p.stem
            pdfs_extraidos = extraer_pdfs_de_zip(p, destino)
            for pdf in pdfs_extraidos:
                lista.append((str(pdf), p.name))

    return lista


def procesar_carpeta_para_extraccion(ruta_carpeta: str) -> Optional[str]:
    """
    Extrae productos, registros sanitarios y resumen de hallazgos de una carpeta.
    Devuelve la ruta del Excel generado, o None si no hay datos.
    """
    carpeta = Path(ruta_carpeta)
    if not carpeta.exists():
        logging.error(f"⚠ No existe la carpeta: {carpeta}")
        return None

    lista_pdfs = _recopilar_pdfs_de_carpeta(carpeta)
    zipcache   = carpeta / "__zipcache"

    if not lista_pdfs:
        logging.info(f"⏳ {carpeta.name}: sin PDFs ni ZIPs, se omite")
        return None

    logging.info(f"⏳ Extrayendo: {carpeta.name}")

    extractor_productos = crear_extractor_productos()
    extractor_registro  = RegistroSanitarioExtractor()
    locator             = ProductPositionLocator()

    productos_por_pdf: Dict[str, List[Dict]] = {}
    registros_sanitarios: List[Dict] = []

    # -------------------------------------------------------------------------
    # Extracción PDF por PDF
    # -------------------------------------------------------------------------
    for ruta_pdf, origen in lista_pdfs:
        nombre_pdf = Path(ruta_pdf).name

        # Productos
        try:
            productos = extractor_productos.extraer_productos(ruta_pdf) or []
        except Exception as e:
            logging.error(f"  Error extrayendo productos en {nombre_pdf}: {e}")
            traceback.print_exc()
            productos = []

        # Registro sanitario + nombre de producto
        try:
            info_rs      = extractor_registro.extraer_registro_y_producto(ruta_pdf) or {}
            registro     = info_rs.get("registro") or ""
            producto_rs  = info_rs.get("producto") or ""
        except Exception as e:
            logging.error(f"  Error extrayendo registro/nombre en {nombre_pdf}: {e}")
            registro    = ""
            producto_rs = ""

        # Enriquecer cada producto con registro sanitario y origen
        for item in productos:
            item["Registro_Sanitario"] = registro
            item["Origen"]             = origen
            if producto_rs:
                item["Producto_RS"] = producto_rs

        if productos:
            productos_por_pdf[nombre_pdf] = productos

        if registro or producto_rs:
            registros_sanitarios.append({
                "Pdf":                nombre_pdf,
                "Origen":             origen,
                "Registro_Sanitario": registro,
                "Producto_RS":        producto_rs,
            })

    if not productos_por_pdf and not registros_sanitarios:
        logging.info(f"✔ Extraída: {carpeta.name} (sin datos)")
        limpiar_zipcache(zipcache)
        return None

    # -------------------------------------------------------------------------
    # Unificar y ordenar todos los productos
    # -------------------------------------------------------------------------
    todos_prod: List[Dict] = [p for lista in productos_por_pdf.values() for p in lista]
    todos_prod.sort(key=lambda r: (
        str(r.get("Pdf") or ""),
        convertir_a_entero(r.get("Pagina")),
        convertir_a_entero(r.get("Fila")),
    ))

    carpeta_salida = carpeta / "resultados"
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    # -------------------------------------------------------------------------
    # Exportar Excel principal de productos
    # -------------------------------------------------------------------------
    try:
        ruta_excel = extractor_productos.exportar_excel(todos_prod, str(carpeta_salida))
    except Exception as e:
        logging.error(f"  Error exportando Excel en {carpeta.name}: {e}")
        traceback.print_exc()
        limpiar_zipcache(zipcache)
        return None

    # -------------------------------------------------------------------------
    # Hoja "Resumen_Hallazgos"
    # -------------------------------------------------------------------------
    resumen_por_pdf: List[Dict] = []
    for ruta_pdf, _ in lista_pdfs:
        nombre_pdf = Path(ruta_pdf).name
        productos  = productos_por_pdf.get(nombre_pdf, [])
        resumen_por_pdf.extend(locator.localizar_resumen(ruta_pdf, productos))

    if resumen_por_pdf:
        ruta_resumen = carpeta_salida / "Resumen_Hallazgos.xlsx"
        locator.exportar_resumen_excel(ruta_resumen, resumen_por_pdf)
        logging.info(f"✔ Resumen exportado: {ruta_resumen}")

    # -------------------------------------------------------------------------
    # Hoja "REGISTROS_SANITARIOS" dentro del mismo Excel
    # -------------------------------------------------------------------------
    if registros_sanitarios:
        try:
            df_rs = pd.DataFrame(registros_sanitarios)[
                ["Pdf", "Origen", "Registro_Sanitario", "Producto_RS"]
            ]
            with pd.ExcelWriter(ruta_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
                df_rs.to_excel(w, index=False, sheet_name="REGISTROS_SANITARIOS")
        except Exception:
            logging.warning("No se pudo escribir la hoja REGISTROS_SANITARIOS")

    limpiar_zipcache(zipcache)
    logging.info(f"✔ Extraída: {carpeta.name}")
    return ruta_excel


# =============================================================================
# Orquestador principal
# =============================================================================

def ejecutar_principal(argv=None) -> None:
    """
    Punto de entrada principal.

    Argumentos opcionales por línea de comandos:
    --ruta_base    Ruta a la carpeta raíz con subcarpetas (si no se pasa, se pide por consola)
    --salida_excel Ruta del Excel resumen global (default: <ruta_base>/resumen_global.xlsx)
    --paginas      Páginas a leer por PDF durante clasificación (default: PAGINAS_A_LEER)
    """
    import argparse

    parser = argparse.ArgumentParser(description="Clasificador y extractor de carpetas")
    parser.add_argument("--ruta_base",    type=str, default=None)
    parser.add_argument("--salida_excel", type=str, default=None)
    parser.add_argument("--paginas",      type=int, default=PAGINAS_A_LEER)
    args = parser.parse_args(argv or sys.argv[1:])

    # Resolver ruta base
    if args.ruta_base:
        ruta_base = Path(args.ruta_base)
        if not ruta_base.exists():
            ruta_base = solicitar_ruta_base_interactiva()
    else:
        ruta_base = solicitar_ruta_base_interactiva()

    salida_global = Path(args.salida_excel) if args.salida_excel else (ruta_base / "resumen_global.xlsx")

    subcarpetas = sorted(
        (p for p in ruta_base.iterdir() if p.is_dir()),
        key=lambda p: p.name,
    )

    if not subcarpetas:
        logging.info("No hay subcarpetas para procesar.")
        return

    clasificador = ClasificadorCarpetas(paginas_a_leer=args.paginas)
    detalles: List[Dict] = []

    # =========================================================================
    # Clasificar → si es apta, extraer inmediatamente
    # =========================================================================
    for carpeta in subcarpetas:
        logging.info(f"\n{'─'*60}")
        logging.info(f"Carpeta: {carpeta.name}")
        logging.info("Clasificando...")

        resultado      = clasificador.clasificar_carpeta_por_frase(carpeta)
        clasificacion  = resultado["clasificacion"]
        frase          = resultado.get("frase_detectada") or ""
        pdf_origen     = resultado.get("pdf_origen") or ""
        pagina_origen  = resultado.get("pagina_origen") or ""

        logging.info(f"Clasificación : {clasificacion}")
        if frase:
            logging.info(f"Frase         : {frase}")
        if pdf_origen:
            logging.info(f"PDF origen    : {pdf_origen}  |  Página: {pagina_origen}")

        debe_extraer = _debe_extraer(clasificacion)

        if debe_extraer:
            logging.info("→ Extrayendo...")
            try:
                procesar_carpeta_para_extraccion(str(carpeta))
            except Exception as e:
                logging.error(f"Error durante extracción de '{carpeta.name}': {e}")
                traceback.print_exc()
        else:
            logging.info("→ No se extrae según reglas configuradas.")

        # Guardar todo lo que devolvió el clasificador + decisión de extracción
        detalles.append({
            **resultado,                         # carpeta, ruta, clasificacion, frase, motivo, pdf_origen, pagina_origen
            "debe_extraerse": debe_extraer,      # sobreescribe con la decisión real del main
        })

    # =========================================================================
    # Excel resumen global
    # =========================================================================
    generar_resumen_excel_desde_detalles(detalles, salida_global)
    logging.info("\n✔ Proceso completado.")


if __name__ == "__main__":
    ejecutar_principal()