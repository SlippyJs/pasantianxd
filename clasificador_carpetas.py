from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Dict, List
import pandas as pd
import pdfplumber
import unicodedata
import shutil
import re

# =============================================================================
# Logging
# =============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",   
    handlers=[logging.StreamHandler(sys.stdout)],
)

# =============================================================================
# Columnas Excel Resumen
# =============================================================================
COLUMNAS_DETALLE = (
    "carpeta",
    "clasificacion",
    "molecula_encontrada",
    "pdf_origen",
    "pagina_origen",
    "ruta_origen",
    "ruta_destino",
)

Detalle = Dict[str, object]

# =============================================================================
# Utilidades
# =============================================================================
def normalizar(texto: str) -> str:
    if not isinstance(texto, str):
        return ""
    texto = "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )
    return texto.lower()

def extraer_lineas(pagina) -> List[str]:
    try:
        txt = pagina.extract_text() or ""
        lineas = [l for l in txt.splitlines() if l.strip()]
        if lineas:
            return lineas
        palabras = pagina.extract_words() or []
        linea = " ".join(w.get("text", "") for w in palabras)
        if linea.strip():
            return [linea]
    except Exception as e:
        logging.warning(f"[WARN] No se pudo extraer texto de la página: {e}")
    return []


# =============================================================================
# Clasificador por moléculas
# =============================================================================
class ClasificadorCarpetas:

    def __init__(self, paginas_a_leer: int = 3):
        self.paginas_a_leer = max(1, int(paginas_a_leer))
        self.prioridad_nombres = [
            "contrato",
            "resolucion de adjudicacion",
            "factura",
            "tabla de valores y precios",
            "ofertas",
            "especificacion tecnica",
            "informe de necesidad",
            "proforma",
            "registro sanitario",
            "reg san",
        ]
        self.keywords_insumos = [
            "insumo",
            "insumos",
            "kit",
            "kits",
            "reactivo",
            "reactivos",
            "material",
            "materiales",
            "dispositivo",
            "dispositivos",
        ]


        # Ruta fija del Excel con moléculas
        ruta_excel = Path(
            r"C:\Users\pasante.oper\OneDrive - Close-up International\Escritorio\maestro moleculas"
        )
        self.moleculas = self._leer_excel_moleculas(ruta_excel)
        

    def _leer_excel_moleculas(self, ruta_archivo: Path) -> List[Dict[str, str]]:
        df = pd.read_csv(ruta_archivo, sep="|", dtype=str)
        
        df.columns = ["COD_MOLECULA_ASSOC", "DESC_MOLECULA_ASOC", "COD_MOLECULA_ASSOC_2",
                    "DESC_MOLECULA_ASSOC_ENG", "DESC_MOLECULA_ASSOC_POR"]
        
        df["DESC_MOLECULA_ASOC"] = df["DESC_MOLECULA_ASOC"].str.strip().str.lower()
        df = df[df["DESC_MOLECULA_ASOC"].notna() & (df["DESC_MOLECULA_ASOC"].str.len() >= 4)]

        moleculas = df.to_dict(orient="records")
        logging.info(f"Moléculas cargadas: {len(moleculas)}")
        return moleculas
    
    def _prioridad_pdf(self, nombre_pdf: str) -> int:
        nombre_norm = normalizar(nombre_pdf)

        for i, palabra in enumerate(self.prioridad_nombres):
            if palabra in nombre_norm:
                return i  # mientras más pequeño, más prioridad

        return len(self.prioridad_nombres)  # sin prioridad



    def clasificar_carpeta_por_frase(self, carpeta: Path, destino_base: Path = Path.cwd () ) -> Detalle:
        return self.clasificar_carpeta(carpeta, destino_base)
    
    def clasificar_carpeta(self, ruta_carpeta: Path, destino_base: Path) -> Detalle:
        ruta_carpeta = Path(ruta_carpeta)
        nombre = ruta_carpeta.name

        # Crear carpetas globales de destino si no existen
        destino_medicamentos = destino_base / "medicamentos"
        destino_medicamentos.mkdir(parents=True, exist_ok=True)

        destino_ambiguos = destino_base / "ambiguos"
        destino_ambiguos.mkdir(parents=True, exist_ok=True)

        destino_insumos = destino_base / "insumos"
        destino_insumos.mkdir(parents=True, exist_ok=True)

        # Variables para registrar resultados
        clasificacion = "AMBIGUOS"
        molecula_encontrada = ""
        pdf_origen = ""
        pagina_origen = ""

        # Obtener PDFs
        pdfs = [p for p in ruta_carpeta.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]

        # Separar PDFs prioritarios
        pdfs_prioritarios = sorted(
            [p for p in pdfs if self._prioridad_pdf(p.name) < len(self.prioridad_nombres)],
            key=lambda p: (self._prioridad_pdf(p.name), p.name.lower())
        )

        # Separar PDFs no prioritarios
        pdfs_no_prioritarios = sorted(
            [p for p in pdfs if self._prioridad_pdf(p.name) == len(self.prioridad_nombres)],
            key=lambda p: p.name.lower()
        )

        # Unir: primero prioritarios
        pdfs = pdfs_prioritarios + pdfs_no_prioritarios

        ya_aviso_no_prioritarios = False

        for pdf_path in pdfs:
            prioridad = self._prioridad_pdf(pdf_path.name)
            es_prioritario = prioridad < len(self.prioridad_nombres)

            # Logs controlados
            if es_prioritario:
                logging.info(f'🔍 Buscando en PDF prioritario: "{pdf_path.name}"')
            elif not ya_aviso_no_prioritarios:
                logging.info("🔎 Buscando en PDFs restantes...")
                ya_aviso_no_prioritarios = True

            try:
                with pdfplumber.open(str(pdf_path)) as pdf:
                    paginas_a_leer = min(self.paginas_a_leer, len(pdf.pages))

                    for num_pagina in range(paginas_a_leer):
                        pagina = pdf.pages[num_pagina]

                        for linea in extraer_lineas(pagina):
                            linea_norm = normalizar(linea)
                            linea_norm = re.sub(r'\s+', ' ', linea_norm)

                            #  1. BUSCAR INSUMOS (PRIORIDAD MÁS ALTA)
                            for kw in self.keywords_insumos:
                                if kw in linea_norm:
                                    clasificacion = "INSUMOS"
                                    molecula_encontrada = kw
                                    pdf_origen = pdf_path.name
                                    pagina_origen = num_pagina + 1

                                    if es_prioritario:
                                        logging.info(
                                            f"Insumo detectado: {kw} "
                                            f"(PDF: {pdf_origen}, Página: {pagina_origen})"
                                        )
                                    break

                            if clasificacion == "INSUMOS":
                                break

                            # 2. BUSCAR MOLÉCULAS
                            for molecula in self.moleculas:
                                patron = r'\b' + re.escape(molecula['DESC_MOLECULA_ASOC']) + r'\b'
                                if re.search(patron, linea_norm):

                                    clasificacion = "MEDICAMENTOS"
                                    molecula_encontrada = molecula
                                    pdf_origen = pdf_path.name
                                    pagina_origen = num_pagina + 1

                                    if es_prioritario:
                                        logging.info(
                                            f" Molécula encontrada: {molecula['DESC_MOLECULA_ASOC']} "
                                            f"(PDF: {pdf_origen}, Página: {pagina_origen})"
                                        )
                                    break

                            if clasificacion == "MEDICAMENTOS":
                                break

                        if clasificacion in ["MEDICAMENTOS", "INSUMOS"]:
                            break

                    if clasificacion in ["MEDICAMENTOS", "INSUMOS"]:
                        break

            except Exception as e_pdf:
                logging.error(f"[ERROR] No se pudo abrir '{pdf_path.name}': {e_pdf}")

        # Determinar destino final
        if clasificacion == "MEDICAMENTOS":
            destino_final = destino_medicamentos
        elif clasificacion == "INSUMOS":
            destino_final = destino_insumos
        else:
            destino_final = destino_ambiguos

        destino_carpeta = destino_final / nombre

        # Evitar sobreescritura
        if destino_carpeta.exists():
            contador = 1
            while True:
                destino_carpeta_temp = destino_carpeta.parent / f"{nombre}_{contador}"
                if not destino_carpeta_temp.exists():
                    destino_carpeta = destino_carpeta_temp
                    break
                contador += 1

        # Mover carpeta
        logging.info(f"Moviendo carpeta '{nombre}' → {destino_carpeta}")
        shutil.move(str(ruta_carpeta), str(destino_carpeta))

        return {
            "carpeta": nombre,
            "clasificacion": clasificacion,
            "molecula_encontrada": molecula_encontrada,
            "pdf_origen": pdf_origen,
            "pagina_origen": pagina_origen,
            "ruta_origen": str(ruta_carpeta),
            "ruta_destino": str(destino_carpeta),
        }
# =============================================================================
# Generar Excel Resumen
# =============================================================================
def generar_resumen_excel(detalles: List[Detalle], ruta_excel: Path):
    ruta_excel.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(detalles)
    for col in COLUMNAS_DETALLE:
        if col not in df.columns:
            df[col] = ""
    df = df[list(COLUMNAS_DETALLE)].sort_values(by=["clasificacion", "carpeta"])
    df.to_excel(ruta_excel, index=False)
    logging.info(f"✔ Resumen global generado: {ruta_excel}")