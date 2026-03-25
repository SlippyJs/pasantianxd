from __future__ import annotations

import logging
import re
import shutil
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import pdfplumber

# =============================================================================
# Logging  —  formato limpio, sin prefijos de nivel
# =============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

# =============================================================================
# Columnas del Excel de resumen
# =============================================================================
COLUMNAS_RESUMEN = (
    "nombre_carpeta",
    "medicamentos",
    "insumos",
    "ambiguos",
    "molecula_encontrada",
    "pagina",
    "insumo_encontrado",
)

Detalle = Dict[str, object]


# =============================================================================
# Utilidades
# =============================================================================
def normalizar(texto: str) -> str:
    """Quita tildes y convierte a minúsculas para comparación."""
    if not isinstance(texto, str):
        return ""
    texto = "".join(
        c
        for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )
    return texto.lower()


def extraer_lineas(pagina) -> List[str]:
    """Extrae líneas de texto de una página PDF."""
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


def _separador(caracter: str = "─", ancho: int = 80) -> str:
    return caracter * ancho


# =============================================================================
# Clasificador por moléculas
# =============================================================================
class ClasificadorCarpetas:

    def __init__(self):
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
            r"C:\Users\pasante.oper\OneDrive - Close-up International"
            r"\Escritorio\maestro moleculas\rm_cat_mol_uni.csv"
        )
        self.moleculas = self._leer_excel_moleculas(ruta_excel)

    # ------------------------------------------------------------------
    # Carga de moléculas
    # ------------------------------------------------------------------
    def _leer_excel_moleculas(self, ruta_archivo: Path) -> List[Dict[str, str]]:
        df = pd.read_csv(ruta_archivo, sep="|", dtype=str, encoding="latin-1")
        df["DESC_MOLECULA_UNI"] = df["DESC_MOLECULA_UNI"].str.strip().str.lower()
        df = df[
            df["DESC_MOLECULA_UNI"].notna()
            & (df["DESC_MOLECULA_UNI"].str.len() >= 4)
        ]
        moleculas = df.to_dict(orient="records")
        logging.info(f"Moléculas cargadas: {len(moleculas)}")
        return moleculas

    # ------------------------------------------------------------------
    # Prioridad de PDF
    # ------------------------------------------------------------------
    def _prioridad_pdf(self, nombre_pdf: str) -> int:
        nombre_norm = normalizar(nombre_pdf)
        for i, palabra in enumerate(self.prioridad_nombres):
            if palabra in nombre_norm:
                return i
        return len(self.prioridad_nombres)

    # ------------------------------------------------------------------
    # Búsqueda de moléculas en un PDF  (FASE 1)
    # Retorna: Dict[nombre_molecula, {"conteo": int, "paginas": [int]}]
    # ------------------------------------------------------------------
    def _buscar_moleculas_en_pdf(
        self, pdf_path: Path
    ) -> Dict[str, Dict]:
        encontradas: Dict[str, Dict] = {}

        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                total_paginas = len(pdf.pages)
                for num_pagina in range(total_paginas):          # ← TODAS las páginas
                    pagina = pdf.pages[num_pagina]
                    for linea in extraer_lineas(pagina):
                        linea_norm = re.sub(r"\s+", " ", normalizar(linea))

                        for mol in self.moleculas:
                            nombre_mol = mol["DESC_MOLECULA_UNI"]
                            # Coincidencia exacta con límites de palabra
                            patron = r"\b" + re.escape(nombre_mol) + r"\b"
                            if re.search(patron, linea_norm):
                                if nombre_mol not in encontradas:
                                    encontradas[nombre_mol] = {
                                        "conteo": 0,
                                        "paginas": [],
                                    }
                                encontradas[nombre_mol]["conteo"] += 1
                                pag_num = num_pagina + 1
                                if pag_num not in encontradas[nombre_mol]["paginas"]:
                                    encontradas[nombre_mol]["paginas"].append(pag_num)

        except Exception as e:
            logging.error(f"[ERROR] No se pudo abrir '{pdf_path.name}': {e}")

        return encontradas

    # ------------------------------------------------------------------
    # Búsqueda de insumos en un PDF  (FASE 2 — solo si no hay moléculas)
    # Retorna: Dict[keyword_insumo, {"conteo": int, "paginas": [int]}]
    # ------------------------------------------------------------------
    def _buscar_insumos_en_pdf(
        self, pdf_path: Path
    ) -> Dict[str, Dict]:
        encontrados: Dict[str, Dict] = {}

        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                total_paginas = len(pdf.pages)
                for num_pagina in range(total_paginas):          # ← TODAS las páginas
                    pagina = pdf.pages[num_pagina]
                    for linea in extraer_lineas(pagina):
                        linea_norm = re.sub(r"\s+", " ", normalizar(linea))

                        for kw in self.keywords_insumos:
                            patron = r"\b" + re.escape(kw) + r"\b"
                            if re.search(patron, linea_norm):
                                if kw not in encontrados:
                                    encontrados[kw] = {"conteo": 0, "paginas": []}
                                encontrados[kw]["conteo"] += 1
                                pag_num = num_pagina + 1
                                if pag_num not in encontrados[kw]["paginas"]:
                                    encontrados[kw]["paginas"].append(pag_num)

        except Exception as e:
            logging.error(f"[ERROR] No se pudo abrir '{pdf_path.name}': {e}")

        return encontrados

    # ------------------------------------------------------------------
    # Clasificación de una carpeta
    # ------------------------------------------------------------------
    def clasificar_carpeta(
        self, ruta_carpeta: Path, destino_base: Path
    ) -> Detalle:
        ruta_carpeta = Path(ruta_carpeta)
        nombre = ruta_carpeta.name

        # Crear carpetas destino
        destino_medicamentos = destino_base / "medicamentos"
        destino_insumos = destino_base / "insumos"
        destino_ambiguos = destino_base / "ambiguos"
        for d in (destino_medicamentos, destino_insumos, destino_ambiguos):
            d.mkdir(parents=True, exist_ok=True)

        clasificacion = "AMBIGUOS"
        molecula_encontrada = ""
        pagina_encontrada = ""
        insumo_encontrado = ""
        pdf_origen = ""

        # Ordenar PDFs por prioridad
        pdfs = sorted(
            [p for p in ruta_carpeta.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"],
            key=lambda p: (self._prioridad_pdf(p.name), p.name.lower()),
        )

        logging.info(_separador())
        logging.info(f"CARPETA: {nombre}  ({len(pdfs)} PDF{'s' if len(pdfs) != 1 else ''})")
        logging.info(_separador())

        ya_aviso_no_prioritarios = False

        for pdf_path in pdfs:
            es_prioritario = self._prioridad_pdf(pdf_path.name) < len(self.prioridad_nombres)

            if es_prioritario:
                logging.info(f'▶ PDF prioritario : "{pdf_path.name}"')
            else:
                if not ya_aviso_no_prioritarios:
                    logging.info("▶ PDFs restantes :")
                    ya_aviso_no_prioritarios = True
                logging.info(f'  · "{pdf_path.name}"')

            # ── FASE 1: Buscar moléculas ─────────────────────────────────
            mols = self._buscar_moleculas_en_pdf(pdf_path)

            if mols:
                # Imprimir cada molécula encontrada en el formato requerido
                for mol_nombre, datos in mols.items():
                    paginas_str = ",".join(str(p) for p in datos["paginas"])
                    logging.info(
                        f"  codigo del proceso: {nombre} ; "
                        f"molecula : {mol_nombre} ; "
                        f"n° de veces : {datos['conteo']} ; "
                        f"documento : {pdf_path.name} ; "
                        f"paginas : {paginas_str}"
                    )

                # Tomar la primera molécula encontrada (orden de aparición)
                primera_mol = next(iter(mols))
                primera_pag = mols[primera_mol]["paginas"][0]

                clasificacion = "MEDICAMENTOS"           # ← clasificación inmediata
                molecula_encontrada = primera_mol
                pagina_encontrada = str(primera_pag)
                pdf_origen = pdf_path.name

                logging.info(
                    f"  → MEDICAMENTOS  (molécula: '{primera_mol}', "
                    f"pág. {primera_pag})  — búsqueda detenida."
                )
                break   # ← detener al primer PDF con molécula

            # ── FASE 2: Buscar insumos (solo si no hubo moléculas) ───────
            insumos = self._buscar_insumos_en_pdf(pdf_path)

            if insumos:
                for kw, datos in insumos.items():
                    paginas_str = ",".join(str(p) for p in datos["paginas"])
                    logging.info(
                        f"  codigo del proceso: {nombre} ; "
                        f"insumo : {kw} ; "
                        f"n° de veces : {datos['conteo']} ; "
                        f"documento : {pdf_path.name} ; "
                        f"paginas : {paginas_str}"
                    )

                primer_insumo = next(iter(insumos))
                clasificacion = "INSUMOS"
                insumo_encontrado = primer_insumo
                pdf_origen = pdf_path.name

                logging.info(
                    f"  → INSUMOS  (keyword: '{primer_insumo}')  — búsqueda detenida."
                )
                break

            logging.info(f"  Sin coincidencias en '{pdf_path.name}' — siguiente PDF...")

        logging.info(f"Resultado final: {clasificacion}")
        logging.info(_separador("═"))

        # ── Seleccionar destino ──────────────────────────────────────────
        destino_final = {
            "MEDICAMENTOS": destino_medicamentos,
            "INSUMOS": destino_insumos,
            "AMBIGUOS": destino_ambiguos,
        }[clasificacion]

        destino_carpeta = destino_final / nombre
        if destino_carpeta.exists():
            contador = 1
            while True:
                tmp = destino_carpeta.parent / f"{nombre}_{contador}"
                if not tmp.exists():
                    destino_carpeta = tmp
                    break
                contador += 1

        shutil.move(str(ruta_carpeta), str(destino_carpeta))

        return {
            "nombre_carpeta": nombre,
            "medicamentos": nombre if clasificacion == "MEDICAMENTOS" else "",
            "insumos": nombre if clasificacion == "INSUMOS" else "",
            "ambiguos": nombre if clasificacion == "AMBIGUOS" else "",
            "molecula_encontrada": molecula_encontrada,
            "pagina": pagina_encontrada,
            "insumo_encontrado": insumo_encontrado,
        }


# =============================================================================
# Generar Excel de resumen
# =============================================================================
def generar_resumen_excel(detalles: List[Detalle], ruta_excel: Path) -> None:
    ruta_excel.parent.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame(detalles)

    # Garantizar todas las columnas aunque el DataFrame venga parcial
    for col in COLUMNAS_RESUMEN:
        if col not in df.columns:
            df[col] = ""

    df = df[list(COLUMNAS_RESUMEN)]

    # Ordenar: primero medicamentos, luego insumos, luego ambiguos
    orden_clasificacion = {"medicamentos": 0, "insumos": 1, "ambiguos": 2}
    df["_orden"] = df.apply(
        lambda r: orden_clasificacion.get(
            "medicamentos" if r["medicamentos"]
            else "insumos" if r["insumos"]
            else "ambiguos",
            3,
        ),
        axis=1,
    )
    df = df.sort_values(by=["_orden", "nombre_carpeta"]).drop(columns=["_orden"])

    df.to_excel(ruta_excel, index=False)

    # Imprimir resumen en consola
    total = len(df)
    n_med = df["medicamentos"].astype(bool).sum()
    n_ins = df["insumos"].astype(bool).sum()
    n_amb = df["ambiguos"].astype(bool).sum()

    logging.info("")
    logging.info(_separador("═"))
    logging.info("RESUMEN GLOBAL")
    logging.info(_separador("═"))
    logging.info(f"  Total carpetas procesadas : {total}")
    logging.info(f"  MEDICAMENTOS              : {n_med}")
    logging.info(f"  INSUMOS                   : {n_ins}")
    logging.info(f"  AMBIGUOS                  : {n_amb}")
    logging.info(_separador())
    logging.info(f"✔ Excel generado en: {ruta_excel}")
