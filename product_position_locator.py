import re
from pathlib import Path
import pdfplumber
import pandas as pd

class ProductPositionLocator:
    # ============================
    #  MÉTODO PRINCIPAL (RESUMEN)
    # ============================
    def localizar_resumen(self, ruta_pdf, productos):
        """
        Recorre el PDF y genera un resumen por tipo de hallazgo:
          - 'productos'
          - 'registro sanitario'
        Devuelve:
          resumen (list[dict]): una fila por hallazgo en ese PDF, con:
              Pdf, Fila (página), Descripcion
        """
        pdf_name = Path(ruta_pdf).name
        resumen = []

        # Prepara expresiones / patrones
        patrones_productos = self._preparar_patrones_productos(productos)
        patron_registro = re.compile(r"REGISTRO\s+SANITARIO", re.IGNORECASE)

        # “slots” para guardar la primera ocurrencia de cada tipo
        hallazgo_productos = None
        hallazgo_registro = None

        # Si el extractor ya encontró productos, registramos el hallazgo directamente
        if productos:
            hallazgo_productos = {
                "Pdf": pdf_name,
                "Fila": None,
                "Descripcion": "En este PDF se encontraron productos",
            }

        try:
            with pdfplumber.open(ruta_pdf) as pdf:
                for numero_pagina, pagina in enumerate(pdf.pages, start=1):

                    # -------- Texto por líneas --------
                    lineas = (pagina.extract_text() or "").split("\n")
                    for linea in lineas:
                        texto_norm = self._normalizar_linea(linea)

                        # 1) Registro sanitario por texto
                        if hallazgo_registro is None and patron_registro.search(linea or ""):
                            hallazgo_registro = {
                                "Pdf": pdf_name,
                                "Fila": numero_pagina,  # ahora indica la página
                                "Descripcion": "En este PDF se encontró el registro sanitario",
                            }

                        # 2) Productos por texto
                        if hallazgo_productos is None and self._coincide_con_cualquier_patron(
                            patrones_productos, texto_norm
                        ):
                            hallazgo_productos = {
                                "Pdf": pdf_name,
                                "Fila": numero_pagina,  # ahora indica la página
                                "Descripcion": "En este PDF se encontraron productos",
                            }

                        if hallazgo_productos and hallazgo_registro:
                            break

                    # -------- Tablas (si existen) --------
                    if not (hallazgo_productos and hallazgo_registro):
                        tablas = pagina.extract_tables() or []
                        for tabla in tablas:
                            for fila in tabla:
                                fila_texto = " ".join([c for c in (fila or []) if c]) if fila else ""
                                if not fila_texto:
                                    continue

                                fila_norm = self._normalizar_linea(fila_texto)

                                # 1) Registro sanitario en tabla
                                if hallazgo_registro is None and patron_registro.search(fila_texto):
                                    hallazgo_registro = {
                                        "Pdf": pdf_name,
                                        "Fila": numero_pagina,  # reemplaza fila de tabla
                                        "Descripcion": "En este PDF se encontró el registro sanitario",
                                    }

                                # 2) Productos en tabla
                                if hallazgo_productos is None and self._coincide_con_cualquier_patron(
                                    patrones_productos, fila_norm
                                ):
                                    hallazgo_productos = {
                                        "Pdf": pdf_name,
                                        "Fila": numero_pagina,  # reemplaza fila de tabla
                                        "Descripcion": "En este PDF se encontraron productos",
                                    }

                                if hallazgo_productos and hallazgo_registro:
                                    break
                            if hallazgo_productos and hallazgo_registro:
                                break

                    if hallazgo_productos and hallazgo_registro:
                        break

            # Construye el resumen (1 fila por cada tipo encontrado)
            if hallazgo_productos:
                resumen.append(hallazgo_productos)
            if hallazgo_registro:
                resumen.append(hallazgo_registro)

        except Exception as e:
            print(f"[ERROR en localizar_resumen]: {e}")

        return resumen

    # ============================
    #  EXPORTAR A EXCEL
    # ============================
    def exportar_resumen_excel(self, ruta_excel, resumen, hoja="Resumen"):
        """
        Escribe/actualiza un Excel con una única hoja 'Resumen' con columnas:
        Pdf, Fila, Descripcion
        """
        cols = ["Pdf", "Fila", "Descripcion"]
        df = pd.DataFrame(resumen) if resumen else pd.DataFrame(columns=cols)
        df = df[[c for c in cols if c in df.columns]]

        # Corrección para evitar error de if_sheet_exists
        if Path(ruta_excel).exists():
            modo = "a"
            if_sheet_exists = "replace"
        else:
            modo = "w"
            if_sheet_exists = None

        with pd.ExcelWriter(ruta_excel, engine="openpyxl", mode=modo, if_sheet_exists=if_sheet_exists) as writer:
            df.to_excel(writer, sheet_name=hoja, index=False)

    # ============================
    #        AUXILIARES
    # ============================

    def _preparar_patrones_productos(self, productos):
        patrones = []
        for prod in productos or []:
            texto = self._texto_busqueda(prod)
            if texto:
                patrones.append(texto)
        return patrones

    def _texto_busqueda(self, prod):
        campos = ["Nombre_Generico", "Cantidad", "Valor_Unitario", "Valor_Total"]
        partes = []
        for c in campos:
            v = prod.get(c)
            if v is None:
                continue
            partes.append(str(v))

        texto = " ".join(partes).upper()
        texto = re.sub(r"\s+", " ", texto).strip()
        texto = texto.replace(".", "").replace(",", ".")
        return texto

    def _normalizar_linea(self, s):
        s = (s or "").upper()
        s = re.sub(r"\s+", " ", s).strip()
        s = s.replace(".", "").replace(",", ".")
        return s

    def _coincide_con_cualquier_patron(self, patrones, linea_norm):
        for patron in patrones:
            tokens = [t for t in patron.split(" ") if t]
            if not tokens:
                continue
            encontrados = sum(1 for t in tokens if t in linea_norm)
            if encontrados >= max(2, len(tokens) // 2):
                return True
        return False