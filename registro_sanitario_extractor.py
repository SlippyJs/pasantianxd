from __future__ import annotations
import re
from pathlib import Path
from typing import List, Dict, Optional, Set
import pdfplumber


class RegistroSanitarioExtractor:

    def __init__(self) -> None:

        # -------------------------------
        # PRODUCTO
        # -------------------------------
        self.rx_producto = re.compile(
            r"(?:Producto\s+denominado|Nombre\s+de(?:l)?\s+Producto|Denominaci[oó]n\s+de(?:l)?\s+Producto)\s*:\s*(?P<val>.+)",
            re.IGNORECASE
        )

        # -------------------------------
        # REGISTRO SANITARIO
        # -------------------------------
        self.rx_registro_linea = re.compile(
            r"REGISTRO\s+SANITARIO\s*No\.?\s*:?\s*(?P<code>[A-Z0-9\-.]+)",
            re.IGNORECASE
        )

        # -------------------------------
        # FORMATO DE CODIGOS ARCSA
        # -------------------------------
        self.rx_codigo = re.compile(
            r"(?:"
            r"[A-Z]{1,3}\d{6,10}"
            r"|GN\d{6,10}"
            r"|\d{2,6}-M[AE]N-\d{2,6}"
            r"|\d{2,6}-MAC-\d{2,6}"
            r"|GBN-\d{3,4}-\d{2}-\d{2}"
            r"|\d{2,5}-\d{2}-\d{2}"
            r")",
            re.IGNORECASE
        )

    # -------------------------------------------------
    # LIMPIEZA
    # -------------------------------------------------

    @staticmethod
    def _clean(text: str) -> str:
        text = (text or "").replace("*", " ")
        text = text.replace("\u200f", " ").replace("\u200e", " ")
        return re.sub(r"[ \t]+", " ", text).strip()

    @staticmethod
    def _norm_producto(text: str) -> str:
        text = re.sub(r"\s+", " ", text)
        return text.strip(" .:;\"'*_[]")

    @staticmethod
    def _norm_codigo(text: str) -> str:
        text = re.sub(r"\s+", "-", text)
        text = re.sub(r"\s*-\s*", "-", text)
        return text.upper().strip()

    # -------------------------------------------------
    # BUSCAR CODIGO CERCA DEL ENCABEZADO
    # -------------------------------------------------

    def _buscar_codigo_cercano(self, texto: str) -> Optional[str]:

        m = self.rx_registro_linea.search(texto)
        if m:
            return self._norm_codigo(m.group("code"))

        pos = texto.upper().find("REGISTRO SANITARIO")

        if pos == -1:
            return None

        fragmento = texto[pos:pos + 120]

        m2 = self.rx_codigo.search(fragmento)

        if m2:
            return self._norm_codigo(m2.group())

        return None

    # -------------------------------------------------
    # EXTRACCION PRINCIPAL
    # -------------------------------------------------

    def extraer_registros_y_productos_multi(self, ruta_pdf: str) -> List[Dict[str, str]]:

        resultados: List[Dict[str, str]] = []
        vistos: Set[tuple[str, str]] = set()

        path = Path(ruta_pdf)

        if not path.exists():
            print(f"No existe archivo: {ruta_pdf}")
            return resultados

        try:

            with pdfplumber.open(path) as pdf:

                registro_actual: Optional[str] = None
                productos: List[str] = []

                for page in pdf.pages:

                    texto = self._clean(page.extract_text() or "")

                    # -------------------------------------
                    # BUSCAR REGISTRO EN TEXTO
                    # -------------------------------------
                    codigo = self._buscar_codigo_cercano(texto)

                    if codigo:
                        registro_actual = f"REGISTRO SANITARIO No. {codigo}"

                    # -------------------------------------
                    # PRODUCTOS EN TEXTO
                    # -------------------------------------
                    for m in self.rx_producto.finditer(texto):

                        prod = self._norm_producto(m.group("val"))

                        if prod:
                            productos.append(prod)

                    # -------------------------------------
                    # TABLAS
                    # -------------------------------------
                    tablas = page.extract_tables()

                    if tablas:

                        for tabla in tablas:

                            for fila in tabla:

                                for celda in fila:

                                    if not celda:
                                        continue

                                    celda = self._clean(celda)

                                    # producto
                                    m = self.rx_producto.search(celda)

                                    if m:
                                        prod = self._norm_producto(m.group("val"))

                                        if prod:
                                            productos.append(prod)

                                    # registro
                                    codigo = self._buscar_codigo_cercano(celda)

                                    if codigo:
                                        registro_actual = f"REGISTRO SANITARIO No. {codigo}"

                # -------------------------------------
                # CREAR RESULTADOS
                # -------------------------------------
                if registro_actual:

                    for p in productos:

                        key = (p, registro_actual)

                        if key not in vistos:

                            resultados.append({
                                "producto": p,
                                "registro": registro_actual
                            })

                            vistos.add(key)

        except Exception as e:
            print("Error procesando PDF:", e)

        return resultados

    # -------------------------------------------------
    # METODOS RAPIDOS
    # -------------------------------------------------

    def extraer_registro_y_producto(self, ruta_pdf: str) -> Dict[str, Optional[str]]:

        pares = self.extraer_registros_y_productos_multi(ruta_pdf)

        if not pares:
            return {"registro": None, "producto": None}

        return pares[0]

    def extraer_registro(self, ruta_pdf: str) -> Optional[str]:

        r = self.extraer_registro_y_producto(ruta_pdf)

        return r.get("registro")

    def extraer_nombre_producto(self, ruta_pdf: str) -> Optional[str]:

        r = self.extraer_registro_y_producto(ruta_pdf)

        return r.get("producto")