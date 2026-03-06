"""
main.py
=======
Programa principal. Orquesta el flujo completo:

    1. Pide la carpeta de PDFs al usuario
    2. Por cada PDF llama a ExtractorTablasPDF  (extractor.py)
    3. Pasa las filas al UbicadorYGeneradorExcel (ubicador.py)
    4. Genera UN solo Excel con todo en la misma carpeta de los PDFs

El Excel tiene 2 hojas:
    • Datos Extraídos  → ítem, cantidad, descripción, valor unitario,
                         valor total, archivo PDF, página, línea
    • Resumen por PDF  → conteo de filas y páginas por cada archivo
"""

import os
import sys
from extractor import ExtractorTablasPDF
from ubicador  import UbicadorYGeneradorExcel


# ─────────────────────────────────────────────────────────────────────────────
# Interacción con el usuario
# ─────────────────────────────────────────────────────────────────────────────

def pedir_carpeta_pdfs() -> str:
    """Solicita y valida la ruta de la carpeta con los PDFs."""
    # Si se pasó por argumento de línea de comandos, usarla
    if len(sys.argv) > 1:
        ruta = sys.argv[1].strip().strip('"').strip("'")
        if os.path.exists(ruta) and os.path.isdir(ruta):
            return ruta
        else:
            print(f"[!] La carpeta no existe: {ruta}\n")
            sys.exit(1)
    
    while True:
        ruta = input("Ruta de la carpeta con los PDFs:\n> ").strip().strip('"').strip("'")

        if not ruta:
            print("[!] La ruta no puede estar vacía.\n")
            continue
        if not os.path.exists(ruta):
            print(f"[!] La carpeta no existe: {ruta}\n")
            continue
        if not os.path.isdir(ruta):
            print(f"[!] Eso no es una carpeta: {ruta}\n")
            continue
        return ruta


def buscar_pdfs(ruta_carpeta: str) -> list:
    """Retorna lista de rutas completas de todos los PDFs en la carpeta."""
    return sorted([
        os.path.join(ruta_carpeta, nombre)
        for nombre in os.listdir(ruta_carpeta)
        if nombre.lower().endswith(".pdf")
    ])


# ─────────────────────────────────────────────────────────────────────────────
# Función principal
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    """
    Flujo principal:
    1. Obtener carpeta del usuario.
    2. Buscar PDFs en la carpeta.
    3. Por cada PDF: extraer filas (extractor.py) → pasar al ubicador (ubicador.py).
    4. Generar el Excel final con todo.
    """

    # ── 1. Configuración del usuario ──────────────────────────────────────────
    ruta_carpeta = pedir_carpeta_pdfs()

    # ── 2. Buscar PDFs ────────────────────────────────────────────────────────
    lista_pdfs = buscar_pdfs(ruta_carpeta)

    if not lista_pdfs:
        print(f"\n[!] No se encontraron archivos PDF en: {ruta_carpeta}")
        sys.exit(1)

    # ── 3. Procesar cada PDF ──────────────────────────────────────────────────
    ubicador = UbicadorYGeneradorExcel(ruta_carpeta_salida=ruta_carpeta)

    for ruta_pdf in lista_pdfs:
        nombre_pdf = os.path.basename(ruta_pdf)
        try:
            extractor  = ExtractorTablasPDF(ruta_pdf=ruta_pdf)
            filas_pdf  = extractor.extraer()

            if filas_pdf:
                ubicador.agregar_filas(filas_pdf, nombre_pdf)

        except Exception as error:
            import traceback
            print(f"[ERROR] No se pudo procesar {nombre_pdf}: {error}")
            traceback.print_exc()
            sys.exit(1)

    # ── 4. Verificar que haya datos ───────────────────────────────────────────
    if ubicador.total_filas() == 0:
        print("[!] No se extrajo ningún dato de ningún PDF.")
        sys.exit(1)

    # ── 5. Generar el Excel final ─────────────────────────────────────────────
    try:
        ruta_excel = ubicador.generar_excel("resultado_extraccion.xlsx")
        print(f"✓ Archivo guardado en: {ruta_excel}")

    except Exception as error:
        print(f"[ERROR] No se pudo generar el Excel: {error}")
        sys.exit(1)



# ─────────────────────────────────────────────────────────────────────────────
# Punto de entrada
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    main()
