# zip_utils.py
import zipfile
import shutil
from pathlib import Path

def extraer_pdfs_de_zip(ruta_zip: Path, carpeta_destino: Path) -> list[Path]:
    """
    Extrae SOLO PDFs de un .zip, incluyendo PDFs dentro de carpetas internas.
    Devuelve una lista de rutas a PDFs extraídos.
    """
    carpeta_destino.mkdir(parents=True, exist_ok=True)
    pdfs_extraidos = []

    try:
        with zipfile.ZipFile(ruta_zip, 'r') as z:
            for item in z.namelist():
                # ignorar carpetas
                if item.endswith('/'):
                    continue

                # Solo PDFs
                if item.lower().endswith('.pdf'):
                    ruta_salida = carpeta_destino / Path(item).name
                    with z.open(item) as src, open(ruta_salida, 'wb') as dst:
                        shutil.copyfileobj(src, dst)
                    pdfs_extraidos.append(ruta_salida)

                # ZIP anidado
                if item.lower().endswith('.zip'):
                    nested_zip_path = carpeta_destino / Path(item).name
                    with open(nested_zip_path, 'wb') as dst:
                        dst.write(z.read(item))
                    pdfs_extraidos.extend(
                        extraer_pdfs_de_zip(nested_zip_path, carpeta_destino / "nested")
                    )

    except Exception as e:
        print(f"[ERROR] No se pudo extraer ZIP {ruta_zip}: {e}")

    return pdfs_extraidos


def limpiar_zipcache(carpeta_zipcache: Path):
    """Borra toda la carpeta temporal zipcache."""
    try:
        if carpeta_zipcache.exists():
            shutil.rmtree(carpeta_zipcache)
    except Exception as e:
        print(f"[ADVERTENCIA] No se pudo borrar zipcache: {e}")
