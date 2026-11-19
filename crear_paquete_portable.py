"""Crea una copia comprimida del sistema sin incluir datos locales.

Uso:
    python crear_paquete_portable.py [ruta_de_salida]

Si no se indica ruta de salida, se crea en el directorio superior al proyecto:
    ../servicio-comunitario-horarios-portable.zip
"""
import os
import sys
import zipfile
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUTPUT = os.path.normpath(os.path.join(BASE_DIR, os.pardir, "servicio-comunitario-horarios-portable.zip"))
PROJECT_NAME = os.path.basename(BASE_DIR)

EXCLUDE_DIRS = {
    ".git",
    ".venv",
    "venv",
    "env",
    "__pycache__",
    "reportes_imprimibles",
    "reportes_generados",
    "uploads",
    ".github",
}

EXCLUDE_PREFIXES = (
    "universidad_horarios.db.backup-",
    "universidad_horarios.backup-",
)
EXCLUDE_SUFFIXES = (
    ".pyc",
    ".pyo",
    ".zip",
    ".db",
    ".db-journal",
    ".db-shm",
    ".db-wal",
)
EXCLUDE_FILES = {
    ".iniciar_en_blanco",
    ".DS_Store",
    "Thumbs.db",
}


def excluir(ruta_relativa: str) -> bool:
    partes = ruta_relativa.split(os.sep)
    nombre = os.path.basename(ruta_relativa)
    if nombre in EXCLUDE_FILES:
        return True
    if any(parte in EXCLUDE_DIRS for parte in partes[:-1]):
        return True
    if any(nombre.startswith(prefijo) for prefijo in EXCLUDE_PREFIXES):
        return True
    if any(nombre.endswith(sufijo) for sufijo in EXCLUDE_SUFFIXES):
        return True
    if nombre.startswith("~$"):
        return True
    if os.path.basename(BASE_DIR) == nombre:
        return True
    return False


def crear_paquete(salida: str) -> None:
    salida = os.path.abspath(salida)
    with zipfile.ZipFile(salida, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as paquete:
        archivos = 0
        for raiz, directorios, ficheros in os.walk(BASE_DIR):
            rel_dir = os.path.relpath(raiz, BASE_DIR)
            if rel_dir == ".":
                rel_dir = ""
            directorios[:] = [d for d in directorios if os.path.normpath(os.path.join(rel_dir, d)) not in EXCLUDE_DIRS]

            for nombre in ficheros:
                ruta_absoluta = os.path.join(raiz, nombre)
                rel = os.path.normpath(os.path.join(rel_dir, nombre)) if rel_dir else nombre
                if excluir(rel):
                    continue
                arcname = os.path.join(PROJECT_NAME, rel).replace(os.sep, "/")
                paquete.write(ruta_absoluta, arcname)
                archivos += 1

        contenido = (
            f"Sistema empaquetado el {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n"
            "No incluye bases de datos ni datos cargados; cada instalacion crea su propia base local.\n\n"
            "Para usarlo en otra computadora:\n"
            "1. Descomprimir este archivo.\n"
            "2. Abrir la carpeta resultante.\n"
            "3. Este paquete contiene codigo fuente; para Windows, descarga el ejecutable compilado desde GitHub Actions.\n"
            "4. Para ejecutar el codigo fuente se requiere Python y las dependencias de requirements.txt.\n"
            "   En Linux/Canaima, instala las dependencias y ejecuta iniciar_sistema.sh.\n"
        )
        paquete.writestr(f"{PROJECT_NAME}/PAQUETE_PORTABLE.txt", contenido)

    tamano_mb = os.path.getsize(salida) / (1024 * 1024)
    print(f"Paquete creado: {salida}")
    print(f"Tamaño: {tamano_mb:.2f} MB")


if __name__ == "__main__":
    salida = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUTPUT
    crear_paquete(salida)
