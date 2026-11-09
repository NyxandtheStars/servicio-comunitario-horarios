#!/usr/bin/env bash
# ==============================================================================
# UNESR - Sistema de Gestion de Horarios Universitarios (SGHU)
# Lanzador para Canaima GNU/Linux, Debian, Ubuntu, Arch y derivados
# ==============================================================================

# Ir al directorio donde se encuentra este script
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$DIR"

echo "=========================================================="
echo "  Iniciando Sistema de Horarios Universitarios (UNESR)    "
echo "  Plataforma: GNU/Linux (Canaima / Debian / Arch)        "
echo "=========================================================="

# Verificar que Python 3 este disponible
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python 3 no esta instalado en este sistema."
    echo "En Canaima / Debian puede instalarlo abriendo una terminal y ejecutando:"
    echo "  sudo apt update && sudo apt install -y python3 python3-venv python3-pip"
    read -p "Presione [Enter] para salir..."
    exit 1
fi

# Si no existe entorno virtual .venv, crearlo
if [ ! -d ".venv" ]; then
    echo "[INFO] Creando entorno virtual local (.venv)..."
    python3 -m venv .venv
fi

# Activar entorno virtual
if [ -f ".venv/bin/activate" ]; then
    source .venv/bin/activate
fi

# Verificar si Flet esta instalado en el entorno
if ! python -c "import flet" &> /dev/null; then
    echo "[INFO] Instalando dependencias necesarias (primera ejecucion)..."
    pip install -r requirements.txt
fi

# Verificar si falta libmpv.so.1 en el sistema (comun en Arch Linux / CachyOS)
if [ "$1" != "--web" ]; then
    if ! ldconfig -p 2>/dev/null | grep -q "libmpv\.so\.1" && [ ! -f /usr/lib/libmpv.so.1 ] && [ ! -f /usr/lib64/libmpv.so.1 ]; then
        echo ""
        echo "----------------------------------------------------------"
        echo "[AVISO] No se detecto 'libmpv.so.1' en el sistema."
        echo "Flet Desktop requiere esta libreria para la ventana nativa."
        echo ""
        echo "Para habilitar la ventana nativa en Arch Linux ejecute:"
        echo "  sudo pacman -S mpv"
        echo "  sudo ln -s /usr/lib/libmpv.so.2 /usr/lib/libmpv.so.1"
        echo ""
        echo "Iniciando temporalmente en modo navegador web..."
        echo "----------------------------------------------------------"
        export FLET_VIEW=web
    fi
fi

echo "[INFO] Lanzando interfaz grafica..."
python SGHU_Universidad.py "$@"

# Si la aplicacion cerro con error, mantener la ventana abierta para ver el detalle
if [ $? -ne 0 ]; then
    echo ""
    echo "[AVISO] La aplicacion se cerro con un codigo de salida diferente de 0."
    read -p "Presione [Enter] para cerrar esta ventana..."
fi
