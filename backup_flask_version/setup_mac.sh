#!/bin/bash
# ============================================
#  EDTECH DOC TEMPLATER — Setup macOS
#  Ejecutar: chmod +x setup_mac.sh && ./setup_mac.sh
# ============================================
set -e

echo "=========================================="
echo "  EDTECH DOC TEMPLATER — Instalador macOS"
echo "=========================================="
echo ""

# 1. Homebrew
if ! command -v brew &> /dev/null; then
    echo "[1/5] Instalando Homebrew..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    # Add to path for Apple Silicon
    if [ -f /opt/homebrew/bin/brew ]; then
        eval "$(/opt/homebrew/bin/brew shellenv)"
    fi
else
    echo "[1/5] ✅ Homebrew ya instalado"
fi

# 2. Python
if ! command -v python3 &> /dev/null; then
    echo "[2/5] Instalando Python 3..."
    brew install python
else
    echo "[2/5] ✅ Python $(python3 --version) ya instalado"
fi

# 3. Tesseract OCR
if ! command -v tesseract &> /dev/null; then
    echo "[3/5] Instalando Tesseract OCR..."
    brew install tesseract
    brew install tesseract-lang
else
    echo "[3/5] ✅ Tesseract $(tesseract --version 2>&1 | head -1) ya instalado"
fi

# 4. Virtual environment + dependencies
echo "[4/5] Configurando entorno Python..."
cd "$(dirname "$0")"

if [ ! -d "venv" ]; then
    python3 -m venv venv
    echo "     Entorno virtual creado"
fi

source venv/bin/activate
pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet
pip install pyinstaller --quiet
echo "     ✅ Dependencias Python instaladas"

# 5. Create working directories
echo "[5/5] Creando carpetas de trabajo..."
mkdir -p uploads outputs
echo "     ✅ Listo"

echo ""
echo "=========================================="
echo "  ✅ INSTALACIÓN COMPLETADA"
echo "=========================================="
echo ""
echo "  Para ejecutar la aplicación:"
echo "    source venv/bin/activate"
echo "    python test_app.py"
echo ""
echo "  Para compilar el instalador .dmg:"
echo "    source venv/bin/activate"
echo "    python build.py --dmg"
echo ""
echo "=========================================="
