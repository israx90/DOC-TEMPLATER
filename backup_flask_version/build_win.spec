# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for Windows build of EDTECH DOC TEMPLATER."""
import os

block_cipher = None

# --- Paths ---
PROJECT = os.path.abspath('.')

# Tesseract on Windows (default Chocolatey install path)
TESS_DIR = r'C:\Program Files\Tesseract-OCR'
TESS_BIN = os.path.join(TESS_DIR, 'tesseract.exe')
TESS_DATA = os.path.join(TESS_DIR, 'tessdata')

# --- Tesseract binaries + data ---
tess_binaries = []
tess_datas = []
if os.path.exists(TESS_BIN):
    tess_binaries.append((TESS_BIN, 'tesseract'))
    # Bundle DLLs from Tesseract install
    for f in os.listdir(TESS_DIR):
        if f.lower().endswith('.dll'):
            tess_binaries.append((os.path.join(TESS_DIR, f), 'tesseract'))
    # Only bundle Spanish + English tessdata
    for lang_file in ['eng.traineddata', 'spa.traineddata', 'osd.traineddata']:
        src = os.path.join(TESS_DATA, lang_file)
        if os.path.exists(src):
            tess_datas.append((src, 'tesseract/tessdata'))

a = Analysis(
    ['test_app.py'],
    pathex=[PROJECT],
    binaries=tess_binaries,
    datas=[
        ('static', 'static'),
        ('templates', 'templates'),
    ] + tess_datas,
    hiddenimports=[
        'flask',
        'werkzeug',
        'jinja2',
        'webview',
        'docx',
        'PIL',
        'cv2',
        'pytesseract',
        'img2table',
        'img2table.ocr',
        'img2table.document',
        'requests',
        'engineio.async_drivers.threading',
        'clr_loader',
        'pythonnet',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'scipy', 'numpy.testing'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EDTECH DOC TEMPLATER',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='app_icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EDTECH DOC TEMPLATER',
)
