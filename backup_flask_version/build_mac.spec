# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for macOS build of EDTECH DOC TEMPLATER."""
import os
import shutil
import subprocess

block_cipher = None

# --- Paths ---
import platform
try:
    BREW_PREFIX = subprocess.check_output(['brew', '--prefix'], text=True).strip()
except Exception:
    if platform.machine() == 'arm64':
        BREW_PREFIX = '/opt/homebrew'
    else:
        BREW_PREFIX = '/usr/local'

PROJECT = os.path.abspath('.')
TESS_BIN = os.path.join(BREW_PREFIX, 'bin', 'tesseract')
TESS_DATA = os.path.join(BREW_PREFIX, 'share', 'tessdata')
TESS_LIB = os.path.join(BREW_PREFIX, 'lib')

# --- Tesseract dylibs (ALL transitive Homebrew deps) ---
# Automatically discover all Homebrew dylib dependencies
import re

def get_homebrew_deps(binary_path, visited=None):
    """Recursively find all Homebrew dylib dependencies."""
    if visited is None:
        visited = set()
    if binary_path in visited:
        return []
    visited.add(binary_path)
    deps = []
    try:
        result = subprocess.run(['otool', '-L', binary_path], capture_output=True, text=True)
        for line in result.stdout.split('\n'):
            line = line.strip()
            match = re.search(r'((?:/opt/homebrew|/usr/local)/\S+\.dylib)', line)
            if match:
                dep_path = match.group(1)
                real_path = os.path.realpath(dep_path)
                if real_path not in visited and os.path.exists(real_path):
                    deps.append(real_path)
                    deps.extend(get_homebrew_deps(real_path, visited))
    except Exception:
        pass
    return deps

tess_dylibs = []
all_deps = get_homebrew_deps(TESS_BIN)
seen = set()
for dep in all_deps:
    base = os.path.basename(dep)
    if base not in seen:
        seen.add(base)
        tess_dylibs.append((dep, 'tesseract'))
        print(f'  [TESS] Bundling: {base}')

# --- Tesseract binary + lang data ---
tess_binaries = []
if os.path.exists(TESS_BIN):
    tess_binaries.append((TESS_BIN, 'tesseract'))
# Only bundle Spanish + English tessdata (keep bundle small)
tess_datas = []
for lang_file in ['eng.traineddata', 'spa.traineddata', 'osd.traineddata']:
    src = os.path.join(TESS_DATA, lang_file)
    if os.path.exists(src):
        tess_datas.append((src, 'tesseract/tessdata'))

a = Analysis(
    ['test_app.py'],
    pathex=[PROJECT],
    binaries=tess_binaries + tess_dylibs,
    datas=[
        ('static', 'static'),
        ('templates', 'templates'),
        # python-docx needs its template XML files at runtime (default-header.xml, etc.)
        (os.path.join(os.path.dirname(__import__('docx').__file__), 'templates'), 'docx/templates'),
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
    icon='app_icon.icns',
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

app = BUNDLE(
    coll,
    name='EDTECH DOC TEMPLATER.app',
    icon='app_icon.icns',
    bundle_identifier='com.edtech.doctemplater',
    info_plist={
        'CFBundleDisplayName': 'EDTECH DOC TEMPLATER',
        'CFBundleShortVersionString': '3.5.0',
        'CFBundleVersion': '3.5.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.15',
    },
)
