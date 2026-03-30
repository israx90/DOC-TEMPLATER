#!/usr/bin/env python3
"""
Build script for EDTECH DOC TEMPLATER.
Detects OS and runs the appropriate PyInstaller spec.

Usage:
    python build.py          # Build for current platform
    python build.py --dmg    # macOS only: also create .dmg
"""
import sys
import os
import platform
import subprocess
import shutil

APP_NAME = 'EDTECH DOC TEMPLATER'
VERSION = '3.5.0'


def check_pyinstaller():
    """Ensure PyInstaller is installed."""
    try:
        import PyInstaller
        print('[OK] PyInstaller {} found'.format(PyInstaller.__version__))
    except ImportError:
        print('[!] PyInstaller not found. Installing...')
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
        print('[OK] PyInstaller installed')


def build_mac(target_arch=None):
    """Build macOS .app bundle."""
    spec = os.path.join(os.path.dirname(__file__), 'build_mac.spec')
    if not os.path.exists(spec):
        print('[ERROR] build_mac.spec not found')
        return False

    # Check Tesseract
    tess = '/opt/homebrew/bin/tesseract'
    if not os.path.exists(tess):
        tess = shutil.which('tesseract')
    if tess:
        print('[OK] Tesseract found: {}'.format(tess))
    else:
        print('[WARN] Tesseract not found - OCR will not be available')

    print('\n=== Building macOS .app ===')
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--clean', '--noconfirm'
    ]
    cmd.append(spec)

    result = subprocess.run(cmd)
    if result.returncode != 0:
        print('[ERROR] Build failed')
        return False

    app_path = os.path.join('dist', '{}.app'.format(APP_NAME))
    if os.path.exists(app_path):
        print('\n[SUCCESS] Built: {}'.format(app_path))
        size_mb = sum(
            os.path.getsize(os.path.join(dp, f))
            for dp, dn, filenames in os.walk(app_path)
            for f in filenames
        ) / (1024 * 1024)
        print('[INFO] App size: {:.1f} MB'.format(size_mb))
        return True
    return False


def create_dmg():
    """Create a .dmg from the .app bundle (macOS only)."""
    app_path = os.path.join('dist', '{}.app'.format(APP_NAME))
    
    # Determine arch suffix from args
    arch_suffix = ""
    if '--arch' in sys.argv:
        idx = sys.argv.index('--arch')
        if idx + 1 < len(sys.argv):
            arch_suffix = " Intel" if sys.argv[idx + 1] == 'x86_64' else " " + sys.argv[idx + 1]

    dmg_path = os.path.join('dist', '{} v{}{}.dmg'.format(APP_NAME, VERSION, arch_suffix))

    if not os.path.exists(app_path):
        print('[ERROR] .app not found, build first')
        return False

    # Remove old DMG
    if os.path.exists(dmg_path):
        os.remove(dmg_path)

    print('\n=== Creating DMG ===')
    # Create a temp folder with .app + alias to /Applications
    tmp_dir = os.path.join('dist', 'dmg_tmp')
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)

    # Copy .app
    shutil.copytree(app_path, os.path.join(tmp_dir, '{}.app'.format(APP_NAME)))

    # Create Applications symlink
    os.symlink('/Applications', os.path.join(tmp_dir, 'Applications'))

    # Create DMG
    result = subprocess.run([
        'hdiutil', 'create',
        '-volname', APP_NAME,
        '-srcfolder', tmp_dir,
        '-ov', '-format', 'UDZO',
        dmg_path
    ])

    # Cleanup
    shutil.rmtree(tmp_dir)

    if result.returncode == 0 and os.path.exists(dmg_path):
        size_mb = os.path.getsize(dmg_path) / (1024 * 1024)
        print('[SUCCESS] DMG created: {} ({:.1f} MB)'.format(dmg_path, size_mb))
        return True
    else:
        print('[ERROR] DMG creation failed')
        return False


def build_win():
    """Build Windows .exe."""
    spec = os.path.join(os.path.dirname(__file__), 'build_win.spec')
    if not os.path.exists(spec):
        print('[ERROR] build_win.spec not found')
        return False

    # Check Tesseract
    tess_dir = r'C:\Program Files\Tesseract-OCR'
    if os.path.exists(os.path.join(tess_dir, 'tesseract.exe')):
        print('[OK] Tesseract found: {}'.format(tess_dir))
    else:
        print('[WARN] Tesseract not found at {}'.format(tess_dir))
        print('       Install: choco install tesseract')

    print('\n=== Building Windows EXE ===')
    result = subprocess.run([
        sys.executable, '-m', 'PyInstaller',
        '--clean', '--noconfirm', spec
    ])
    if result.returncode != 0:
        print('[ERROR] Build failed')
        return False

    exe_path = os.path.join('dist', APP_NAME, '{}.exe'.format(APP_NAME))
    if os.path.exists(exe_path):
        print('\n[SUCCESS] Built: {}'.format(exe_path))
        return True
    return False


if __name__ == '__main__':
    print('=' * 50)
    print('  {} v{} — Build System'.format(APP_NAME, VERSION))
    print('  Platform: {} ({})'.format(platform.system(), platform.machine()))
    print('=' * 50)

    check_pyinstaller()

    system = platform.system()
    if system == 'Darwin':
        target_arch = None
        if '--arch' in sys.argv:
            idx = sys.argv.index('--arch')
            if idx + 1 < len(sys.argv):
                target_arch = sys.argv[idx + 1]
        
        ok = build_mac(target_arch)
        if ok and '--dmg' in sys.argv:
            create_dmg()
    elif system == 'Windows':
        build_win()
    else:
        print('[ERROR] Unsupported platform: {}'.format(system))
        sys.exit(1)
