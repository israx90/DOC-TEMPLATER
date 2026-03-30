#!/usr/bin/env python3
"""Launch DOC TEMPLATER in a native window (no browser needed).
Usage:
    python test_app.py          # Normal mode (native window)
    python test_app.py --dev    # Dev mode: kills existing port, runs Flask with hot reload
"""
import sys
import os
import signal
import subprocess
import threading
import time

PORT = 3000

def kill_port(port):
    """Kill any process using the given port."""
    try:
        result = subprocess.run(
            ['lsof', '-ti', ':{}'.format(port)],
            capture_output=True, text=True
        )
        pids = result.stdout.strip().split('\n')
        my_pid = str(os.getpid())
        for pid in pids:
            if pid and pid != my_pid:
                os.kill(int(pid), signal.SIGKILL)
                print('Killed process {} on port {}'.format(pid, port))
        time.sleep(0.5)
    except Exception:
        pass

def start_flask():
    from app import app
    app.run(port=PORT, use_reloader=False)


class Api:
    """Exposed to JS as window.pywebview.api"""

    def select_folder(self):
        """Open native folder dialog and return the selected path."""
        import webview
        try:
            active = webview.windows[0] if webview.windows else None
            if not active:
                return None
            result = active.create_file_dialog(webview.FOLDER_DIALOG)
            if not result:
                return None
            # pywebview may return a tuple, list, or string depending on platform
            if isinstance(result, (list, tuple)):
                return result[0] if len(result) > 0 else None
            return str(result)
        except Exception as e:
            print('select_folder error: {}'.format(e))
            return None

    def select_image(self, slot):
        """Open native image file dialog. Upload to Flask. Return JSON response.
        slot: 'cover', 'header', 'footer', or 'backpage'
        """
        import webview
        import requests as _req

        active = webview.windows[0] if webview.windows else None
        if not active:
            return None

        file_types = ('Image Files (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.emf;*.wmf)',)
        result = active.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=file_types
        )
        if not result or len(result) == 0:
            return None

        filepath = result[0]
        filename = os.path.basename(filepath)

        # Map slot to Flask endpoint and form field name
        slot_map = {
            'cover':    ('upload_cover',    'coverUtils'),
            'header':   ('upload_header',   'headerUtils'),
            'footer':   ('upload_footer',   'footerUtils'),
            'backpage': ('upload_backpage', 'backpageUtils'),
        }
        if slot not in slot_map:
            return None

        endpoint, field = slot_map[slot]
        url = 'http://127.0.0.1:{}/{}'.format(PORT, endpoint)

        with open(filepath, 'rb') as f:
            resp = _req.post(url, files={field: (filename, f)})

        try:
            data = resp.json()
            data['local_path'] = filepath
            return data
        except Exception:
            return None

    def save_template(self, config_json):
        """Save .edd template: send config to Flask, get ZIP, save with native dialog."""
        import webview
        import requests as _req
        import json as _json

        active = webview.windows[0] if webview.windows else None
        if not active:
            return None

        # Ask user where to save
        result = active.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename='plantilla.edd',
            file_types=('EDTech Template (*.edd)',)
        )
        if not result:
            return None

        save_path = result if isinstance(result, str) else result[0] if result else None
        if not save_path:
            return None

        # Ensure .edd extension
        if not save_path.lower().endswith('.edd'):
            save_path += '.edd'

        # Fetch the .edd ZIP from Flask
        url = 'http://127.0.0.1:{}/save_template'.format(PORT)
        resp = _req.post(url, json=_json.loads(config_json) if isinstance(config_json, str) else config_json)

        if resp.status_code != 200:
            return {'error': 'Server error: {}'.format(resp.status_code)}

        with open(save_path, 'wb') as f:
            f.write(resp.content)

        return {'success': True, 'path': save_path}

    def select_docs(self):
        """Open native file dialog for DOCX files. Returns list of paths."""
        import webview
        active = webview.windows[0] if webview.windows else None
        if not active:
            return None
        file_types = ('Word Documents (*.docx)',)
        result = active.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=file_types
        )
        return list(result) if result else None


if __name__ == '__main__':
    dev_mode = '--dev' in sys.argv

    # Only kill port if not a Flask reloader child process
    if not os.environ.get('WERKZEUG_RUN_MAIN'):
        kill_port(PORT)

    if dev_mode:
        print('=== DEV MODE: Flask at http://127.0.0.1:{} ==='.format(PORT))
        from app import app
        app.run(port=PORT, debug=True, use_reloader=True)
    else:
        import webview
        api = Api()
        t = threading.Thread(target=start_flask, daemon=True)
        t.start()
        time.sleep(1.5)
        webview.create_window(
            'EDTECH DOC TEMPLATER v3.5',
            'http://127.0.0.1:{}'.format(PORT),
            width=1400, height=900,
            js_api=api
        )
        webview.start()
