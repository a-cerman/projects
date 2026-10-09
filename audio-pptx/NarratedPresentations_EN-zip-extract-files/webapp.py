"""Loopback-only UI, no external services or dependencies."""
import json
import secrets
import threading
import subprocess
import webbrowser
from pathlib import Path
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from app import parse_docx, narrate

BASE = Path(__file__).resolve().parent

def choose(kind):
    filters = {'ppt': 'PowerPoint|*.pptx;*.ppt', 'doc': 'Word|*.docx', 'model': 'Piper|*.onnx'}
    if kind == 'target':
        code = '$d=New-Object System.Windows.Forms.FolderBrowserDialog'
        result = '$d.SelectedPath'
    elif kind in filters:
        code = '$d=New-Object System.Windows.Forms.OpenFileDialog; $d.Filter="' + filters[kind] + '"'
        result = '$d.FileName'
    else:
        raise ValueError('Invalid selection.')
    script = '[Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Add-Type -AssemblyName System.Windows.Forms; ' + code
    # Give the native dialog a visible, topmost owner so it does not remain
    # behind the browser when launched by a hidden Python process.
    script += '; $owner=New-Object System.Windows.Forms.Form; $owner.Text="Narrated Presentations - Select a file"'
    script += '; $owner.TopMost=$true; $owner.Width=420; $owner.Height=120; $owner.StartPosition="CenterScreen"'
    script += '; try { $owner.Show(); $owner.Activate(); if($d.ShowDialog($owner) -eq "OK"){[Console]::Write(' + result + ')} } finally { $d.Dispose(); $owner.Dispose() }'
    try:
        proc = subprocess.run(['powershell.exe', '-NoProfile', '-STA', '-Command', script],
                              capture_output=True, encoding='utf-8', creationflags=subprocess.CREATE_NO_WINDOW,
                              timeout=120)
    except subprocess.TimeoutExpired:
        raise RuntimeError('File selection timed out. Try Browse again or paste the full path into the field.')
    if proc.returncode:
        raise RuntimeError('The file picker could not open. You can paste the file path into the field.')
    return proc.stdout.strip().lstrip('\ufeff')

def make_server():
    token = secrets.token_urlsafe(32)
    state = {'busy': False, 'logs': [], 'output': None}
    lock = threading.Lock()
    def log(value):
        with lock:
            state['logs'].append(value)
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass
        def send(self, data, status=200, mime='application/json; charset=utf-8'):
            raw = data if isinstance(data, bytes) else json.dumps(data, ensure_ascii=False).encode('utf-8')
            self.send_response(status)
            self.send_header('Content-Type', mime)
            self.send_header('Content-Length', str(len(raw)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; connect-src 'self'; frame-ancestors 'none'")
            self.end_headers()
            self.wfile.write(raw)
        def allowed(self):
            expected = f'127.0.0.1:{self.server.server_port}'
            if self.headers.get('Host') != expected:
                self.send({'error': 'Invalid address.'}, 403)
                return False
            return True
        def do_GET(self):
            if not self.allowed(): return
            if self.path == '/':
                raw = (BASE / 'index.html').read_text(encoding='utf-8').replace('__TOKEN__', token)
                self.send(raw.encode('utf-8'), mime='text/html; charset=utf-8')
            else:
                self.send({'error': 'Not found.'}, 404)
        def do_POST(self):
            if not self.allowed(): return
            if self.headers.get('X-Session-Token') != token:
                self.send({'error': 'Invalid session.'}, 403); return
            if self.headers.get('Origin') not in (None, f'http://127.0.0.1:{self.server.server_port}'):
                self.send({'error': 'Invalid origin.'}, 403); return
            try:
                length = int(self.headers.get('Content-Length', '0'))
                if not 0 <= length <= 65536: raise ValueError('Request is too large.')
                data = json.loads(self.rfile.read(length) or b'{}')
                if self.path == '/api/status':
                    with lock: snapshot = dict(state, logs=list(state['logs']))
                    self.send(snapshot)
                elif self.path == '/api/choose':
                    self.send({'path': choose(data['kind'])})
                elif self.path == '/api/preview':
                    self.send({'slides': parse_docx(data['doc'])})
                elif self.path == '/api/start':
                    values = [str(data[k]).strip() for k in ('ppt', 'doc', 'model', 'target')]
                    for value, suffixes in zip(values[:3], (('.ppt', '.pptx'), ('.docx',), ('.onnx',))):
                        p = Path(value)
                        if not p.is_file() or p.suffix.lower() not in suffixes:
                            raise ValueError('Check the file path and extension.')
                    if not Path(values[3]).is_dir(): raise ValueError('Output folder not found.')
                    parse_docx(values[1])
                    with lock:
                        if state['busy']: raise ValueError('Processing is already in progress.')
                        state.update(busy=True, logs=[], output=None)
                    def worker():
                        try:
                            narrate(*values, log)
                        except Exception as e:
                            log(f'ERROR: {e}')
                        finally:
                            with lock: state['busy'] = False
                    threading.Thread(target=worker, daemon=False).start()
                    self.send({'ok': True})
                elif self.path == '/api/stop':
                    with lock:
                        if state['busy']: raise ValueError('Wait for narration to finish before closing the application.')
                    self.send({'ok': True})
                    threading.Thread(target=self.server.shutdown, daemon=True).start()
                else:
                    self.send({'error': 'Not found.'}, 404)
            except Exception as e:
                self.send({'error': str(e)}, 400)
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    return server, token

def main():
    server, _ = make_server()
    url = f'http://127.0.0.1:{server.server_port}/'
    try:
        webbrowser.open(url)
        server.serve_forever()
    finally:
        server.server_close()

if __name__ == '__main__':
    main()
