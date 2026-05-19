#!/usr/bin/env python3
"""
ptx_cuesheet.py — orchestrator for the right-click PTX → Cue Sheet flow.

Usage:
    python3 ptx_cuesheet.py /path/to/session.ptx

Pipeline:
    1. Parse .ptx via ptx_reader → session dict + HTML report
    2. Inject our export-button handler into the HTML
    3. Start a localhost HTTP server, serve the HTML
    4. Open the report in the default browser
    5. User picks tracks, clicks "Export Session Info as Text…"
    6. Browser POSTs selected track names → /generate
    7. Server bridges selected tracks → .txt, runs cuesheet pipeline
    8. CSV + PDF appear next to the .ptx, server shuts down

Runs both as a regular Python script and as a PyInstaller bundle.
Requires Python 3.10+ when run as a script.
"""
from __future__ import annotations
import http.server
import json
import socket
import sys
import threading
import traceback
import webbrowser
from pathlib import Path
from urllib.parse import urlparse


def _resolve_resource_dir() -> Path:
    """Locate the folder holding our resource files.

    When frozen by PyInstaller, data files live in sys._MEIPASS.
    When running as a normal script, they live next to this file.
    """
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


HERE = _resolve_resource_dir()
sys.path.insert(0, str(HERE))

try:
    from ptx_reader import read_session, render_html
except ImportError as e:
    print(f"error: cannot import ptx_reader from {HERE}: {e}", file=sys.stderr)
    sys.exit(2)

try:
    from ptx_to_txt import write_pt_text_from_session
except ImportError as e:
    print(f"error: cannot import ptx_to_txt from {HERE}: {e}", file=sys.stderr)
    sys.exit(2)

try:
    import cuesheet
except ImportError as e:
    print(f"error: cannot import cuesheet from {HERE}: {e}", file=sys.stderr)
    sys.exit(2)


TEMPLATE_PATH = HERE / "report_template.html"


# Injected JS overrides the sel-export button to POST selected tracks.
INJECTED_SCRIPT = """
<script>
(function() {
  function init() {
    const btn = document.getElementById('sel-export');
    if (!btn) return;

    const clone = btn.cloneNode(true);
    btn.parentNode.replaceChild(clone, btn);

    clone.addEventListener('click', async (e) => {
      e.preventDefault();
      const cboxes = document.querySelectorAll('input[type="checkbox"][aria-label^="Select "]');
      const selected = [];
      for (const cb of cboxes) {
        if (!cb.checked) continue;
        if (cb.id === 'select-all') continue;
        const name = cb.getAttribute('aria-label').replace(/^Select /, '');
        selected.push(name);
      }
      if (selected.length === 0) {
        alert('No tracks selected. Tick the tracks you want in the cue sheet.');
        return;
      }
      clone.disabled = true;
      const originalText = clone.textContent;
      clone.textContent = 'Generating…';

      try {
        const res = await fetch('/generate', {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify({tracks: selected}),
        });
        const data = await res.json();
        if (data.ok) {
          clone.textContent = '\u2713 Done \u2014 cue sheet saved next to the session';
          setTimeout(() => { try { window.close(); } catch (_) {} }, 2000);
        } else {
          clone.textContent = originalText;
          clone.disabled = false;
          alert('Cue sheet generation failed:\\n' + (data.error || 'unknown error'));
        }
      } catch (err) {
        clone.textContent = originalText;
        clone.disabled = false;
        alert('Error contacting generator:\\n' + err.message);
      }
    });
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
</script>
"""


class CueSheetHandler(http.server.BaseHTTPRequestHandler):
    html_content: str = ""
    session: dict | None = None
    ptx_path: Path | None = None
    shutdown_event: threading.Event | None = None

    def log_message(self, fmt, *args):
        return

    def do_GET(self):
        url = urlparse(self.path)
        if url.path in ('/', '/index.html'):
            body = self.html_content.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_error(404)

    def do_POST(self):
        url = urlparse(self.path)
        if url.path != '/generate':
            self.send_error(404)
            return
        try:
            length = int(self.headers.get('Content-Length', 0))
            payload = json.loads(self.rfile.read(length))
            selected = payload.get('tracks', [])
            result = self._generate(selected)

            body = json.dumps(result).encode('utf-8')
            self.send_response(200 if result['ok'] else 500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

            if result['ok']:
                threading.Timer(0.5, self.shutdown_event.set).start()
        except Exception as e:
            traceback.print_exc()
            body = json.dumps({'ok': False, 'error': str(e)}).encode('utf-8')
            try:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception:
                pass

    def _generate(self, selected_names: list[str]) -> dict:
        ptx = self.ptx_path
        base = ptx.stem
        work_dir = ptx.parent
        intermediate_txt = work_dir / f"{base}.cuesheet-tmp.txt"

        try:
            n_clips = write_pt_text_from_session(
                self.session, selected_names, intermediate_txt,
            )
            if n_clips == 0:
                return {'ok': False, 'error': 'No clips found on the selected tracks.'}

            # In-process call: works whether script or bundled.
            csv_src, pdf_src = cuesheet.main(str(intermediate_txt))
            csv_src = Path(csv_src)
            pdf_src = Path(pdf_src)

            if not csv_src.exists() or not pdf_src.exists():
                return {'ok': False, 'error': 'cuesheet pipeline ran but expected outputs were not found.'}

            dst_csv = work_dir / f"{base} Cuesheet.csv"
            dst_pdf = work_dir / f"{base} Cuesheet.pdf"

            if dst_csv.exists():
                dst_csv.unlink()
            if dst_pdf.exists():
                dst_pdf.unlink()
            csv_src.rename(dst_csv)
            pdf_src.rename(dst_pdf)

            return {
                'ok': True,
                'csv': str(dst_csv),
                'pdf': str(dst_pdf),
                'tracks_used': len(selected_names),
                'clips': n_clips,
            }
        finally:
            if intermediate_txt.exists():
                try:
                    intermediate_txt.unlink()
                except OSError:
                    pass


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]


def show_error_dialog(message: str) -> None:
    """Surface a macOS alert when running headless (frozen) with no console."""
    if sys.platform != 'darwin':
        return
    try:
        import subprocess
        safe = message.replace('"', "'")
        subprocess.run(
            ['osascript', '-e', f'display alert "Music Cue Sheet" message "{safe}" as critical'],
            check=False, timeout=10,
        )
    except Exception:
        pass


def main_inner() -> int:
    if len(sys.argv) < 2:
        msg = "No .ptx file provided. Right-click a .ptx in Finder and choose 'Music Cue Sheet'."
        print(msg, file=sys.stderr)
        show_error_dialog(msg)
        return 1

    ptx_path = Path(sys.argv[1]).expanduser().resolve()
    if not ptx_path.exists():
        msg = f"Could not find {ptx_path}."
        print(f"error: {msg}", file=sys.stderr)
        show_error_dialog(msg)
        return 2
    if not TEMPLATE_PATH.exists():
        msg = f"report_template.html not found at {TEMPLATE_PATH}."
        print(f"error: {msg}", file=sys.stderr)
        show_error_dialog(msg)
        return 2

    print(f"Parsing {ptx_path.name}…")
    session = read_session(ptx_path)
    if not session.get('track_clips'):
        msg = (
            "No track clips extracted from this .ptx.\\n\\n"
            "This usually means the ptunxor helper binary isn't available.\\n"
            "Make sure ptunxor is bundled with the app or on your PATH."
        )
        print(f"error: {msg}", file=sys.stderr)
        show_error_dialog(msg)
        return 3

    html = render_html(session, TEMPLATE_PATH)
    html_with_handler = (
        html.replace('</body>', INJECTED_SCRIPT + '</body>', 1)
        if '</body>' in html else html + INJECTED_SCRIPT
    )

    CueSheetHandler.html_content = html_with_handler
    CueSheetHandler.session = session
    CueSheetHandler.ptx_path = ptx_path
    CueSheetHandler.shutdown_event = threading.Event()

    port = find_free_port()
    server = http.server.HTTPServer(('127.0.0.1', port), CueSheetHandler)
    threading.Thread(target=server.serve_forever, daemon=True).start()

    url = f"http://127.0.0.1:{port}/"
    print(f"Opening {url}")
    webbrowser.open(url)

    try:
        CueSheetHandler.shutdown_event.wait()
    except KeyboardInterrupt:
        print("\nCancelled.")
    finally:
        server.shutdown()

    print("Done.")
    return 0


def main() -> int:
    try:
        return main_inner()
    except Exception as e:
        tb = traceback.format_exc()
        print(tb, file=sys.stderr)
        show_error_dialog(f"Unexpected error:\\n{e}")
        return 99


if __name__ == '__main__':
    raise SystemExit(main())
