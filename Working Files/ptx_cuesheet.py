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
    7. Server bridges selected tracks → .txt, runs cuesheet.py
    8. CSV + PDF appear next to the .ptx, server shuts down

All paths derive from the .ptx — nothing hardcoded.
Requires Python 3.10+ (uses `X | None` type hints).
"""
from __future__ import annotations
import http.server
import json
import socket
import subprocess
import sys
import threading
import webbrowser
from pathlib import Path
from urllib.parse import urlparse

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

try:
    from ptx_reader import read_session, render_html
except ImportError as e:
    print(f"error: cannot import ptx_reader from {HERE}", file=sys.stderr)
    print(f"  {e}", file=sys.stderr)
    sys.exit(2)

try:
    from ptx_to_txt import write_pt_text_from_session
except ImportError as e:
    print(f"error: cannot import ptx_to_txt from {HERE}", file=sys.stderr)
    print(f"  {e}", file=sys.stderr)
    sys.exit(2)


TEMPLATE_PATH = HERE / "report_template.html"
CUESHEET_SCRIPT = HERE / "cuesheet.py"


# Injected JS overrides the sel-export button to POST selected tracks.
INJECTED_SCRIPT = """
<script>
(function() {
  function init() {
    const btn = document.getElementById('sel-export');
    if (!btn) return;

    // Replace existing handlers by cloning + swapping
    const clone = btn.cloneNode(true);
    btn.parentNode.replaceChild(clone, btn);

    clone.addEventListener('click', async (e) => {
      e.preventDefault();

      // Find every checked track checkbox via its aria-label
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
    # set by the orchestrator on the class itself
    html_content: str = ""
    session: dict | None = None
    ptx_path: Path | None = None
    shutdown_event: threading.Event | None = None

    def log_message(self, fmt, *args):
        return  # silence default request logging

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
                # let the response flush, then trigger shutdown
                threading.Timer(0.5, self.shutdown_event.set).start()
        except Exception as e:
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

            # cuesheet.py reads input from sys.argv[1] and writes outputs
            # next to the input file with "_cuesheet" suffix.
            proc = subprocess.run(
                [sys.executable, str(CUESHEET_SCRIPT), str(intermediate_txt)],
                cwd=str(work_dir),
                capture_output=True,
                text=True,
            )
            if proc.returncode != 0:
                return {
                    'ok': False,
                    'error': f'cuesheet.py failed (exit {proc.returncode}):\n{proc.stderr.strip()}',
                }

            # cuesheet.py output filenames are derived from the .txt stem
            tmp_stem = intermediate_txt.stem  # "<base>.cuesheet-tmp"
            src_csv = work_dir / f"{tmp_stem}_cuesheet.csv"
            src_pdf = work_dir / f"{tmp_stem}_cuesheet.pdf"
            dst_csv = work_dir / f"{base} Cuesheet.csv"
            dst_pdf = work_dir / f"{base} Cuesheet.pdf"

            if not src_csv.exists() or not src_pdf.exists():
                return {
                    'ok': False,
                    'error': 'cuesheet.py ran but expected outputs were not found.',
                }

            # replace existing outputs cleanly
            if dst_csv.exists():
                dst_csv.unlink()
            if dst_pdf.exists():
                dst_pdf.unlink()
            src_csv.rename(dst_csv)
            src_pdf.rename(dst_pdf)

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


def main() -> int:
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <session.ptx>", file=sys.stderr)
        return 1

    ptx_path = Path(sys.argv[1]).expanduser().resolve()
    if not ptx_path.exists():
        print(f"error: {ptx_path} not found", file=sys.stderr)
        return 2
    if not TEMPLATE_PATH.exists():
        print(f"error: report_template.html not found at {TEMPLATE_PATH}", file=sys.stderr)
        return 2
    if not CUESHEET_SCRIPT.exists():
        print(f"error: cuesheet.py not found at {CUESHEET_SCRIPT}", file=sys.stderr)
        return 2

    print(f"Parsing {ptx_path.name}…")
    session = read_session(ptx_path)
    if not session.get('track_clips'):
        print(
            "error: no track_clips extracted. Make sure ptunxor is built and on PATH.",
            file=sys.stderr,
        )
        return 3

    html = render_html(session, TEMPLATE_PATH)
    if '</body>' in html:
        html_with_handler = html.replace('</body>', INJECTED_SCRIPT + '</body>', 1)
    else:
        html_with_handler = html + INJECTED_SCRIPT

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


if __name__ == '__main__':
    raise SystemExit(main())
