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
import io
import json
import os
import re
import socket
import sys
import threading
import traceback
import webbrowser
from contextlib import redirect_stdout, redirect_stderr
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


# Injected JS adds:
#   1) A timeline panel showing clips on selected tracks (updates live)
#   2) Time range inputs (HH:MM:SS:FF format) to slice the cue sheet
#   3) A merge-gap slider (default 5s, 0-30)
#   4) Overrides the sel-export button to POST tracks + gap + time range
#
# Session clip data is embedded as JSON in <script id="ms-session-data">.
INJECTED_SCRIPT = """
<script>
(function() {
  // ─── Family → colour map ─────────────────────────────────────
  const COLOR_MAP = {
    'MX': '#fbbf24', 'MUSIC': '#fbbf24', 'M': '#fbbf24', 'MUS': '#fbbf24',
    'NAR': '#60a5fa', 'NARR': '#60a5fa', 'VO': '#60a5fa', 'VOICE': '#60a5fa',
    'DX': '#a78bfa', 'DIALOG': '#a78bfa', 'DIALOGUE': '#a78bfa',
    'FX': '#f87171', 'SFX': '#f87171',
    'FOLEY': '#fb923c', 'BG': '#fb923c', 'AMB': '#fb923c',
    'OPT': '#34d399', 'OPTIONAL': '#34d399',
  };
  function colourFor(family) {
    const key = (family || '').toUpperCase().split(/[\\s_]/)[0];
    return COLOR_MAP[key] || '#9ca3af';
  }

  // ─── Timecode helpers (25fps) ────────────────────────────────
  const FPS = 25;
  function tcToFrames(tc) {
    const m = /^(\\d{1,2}):(\\d{1,2}):(\\d{1,2}):(\\d{1,2})$/.exec((tc||'').trim());
    if (!m) return null;
    const h=+m[1], mi=+m[2], s=+m[3], f=+m[4];
    if (mi>59 || s>59 || f>=FPS) return null;
    return h*90000 + mi*1500 + s*FPS + f;
  }
  function framesToTc(frames) {
    const h = Math.floor(frames/90000); frames %= 90000;
    const mi = Math.floor(frames/1500); frames %= 1500;
    const s = Math.floor(frames/FPS);
    const f = frames % FPS;
    const p = n => String(n).padStart(2,'0');
    return `${p(h)}:${p(mi)}:${p(s)}:${p(f)}`;
  }

  function init() {
    const btn = document.getElementById('sel-export');
    if (!btn) return;

    // ─── Parse the embedded session data ───────────────────────
    let session = null;
    try {
      const dataEl = document.getElementById('ms-session-data');
      if (dataEl) session = JSON.parse(dataEl.textContent);
    } catch (e) {
      console.warn('No session data for timeline:', e);
    }

    let maxSample = 0;
    let samplesPerFrame = 1920; // 48k/25
    if (session) {
      samplesPerFrame = session.sample_rate / FPS;
      for (const t of session.tracks) {
        for (const c of t.clips) {
          const end = c.s + c.l;
          if (end > maxSample) maxSample = end;
        }
      }
    }
    const samplesToFrames = sam => Math.floor(sam / samplesPerFrame);

    // ─── Container panel ───────────────────────────────────────
    const panel = document.createElement('div');
    panel.id = 'ms-panel';
    panel.style.cssText = (
      'margin: 16px 0; padding: 14px 16px; ' +
      'border: 1px solid rgba(127,127,127,0.25); ' +
      'border-radius: 10px; font-family: inherit; font-size: 13px;'
    );

    // ─── Time range row ───────────────────────────────────────
    const rangeRow = document.createElement('div');
    rangeRow.style.cssText = 'display: flex; align-items: center; gap: 10px; flex-wrap: wrap;';

    const rangeLabel = document.createElement('span');
    rangeLabel.textContent = 'Time range';
    rangeLabel.style.cssText = 'font-weight: 600; min-width: 80px;';

    function makeTcInput(placeholder) {
      const i = document.createElement('input');
      i.type = 'text';
      i.placeholder = placeholder;
      i.pattern = '\\\\d{1,2}:\\\\d{2}:\\\\d{2}:\\\\d{2}';
      i.style.cssText = (
        'width: 110px; padding: 6px 8px; ' +
        'font-family: ui-monospace, Menlo, Consolas, monospace; ' +
        'font-size: 13px; ' +
        'border: 1px solid rgba(127,127,127,0.4); ' +
        'border-radius: 6px; background: transparent; color: inherit;'
      );
      return i;
    }
    const fromIn = makeTcInput('00:00:00:00');
    const toIn = makeTcInput(maxSample ? framesToTc(samplesToFrames(maxSample)) : '00:00:00:00');
    fromIn.id = 'ms-range-from';
    toIn.id = 'ms-range-to';

    const clearBtn = document.createElement('button');
    clearBtn.type = 'button';
    clearBtn.textContent = 'Clear';
    clearBtn.style.cssText = (
      'padding: 6px 12px; font-size: 12px; ' +
      'border: 1px solid rgba(127,127,127,0.4); ' +
      'border-radius: 6px; background: transparent; color: inherit; cursor: pointer;'
    );
    clearBtn.addEventListener('click', () => {
      fromIn.value = '';
      toIn.value = '';
      renderTimeline();
    });

    const fullLabel = document.createElement('span');
    fullLabel.textContent = (
      maxSample
        ? `(session: 00:00:00:00 — ${framesToTc(samplesToFrames(maxSample))})`
        : ''
    );
    fullLabel.style.cssText = 'opacity: 0.55; font-size: 11px; margin-left: 4px;';

    rangeRow.appendChild(rangeLabel);
    rangeRow.appendChild(fromIn);
    const sep = document.createElement('span'); sep.textContent = '→'; sep.style.opacity='0.6';
    rangeRow.appendChild(sep);
    rangeRow.appendChild(toIn);
    rangeRow.appendChild(clearBtn);
    rangeRow.appendChild(fullLabel);
    panel.appendChild(rangeRow);

    // ─── Timeline SVG ──────────────────────────────────────────
    const svgWrap = document.createElement('div');
    svgWrap.style.cssText = (
      'margin-top: 12px; padding: 8px; ' +
      'background: rgba(127,127,127,0.06); ' +
      'border-radius: 6px; overflow-x: auto;'
    );
    const svgNs = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(svgNs, 'svg');
    svg.setAttribute('width', '100%');
    svg.style.cssText = 'display: block;';
    svgWrap.appendChild(svg);
    panel.appendChild(svgWrap);

    const emptyMsg = document.createElement('div');
    emptyMsg.textContent = 'Tick tracks above to preview them on the timeline.';
    emptyMsg.style.cssText = 'padding: 24px; text-align: center; opacity: 0.5;';
    svgWrap.appendChild(emptyMsg);

    // ─── Merge-gap row ─────────────────────────────────────────
    const gapRow = document.createElement('div');
    gapRow.style.cssText = (
      'display: flex; align-items: center; gap: 12px; ' +
      'margin-top: 12px; padding-top: 12px; ' +
      'border-top: 1px solid rgba(127,127,127,0.15);'
    );

    const gapLabel = document.createElement('span');
    gapLabel.textContent = 'Merge gap';
    gapLabel.style.cssText = 'font-weight: 600; min-width: 80px;';

    const slider = document.createElement('input');
    slider.type = 'range';
    slider.id = 'merge-gap-slider';
    slider.min = '0'; slider.max = '30'; slider.step = '1'; slider.value = '5';
    slider.style.cssText = 'flex: 1; accent-color: #1a1a1a;';

    const gapValue = document.createElement('span');
    gapValue.id = 'merge-gap-value';
    gapValue.textContent = '5s';
    gapValue.style.cssText = (
      'min-width: 44px; text-align: right; ' +
      'font-variant-numeric: tabular-nums; font-weight: 600;'
    );

    const gapHelp = document.createElement('span');
    gapHelp.textContent = 'consolidates adjacent clips';
    gapHelp.style.cssText = 'opacity: 0.55; font-size: 11px;';

    slider.addEventListener('input', () => { gapValue.textContent = slider.value + 's'; });
    gapRow.appendChild(gapLabel);
    gapRow.appendChild(slider);
    gapRow.appendChild(gapValue);
    gapRow.appendChild(gapHelp);
    panel.appendChild(gapRow);

    // ─── Insert the panel above the export button row ─────────
    const host = btn.parentNode;
    host.parentNode.insertBefore(panel, host);

    // ─── Timeline renderer ─────────────────────────────────────
    const SVG_HEIGHT_PER_ROW = 16;
    const SVG_GAP = 3;
    const TICKS_HEIGHT = 18;

    function getSelectedTrackNames() {
      const cboxes = document.querySelectorAll('input[type="checkbox"][aria-label^="Select "]');
      const out = [];
      for (const cb of cboxes) {
        if (!cb.checked) continue;
        if (cb.id === 'select-all') continue;
        out.push(cb.getAttribute('aria-label').replace(/^Select /, ''));
      }
      return out;
    }

    function chooseTickIntervalSec(totalSec) {
      const tgt = totalSec / 8;
      const steps = [10, 30, 60, 120, 300, 600, 1800, 3600];
      for (const s of steps) if (s >= tgt) return s;
      return 3600;
    }

    function getRange() {
      const fromFrames = tcToFrames(fromIn.value);
      const toFrames = tcToFrames(toIn.value);
      if (fromFrames == null && toFrames == null) return null;
      if (fromFrames == null || toFrames == null) return null;
      if (toFrames <= fromFrames) return null;
      return {
        startSamples: fromFrames * samplesPerFrame,
        endSamples: toFrames * samplesPerFrame,
      };
    }

    function renderTimeline() {
      while (svg.firstChild) svg.removeChild(svg.firstChild);
      const selected = getSelectedTrackNames();
      if (!session || selected.length === 0 || maxSample === 0) {
        svg.setAttribute('height', '0');
        emptyMsg.style.display = '';
        return;
      }
      emptyMsg.style.display = 'none';

      // Width: SVG fills container; we use a virtual viewBox
      const W = 1000;
      const matched = session.tracks.filter(t => selected.includes(t.name));
      const rows = matched.length;
      const H = rows * (SVG_HEIGHT_PER_ROW + SVG_GAP) + TICKS_HEIGHT;
      svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
      svg.setAttribute('height', H);
      svg.setAttribute('preserveAspectRatio', 'none');

      const range = getRange();

      matched.forEach((t, i) => {
        const y = i * (SVG_HEIGHT_PER_ROW + SVG_GAP);
        const colour = colourFor(t.family || t.name);
        // background lane
        const lane = document.createElementNS(svgNs, 'rect');
        lane.setAttribute('x', '0');
        lane.setAttribute('y', y);
        lane.setAttribute('width', W);
        lane.setAttribute('height', SVG_HEIGHT_PER_ROW);
        lane.setAttribute('fill', 'rgba(127,127,127,0.06)');
        svg.appendChild(lane);
        // clip rects
        for (const c of t.clips) {
          const x = (c.s / maxSample) * W;
          const w = Math.max(1, (c.l / maxSample) * W);
          // dim if outside range
          let opacity = 1;
          if (range) {
            const cEnd = c.s + c.l;
            const overlaps = !(cEnd <= range.startSamples || c.s >= range.endSamples);
            opacity = overlaps ? 1 : 0.18;
          }
          const r = document.createElementNS(svgNs, 'rect');
          r.setAttribute('x', x);
          r.setAttribute('y', y + 1);
          r.setAttribute('width', w);
          r.setAttribute('height', SVG_HEIGHT_PER_ROW - 2);
          r.setAttribute('fill', colour);
          r.setAttribute('opacity', opacity);
          svg.appendChild(r);
        }
      });

      // Range overlay lines (vertical)
      if (range) {
        const fromX = (range.startSamples / maxSample) * W;
        const toX = (range.endSamples / maxSample) * W;
        for (const x of [fromX, toX]) {
          const line = document.createElementNS(svgNs, 'line');
          line.setAttribute('x1', x); line.setAttribute('x2', x);
          line.setAttribute('y1', '0'); line.setAttribute('y2', H - TICKS_HEIGHT);
          line.setAttribute('stroke', '#1a1a1a');
          line.setAttribute('stroke-width', '1');
          line.setAttribute('stroke-dasharray', '3 2');
          svg.appendChild(line);
        }
      }

      // Time ticks
      const totalSec = maxSample / session.sample_rate;
      const tickSec = chooseTickIntervalSec(totalSec);
      const tickY = rows * (SVG_HEIGHT_PER_ROW + SVG_GAP);
      for (let t = 0; t <= totalSec + 1e-6; t += tickSec) {
        const x = (t / totalSec) * W;
        const tk = document.createElementNS(svgNs, 'line');
        tk.setAttribute('x1', x); tk.setAttribute('x2', x);
        tk.setAttribute('y1', tickY); tk.setAttribute('y2', tickY + 4);
        tk.setAttribute('stroke', 'currentColor'); tk.setAttribute('opacity', '0.4');
        svg.appendChild(tk);
        const tx = document.createElementNS(svgNs, 'text');
        const frames = Math.round(t * FPS);
        tx.setAttribute('x', x + 3);
        tx.setAttribute('y', tickY + 13);
        tx.setAttribute('font-size', '9');
        tx.setAttribute('font-family', 'ui-monospace, monospace');
        tx.setAttribute('fill', 'currentColor');
        tx.setAttribute('opacity', '0.6');
        tx.textContent = framesToTc(frames);
        svg.appendChild(tx);
      }
    }

    // Re-render whenever any checkbox state changes or range inputs change.
    // Use bubble phase + rAF so we run AFTER the template's own change
    // handlers have propagated state to child checkboxes.
    document.addEventListener('change', () => requestAnimationFrame(renderTimeline));
    document.addEventListener('click', () => requestAnimationFrame(renderTimeline));
    fromIn.addEventListener('input', renderTimeline);
    toIn.addEventListener('input', renderTimeline);
    renderTimeline();

    // ─── Replace the export button's click handler ─────────────
    const clone = btn.cloneNode(true);
    btn.parentNode.replaceChild(clone, btn);

    clone.addEventListener('click', async (e) => {
      e.preventDefault();
      const selected = getSelectedTrackNames();
      if (selected.length === 0) {
        alert('No tracks selected. Tick the tracks you want in the cue sheet.');
        return;
      }

      const mergeGap = parseInt(slider.value, 10);

      // Validate optional time range
      let rangeStart = null, rangeEnd = null;
      const fromVal = fromIn.value.trim();
      const toVal = toIn.value.trim();
      if (fromVal || toVal) {
        const f = tcToFrames(fromVal);
        const t = tcToFrames(toVal);
        if (f == null || t == null) {
          alert('Time range must be in HH:MM:SS:FF format (e.g. 00:01:30:00).');
          return;
        }
        if (t <= f) {
          alert('End time must be after start time.');
          return;
        }
        rangeStart = fromVal;
        rangeEnd = toVal;
      }

      clone.disabled = true;
      const originalText = clone.textContent;
      clone.textContent = 'Generating…';

      const payload = {
        tracks: selected,
        merge_gap_seconds: mergeGap,
      };
      if (rangeStart) payload.range_start = rangeStart;
      if (rangeEnd) payload.range_end = rangeEnd;

      try {
        const res = await fetch('/generate', {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify(payload),
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


def build_session_data_script(session: dict) -> str:
    """Build a slim JSON blob with just enough for the timeline to render.

    Embedded into the HTML as <script id="ms-session-data" type="application/json">.
    """
    # Build reverse map: track name → family
    fam_lookup = {}
    for fam, names in (session.get('track_families') or {}).items():
        for n in names:
            fam_lookup[n] = fam

    payload = {
        'sample_rate': session.get('sample_rate', 48000),
        'tracks': [],
    }
    for tname, clips in (session.get('track_clips') or {}).items():
        payload['tracks'].append({
            'name': tname,
            'family': fam_lookup.get(tname, ''),
            'clips': [
                {'s': c.get('start_samples', 0), 'l': c.get('length_samples', 0)}
                for c in clips
            ],
        })

    # Escape </script> in JSON to prevent breaking out of the script tag
    blob = json.dumps(payload).replace('</', '<\\/')
    return f'<script id="ms-session-data" type="application/json">{blob}</script>'


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
            # merge_gap_seconds: 0-30. Default 5s if missing. Clamped.
            try:
                gap_seconds = int(payload.get('merge_gap_seconds', 5))
            except (TypeError, ValueError):
                gap_seconds = 5
            gap_seconds = max(0, min(30, gap_seconds))
            # optional time range in HH:MM:SS:FF
            range_start = payload.get('range_start')
            range_end = payload.get('range_end')
            result = self._generate(selected, gap_seconds, range_start, range_end)

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

    def _generate(
        self,
        selected_names: list[str],
        gap_seconds: int = 5,
        range_start: str | None = None,
        range_end: str | None = None,
    ) -> dict:
        ptx = self.ptx_path
        base = ptx.stem
        work_dir = ptx.parent
        intermediate_txt = work_dir / f"{base}.cuesheet-tmp.txt"
        # 25fps → frames per second
        merge_gap_frames = gap_seconds * 25

        # Parse time-range strings to sample positions
        sr = self.session.get('sample_rate', 48000) if self.session else 48000
        samples_per_frame = sr / 25
        def tc_to_samples(tc: str):
            m = re.match(r'^(\d{1,2}):(\d{2}):(\d{2}):(\d{2})$', (tc or '').strip())
            if not m:
                return None
            h, mi, s, f = (int(x) for x in m.groups())
            if mi > 59 or s > 59 or f >= 25:
                return None
            return int((h*90000 + mi*1500 + s*25 + f) * samples_per_frame)

        start_samples = tc_to_samples(range_start) if range_start else None
        end_samples = tc_to_samples(range_end) if range_end else None

        try:
            n_clips = write_pt_text_from_session(
                self.session, selected_names, intermediate_txt,
                start_samples=start_samples,
                end_samples=end_samples,
            )
            if n_clips == 0:
                msg = 'No clips found on the selected tracks.'
                if start_samples is not None or end_samples is not None:
                    msg = 'No clips fall within the selected time range.'
                return {'ok': False, 'error': msg}

            # In-process call: works whether script or bundled.
            csv_src, pdf_src = cuesheet.main(
                str(intermediate_txt),
                merge_gap_frames=merge_gap_frames,
            )
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
    if ptx_path.is_dir():
        msg = (
            f"This is a folder, not a Pro Tools session:\\n{ptx_path}\\n\\n"
            "Right-click a .ptx file in Finder, not a folder."
        )
        print(f"error: {msg}", file=sys.stderr)
        show_error_dialog(msg)
        return 2
    if ptx_path.suffix.lower() != '.ptx':
        msg = (
            f"This is not a Pro Tools session file:\\n{ptx_path.name}\\n\\n"
            "Right-click a .ptx file in Finder."
        )
        print(f"error: {msg}", file=sys.stderr)
        show_error_dialog(msg)
        return 2
    if not TEMPLATE_PATH.exists():
        msg = f"report_template.html not found at {TEMPLATE_PATH}."
        print(f"error: {msg}", file=sys.stderr)
        show_error_dialog(msg)
        return 2

    print(f"Parsing {ptx_path.name}…")

    # When bundled, ptunxor lives in sys._MEIPASS (HERE) but the reader looks
    # for it relative to CWD. Make it findable both ways: prepend HERE to PATH,
    # and chdir into the bundle dir during the parse.
    old_cwd = os.getcwd()
    if getattr(sys, 'frozen', False):
        os.environ['PATH'] = str(HERE) + os.pathsep + os.environ.get('PATH', '')
        # Ensure ptunxor is executable (bundled binaries can lose +x)
        ptunxor_path = HERE / 'ptunxor'
        if ptunxor_path.exists():
            try:
                ptunxor_path.chmod(0o755)
            except OSError:
                pass
        os.chdir(str(HERE))

    # Capture stdout/stderr from read_session so failures can be surfaced —
    # bundled apps have no visible console.
    parse_log = io.StringIO()
    try:
        with redirect_stdout(parse_log), redirect_stderr(parse_log):
            session = read_session(ptx_path)
    finally:
        if getattr(sys, 'frozen', False):
            os.chdir(old_cwd)

    diagnostic = parse_log.getvalue()
    if diagnostic.strip():
        print(diagnostic, file=sys.stderr)

    if not session.get('track_clips'):
        msg = (
            "No track clips extracted from this .ptx.\\n\\n"
            "Most likely causes:\\n"
            " - The bundled ptunxor binary can't run on this Mac (Apple Silicon build "
            "won't run on Intel, or vice versa)\\n"
            " - This .ptx is from a Pro Tools version the parser doesn't support\\n"
            " - The .ptx is corrupted or empty\\n\\n"
            "Diagnostic output:\\n"
            + (diagnostic[:600] if diagnostic else "(no output)")
        )
        print(f"error: {msg}", file=sys.stderr)
        show_error_dialog(msg)
        return 3

    html = render_html(session, TEMPLATE_PATH)
    data_script = build_session_data_script(session)
    injection = data_script + INJECTED_SCRIPT
    html_with_handler = (
        html.replace('</body>', injection + '</body>', 1)
        if '</body>' in html else html + injection
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
