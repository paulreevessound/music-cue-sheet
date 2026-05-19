#!/bin/bash
# build.sh — build Music Cue Sheet.app on macOS
#
# Run this once on your work machine to produce a shareable .app + .zip.
# Output ends up in ./dist/.
#
# Assumes you're inside the Working Files folder.

set -euo pipefail

cd "$(dirname "$0")"

echo "Building Music Cue Sheet.app"
echo "=============================="
echo

# ── 1. Dependency checks ─────────────────────────────────────────
echo "[1/5] Checking dependencies…"

command -v python3 >/dev/null || { echo "  ERROR: python3 not found. Install via 'brew install python' or python.org."; exit 1; }

PY_OK=$(python3 -c 'import sys; print(1 if sys.version_info >= (3, 10) else 0)')
if [ "$PY_OK" != "1" ]; then
    echo "  ERROR: Python 3.10+ required. Found: $(python3 --version)"
    echo "  Install a newer Python with: brew install python"
    exit 1
fi
echo "  Python: $(python3 --version)"

command -v git >/dev/null || { echo "  ERROR: git not found. Install with: xcode-select --install"; exit 1; }
command -v make >/dev/null || { echo "  ERROR: make not found. Install with: xcode-select --install"; exit 1; }


# ── 2. Python packages ───────────────────────────────────────────
echo
echo "[2/5] Installing Python packages (pyinstaller, reportlab)…"
python3 -m pip install --quiet --upgrade pip
python3 -m pip install --quiet pyinstaller reportlab
echo "  Done."


# ── 3. ptunxor helper binary ─────────────────────────────────────
echo
echo "[3/5] Checking ptunxor helper binary…"
if [ -f ./ptunxor ]; then
    echo "  Found ./ptunxor"
else
    echo "  Building ptunxor from zamaudio/ptformat…"
    BUILDDIR=$(mktemp -d)
    git clone --depth 1 https://github.com/zamaudio/ptformat "$BUILDDIR/ptformat"
    ( cd "$BUILDDIR/ptformat" && make )
    cp "$BUILDDIR/ptformat/ptunxor" ./ptunxor
    chmod +x ./ptunxor
    rm -rf "$BUILDDIR"
    echo "  Built and copied to ./ptunxor"
fi


# ── 4. PyInstaller build ─────────────────────────────────────────
echo
echo "[4/5] Building app bundle…"
rm -rf build dist *.spec

python3 -m PyInstaller \
    --windowed \
    --noconfirm \
    --name "Music Cue Sheet" \
    --osx-bundle-identifier "studio.mightysound.cuesheet" \
    --add-data "report_template.html:." \
    --add-data "logo.png:." \
    --add-binary "ptunxor:." \
    --hidden-import "ptblocks" \
    --hidden-import "ptx_clips" \
    --hidden-import "ptx_reader" \
    --hidden-import "ptx_to_txt" \
    --hidden-import "cuesheet" \
    --hidden-import "reportlab" \
    --hidden-import "reportlab.pdfgen" \
    --hidden-import "reportlab.platypus" \
    --hidden-import "reportlab.lib" \
    --hidden-import "reportlab.lib.pagesizes" \
    --hidden-import "reportlab.lib.styles" \
    --hidden-import "reportlab.lib.units" \
    --hidden-import "reportlab.lib.colors" \
    ptx_cuesheet.py

echo "  Bundled to dist/Music Cue Sheet.app"


# ── 5. Create shareable zip ──────────────────────────────────────
echo
echo "[5/5] Creating shareable zip…"
cd dist
rm -f "Music Cue Sheet.zip"
# Use `ditto` (not Finder Compress) to preserve permissions/attributes
ditto -c -k --sequesterRsrc --keepParent "Music Cue Sheet.app" "Music Cue Sheet.zip"
cd ..

SIZE_APP=$(du -sh "dist/Music Cue Sheet.app" | cut -f1)
SIZE_ZIP=$(du -sh "dist/Music Cue Sheet.zip" | cut -f1)

echo
echo "=============================="
echo "✓ Build complete"
echo
echo "  App:  dist/Music Cue Sheet.app   ($SIZE_APP)"
echo "  Zip:  dist/Music Cue Sheet.zip   ($SIZE_ZIP)"
echo
echo "Test locally:"
echo "  open 'dist/Music Cue Sheet.app' --args '/path/to/some test.ptx'"
echo
echo "Share the zip with testers along with the user guide PDF."
