#!/usr/bin/env python3
"""
ptx_reader.py — read a Pro Tools .ptx session file and pull out
everything that isn't behind the format's XOR obfuscation.

Pro Tools sessions are a partly reverse-engineered binary format (zamaudio
ptformat is the reference). Track lists, region/clip blocks, and plugin
chains are stored XOR-obfuscated with per-block keys, so a clean extractor
isn't realistic without that work. What ISN'T obfuscated, and what this
script pulls out, is plenty useful on its own:

  - Session name and path metadata
  - Audio file references (clip filenames including extension)
  - Other session references (templates, archived sessions)
  - Track names that follow common post-production naming conventions
  - File-system paths
  - ElevenLabs-generated VO clip metadata (voice, params, timestamps)

Usage:
    python3 ptx_reader.py <path/to/session.ptx> [--json out.json] [--html report.html]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import struct
import subprocess
import sys
from collections import Counter, defaultdict
from pathlib import Path

try:
    from ptx_clips import extract_clips
except ImportError:
    extract_clips = None  # clip extraction is optional

# --- string extraction ------------------------------------------------------

PRINTABLE = set(range(0x20, 0x7F))


def is_obfuscated(data: bytes) -> bool:
    """A raw .ptx file keeps its first 4KB header in the clear (where the
    version string lives) and obfuscates blocks from byte 0x1000 onward.
    So we sniff a window past the header where raw and decoded actually
    differ. Decoded files have ~20-40% printable bytes there; encoded ones
    drop below 15%."""
    if len(data) < 0x12000:
        return False
    window = data[0x10000:0x11000]
    printable_ratio = sum(1 for b in window if b in PRINTABLE) / len(window)
    return printable_ratio < 0.18


def find_ptunxor() -> Path | None:
    """Look for the ptunxor binary alongside this script, in PATH, or in a
    sibling 'ptformat' directory."""
    here = Path(__file__).resolve().parent
    candidates = [
        here / "ptunxor",
        here / "bin" / "ptunxor",
        here.parent / "ptformat" / "ptunxor",
        Path("/home/claude/ptformat/ptunxor"),  # dev path
    ]
    for c in candidates:
        if c.exists() and os.access(c, os.X_OK):
            return c
    which = shutil.which("ptunxor")
    return Path(which) if which else None


# Hard ceiling on how long ptunxor may run before we give up and fall back to
# raw bytes. Without this, a session ptunxor can't handle hangs the whole app
# forever (no browser, no error dialog).
PTUNXOR_TIMEOUT_SEC = 120


def _audio_ref_count(data: bytes) -> int:
    """Cheap richness signal: how many audio-file references the byte stream
    contains. A correctly decoded session has hundreds/thousands; an
    obfuscated (or mis-decoded) one has almost none. Used to decide whether
    ptunxor's output is actually better than the raw bytes — far more reliable
    than the printable-ratio heuristic, which false-negatives on some sessions
    (e.g. an obfuscated file whose 0x10000 window happens to be printable)."""
    return sum(data.count(ext) for ext in
               (b".wav", b".WAV", b".aif", b".AIF", b".aiff", b".mp3", b".flac"))


def decode_ptx(path: Path, force: bool = False) -> bytes:
    """Read a .ptx file and return its plain-text byte stream.

    Runs ptunxor (if available) and keeps whichever of {raw, decoded} actually
    has the richer session content — measured by audio-file references. We no
    longer let the is_obfuscated() printable-ratio heuristic *decide* whether
    to decode, because it misfires on real sessions (some obfuscated files read
    as "already decoded" and never get un-XORed, yielding 0 regions). force=True
    keeps ptunxor's output regardless of the comparison."""
    raw = path.read_bytes()
    binary = find_ptunxor()
    if binary is None:
        if is_obfuscated(raw):
            print("  format : OBFUSCATED — ptunxor not found, running with raw bytes "
                  "(extraction will be partial)", file=sys.stderr)
            print("           Build ptformat from https://github.com/zamaudio/ptformat",
                  file=sys.stderr)
        else:
            print(f"  format : no ptunxor; using raw bytes ({len(raw):,})", file=sys.stderr)
        return raw
    try:
        res = subprocess.run(
            [str(binary), str(path)],
            capture_output=True, check=True, timeout=PTUNXOR_TIMEOUT_SEC,
        )
        decoded = res.stdout
    except subprocess.TimeoutExpired:
        print(f"  format : ptunxor timed out after {PTUNXOR_TIMEOUT_SEC}s; "
              "using raw bytes (extraction will be partial)", file=sys.stderr)
        return raw
    except subprocess.CalledProcessError as e:
        print(f"  format : ptunxor failed ({e.returncode}); using raw bytes",
              file=sys.stderr)
        return raw

    raw_refs = _audio_ref_count(raw)
    dec_refs = _audio_ref_count(decoded)
    if force or dec_refs > raw_refs:
        print(f"  format : decoded via {binary.name} "
              f"({len(decoded):,} bytes, audio refs {raw_refs}→{dec_refs})", file=sys.stderr)
        return decoded
    print(f"  format : already decoded; raw kept "
          f"({len(raw):,} bytes, audio refs {raw_refs} ≥ {dec_refs})", file=sys.stderr)
    return raw


def extract_length_prefixed(data: bytes) -> dict[str, int]:
    """Walk the file looking for <4-byte LE length><ASCII bytes> records.

    Returns {string: first_offset}. This is the precise extractor — it
    catches records the format actually wrote with a length prefix.
    """
    out: dict[str, int] = {}
    i, n = 0, len(data)
    while i < n - 4:
        L = struct.unpack_from("<I", data, i)[0]
        if 2 <= L <= 1024 and i + 4 + L <= n:
            chunk = data[i + 4 : i + 4 + L]
            if all(b in PRINTABLE for b in chunk):
                s = chunk.decode("ascii")
                if re.search(r"[A-Za-z0-9]", s) and s not in out:
                    out[s] = i
                i += 4 + L
                continue
        i += 1
    return out


def extract_printable_runs(data: bytes, min_len: int = 6) -> set[str]:
    """Permissive fallback — every run of printable ASCII >= min_len.

    Catches strings that sit inside record blocks where the prefix isn't a
    clean 4-byte length we can detect.
    """
    out: set[str] = set()
    for run in re.findall(rb"[\x20-\x7e]{%d,}" % min_len, data):
        s = run.decode("ascii", "ignore").strip()
        if len(s) >= min_len and re.search(r"[A-Za-z]", s):
            out.add(s)
    return out


# --- classification ---------------------------------------------------------

AUDIO_RE = re.compile(r"\.(wav|aif|aiff|mp3|m4a|caf|sd2|flac|ogg)", re.I)
SESSION_RE = re.compile(r"\.(ptx|pts|ptt)\b", re.I)
PATH_RE = re.compile(r"^/|^[A-Z]:|Macintosh HD|Volumes/|Users/")

# Strict post-production track-name whitelist.
#
# After ptunxor, the binary contains lots of strings that LOOK trackish — e.g.
# "DX 67", "DX 67-68", "DX 612" — but those are bus numbers and stereo bus
# pairs, not track names. The real naming convention in these templates is
# strict, so we whitelist specifically:
#
#   v<FAM> All / v<FAM> Effects All     — virtual master folders
#   <FAM> All / <FAM> Effects All       — top-level master Aux
#   <FAM> Master Bus / Master Bed / Bed / Global Effects
#   v<FAM> <letter>                      — virtual subgroup folder
#   <FAM> <letter>                       — sub-master Aux
#   <FAM> <letter><digit 1-9>            — leaf audio (a1..a8 etc.)
#   m/s/i <FAM> <digit>                  — mix/source/print bus
#   Boom <FAM>                           — boom feed
#
# Plus the non-family-prefixed utility/record tracks separately.

_FAM = r"(?:NAR|DX|SFX|MX|FX|VO|BG|FOL|FOLEY|MUS|MUSIC|REV|SUB|OPT)"
_TRACK_PATTERN = (
    r"^(?:"
    rf"v?{_FAM}\s(?:All|Effects\sAll)"
    rf"|{_FAM}\s(?:Master\sBus|Master\sBed|Bed|Global\sEffects)"
    rf"|v?{_FAM}\s[a-z]"
    rf"|{_FAM}\s[a-z][1-9]"
    rf"|[msi]\s{_FAM}\s\d"
    rf"|Boom\s{_FAM}"
    r"|Source Connect.*"
    r"|VO Booth.*"
    r"|Stereo VO.*"
    r"|ADR Record.*"
    r"|Mac(?:\s\w+)?"
    r"|Video\s*\d+"
    r"|Talkback"
    r"|Timecode"
    r"|Reverb\s+\w+"
    r"|Basehead"
    r"|Ext\sClients"
    r")$"
)
TRACK_RE = re.compile(_TRACK_PATTERN, re.I)


VIDEO_RE = re.compile(r"^[\w \-_.()\[\]]+\.(?:mov|mp4|mxf|m2v)\b", re.I)
# Plugin signatures we look for in the decoded binary
PLUGIN_HINTS = [
    "FabFilter Pro-Q 3", "FabFilter Pro-Q 4", "FabFilter Pro-C 2",
    "FabFilter Pro-L 2", "FabFilter Pro-DS", "FabFilter Pro-G",
    "FabFilter Pro-MB", "FabFilter Pro-R", "FabFilter Saturn 2",
    "iZotope RX", "iZotope Ozone", "iZotope Neutron",
    "Soothe 2", "Soothe 3", "Gullfoss", "DeBleeder",
    "Waves SSL", "Waves L2", "Waves L1", "Waves Q10", "Waves CLA",
    "Avid De-Esser", "Avid Compressor", "Avid EQ", "Avid Reverb",
    "Pro Tools EQ", "Pro Tools Compressor", "Pro Tools De-Esser",
    "AltiVerb", "Reverb Profiling", "Trim", "Channel Strip", "Eleven",
    "BX_", "Slate Digital", "UAD",
]


def extract_sample_rate(data: bytes) -> int | None:
    """Pro Tools stores the sample rate as a uint32 LE near offset 0x18000-
    0x19000 in the un-obfuscated file. Verify by looking for the canonical
    rates at plausible offsets."""
    for sr in (44100, 48000, 88200, 96000, 176400, 192000):
        b = struct.pack("<I", sr)
        # Look in the first ~256 KB where session header lives
        pos = data[:0x40000].find(b)
        if pos >= 0:
            return sr
    return None


def extract_video_refs(data: bytes) -> list[str]:
    """Pull standalone video filenames out of the decoded stream.

    Pro Tools writes references like '...v2.0_.movVooM' where the file
    extension is immediately followed by a trailing marker, so \\b doesn't
    fire — we anchor with a lookahead instead."""
    runs = re.findall(rb"[\x20-\x7e]{8,}", data)
    out = []
    seen = set()
    # filename: word chars, spaces, dashes, dots, parens — then .ext, then a non-alnum
    # (?i:...) makes only the inner part case-insensitive; the outer
    # lookahead stays case-sensitive so "movVooM" doesn't trigger it.
    pat = re.compile(r"([\w \-.()\[\]]+?(?i:\.(?:mov|mp4|mxf|m2v|prproj)))(?![a-z])")
    for r in runs:
        s = r.decode("ascii", "ignore")
        for m in pat.finditer(s):
            v = m.group(1).strip(" .-_")
            if v not in seen and len(v) >= 5:
                seen.add(v)
                out.append(v)
    return sorted(out)


def extract_plugin_names(data: bytes) -> list[str]:
    """Scan the decoded bytes for known plugin name signatures."""
    found = Counter()
    text = data.decode("latin-1", "ignore")
    for hint in PLUGIN_HINTS:
        c = text.count(hint)
        if c:
            found[hint] = c
    # Order by appearance frequency in the file
    return [name for name, _ in found.most_common()]


def looks_like_xor_noise(s: str) -> bool:
    """Reject strings that look like XOR-deobfuscation garbage.

    Symptoms: long runs of identical chars (e.g. 'pppppp', 'TTTTk'),
    low-information patterns, no meaningful word structure.
    """
    if re.search(r"(.)\1{3,}", s):  # 4+ identical chars in a row
        return True
    # All-lowercase blob with no spaces and no recognised words
    letters = re.sub(r"[^A-Za-z]", "", s)
    if len(letters) >= 5 and letters.islower() and " " not in s and "_" not in s and "-" not in s:
        # things like "fxeohhh", "bgttttk", "mxxib" — no English-ish patterns
        if not re.search(r"[aeiou]{1,2}[bcdfghjklmnpqrstvwxyz]", letters):
            return True
    return False


def looks_like_real_path(s: str) -> bool:
    """Filter out XOR-noise that begins with '/' but isn't actually a path."""
    if not (s.startswith("/") or re.match(r"^[A-Z]:", s) or s.startswith("Macintosh")):
        return False
    # require either a recognisable root, multiple path segments,
    # or a file extension
    if any(root in s for root in ("Volumes/", "Users/", "Library/", "Macintosh HD",
                                   "/Applications", "/System", "/Audio Files",
                                   "/Session File", "/Desktop", "/Documents")):
        return True
    # at least 2 slashes and mostly alphanumeric segments
    parts = [p for p in s.split("/") if p]
    if len(parts) >= 2 and all(re.match(r"^[\w\-. ]+$", p) for p in parts):
        return True
    return False

ELEVENLABS_RE = re.compile(
    r"ElevenLabs_(\d{4}-\d{2}-\d{2})T([\d_]+)_"
    r"([^_]+(?:\s+-\s+[^_]+)?)_"
    r"pvc_([\w_]+)\.mp3(.*)"
)


def family_of(track: str) -> str:
    """Bucket a track into its post-prod family.

    The naming covers many variations all referring to the same family:
      NAR / vNAR / Boom NAR / m NAR 1 / s NAR 2 / i NAR 1  → "NAR"
      MUS / MUSIC                                           → "MUSIC"
      FOL / FOLEY / vFOL                                    → "FOLEY"
      OPT / vOPT / Boom OPT                                 → "OPT"
    """
    t = track.strip()
    upper = t.upper()

    # Longer prefixes first so MUSIC beats MUS, FOLEY beats FOL.
    FAM_MAP = [
        ("MUSIC", "MUSIC"), ("MUS", "MUSIC"),
        ("FOLEY", "FOLEY"), ("FOL", "FOLEY"),
        ("NAR", "NAR"),
        ("OPT", "OPT"),
        ("SFX", "SFX"),
        ("DX", "DX"),
        ("MX", "MX"),
        ("FX", "FX"),
        ("BG", "BG"),
        ("VO", "VO"),
    ]
    def lookup(s: str) -> str | None:
        for prefix, normalized in FAM_MAP:
            if s.startswith(prefix):
                return normalized
        return None

    # m/s/i mix bus: "m NAR 1", "s DX 2"
    m = re.match(r"^[MSI]\s+([A-Z]+)\s+\d+$", upper)
    if m:
        fam = lookup(m.group(1))
        return f"{fam} (mix bus)" if fam else "Other"

    # Boom <FAM>
    if upper.startswith("BOOM "):
        fam = lookup(upper[5:].strip())
        return fam or "Other"

    # v<FAM> ...
    if upper.startswith("V") and len(upper) > 1 and upper[1].isalpha():
        # Try treating the rest as a family; if not a v-prefixed family fall
        # through to the plain check (e.g. "VO" is its own family, not v+O)
        if upper.startswith("VO ") or upper == "VO":
            return "VO"
        fam = lookup(upper[1:])
        if fam:
            return fam

    fam = lookup(upper)
    if fam:
        return fam

    # Reverb sends/buses
    if upper.startswith("REVERB ") or upper.startswith("REV "):
        return "Reverb"

    # Utility / record / external buckets
    if "Source Connect" in t:
        return "External Sources"
    if "VO Booth" in t or "Stereo VO" in t or "ADR" in t:
        return "Record"
    if t.startswith("Mac") or t in ("Basehead", "Ext Clients"):
        return "External Sources"
    if t in ("Talkback", "Timecode") or t.startswith("Video"):
        return "Utility"
    return "Other"


# --- track type & hierarchy inference --------------------------------------
#
# We don't know the actual Pro Tools track type from the binary — that data
# is XOR'd. But naming conventions in post-production templates are strong
# enough to make a confident guess:
#
#   "vNAR All" / "vDX All"          — virtual master bus     → Folder
#   "NAR Master Bus"                — top-level master         → Aux Master
#   "NAR Global Effects"            — fx send                 → Aux
#   "NAR a", "NAR b"                — sub-folder per group     → Folder/Aux
#   "NAR a1", "NAR a2"              — actual narration tracks  → Audio
#   "m NAR 1", "s NAR 1", "i NAR 1" — mix/source/print buses  → Aux
#   "Boom NAR", "Boom DX"           — boom feed                → Audio
#   "VO Booth Record"               — record path              → Audio
#   "Source Connect"                — external feed            → Aux
#   "Talkback"                      — utility                  → Aux
#   "Timecode" / "Video N"          — utility / video          → Video/Aux

LEAF_RE = re.compile(r"^(NAR|DX|SFX|MX|FX|VO|BG|MUS|MUSIC)\s+[a-z]\d+$", re.I)
SUBGROUP_RE = re.compile(r"^v?(NAR|DX|SFX|MX|FX|VO|BG|MUS|MUSIC)\s+[a-z]$", re.I)


def infer_type(track: str) -> str:
    t = track.strip()
    if t.startswith(("vNAR", "vDX", "vSFX", "vMX", "vFX", "vVO", "vBG")):
        return "Folder"
    if "Master Bus" in t or "Master Bed" in t:
        return "Aux Master"
    if "Global Effects" in t or "Effects All" in t:
        return "Aux"
    if SUBGROUP_RE.match(t):
        return "Folder"
    if LEAF_RE.match(t):
        return "Audio"
    if re.match(r"^[msi]\s+(NAR|DX|SFX|MX|FX)\s*\d*$", t):
        return "Aux"
    if t.startswith("Boom"):
        return "Audio"
    if "Booth Record" in t or "VO Record" in t or "ADR Record" in t or t == "Mac Record":
        return "Audio"
    if t in ("Source Connect", "Source Connect VO Record", "Mac Dante", "Talkback", "Basehead", "Ext Clients"):
        return "Aux"
    if t == "Timecode":
        return "Aux"
    if t.startswith("Video"):
        return "Video"
    if "Trim" in t or "Reverb" in t or t == "Utility":
        return "Aux"
    return "Aux"


def infer_format(track: str, ttype: str) -> str:
    """Best-guess channel format. Mono for VO booth/boom; stereo otherwise."""
    if ttype == "Folder":
        return "Mono"  # folder containers don't have a real format
    if "Booth Record" in track and "Stereo" not in track:
        return "Mono"
    if "Boom" in track:
        return "Mono"
    return "Stereo"


# Pro Tools-ish track color palette — keyed by family
FAMILY_COLOR = {
    "NAR":     "#c54545",   # red
    "VNAR":    "#a64545",
    "DX":      "#2a8a8a",   # teal
    "MX":      "#d2a13d",   # tape yellow / orange
    "MUSIC":   "#d2a13d",
    "MUS":     "#d2a13d",
    "FX":      "#7c5dc7",   # violet
    "SFX":     "#7c5dc7",
    "VO":      "#c25b9b",   # magenta
    "BG":      "#5a86c4",   # blue
    "FOLEY":   "#86b049",   # green
    "External Sources": "#6b6b6b",
    "Record":  "#c25b9b",
    "Utility": "#9aa0a6",
    "Other":   "#9aa0a6",
}
def color_for(family: str) -> str:
    f = family.upper().split(" (")[0]
    return FAMILY_COLOR.get(f, FAMILY_COLOR["Other"])


# Shared family regex group for hierarchy patterns. Uses the same set as
# _FAM above (track-name whitelist) so the tree builder agrees on which
# prefixes count as families.
_FAM_HIER = r"(?:NAR|DX|SFX|MX|FX|VO|BG|FOL|FOLEY|MUS|MUSIC|OPT)"


def build_hierarchy(tracks: list[str]) -> list[dict]:
    """Build a nested tree from flat track names using naming conventions.

    Levels:
      v<FAM> All                              (root folder per family)
        <FAM> Master Bus / Master Bed / Global Effects
        v<FAM> Effects All / <FAM> Effects All
        <FAM> Bed
        v<FAM> a / b / c / d                  (virtual subgroup folders)
          <FAM> a / b / c / d                 (sub-master Aux)
            <FAM> a1 .. a8                    (leaf audio)
        Boom <FAM>                            (leaf audio under root)
      m/s/i <FAM> N                           (mix-bus group, separate root)
      Top-level groups (no virtual parent):
        Utility, External Sources, Record Tracks, Video
    """
    # Index tracks by category
    by_family_root = {}    # "NAR" -> root track node
    subgroups = {}         # ("NAR","a") -> subgroup node
    leaves_by_sub = defaultdict(list)
    floating_under_root = defaultdict(list)
    mix_bus = defaultdict(list)
    utility = []
    external = []
    record = []
    video = []
    generic = defaultdict(list)   # non-standard but real tracks, grouped by prefix
    other = []

    def normalize_fam(s: str) -> str:
        s = s.upper()
        return {"FOL": "FOLEY", "MUS": "MUSIC"}.get(s, s)

    def node(name, ttype, family, children=None, depth=0, fmt=None):
        return {
            "name": name,
            "type": ttype,
            "format": fmt if fmt is not None else infer_format(name, ttype),
            "family": family,
            "color": color_for(family),
            "depth": depth,
            "children": children or [],
        }

    for t in tracks:
        upper = t.upper()
        # virtual root: "vNAR All", "vDX All", "vFOL All"...
        m = re.match(rf"^v({_FAM_HIER})\s+All$", t, re.I)
        if m:
            fam = normalize_fam(m.group(1))
            by_family_root[fam] = node(t, "Folder", fam, depth=0, fmt="Mono")
            continue
        # virtual effects-all folder: "vDX Effects All" — child of family root
        m = re.match(rf"^v({_FAM_HIER})\s+Effects\s+All$", t, re.I)
        if m:
            fam = normalize_fam(m.group(1))
            floating_under_root[fam].append(node(t, "Folder", fam, depth=1, fmt="Mono"))
            continue
        # m/s/i mix bus: "m NAR 1", "s DX 2"...
        m = re.match(rf"^([msi])\s+({_FAM_HIER})\s*(\d*)$", t, re.I)
        if m:
            tag, fam_raw, num = m.groups()
            fam = normalize_fam(fam_raw)
            mix_bus[fam].append(node(t, "Aux", fam, depth=1))
            continue
        # leaf: "NAR a1" — single letter + single digit
        m = re.match(rf"^({_FAM_HIER})\s+([a-z])([1-9])$", t, re.I)
        if m:
            fam_raw, sub, num = m.groups()
            fam = normalize_fam(fam_raw)
            key = (fam, sub.lower())
            leaves_by_sub[key].append(node(t, "Audio", fam, depth=2))
            continue
        # virtual subgroup folder: "vNAR a", "vDX b"
        m = re.match(rf"^v({_FAM_HIER})\s+([a-z])$", t, re.I)
        if m:
            fam_raw, sub = m.groups()
            fam = normalize_fam(fam_raw)
            key = (fam, sub.lower())
            new_sg = node(t, "Folder", fam, depth=1, fmt="Mono")
            existing = subgroups.get(key)
            if existing and "_sub_masters" in existing:
                new_sg["_sub_masters"] = existing["_sub_masters"]
            subgroups[key] = new_sg
            continue
        # sub-master aux (no v prefix): "NAR a", "DX b"
        m = re.match(rf"^({_FAM_HIER})\s+([a-z])$", t, re.I)
        if m:
            fam_raw, sub = m.groups()
            fam = normalize_fam(fam_raw)
            sub_master = node(t, "Aux", fam, depth=2)
            key = (fam, sub.lower())
            if key in subgroups:
                subgroups[key].setdefault("_sub_masters", []).append(sub_master)
            else:
                subgroups[key] = node(f"{fam_raw} {sub}", "Folder", fam, depth=1, fmt="Mono")
                subgroups[key]["_sub_masters"] = [sub_master]
                subgroups[key]["synthetic"] = True
            continue
        # Family-prefixed but not a leaf or subgroup (e.g. "NAR Master Bus")
        m = re.match(rf"^({_FAM_HIER})\b", t, re.I)
        if m:
            fam = normalize_fam(m.group(1))
            ttype = infer_type(t)
            floating_under_root[fam].append(node(t, ttype, fam, depth=1))
            continue
        if t.startswith("Boom"):
            # find which family it belongs to
            suffix = t[len("Boom"):].strip().upper()
            fam = normalize_fam(suffix) if re.match(rf"^{_FAM_HIER}$", suffix) else suffix
            floating_under_root.setdefault(fam, []).append(node(t, "Audio", fam, depth=1))
            continue
        if t in ("Talkback", "Timecode") or "Trim" in t or t.startswith("Reverb"):
            utility.append(node(t, infer_type(t), "Utility", depth=0))
            continue
        if t.startswith("Video"):
            video.append(node(t, "Video", "Utility", depth=0, fmt="Stereo"))
            continue
        if t.startswith("Mac") or "Source Connect" in t or t in ("Basehead", "Ext Clients"):
            external.append(node(t, infer_type(t), "External Sources", depth=0))
            continue
        if "Booth Record" in t or "Stereo VO" in t or "ADR Record" in t:
            record.append(node(t, "Audio", "Record", depth=0))
            continue
        # Non-standard name (e.g. "Sync FX a1", "Spot FX b2", "GFX FX e5",
        # "Editor BG d3") — the studio uses prefixes the family whitelist
        # doesn't know. Group by the prefix left of a trailing
        # "<letter(s)><digits>" channel designator so these still appear as
        # tidy folders rather than being dropped or dumped flat.
        m = re.match(r'^(.*\S)\s+[A-Za-z]{1,3}\s*\d{0,3}$', t)
        grp = m.group(1).strip() if m else None
        if grp and len(grp) >= 2:
            generic[grp].append(node(t, "Audio", grp, depth=1))
        else:
            other.append(node(t, infer_type(t), "Other", depth=0))

    # Stitch tree together
    roots = []

    if utility:
        roots.append({"name": "Utility", "type": "Folder", "format": "Mono",
                      "family": "Utility", "color": color_for("Utility"),
                      "depth": 0, "children": utility, "synthetic": True})
    if video:
        roots.append({"name": "Video", "type": "Folder", "format": "Stereo",
                      "family": "Utility", "color": color_for("Utility"),
                      "depth": 0, "children": video, "synthetic": True})
    if external:
        roots.append({"name": "External Sources", "type": "Folder", "format": "Mono",
                      "family": "External Sources", "color": color_for("External Sources"),
                      "depth": 0, "children": external, "synthetic": True})
    if record:
        roots.append({"name": "Record Tracks", "type": "Folder", "format": "Mono",
                      "family": "Record", "color": color_for("Record"),
                      "depth": 0, "children": record, "synthetic": True})

    # Family roots: now includes the full set
    for fam in ("NAR", "DX", "MX", "MUSIC", "FX", "SFX", "VO", "BG", "FOLEY", "OPT"):
        root = by_family_root.get(fam)
        if root is None and fam not in floating_under_root and fam not in mix_bus:
            # are there subgroups for this family?
            if not any(f == fam for (f, _) in subgroups):
                continue
        if root is None:
            root = {"name": fam, "type": "Folder", "format": "Mono",
                    "family": fam, "color": color_for(fam),
                    "depth": 0, "children": [], "synthetic": True}
        for n in floating_under_root.get(fam, []):
            root["children"].append(n)
        for (f, sub), sg in sorted(subgroups.items()):
            if f != fam:
                continue
            sg = dict(sg)
            children = []
            for sm in sg.pop("_sub_masters", []):
                children.append(sm)
            children.extend(sorted(leaves_by_sub.get((f, sub), []), key=lambda x: x["name"]))
            sg["children"] = children
            root["children"].append(sg)
        for (f, sub), leaves in leaves_by_sub.items():
            if f == fam and (f, sub) not in subgroups:
                for n in leaves:
                    root["children"].append(n)
        roots.append(root)

    for fam, items in mix_bus.items():
        roots.append({"name": f"{fam} mix bus", "type": "Folder", "format": "Mono",
                      "family": fam, "color": color_for(fam),
                      "depth": 0, "children": sorted(items, key=lambda x: x["name"]),
                      "synthetic": True})

    for grp, items in sorted(generic.items()):
        roots.append({"name": grp, "type": "Folder", "format": "Mono",
                      "family": grp, "color": color_for(grp),
                      "depth": 0, "children": sorted(items, key=lambda x: x["name"]),
                      "synthetic": True})

    if other:
        roots.append({"name": "Other", "type": "Folder", "format": "Mono",
                      "family": "Other", "color": color_for("Other"),
                      "depth": 0, "children": other, "synthetic": True})

    return roots


def flatten_tree(roots: list[dict]) -> list[dict]:
    """Walk the tree depth-first, emitting flat rows with depth info."""
    out = []
    def walk(node, depth):
        n = {k: v for k, v in node.items() if k != "children"}
        n["depth"] = depth
        n["has_children"] = bool(node.get("children"))
        out.append(n)
        for c in node.get("children", []):
            walk(c, depth + 1)
    for r in roots:
        walk(r, 0)
    return out


def parse_elevenlabs(filename: str) -> dict | None:
    m = ELEVENLABS_RE.search(filename)
    if not m:
        return None
    date, t, voice, params, suffix = m.groups()
    return {
        "date": date,
        "time": t.replace("_", ":"),
        "voice": voice.strip(),
        "params": params,
        "suffix": suffix.strip(" -_"),
        "file": filename,
    }


# --- driver -----------------------------------------------------------------


def read_session(path: Path, force_decode: bool = False) -> dict:
    data = decode_ptx(path, force=force_decode)
    raw_size = path.stat().st_size

    lp = extract_length_prefixed(data)
    runs = extract_printable_runs(data)
    all_strings = set(lp) | runs

    audio_clips = sorted({s for s in all_strings if AUDIO_RE.search(s)})
    session_refs = sorted({s for s in all_strings if SESSION_RE.search(s) and not AUDIO_RE.search(s)})
    paths = sorted({s for s in all_strings if PATH_RE.match(s) and looks_like_real_path(s)})
    video_refs = extract_video_refs(data)

    # --- track -> clip -> file mapping (ported ptformat block walk) ---
    # Extracted early: clip-bearing track names are authoritative and feed the
    # track list below.
    clip_data = {"wav_files": [], "regions": [], "tracks": {}}
    if extract_clips is not None:
        try:
            clip_data = extract_clips(data)
        except Exception as e:  # pragma: no cover - defensive
            print(f"  clips  : extraction failed ({e})", file=sys.stderr)
    track_clips = clip_data.get("tracks", {})

    # tracks: short strings (<=60c) that look like post-prod naming, UNIONed
    # with every track that actually has clips. A track with audio must be
    # selectable even if its name doesn't match the family whitelist (the
    # studio uses prefixes like "Sync FX", "Spot FX", "GFX", "Editor BG"…).
    short = [s for s in all_strings if len(s) <= 60 and not AUDIO_RE.search(s) and not SESSION_RE.search(s) and not PATH_RE.match(s)]
    whitelisted = {s for s in short if TRACK_RE.match(s) and not looks_like_xor_noise(s)}
    tracks = sorted(whitelisted | set(track_clips))

    # session metadata that's now readable
    sample_rate = extract_sample_rate(data)
    plugin_names = extract_plugin_names(data)

    families: dict[str, list[str]] = defaultdict(list)
    for t in tracks:
        families[family_of(t)].append(t)
    # keep families sorted by track count desc
    families = dict(sorted(families.items(), key=lambda kv: -len(kv[1])))

    hierarchy = build_hierarchy(tracks)
    flat_tree = flatten_tree(hierarchy)

    total_clips = sum(len(v) for v in track_clips.values())
    # attach clips to matching rows in the flat tree
    for row in flat_tree:
        clips = track_clips.get(row["name"])
        if clips:
            row["clips"] = clips
    print(f"  clips  : {total_clips} clips across {len(track_clips)} tracks "
          f"({len(clip_data.get('wav_files', []))} WAV files, "
          f"{len(clip_data.get('regions', []))} regions)", file=sys.stderr)

    eleven = [parsed for c in audio_clips if (parsed := parse_elevenlabs(c))]

    ext_count: Counter[str] = Counter()
    for c in audio_clips:
        m = re.search(r"\.(\w{2,4})(?:[\-_\s]|$)", c)
        if m:
            ext_count[m.group(1).lower()] += 1

    # detect session display name from the start of the file (the metadata
    # block keeps it readable even when the rest is obfuscated)
    name_match = re.search(rb"\x1f\x00\x00\x00([\x20-\x7e]{31})", data)
    session_name = name_match.group(1).decode("ascii") if name_match else path.stem

    return {
        "file": path.name,
        "size_bytes": raw_size,
        "session_name": session_name,
        "sample_rate": sample_rate,
        "video_files": video_refs,
        "plugins": plugin_names,
        "track_clips": track_clips,
        "wav_files": clip_data.get("wav_files", []),
        "counts": {
            "unique_strings": len(all_strings),
            "length_prefixed_strings": len(lp),
            "audio_clips": len(audio_clips),
            "elevenlabs_clips": len(eleven),
            "tracks": len(tracks),
            "tracks_with_clips": len(track_clips),
            "total_clips": total_clips,
            "wav_files": len(clip_data.get("wav_files", [])),
            "regions": len(clip_data.get("regions", [])),
            "session_refs": len(session_refs),
            "paths": len(paths),
            "video_files": len(video_refs),
            "plugins": len(plugin_names),
        },
        "audio_clips": audio_clips,
        "session_refs": session_refs,
        "paths": paths,
        "tracks": tracks,
        "track_families": families,
        "track_tree": hierarchy,
        "track_rows": flat_tree,
        "elevenlabs": eleven,
        "elevenlabs_voices": dict(Counter(e["voice"] for e in eleven)),
        "elevenlabs_by_date": dict(sorted(Counter(e["date"] for e in eleven).items())),
        "audio_extensions": dict(ext_count),
    }


def render_html(session: dict, template_path: Path) -> str:
    """Inline the parsed session JSON into the HTML template."""
    template = template_path.read_text()
    blob = json.dumps(session)
    # the template has a marker <!--SESSION_DATA--> that we replace with
    # a script tag carrying the JSON
    return template.replace(
        "<!--SESSION_DATA-->",
        f'<script id="session-data" type="application/json">{blob}</script>',
    )


def main() -> int:
    p = argparse.ArgumentParser(description="Parse a Pro Tools .ptx session file")
    p.add_argument("ptx", type=Path, help="path to the .ptx file")
    p.add_argument("--json", type=Path, help="write parsed result here (default: stdout)")
    p.add_argument("--html", type=Path, help="write standalone HTML report here")
    p.add_argument("--template", type=Path, default=Path(__file__).with_name("report_template.html"),
                   help="HTML template to inline data into (used with --html)")
    args = p.parse_args()

    if not args.ptx.exists():
        print(f"error: {args.ptx} not found", file=sys.stderr)
        return 2

    session = read_session(args.ptx)

    if args.json:
        args.json.write_text(json.dumps(session, indent=2))
        print(f"wrote {args.json}", file=sys.stderr)
    elif not args.html:
        json.dump(session, sys.stdout, indent=2)
        sys.stdout.write("\n")

    if args.html:
        if not args.template.exists():
            print(f"error: template {args.template} not found", file=sys.stderr)
            return 2
        args.html.write_text(render_html(session, args.template))
        print(f"wrote {args.html}", file=sys.stderr)

    print(
        f"  session : {session['session_name']}\n"
        f"  size    : {session['size_bytes']:,} bytes\n"
        f"  tracks  : {session['counts']['tracks']}\n"
        f"  clips   : {session['counts']['audio_clips']} ({session['counts']['elevenlabs_clips']} ElevenLabs)\n"
        f"  strings : {session['counts']['unique_strings']:,} unique",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
