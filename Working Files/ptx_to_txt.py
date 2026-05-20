#!/usr/bin/env python3
"""
ptx_to_txt.py — bridge: PTX-extracted session dict → Pro Tools text export format

Provides `write_pt_text_from_session(session, selected_track_names, output_path)`
which writes a Pro Tools text export shaped file containing only the selected
tracks' clips. The resulting file feeds directly into cuesheet.py.

Can also be run as a script for ad-hoc testing:
    python3 ptx_to_txt.py session.json [output.txt]
    (in script mode, includes all MX-prefix tracks by default)
"""
from __future__ import annotations
import json
import re
import sys
from pathlib import Path

# 25fps, 48kHz → 1920 samples per frame
FRAMES_PER_SECOND = 25
SAMPLE_RATE = 48000
SAMPLES_PER_FRAME = SAMPLE_RATE // FRAMES_PER_SECOND  # 1920


def samples_to_tc(samples: int) -> str:
    """Convert samples (48kHz) to HH:MM:SS:FF at 25fps."""
    total_frames = samples // SAMPLES_PER_FRAME
    h = total_frames // 90000
    total_frames %= 90000
    m = total_frames // 1500
    total_frames %= 1500
    s = total_frames // 25
    f = total_frames % 25
    return f"{h:02d}:{m:02d}:{s:02d}:{f:02d}"


def write_pt_text_from_session(
    session: dict,
    selected_track_names: list[str],
    output_path: Path,
    start_samples: int | None = None,
    end_samples: int | None = None,
) -> int:
    """Render selected tracks of a session dict as a Pro Tools text export.

    Args:
        session: parsed session dict from ptx_reader.read_session
        selected_track_names: list of track names (keys into session['track_clips'])
        output_path: where to write the .txt
        start_samples: optional inclusive lower bound; clips ending before this are dropped
        end_samples: optional inclusive upper bound; clips starting after this are dropped

    Returns: number of clip lines written.
    """
    track_clips = session.get('track_clips', {})
    session_name = session.get('session_name', 'Unknown')
    sample_rate = session.get('sample_rate', SAMPLE_RATE)

    # Filter: only tracks that exist in track_clips and have clips
    eligible = [
        (name, track_clips[name])
        for name in selected_track_names
        if name in track_clips and track_clips[name]
    ]

    lines = [
        f"SESSION NAME:\t{session_name}",
        f"SAMPLE RATE:\t{float(sample_rate):.6f}",
        f"BIT DEPTH:\t24-bit",
        f"SESSION START TIMECODE:\t00:00:00:00",
        f"TIMECODE FORMAT:\t25 Frame",
        f"# OF AUDIO TRACKS:\t{len(eligible)}",
        f"# OF AUDIO CLIPS:\t{sum(len(c) for c in track_clips.values())}",
        f"# OF AUDIO FILES:\t{len(session.get('wav_files', []))}",
        "",
        "",
        "T R A C K  L I S T I N G",
        "",
    ]

    clip_total = 0
    for track_name, clips in eligible:
        lines.append(f"TRACK NAME:\t{track_name} (Stereo)")
        lines.append(f"COMMENTS:\t")
        lines.append(f"USER DELAY:\t0 Samples")
        lines.append(f"STATE: ")

        sorted_clips = sorted(clips, key=lambda c: c.get('start_samples', 0))

        for event_num, clip in enumerate(sorted_clips, start=1):
            name = clip['clip_name']
            clip_start = clip.get('start_samples', 0)
            clip_length = clip.get('length_samples', 0)
            clip_end = clip_start + clip_length

            if clip_length == 0:
                continue

            # Time-range filter: skip clips entirely outside the requested range.
            # Keep any clip whose timespan overlaps the range.
            if end_samples is not None and clip_start >= end_samples:
                continue
            if start_samples is not None and clip_end <= start_samples:
                continue

            start_tc = samples_to_tc(clip_start)
            end_tc = samples_to_tc(clip_end)
            duration_tc = samples_to_tc(clip_length)

            line = (
                f"1       \t"
                f"{event_num:<8d}\t"
                f"{name}\t"
                f"   {start_tc}\t"
                f"   {end_tc}\t"
                f"   {duration_tc}\t"
                f"Unmuted"
            )
            lines.append(line)
            clip_total += 1

        lines.append("")
        lines.append("")
        lines.append("")

    output_path.write_text("\n".join(lines), encoding="utf-8")
    return clip_total


# ── Legacy CLI mode ─────────────────────────────────────────────
def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python3 ptx_to_txt.py session.json [output.txt]", file=sys.stderr)
        return 1
    json_path = Path(sys.argv[1])
    if not json_path.exists():
        print(f"error: {json_path} not found", file=sys.stderr)
        return 2
    output_path = Path(sys.argv[2]) if len(sys.argv) >= 3 else json_path.with_suffix('.txt')

    session = json.loads(json_path.read_text())
    MX = re.compile(r'^MX\b', re.IGNORECASE)
    selected = [n for n in session.get('track_clips', {}) if MX.match(n)]

    n = write_pt_text_from_session(session, selected, output_path)
    print(f"Wrote {output_path} ({n} clips from {len(selected)} tracks)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
