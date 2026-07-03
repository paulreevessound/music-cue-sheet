#!/usr/bin/env python3
"""
ptx_clips.py — reconstruct track → clip → audio-file relationships from a
decoded (un-xored) Pro Tools session.

Ports the relevant block-walking logic from zamaudio/ptformat. ptformat's
own version gate rejects PT 13+, but the block *structure* is unchanged —
only the content-type codes shift (this build uses the 0x262a/0x2629 region
codes and the 0x1054/0x1052/0x1050/0x104f region->track map, all of which
ptformat already recognises as alternates).

Chain:
  WAV list   0x1004 -> 0x103a            : indexed audio filenames
  Regions    0x262a -> 0x2629            : region name + 3-point + WAV findex
  Map        0x1054 -> 0x1052 -> 0x1050  : per-track region entries
                       -> 0x104f         : region rawindex + timeline start
"""
from __future__ import annotations
import struct
from ptblocks import parseblocks


def _u(data, pos, n):
    """little-endian unsigned read of n bytes"""
    if pos + n > len(data):
        return 0
    return int.from_bytes(data[pos:pos+n], "little")


def _rstr(data, pos):
    """PT length-prefixed string (4-byte LE length + bytes)"""
    if pos + 4 > len(data):
        return ""
    n = struct.unpack_from("<I", data, pos)[0]
    if n > 4096 or pos + 4 + n > len(data):
        return ""
    return data[pos+4:pos+4+n].decode("latin-1", "replace")


def _walk(node, ct):
    if node["content_type"] == ct:
        yield node
    for c in node["children"]:
        yield from _walk(c, ct)


def _parse_three_point(data, j):
    """Returns (start, offset, length) in samples. Little-endian layout."""
    offsetbytes = (data[j+1] & 0xf0) >> 4
    lengthbytes = (data[j+2] & 0xf0) >> 4
    startbytes  = (data[j+3] & 0xf0) >> 4
    offset = _u(data, j+5, offsetbytes) if 1 <= offsetbytes <= 5 else 0
    j += offsetbytes
    length = _u(data, j+5, lengthbytes) if 1 <= lengthbytes <= 5 else 0
    j += lengthbytes
    start = _u(data, j+5, startbytes) if 1 <= startbytes <= 5 else 0
    return start, offset, length


# Track types as stored in the header track list (low byte of the 2-byte tag).
#   0x00 audio   0x02 bus   0x05 bed   0x08 video/cuts   0x09 folder   0x0b aux
TRACK_TYPES = {0x00, 0x02, 0x05, 0x08, 0x09, 0x0b}


def _track_rec_at(data: bytes, i: int):
    """Parse a header track record at i: [type:2 LE][namelen:4 LE][name].
    Returns (type_byte, name, end_offset) or None."""
    if i + 6 > len(data):
        return None
    if data[i + 1] != 0 or data[i] not in TRACK_TYPES:
        return None
    L = struct.unpack_from("<I", data, i + 2)[0]
    if not (1 <= L <= 48):
        return None
    nm = data[i + 4 + 2:i + 4 + 2 + L]
    if not all(0x20 <= b < 0x7f for b in nm):
        return None
    return (data[i], nm.decode("ascii"), i + 6 + L)


def _next_track_rec(data: bytes, frm: int):
    """Records have variable trailing bytes; find the next one within a window."""
    for k in range(frm, min(frm + 48, len(data))):
        if _track_rec_at(data, k):
            return k
    return None


def extract_track_list(decoded: bytes) -> list:
    """Read the session's own ordered track list from the header.

    Returns [(type_byte, name), ...] in true Pro Tools track order, where
    type is 0x09 folder, 0x0b aux/submaster, 0x00 audio, 0x02 bus. This is the
    authoritative structure as the user arranged it — relabelled folders and
    all — so the tree can be read from the session instead of guessed from
    names. (The block is the template structure and can omit later-added
    tracks; callers should append any clip-bearing track not found here.)"""
    # Find the start: the first offset that begins a run of >=10 records.
    # The list begins inside the cleartext header (well before 0x1000).
    start = None
    i = 0x400
    while i < 0x4000 and start is None:
        j, count = i, 0
        p = _track_rec_at(decoded, j)
        while p:
            count += 1
            if count >= 10:
                start = i
                break
            nxt = _next_track_rec(decoded, p[2])
            if nxt is None:
                break
            j = nxt
            p = _track_rec_at(decoded, j)
        i += 1
    if start is None:
        return []

    out = []
    i = start
    while True:
        p = _track_rec_at(decoded, i)
        if not p:
            break
        out.append((p[0], p[1]))
        nxt = _next_track_rec(decoded, p[2])
        if nxt is None:
            break
        i = nxt
    return out


def extract_clips(decoded: bytes) -> dict:
    """Build the track -> clips mapping. Returns a dict with:
        wav_files:  [filename, ...]                       (indexed)
        regions:    [{name, file, file_index, start,      (indexed)
                      offset, length}, ...]
        tracks:     {track_name: [clip, ...]}  where clip =
                      {clip_name, file, start_samples,
                       length_samples, region_index}
    """
    blocks, _ = parseblocks(decoded)

    def all_of(ct):
        for b in blocks:
            yield from _walk(b, ct)

    # --- 1) WAV file list (0x1004 -> 0x103a) ---
    wav_files: list[str] = []
    for wlb in all_of(0x1004):
        nwavs = _u(decoded, wlb["offset"] + 2, 4)
        for c in wlb["children"]:
            if c["content_type"] != 0x103a:
                continue
            pos = c["offset"] + 11
            end = c["offset"] + c["block_size"]
            count = 0
            while pos < end and count < nwavs + 16:
                nm = _rstr(decoded, pos)
                if not nm:
                    break
                pos += len(nm) + 4
                wavtype = decoded[pos:pos+4].decode("latin-1", "replace")
                pos += 9
                count += 1
                # ptformat skips these pseudo-entries
                if ".grp" in nm or nm in ("Audio Files", "Fade Files"):
                    continue
                if wavtype and not any(t in wavtype for t in ("WAVE", "EVAW", "AIFF", "FFIA")):
                    if not (".wav" in nm.lower() or ".aif" in nm.lower() or ".mp3" in nm.lower()):
                        continue
                wav_files.append(nm)
        break

    # --- 2) Region list (0x262a -> 0x2629) ---
    regions: list[dict] = []
    for rlb in all_of(0x262a):
        for c in rlb["children"]:
            if c["content_type"] != 0x2629:
                continue
            j = c["offset"] + 11
            rname = _rstr(decoded, j)
            j += len(rname) + 4
            start, offset, length = _parse_three_point(decoded, j)
            findex = None
            if c["children"]:
                d = c["children"][0]
                findex = _u(decoded, d["offset"] + d["block_size"], 4)
            fname = wav_files[findex] if (findex is not None and findex < len(wav_files)) else None
            regions.append({
                "name": rname,
                "file_index": findex,
                "file": fname,
                "start": start,
                "offset": offset,
                "length": length,
            })

    # --- 3) Region -> track map (0x1052 -> 0x1050 -> 0x104f) ---
    # 0x1052 entries usually sit under 0x1054, but scan for them directly
    # so we don't miss any that are nested elsewhere.
    tracks: dict[str, list[dict]] = {}
    for entry in all_of(0x1052):
        tname = _rstr(decoded, entry["offset"] + 2)
        if not tname:
            continue
        clips = tracks.setdefault(tname, [])
        for r1050 in entry["children"]:
            if r1050["content_type"] != 0x1050:
                continue
            # byte at offset+46 marks a fade region in ptformat
            is_fade = (r1050["offset"] + 46 < len(decoded)
                       and decoded[r1050["offset"] + 46] == 0x01)
            if is_fade:
                continue
            for sub in r1050["children"]:
                if sub["content_type"] != 0x104f:
                    continue
                # byte at offset+2 of the 0x104f placement is the CLIP MUTE
                # flag (verified against a known-muted track: 1 = muted).
                muted = (sub["offset"] + 2 < len(decoded)
                         and decoded[sub["offset"] + 2] == 0x01)
                j = sub["offset"] + 4
                rawindex = _u(decoded, j, 4)
                j += 4 + 1
                start = _u(decoded, j, 4)
                reg = regions[rawindex] if rawindex < len(regions) else None
                clips.append({
                    "clip_name": reg["name"] if reg else f"(region {rawindex})",
                    "file": reg["file"] if reg else None,
                    "region_index": rawindex,
                    "start_samples": start,
                    "length_samples": reg["length"] if reg else 0,
                    "offset_samples": reg["offset"] if reg else 0,
                    "muted": muted,
                })

    # drop tracks with no clips, keep stable order
    tracks = {k: v for k, v in tracks.items() if v}

    # Canonical AUDIO track list straight from the session (0x1014): the
    # authoritative track names, in session order, with NO name-pattern
    # guessing. Includes tracks that have no clips. Folder tracks aren't here.
    track_names: list[str] = []
    seen_tn: set[str] = set()
    for tnb in all_of(0x1014):
        nm = _rstr(decoded, tnb["offset"] + 2)
        if nm and nm not in seen_tn:
            seen_tn.add(nm)
            track_names.append(nm)

    # --- timecode metadata ---
    # Frame rate is a num/den rational in the single 0x270a block (30000/1001 =
    # 29.97, 25000/1000 = 25, etc.); session start is a frame number in the
    # single 0x204d block (drop-frame-aware).
    fr_num, fr_den = 0, 0
    for e in all_of(0x270a):
        fr_num = _u(decoded, e["offset"] + 43, 4)
        fr_den = _u(decoded, e["offset"] + 47, 4)
        break
    session_start_frames = 0
    for e in all_of(0x204d):
        session_start_frames = _u(decoded, e["offset"] + 11, 4)
        break

    return {
        "wav_files": wav_files,
        "regions": regions,
        "tracks": tracks,
        "track_names": track_names,
        "track_list": extract_track_list(decoded),
        "frame_rate": (fr_num, fr_den),
        "session_start_frames": session_start_frames,
    }


if __name__ == "__main__":
    import sys, json
    data = open(sys.argv[1], "rb").read()
    result = extract_clips(data)
    print(f"WAV files : {len(result['wav_files'])}")
    print(f"regions   : {len(result['regions'])}")
    print(f"tracks w/ clips: {len(result['tracks'])}")
    print()
    # show a few tracks with their clips
    shown = 0
    for tname, clips in result["tracks"].items():
        if shown >= 6:
            break
        if not clips:
            continue
        print(f"TRACK: {tname}  ({len(clips)} clips)")
        for clip in clips[:4]:
            sr = 48000
            start_s = clip["start_samples"] / sr
            len_s = clip["length_samples"] / sr
            print(f"    {clip['clip_name'][:60]:<60}")
            print(f"        file={clip['file']}")
            print(f"        start={start_s:.2f}s  length={len_s:.2f}s")
        if len(clips) > 4:
            print(f"    ... +{len(clips)-4} more")
        print()
        shown += 1
