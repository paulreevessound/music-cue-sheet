#!/usr/bin/env python3
"""Parse the PTX block tree from a decoded (un-xored) Pro Tools session.

Block header layout (from zamaudio/ptformat parse_block_at):
  pos+0       ZMARK  (0x5a)
  pos+1..2    block_type    (uint16 LE)
  pos+3..6    block_size    (uint32 LE)
  pos+7..8    content_type  (uint16 LE)
  content/children start at pos+9, run for block_size-2 bytes
  (offset field in ptformat = pos+7, block ends at offset+block_size)
"""
import struct, sys
from collections import Counter

ZMARK = 0x5A

CONTENT_NAMES = {
    0x0030: "INFO product/version", 0x1001: "WAV samplerate/size",
    0x1003: "WAV metadata", 0x1004: "WAV list full",
    0x1007: "region name/number", 0x1008: "AUDIO region name/num v5",
    0x100b: "AUDIO region list v5", 0x100f: "region->track entry",
    0x1011: "region->track map entries", 0x1012: "region->track full map",
    0x1014: "AUDIO track name/number", 0x1015: "AUDIO tracks",
    0x1017: "PLUGIN entry", 0x1018: "PLUGIN full list",
    0x1021: "I/O channel entry", 0x1022: "I/O channel list",
    0x1028: "INFO sample rate", 0x103a: "WAV names",
    0x104f: "region->track subentry v8", 0x1050: "region->track entry v8",
    0x1052: "region->track map entries v8", 0x1054: "region->track full map v8",
    0x1056: "MIDI region->track entry", 0x1057: "MIDI r->t map entries",
    0x1058: "MIDI r->t full map", 0x1062: "AUDIO region name/num v10",
    0x1063: "AUDIO region list v10", 0x1064: "COMPOUND region full map",
    0x1067: "MIDI region name/num v10", 0x1068: "MIDI regions map v10",
    0x2067: "INFO path of session",
}

def parse_block_at(data, pos, parent_end, level, out):
    if pos >= len(data) or data[pos] != ZMARK:
        return None
    if pos + 9 > len(data):
        return None
    block_type = struct.unpack_from("<H", data, pos+1)[0]
    block_size = struct.unpack_from("<I", data, pos+3)[0]
    content_type = struct.unpack_from("<H", data, pos+7)[0]
    offset = pos + 7
    if block_type & 0xff00:
        return None
    end = offset + block_size
    if end > parent_end:
        return None
    node = {"pos": pos, "block_type": block_type, "block_size": block_size,
            "content_type": content_type, "offset": offset, "children": []}
    out.append(content_type)
    # walk children
    i = 1
    childjump = 0
    while i < block_size and pos + i < end:
        child = parse_block_at(data, pos + i, end, level+1, out)
        if child:
            node["children"].append(child)
            i += child["block_size"] + 7
        else:
            i += 1
    return node

def parseblocks(data):
    blocks = []
    content_hist = []
    i = 20
    n = len(data)
    while i < n:
        b = parse_block_at(data, i, n, 0, content_hist)
        if b:
            blocks.append(b)
            i += b["block_size"] + 7 if b["block_size"] else 1
        else:
            i += 1
    return blocks, content_hist

if __name__ == "__main__":
    data = open(sys.argv[1], "rb").read()
    blocks, hist = parseblocks(data)
    print(f"top-level blocks: {len(blocks)}")
    print(f"total blocks parsed: {len(hist)}")
    print()
    print("content_type histogram (all blocks, nested):")
    for ct, n in Counter(hist).most_common():
        name = CONTENT_NAMES.get(ct, "?")
        print(f"  0x{ct:04x}  {n:6d}  {name}")
