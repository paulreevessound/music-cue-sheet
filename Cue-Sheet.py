import csv
import os
import re
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Image, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors

# ── Paths ────────────────────────────────────────────────────────
input_file = "/Users/paulreeves/Documents/APP stuff/MX Cue sheet generator/MS0811 GABO S9E01 Music Cue Sheet Domestic.txt"

input_dir = os.path.dirname(input_file)
input_name = os.path.basename(input_file)
base_name = os.path.splitext(input_name)[0]
output_file = os.path.join(input_dir, f"{base_name}_cuesheet.csv")
pdf_output = os.path.join(input_dir, f"{base_name}_cuesheet.pdf")
logo_path = os.path.join(input_dir, 'logo.png')

# ── Constants ────────────────────────────────────────────────────
HEADER_LABELS = [
    'SESSION NAME:',
    'SAMPLE RATE:',
    'BIT DEPTH:',
    'SESSION START TIMECODE:',
    'TIMECODE FORMAT:',
]

MERGE_GAP_FRAMES = 125                 # 5 seconds at 25fps
PREFIX_THRESHOLD = 8                   # base prefix length for stem grouping
PREFIX_THRESHOLD_AT_BOUNDARY = 15      # stricter when prefix ends at '_' or space

clips = []
session_info = {}


# ── Helpers ──────────────────────────────────────────────────────
def clean_name(name):
    return re.sub(r'\.\d{2}(-\d{2})?$', '', name)


def tc_to_frames(tc):
    h, m, s, f = tc.split(':')
    return int(h) * 90000 + int(m) * 1500 + int(s) * 25 + int(f)


def frames_to_tc(frames):
    h = frames // 90000
    frames %= 90000
    m = frames // 1500
    frames %= 1500
    s = frames // 25
    f = frames % 25
    return f"{h:02d}:{m:02d}:{s:02d}:{f:02d}"


def common_prefix(a, b):
    i = 0
    while i < len(a) and i < len(b) and a[i] == b[i]:
        i += 1
    return a[:i]


def overlaps_or_within(start_a, end_a, start_b, end_b, gap_frames):
    a_start = tc_to_frames(start_a)
    a_end   = tc_to_frames(end_a)
    b_start = tc_to_frames(start_b)
    b_end   = tc_to_frames(end_b)
    return a_start - gap_frames < b_end and b_start - gap_frames < a_end


def cleanup_display_name(prefix):
    """Tidy a common-prefix display name: strip trailing partial stem
    numbers like ' 0' or '_0', and trailing separators."""
    prefix = prefix.rstrip()
    m = re.search(r'[ _]\d+$', prefix)
    if m:
        prefix = prefix[:m.start()]
    return prefix.rstrip(' _')


# ── Parse Pro Tools .txt ─────────────────────────────────────────
def read_file():
    with open(input_file, 'r', encoding='UTF-8') as f:
        for line in f:
            if not line.strip():
                continue
            parts = line.split('\t')

            # Capture session info header lines
            if len(parts) == 2 and parts[0] in HEADER_LABELS:
                session_info[parts[0]] = parts[1].strip()
                continue

            # Skip anything that's not a 7-field clip line
            if len(parts) < 7:
                continue

            name     = clean_name(parts[2].strip())
            start    = parts[3].strip()
            end      = parts[4].strip()
            duration = parts[5].strip()
            state    = parts[6].strip()

            # Validate timecodes look like HH:MM:SS:FF
            if not re.match(r'^\d{2}:\d{2}:\d{2}:\d{2}$', start):
                continue
            if not re.match(r'^\d{2}:\d{2}:\d{2}:\d{2}$', end):
                continue

            if state != 'Unmuted':
                continue
            if start == end:
                continue

            clips.append((name, start, end, duration))
    print(f"Parsed {len(clips)} clips")


# ── Build PDF ────────────────────────────────────────────────────
def write_pdf(clips, output_path, session_info, logo_path):
    BRAND_DARK = colors.HexColor('#1a1a1a')
    LABEL_GREY = colors.HexColor('#666666')
    DIVIDER    = colors.HexColor('#e0e0e0')
    ROW_ALT    = colors.HexColor('#fafafa')

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=15*mm, bottomMargin=18*mm,
        title="Music Cue Sheet",
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('Title', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=20, textColor=BRAND_DARK,
        leading=22, spaceAfter=2)
    subtitle_style = ParagraphStyle('Subtitle', parent=styles['Normal'],
        fontName='Helvetica', fontSize=11, textColor=LABEL_GREY, leading=14)
    label_style = ParagraphStyle('Label', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=7, textColor=LABEL_GREY, leading=10)
    value_style = ParagraphStyle('Value', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8.5, textColor=BRAND_DARK, leading=10)
    cell_name_style = ParagraphStyle('CellName', parent=styles['Normal'],
        fontName='Helvetica', fontSize=7.5, textColor=BRAND_DARK, leading=9.5)

    # Header strip: title block + logo
    title_block = Table(
        [[Paragraph('MUSIC CUE SHEET', title_style)],
         [Paragraph(session_info.get('SESSION NAME:', ''), subtitle_style)]],
        colWidths=[145*mm],
    )
    title_block.setStyle(TableStyle([
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))

    logo = Image(logo_path, width=22*mm, height=22*mm)

    header_strip = Table([[title_block, logo]], colWidths=[145*mm, 35*mm])
    header_strip.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,0), 'TOP'),
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))

    # Session info: 5-column metadata band
    label_map = {
        'SESSION NAME:': 'SESSION NAME',
        'SAMPLE RATE:': 'SAMPLE RATE',
        'BIT DEPTH:': 'BIT DEPTH',
        'SESSION START TIMECODE:': 'START TC',
        'TIMECODE FORMAT:': 'FRAME RATE',
    }
    info_row = []
    for label in HEADER_LABELS:
        cell = [
            Paragraph(label_map[label], label_style),
            Paragraph(session_info.get(label, ''), value_style),
        ]
        info_row.append(cell)

    info_table = Table([info_row], colWidths=[36*mm]*5)
    info_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LINEABOVE', (0,0), (-1,0), 0.5, DIVIDER),
        ('LINEBELOW', (0,0), (-1,0), 0.5, DIVIDER),
    ]))

    # Cue table
    table_data = [['No.', 'Name', 'Timecode In', 'Timecode Out', 'Duration']]
    for i, clip in enumerate(clips, start=1):
        name, start, end, duration = clip
        table_data.append([str(i), Paragraph(name, cell_name_style), start, end, duration])

    col_widths = [11*mm, 95*mm, 24*mm, 24*mm, 26*mm]
    cue_table = Table(table_data, colWidths=col_widths, repeatRows=1)

    ts = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), BRAND_DARK),
        ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,0), 8),
        ('ALIGN',      (0,0), (1,0),  'LEFT'),
        ('ALIGN',      (2,0), (-1,0), 'RIGHT'),
        ('TOPPADDING', (0,0), (-1,0), 7),
        ('BOTTOMPADDING', (0,0), (-1,0), 7),
        ('LEFTPADDING',  (0,0), (-1,0), 6),
        ('RIGHTPADDING', (0,0), (-1,0), 6),

        ('FONTNAME',  (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE',  (0,1), (-1,-1), 7.5),
        ('TEXTCOLOR', (0,1), (-1,-1), BRAND_DARK),
        ('VALIGN',    (0,1), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',    (0,1), (-1,-1), 4),
        ('BOTTOMPADDING', (0,1), (-1,-1), 4),
        ('LEFTPADDING',   (0,1), (-1,-1), 6),
        ('RIGHTPADDING',  (0,1), (-1,-1), 6),

        ('ALIGN', (0,1), (0,-1),  'RIGHT'),
        ('ALIGN', (1,1), (1,-1),  'LEFT'),
        ('ALIGN', (2,1), (-1,-1), 'RIGHT'),

        ('FONTNAME', (2,1), (-1,-1), 'Courier'),

        ('LINEBELOW', (0,0), (-1,-1), 0.25, DIVIDER),
    ])

    for row in range(1, len(table_data)):
        if row % 2 == 0:
            ts.add('BACKGROUND', (0, row), (-1, row), ROW_ALT)
    cue_table.setStyle(ts)

    def draw_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 7.5)
        canvas.setFillColor(LABEL_GREY)
        canvas.drawString(15*mm, 10*mm, 'Mighty Sound  ·  mightysound.studio')
        canvas.drawRightString(A4[0] - 15*mm, 10*mm, f'Page {canvas.getPageNumber()}')
        canvas.restoreState()

    doc.build(
        [header_strip, Spacer(1, 6*mm), info_table, Spacer(1, 8*mm), cue_table],
        onFirstPage=draw_footer,
        onLaterPages=draw_footer,
    )


# ── Run pipeline ─────────────────────────────────────────────────
read_file()

# Sort by start TC
clips.sort(key=lambda c: c[1])

# Dedupe identical (name, start, end)
seen = set()
unique_clips = []
for clip in clips:
    name, start, end, duration = clip
    key = (name, start, end)
    if key in seen:
        continue
    seen.add(key)
    unique_clips.append(clip)
clips = unique_clips

# Merge adjacent same-name clips within 5s
merged_clips = []
for clip in clips:
    name, start, end, duration = clip
    if merged_clips:
        prev_name, prev_start, prev_end, _ = merged_clips[-1]
        if name == prev_name and tc_to_frames(start) - tc_to_frames(prev_end) <= MERGE_GAP_FRAMES:
            new_duration = frames_to_tc(tc_to_frames(end) - tc_to_frames(prev_start))
            merged_clips[-1] = (name, prev_start, end, new_duration)
            continue
    merged_clips.append(clip)
clips = merged_clips

# Group stems (overlap-or-within + shared prefix, with stricter threshold at word boundaries)
grouped_clips = []
i = 0
while i < len(clips):
    name, start, end, duration = clips[i]
    group_end = end
    group_prefix = name
    j = i + 1
    while j < len(clips):
        next_name, next_start, next_end, _ = clips[j]
        prefix = common_prefix(group_prefix, next_name)
        min_prefix = PREFIX_THRESHOLD_AT_BOUNDARY if (prefix.endswith('_') or prefix.endswith(' ')) else PREFIX_THRESHOLD
        if len(prefix) >= min_prefix and overlaps_or_within(start, group_end, next_start, next_end, MERGE_GAP_FRAMES):
            group_prefix = prefix
            if tc_to_frames(next_end) > tc_to_frames(group_end):
                group_end = next_end
            j += 1
        else:
            break
    merged_duration = frames_to_tc(tc_to_frames(group_end) - tc_to_frames(start))
    grouped_clips.append((cleanup_display_name(group_prefix), start, group_end, merged_duration))
    i = j
clips = grouped_clips

print(f"Final cue count: {len(clips)}")

# Write CSV
with open(output_file, 'w', encoding='UTF-8', newline='') as f:
    writer = csv.writer(f)
    for label in HEADER_LABELS:
        writer.writerow([label, session_info.get(label, '')])
    writer.writerow([])
    writer.writerow(['No.', 'Name', 'Timecode In', 'Timecode Out', 'Duration'])
    for i, clip in enumerate(clips, start=1):
        name, start, end, duration = clip
        writer.writerow([i, name, start, end, duration])
print(f"Wrote {output_file}")

# Write PDF
write_pdf(clips, pdf_output, session_info, logo_path)
print(f"Wrote {pdf_output}")