"""
cuesheet.py — Pro Tools text export → CSV + branded PDF.

CLI:
    python3 cuesheet.py <path-to-protools-export.txt>

Importable:
    import cuesheet
    csv_path, pdf_path = cuesheet.main("/path/to/session.txt")
"""
import os
import re
import sys
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Image, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


HEADER_LABELS = [
    'SESSION NAME:',
    'SAMPLE RATE:',
    'BIT DEPTH:',
    'SESSION START TIMECODE:',
    'TIMECODE FORMAT:',
]

MERGE_GAP_FRAMES = 125              # 5 seconds at 25fps
PREFIX_THRESHOLD = 8                # min common-prefix length for stem grouping
PREFIX_THRESHOLD_AT_BOUNDARY = 15   # stricter when prefix ends at word boundary


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
    prefix = prefix.rstrip()
    m = re.search(r'[ _]\d+$', prefix)
    if m:
        prefix = prefix[:m.start()]
    return prefix.rstrip(' _')


def read_file(input_file):
    """Parse a Pro Tools text export, returning (clips, session_info)."""
    clips = []
    session_info = {}
    with open(input_file, 'r', encoding='UTF-8') as f:
        for line in f:
            if not line.strip():
                continue
            parts = line.split('\t')
            if len(parts) == 2 and parts[0] in HEADER_LABELS:
                session_info[parts[0]] = parts[1].strip()
                continue
            if len(parts) < 7:
                continue
            name     = parts[2].strip()   # raw name; cleaned later for the cue sheet
            start    = parts[3].strip()
            end      = parts[4].strip()
            duration = parts[5].strip()
            state    = parts[6].strip()
            if not re.match(r'^\d{2}:\d{2}:\d{2}:\d{2}$', start):
                continue
            if not re.match(r'^\d{2}:\d{2}:\d{2}:\d{2}$', end):
                continue
            if state != 'Unmuted':
                continue
            if start == end:
                continue
            clips.append((name, start, end, duration))
    return clips, session_info


def write_pdf(clips, output_path, session_info, logo_path):
    BRAND_DARK = colors.HexColor('#1a1a1a')
    LABEL_GREY = colors.HexColor('#666666')
    DIVIDER    = colors.HexColor('#e0e0e0')
    ROW_ALT    = colors.HexColor('#fafafa')

    doc = SimpleDocTemplate(output_path, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=18*mm,
        title="Music Cue Sheet")
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=20, textColor=BRAND_DARK, leading=22, spaceAfter=2)
    subtitle_style = ParagraphStyle('Subtitle', parent=styles['Normal'], fontName='Helvetica', fontSize=11, textColor=LABEL_GREY, leading=14)
    label_style = ParagraphStyle('Label', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7, textColor=LABEL_GREY, leading=10)
    value_style = ParagraphStyle('Value', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, textColor=BRAND_DARK, leading=10)
    cell_name_style = ParagraphStyle('CellName', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, textColor=BRAND_DARK, leading=9.5)

    title_block = Table([[Paragraph('MUSIC CUE SHEET', title_style)], [Paragraph(session_info.get('SESSION NAME:', ''), subtitle_style)]], colWidths=[145*mm])
    title_block.setStyle(TableStyle([('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0)]))

    if os.path.exists(logo_path):
        logo = Image(logo_path, width=22*mm, height=22*mm)
        header_strip = Table([[title_block, logo]], colWidths=[145*mm, 35*mm])
    else:
        header_strip = Table([[title_block]], colWidths=[180*mm])
    header_strip.setStyle(TableStyle([('VALIGN', (0,0), (-1,0), 'TOP'), ('ALIGN', (1,0), (1,0), 'RIGHT'), ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0)]))

    label_map = {'SESSION NAME:': 'SESSION NAME', 'SAMPLE RATE:': 'SAMPLE RATE', 'BIT DEPTH:': 'BIT DEPTH', 'SESSION START TIMECODE:': 'START TC', 'TIMECODE FORMAT:': 'FRAME RATE'}
    info_row = []
    for label in HEADER_LABELS:
        info_row.append([Paragraph(label_map[label], label_style), Paragraph(session_info.get(label, ''), value_style)])
    info_table = Table([info_row], colWidths=[36*mm]*5)
    info_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 6), ('TOPPADDING', (0,0), (-1,-1), 8), ('BOTTOMPADDING', (0,0), (-1,-1), 8), ('LINEABOVE', (0,0), (-1,0), 0.5, DIVIDER), ('LINEBELOW', (0,0), (-1,0), 0.5, DIVIDER)]))

    table_data = [['No.', 'Name', 'Timecode In', 'Timecode Out', 'Duration']]
    for i, clip in enumerate(clips, start=1):
        name, start, end, duration = clip
        table_data.append([str(i), Paragraph(name, cell_name_style), start, end, duration])
    col_widths = [11*mm, 95*mm, 24*mm, 24*mm, 26*mm]
    cue_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    ts = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), BRAND_DARK), ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (-1,0), 8),
        ('ALIGN', (0,0), (1,0), 'LEFT'), ('ALIGN', (2,0), (-1,0), 'RIGHT'),
        ('TOPPADDING', (0,0), (-1,0), 7), ('BOTTOMPADDING', (0,0), (-1,0), 7),
        ('LEFTPADDING', (0,0), (-1,0), 6), ('RIGHTPADDING', (0,0), (-1,0), 6),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'), ('FONTSIZE', (0,1), (-1,-1), 7.5),
        ('TEXTCOLOR', (0,1), (-1,-1), BRAND_DARK), ('VALIGN', (0,1), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,1), (-1,-1), 4), ('BOTTOMPADDING', (0,1), (-1,-1), 4),
        ('LEFTPADDING', (0,1), (-1,-1), 6), ('RIGHTPADDING', (0,1), (-1,-1), 6),
        ('ALIGN', (0,1), (0,-1), 'RIGHT'), ('ALIGN', (1,1), (1,-1), 'LEFT'),
        ('ALIGN', (2,1), (-1,-1), 'RIGHT'), ('FONTNAME', (2,1), (-1,-1), 'Courier'),
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

    doc.build([header_strip, Spacer(1, 6*mm), info_table, Spacer(1, 8*mm), cue_table], onFirstPage=draw_footer, onLaterPages=draw_footer)


XLSX_LABEL_MAP = {
    'SESSION NAME:': 'Session Name',
    'SAMPLE RATE:': 'Sample Rate',
    'BIT DEPTH:': 'Bit Depth',
    'SESSION START TIMECODE:': 'Start TC',
    'TIMECODE FORMAT:': 'Frame Rate',
}
BRAND_DARK_HEX = '1A1A1A'
LABEL_GREY_HEX = '666666'


def _fill_sheet(ws, session_info, clips):
    """Lay out one worksheet: title + session-info block + cue table."""
    ws['A1'] = 'MUSIC CUE SHEET'
    ws['A1'].font = Font(bold=True, size=16, color=BRAND_DARK_HEX)
    ws['A2'] = session_info.get('SESSION NAME:', '')
    ws['A2'].font = Font(size=11, color=LABEL_GREY_HEX)

    r = 4
    for label in HEADER_LABELS:
        if label == 'TIMECODE FORMAT:':
            continue  # frame rate omitted — confusing on delivery
        lab = ws.cell(row=r, column=1, value=XLSX_LABEL_MAP.get(label, label))
        lab.font = Font(bold=True, size=9, color=LABEL_GREY_HEX)
        val = ws.cell(row=r, column=2, value=session_info.get(label, ''))
        val.font = Font(size=10, color=BRAND_DARK_HEX)
        r += 1

    header_row = r + 1
    headers = ['No.', 'Name', 'Timecode In', 'Timecode Out', 'Duration']
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=h)
        cell.font = Font(bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill('solid', fgColor=BRAND_DARK_HEX)
        cell.alignment = Alignment(horizontal=('left' if col <= 2 else 'right'),
                                   vertical='center')

    row = header_row + 1
    for i, clip in enumerate(clips, start=1):
        name, start, end, duration = clip
        ws.cell(row=row, column=1, value=i).alignment = Alignment(horizontal='right')
        ws.cell(row=row, column=2, value=name)
        for col, val in ((3, start), (4, end), (5, duration)):
            c = ws.cell(row=row, column=col, value=val)
            c.number_format = '@'  # keep timecodes as text, not auto-parsed dates
            c.alignment = Alignment(horizontal='right')
            c.font = Font(name='Courier New', size=10)
        row += 1

    for col, width in {1: 6, 2: 62, 3: 15, 4: 15, 5: 13}.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    # Freeze the title/info/header rows so the table header stays on screen.
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)


def write_xlsx(merged_clips, raw_clips, output_path, session_info):
    """Write one workbook with two sheets: 'Cue Sheet' (merged/grouped) and
    'Raw Clips' (every parsed clip, completely unedited)."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = 'Cue Sheet'
    _fill_sheet(ws1, session_info, merged_clips)
    ws2 = wb.create_sheet('Raw Clips')
    _fill_sheet(ws2, session_info, raw_clips)
    wb.save(output_path)


def main(input_file, merge_gap_frames=None):
    """Run the full pipeline. Returns (xlsx_path, pdf_path).

    Writes a single .xlsx workbook (sheet 'Cue Sheet' matches the PDF; sheet
    'All Clips' is the raw deduped list) plus the branded PDF.

    merge_gap_frames: gap (in frames @25fps) used both for merging adjacent
    same-name clips and for stem grouping. Defaults to MERGE_GAP_FRAMES (125
    = 5s) when not supplied by the caller (e.g. the merge-gap slider).
    """
    gap_frames = MERGE_GAP_FRAMES if merge_gap_frames is None else merge_gap_frames
    input_dir = os.path.dirname(input_file)
    input_name = os.path.basename(input_file)
    base_name = os.path.splitext(input_name)[0]
    xlsx_output = os.path.join(input_dir, f"{base_name}_cuesheet.xlsx")

    clips, session_info = read_file(input_file)
    print(f"Parsed {len(clips)} clips")

    # Normalise names (strip version suffixes like .01) so stereo/copy pairs
    # sitting at the same position collapse to one, then sort by start.
    clips = [(clean_name(name), start, end, dur)
             for (name, start, end, dur) in clips]
    clips.sort(key=lambda c: c[1])

    # Dedupe identical (name, start, end) — collapses stereo L/R pairs
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

    # Raw Clips sheet = every DISTINCT clip (deduped, so stereo/copy pairs at
    # the same position don't show twice), unmerged and ungrouped.
    raw_clips = list(unique_clips)

    # Consolidate same-name clips that overlap or sit within the gap — even when
    # OTHER clips are interleaved between them in time. A music cue is usually
    # chopped into segments (edits/crossfades) with FX or other clips landing in
    # the gaps; tracking the open cue per NAME (not just the previous row) keeps
    # the music cue as a single entry instead of splitting it across rows.
    open_by_name = {}   # name -> index of its currently-open cue in merged_clips
    merged_clips = []
    for clip in clips:  # sorted by start
        name, start, end, duration = clip
        idx = open_by_name.get(name)
        if idx is not None:
            _, p_start, p_end, _ = merged_clips[idx]
            if tc_to_frames(start) - tc_to_frames(p_end) <= gap_frames:
                new_end = end if tc_to_frames(end) > tc_to_frames(p_end) else p_end
                new_duration = frames_to_tc(tc_to_frames(new_end) - tc_to_frames(p_start))
                merged_clips[idx] = (name, p_start, new_end, new_duration)
                continue
        merged_clips.append((name, start, end, duration))
        open_by_name[name] = len(merged_clips) - 1
    clips = merged_clips

    # Group stems with common prefix that overlap or are within gap
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
            if len(prefix) >= min_prefix and overlaps_or_within(start, group_end, next_start, next_end, gap_frames):
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

    # One workbook, two sheets: 'Cue Sheet' (merged/grouped) and 'Raw Clips'
    # (every clip, completely unedited). A real .xlsx opens cleanly in Excel
    # and Google Sheets.
    write_xlsx(clips, raw_clips, xlsx_output, session_info)
    print(f"Wrote {xlsx_output} ({len(clips)} cues / {len(raw_clips)} raw clips)")

    return xlsx_output


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 cuesheet.py <path-to-protools-export.txt>")
        sys.exit(1)
    main(sys.argv[1])
