"""
Generate DOCX and PDF versions of the ProMix DJ Studio User Manual.
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ─────────────────────────────────────────────
# COLOR PALETTE
# ─────────────────────────────────────────────
C_PURPLE    = RGBColor(0x7C, 0x3A, 0xED)
C_INDIGO    = RGBColor(0x4F, 0x46, 0xE5)
C_LAVENDER  = RGBColor(0xA7, 0x8B, 0xFA)
C_BLUE      = RGBColor(0x60, 0xA5, 0xFA)
C_GREEN     = RGBColor(0x34, 0xD3, 0x99)
C_YELLOW    = RGBColor(0xFB, 0xBF, 0x24)
C_RED       = RGBColor(0xF8, 0x71, 0x71)
C_DARK      = RGBColor(0x1E, 0x1B, 0x4B)
C_MID       = RGBColor(0x4B, 0x55, 0x63)
C_LIGHT     = RGBColor(0x6B, 0x72, 0x80)
C_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
C_BG_PURPLE = RGBColor(0xED, 0xE9, 0xFE)
C_BG_BLUE   = RGBColor(0xDB, 0xEA, 0xFE)
C_BG_GREEN  = RGBColor(0xD1, 0xFA, 0xE5)
C_BG_YELLOW = RGBColor(0xFE, 0xF3, 0xC7)
C_SIDEBAR   = RGBColor(0x1A, 0x1A, 0x2E)
C_BORDER    = RGBColor(0xE5, 0xE7, 0xEB)
C_THEAD     = RGBColor(0x1A, 0x1A, 0x2E)
C_ROW_ALT   = RGBColor(0xF3, 0xF0, 0xFF)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def set_cell_bg(cell, rgb: RGBColor):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  f'{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}')
    tcPr.append(shd)

def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'),   val.get('val',   'single'))
            el.set(qn('w:sz'),    val.get('sz',    '4'))
            el.set(qn('w:space'), val.get('space', '0'))
            el.set(qn('w:color'), val.get('color', 'auto'))
            tcBorders.append(el)
    tcPr.append(tcBorders)

def add_run(para, text, bold=False, italic=False, color=None, size=None, font_name=None):
    run = para.add_run(text)
    run.bold   = bold
    run.italic = italic
    if color:      run.font.color.rgb = color
    if size:       run.font.size      = Pt(size)
    if font_name:  run.font.name      = font_name
    return run

def add_heading(doc, text, level=1, color=C_DARK):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.color.rgb = color
    return p

def add_para(doc, text='', bold=False, italic=False, color=None, size=None, alignment=None, space_before=None, space_after=None):
    p = doc.add_paragraph()
    if alignment: p.alignment = alignment
    if space_before is not None: p.paragraph_format.space_before = Pt(space_before)
    if space_after  is not None: p.paragraph_format.space_after  = Pt(space_after)
    if text:
        add_run(p, text, bold=bold, italic=italic, color=color, size=size)
    return p

def add_callout(doc, icon, label, body_text, bg_color, border_color):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    # left icon cell
    c0 = tbl.cell(0, 0)
    c0.width = Cm(1.0)
    set_cell_bg(c0, bg_color)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run(icon)
    r0.font.size = Pt(14)
    # right text cell
    c1 = tbl.cell(0, 1)
    set_cell_bg(c1, bg_color)
    p1 = c1.paragraphs[0]
    add_run(p1, label + '  ', bold=True, color=C_DARK, size=10)
    add_run(p1, body_text, color=C_MID, size=10)
    # border left
    set_cell_border(c0, left={'val':'single','sz':'12','color':f'{border_color[0]:02X}{border_color[1]:02X}{border_color[2]:02X}'})
    doc.add_paragraph()

def add_step(doc, num, title, body):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    c0 = tbl.cell(0, 0)
    c0.width = Cm(1.2)
    set_cell_bg(c0, C_PURPLE)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run(str(num))
    r0.bold = True
    r0.font.color.rgb = C_WHITE
    r0.font.size = Pt(11)
    c1 = tbl.cell(0, 1)
    p1 = c1.paragraphs[0]
    add_run(p1, title + '\n', bold=True, color=C_DARK, size=11)
    add_run(p1, body, color=C_MID, size=10)
    doc.add_paragraph()

def add_feature_table(doc, headers, rows, col_widths=None):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    # Header row
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        set_cell_bg(cell, C_THEAD)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.color.rgb = C_WHITE
        r.font.size = Pt(9)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # Data rows
    for ri, row in enumerate(rows):
        bg = C_ROW_ALT if ri % 2 == 0 else C_WHITE
        for ci, cell_text in enumerate(row):
            cell = tbl.cell(ri + 1, ci)
            set_cell_bg(cell, bg)
            p = cell.paragraphs[0]
            r = p.add_run(str(cell_text))
            r.font.size = Pt(9)
            if ci == 0:
                r.bold = True
                r.font.color.rgb = C_DARK
            else:
                r.font.color.rgb = C_MID
    if col_widths:
        for ri in range(len(tbl.rows)):
            for ci, w in enumerate(col_widths):
                tbl.cell(ri, ci).width = Cm(w)
    doc.add_paragraph()

def add_kb_table(doc, rows):
    tbl = doc.add_table(rows=1 + len(rows), cols=2)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, h in enumerate(['Key', 'Action']):
        cell = tbl.cell(0, i)
        set_cell_bg(cell, C_THEAD)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.bold = True; r.font.color.rgb = C_WHITE; r.font.size = Pt(9)
    for ri, (key, action) in enumerate(rows):
        bg = C_BG_PURPLE if ri % 2 == 0 else C_WHITE
        c0 = tbl.cell(ri + 1, 0)
        set_cell_bg(c0, bg)
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(key)
        r0.bold = True; r0.font.color.rgb = C_PURPLE
        r0.font.name = 'Courier New'; r0.font.size = Pt(10)
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c0.width = Cm(3)
        c1 = tbl.cell(ri + 1, 1)
        set_cell_bg(c1, bg)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(action)
        r1.font.size = Pt(9); r1.font.color.rgb = C_MID
    doc.add_paragraph()

def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Inches(0.3 * (level + 1))
    add_run(p, text, color=C_MID, size=10)
    return p

def page_break(doc):
    doc.add_page_break()

# ─────────────────────────────────────────────
# DOCUMENT
# ─────────────────────────────────────────────
doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# Default font
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)
style.font.color.rgb = C_MID

# Heading styles
for lvl, sz, col in [(1, 20, C_PURPLE), (2, 15, C_DARK), (3, 12, C_INDIGO)]:
    h_style = doc.styles[f'Heading {lvl}']
    h_style.font.name  = 'Calibri'
    h_style.font.size  = Pt(sz)
    h_style.font.color.rgb = col
    h_style.font.bold  = True
    h_style.paragraph_format.space_before = Pt(14)
    h_style.paragraph_format.space_after  = Pt(6)

# ══════════════════════════════════════════════
# COVER PAGE
# ══════════════════════════════════════════════
cover_title = doc.add_paragraph()
cover_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
cover_title.paragraph_format.space_before = Pt(60)
r = cover_title.add_run('ProMix DJ Studio')
r.bold = True; r.font.size = Pt(36); r.font.color.rgb = C_PURPLE; r.font.name = 'Calibri'

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = sub.add_run('The Ultimate Browser-Based Party Mixer')
r2.font.size = Pt(16); r2.font.color.rgb = C_MID; r2.italic = True

doc.add_paragraph()
ver = doc.add_paragraph()
ver.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(ver, 'USER MANUAL  ·  Version 1.0', bold=True, color=C_LAVENDER, size=12)

doc.add_paragraph()
line = doc.add_paragraph()
line.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(line, '─' * 55, color=C_LAVENDER, size=11)

doc.add_paragraph()
tagline = doc.add_paragraph()
tagline.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(tagline, 'Professional DJ Tools  ·  Zero Installation  ·  Open Any Browser',
        italic=True, color=C_LIGHT, size=11)

doc.add_paragraph()
doc.add_paragraph()
features_preview = doc.add_paragraph()
features_preview.alignment = WD_ALIGN_PARAGRAPH.CENTER
for feat in ['Dual Decks', 'BPM & Key Detection', 'Camelot Wheel',
             'YouTube Search', 'MP3 Download', 'Mix Recording',
             '3-Band EQ', '5 Effects', 'Auto Mix', 'Stem Separation']:
    add_run(features_preview, f'  ●  {feat}', color=C_LIGHT, size=10)

page_break(doc)

# ══════════════════════════════════════════════
# TABLE OF CONTENTS
# ══════════════════════════════════════════════
add_heading(doc, 'Table of Contents', level=1)
toc_items = [
    ('1.', 'Overview & Feature List'),
    ('2.', 'System Requirements'),
    ('3.', 'Quick Start Guide'),
    ('4.', 'Interface Layout'),
    ('5.', 'Loading Tracks'),
    ('6.', 'Playback Controls, Hot Cues & Loops'),
    ('7.', 'BPM & Key Detection'),
    ('8.', 'Camelot Wheel — Harmonic Mixing'),
    ('9.', 'EQ & Mixing'),
    ('10.', 'Crossfader'),
    ('11.', 'Effects'),
    ('12.', 'YouTube Search & MP3 Download'),
    ('13.', 'Recording & Exporting Your Mix'),
    ('14.', 'Auto Mix'),
    ('15.', 'Stem Separation'),
    ('16.', 'Sync & Pitch'),
    ('17.', 'Keyboard Shortcuts'),
    ('18.', 'Tutorial 1 — Your First Mix'),
    ('19.', 'Tutorial 2 — DJ Party Set with Recording'),
    ('20.', 'Tutorial 3 — YouTube to Mix'),
    ('21.', 'Tutorial 4 — Harmonic Mixing'),
    ('22.', 'FAQ'),
    ('23.', 'Troubleshooting'),
]
for num, title in toc_items:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    add_run(p, f'{num:<5}', bold=True, color=C_PURPLE, size=10, font_name='Courier New')
    add_run(p, title, color=C_MID, size=10)

page_break(doc)

# ══════════════════════════════════════════════
# 1. OVERVIEW
# ══════════════════════════════════════════════
add_heading(doc, '1. Overview', level=1)
add_para(doc,
    'ProMix DJ Studio is a fully self-contained DJ mixing application delivered as a single HTML file. '
    'Open it in any modern web browser — no software installation, no app store, no subscription. '
    'Everything runs locally on your machine using the Web Audio API.',
    color=C_MID, size=10.5)
add_para(doc,
    'It was designed by analyzing and improving upon the world\'s best DJ software: DJ.Studio, '
    'Algoriddim djay Pro, VirtualDJ, and Fadr — combining the best features of all four into one '
    'free, portable tool.',
    color=C_MID, size=10.5)

add_heading(doc, 'Feature Summary', level=2)
add_feature_table(doc,
    ['Feature', 'Description'],
    [
        ['Dual Decks',         'Two independent playback decks (A and B) with animated vinyl turntables'],
        ['Waveform Display',   'Real-time audio waveform with playhead position and click-to-seek'],
        ['BPM Detection',      'Automatic beats-per-minute analysis on every loaded track'],
        ['Key Detection',      'Musical key detection mapped to Camelot Wheel notation'],
        ['Camelot Wheel',      'Interactive SVG harmonic mixing guide for compatible key transitions'],
        ['3-Band EQ',          'Hi / Mid / Lo equalizer per deck with interactive rotary knobs'],
        ['Crossfader',         'Linear, Scratch, and Smooth curve modes'],
        ['5 Effects',          'Reverb, Delay, Filter, Flanger, Echo — all with parameter knobs'],
        ['Hot Cues',           '8 hot cue points per deck — set and jump with one click'],
        ['Loop Control',       '4-beat auto-loop with one button'],
        ['Pitch / Sync',       '±8% pitch slider, key shift (±semitones), one-click BPM sync'],
        ['Tap Tempo',          'Manual BPM entry by tapping the beat'],
        ['Auto Mix',           'One-click automated BPM-synced crossfader transition'],
        ['YouTube Search',     'Search by song or artist, load results directly to a deck'],
        ['MP3 Download',       'Download any YouTube track as MP3 via cobalt.tools'],
        ['Mix Recording',      'Record the master output and download as an audio file'],
        ['Timeline',           'Full-length waveform overview with BPM beat grid'],
        ['Stems Panel',        'Vocal / Drum / Bass / Melody controls + AI stem separation'],
        ['VU Meters',          'Color-coded level meters for both decks and master output'],
    ],
    col_widths=[5, 11]
)

# ══════════════════════════════════════════════
# 2. SYSTEM REQUIREMENTS
# ══════════════════════════════════════════════
add_heading(doc, '2. System Requirements', level=1)
add_feature_table(doc,
    ['Item', 'Requirement'],
    [
        ['Browser',       'Google Chrome 90+, Firefox 88+, Microsoft Edge 90+, Safari 15+'],
        ['Operating System', 'Windows 10/11, macOS 11+, Linux (any modern distro)'],
        ['RAM',           '4 GB minimum, 8 GB recommended for large audio files'],
        ['Audio Files',   'MP3, WAV, FLAC, OGG, AAC, M4A (any format your browser supports)'],
        ['Internet',      'Required only for YouTube search and MP3 download'],
        ['YouTube API Key', 'Optional — enables in-app search (free to obtain)'],
    ],
    col_widths=[5, 11]
)
add_callout(doc, '💡', 'Best Experience:',
    'Use Google Chrome for the best Web Audio API performance and broadest audio format support.',
    C_BG_PURPLE, C_PURPLE)

# ══════════════════════════════════════════════
# 3. QUICK START
# ══════════════════════════════════════════════
add_heading(doc, '3. Quick Start Guide', level=1)
add_para(doc, 'Get mixing in under 2 minutes:', color=C_MID, size=10.5)
steps_qs = [
    ('Open the file',      'Download dj-mixer.html from GitHub. Double-click it — it opens in your default browser. No installation needed.'),
    ('Load Track A',       'Click the 📂 Load button on Deck A (left side). Select any MP3 or audio file. The waveform appears and BPM/key are detected automatically.'),
    ('Load Track B',       'Click the 📂 Load button on Deck B (right side). Load a second track.'),
    ('Play Deck A',        'Press the large green ▶ Play button on Deck A, or press Space on your keyboard.'),
    ('Sync and mix',       'Press SYNC on Deck B to match its tempo. Press Play on Deck B. Slide the crossfader left-to-right to blend between tracks.'),
    ('Record your mix',    'Click ● REC in the top bar before you start mixing. Click ⬇ Download Mix to save the file locally when done.'),
]
for i, (title, body) in enumerate(steps_qs, 1):
    add_step(doc, i, title, body)
add_callout(doc, '✅', 'You\'re mixing!',
    'That\'s all it takes for a basic mix. Read on to explore all the professional features.',
    C_BG_GREEN, C_GREEN)

page_break(doc)

# ══════════════════════════════════════════════
# 4. INTERFACE LAYOUT
# ══════════════════════════════════════════════
add_heading(doc, '4. Interface Layout', level=1)
add_para(doc,
    'The ProMix DJ Studio interface is organized in a professional DJ layout with all controls accessible at a glance.',
    color=C_MID, size=10.5)

# ASCII diagram
p_diag = doc.add_paragraph()
p_diag.paragraph_format.space_before = Pt(6)
p_diag.paragraph_format.space_after  = Pt(6)
diag_text = (
    "┌──────────────────────────────────────────────────────────────┐\n"
    "│  ProMix DJ Studio  │ [API Key] │ [●REC] [⬇Mix] [Camelot] [⚡]│\n"
    "├──────────────────────────────────────────────────────────────┤\n"
    "│  🔍 [Search bar....................]  [SEARCH]               │\n"
    "│  [Search Results / YouTube Results Panel]                    │\n"
    "├─────────────────┬─────────────────────┬──────────────────────┤\n"
    "│  DECK A         │   MIXER CENTER      │  DECK B              │\n"
    "│  ┌───────────┐  │  [EQ A]  [EQ B]     │  ┌───────────┐      │\n"
    "│  │ ◉ Vinyl  │  │  Hi  Mid  Lo        │  │ ◉ Vinyl  │      │\n"
    "│  └───────────┘  │  [VU A][VU M][VU B] │  └───────────┘      │\n"
    "│  ~~~Waveform~~  │  [VolA][Mst][VolB]  │  ~~~Waveform~~      │\n"
    "│  BPM:128  8A    │  ──── Crossfader ───│  BPM:130  8B        │\n"
    "│  [▶][CUE][SYNC] │  [Auto Mix]         │  [▶][CUE][SYNC]     │\n"
    "│  [1][2]..[8]    │                     │  [1][2]..[8]        │\n"
    "├─────────────────┴─────────────────────┴──────────────────────┤\n"
    "│  FX: [REVERB ⚬⚬][DELAY ⚬⚬][FILTER ⚬⚬][FLANGER ⚬⚬][ECHO ⚬] │\n"
    "├──────────────────────────────────────────────────────────────┤\n"
    "│  STEMS: [🎤VOCALS][🥁DRUMS][🎸BASS][🎹MELODY] [⚗ Separate]   │\n"
    "├──────────────────────────────────────────────────────────────┤\n"
    "│  TIMELINE: [Deck A ████████░░░░] [Deck B ░░████████░░]       │\n"
    "└──────────────────────────────────────────────────────────────┘"
)
r_diag = p_diag.add_run(diag_text)
r_diag.font.name = 'Courier New'
r_diag.font.size = Pt(8)
r_diag.font.color.rgb = C_DARK

add_heading(doc, 'Header Bar', level=2)
add_feature_table(doc,
    ['Control', 'Function'],
    [
        ['API Key field',    'Paste your YouTube Data API v3 key here to enable YouTube search'],
        ['● REC button',     'Starts/stops recording the master audio output. Pulses red when active.'],
        ['⬇ Download Mix',   'Downloads your recorded mix as a WebM audio file'],
        ['⟳ Camelot Wheel',  'Opens the interactive harmonic mixing guide modal'],
        ['⚡ Auto Mix',       'Automatically syncs BPM and performs a crossfader transition'],
    ],
    col_widths=[5, 11]
)

add_heading(doc, 'Deck Controls Reference', level=2)
add_feature_table(doc,
    ['Control', 'Function'],
    [
        ['Track Title',     'Shows the loaded file name'],
        ['📂 Load button',  'Opens file browser to load an audio file'],
        ['Vinyl Canvas',    'Animated spinning record — spins when playing'],
        ['Waveform',        'Visual audio representation. Click to jump to that position.'],
        ['BPM display',     'Auto-detected beats per minute'],
        ['KEY display',     'Detected musical key in Camelot notation (e.g., 8A, 8B)'],
        ['TAP button',      'Tap repeatedly on the beat to set BPM manually'],
        ['± Key shift',     'Shift pitch up/down by semitones'],
        ['⏱ Timer',         'Current playback position (minutes:seconds)'],
        ['PITCH slider',    'Adjusts playback speed ±8%'],
        ['▶ Play/Pause',    'Start or pause playback'],
        ['CUE button',      'If paused: sets cue point. If playing: jumps to cue point.'],
        ['SYNC button',     'Matches this deck\'s BPM to the other deck'],
        ['⟳ LOOP button',   'Activates/deactivates a 4-beat loop from current position'],
        ['Hot Cues 1–8',    'Colored buttons for setting and jumping to saved positions'],
    ],
    col_widths=[5, 11]
)

page_break(doc)

# ══════════════════════════════════════════════
# 5. LOADING TRACKS
# ══════════════════════════════════════════════
add_heading(doc, '5. Loading Tracks', level=1)
add_heading(doc, 'Method 1: Local File Upload', level=2)
steps_load = [
    ('Click 📂 Load on your chosen deck', 'The file browser opens. Navigate to your music folder.'),
    ('Select your audio file',            'Supported: MP3, WAV, FLAC, OGG, AAC, M4A, and any format your browser supports.'),
    ('Wait for analysis',                 'The app decodes the audio, draws the waveform, and detects BPM and key. Takes 1–3 seconds.'),
    ('Check the readout',                 'Verify the BPM and KEY values appear. If BPM looks wrong, use TAP tempo to correct it.'),
]
for i, (t, b) in enumerate(steps_load, 1): add_step(doc, i, t, b)
add_callout(doc, '⚠️', 'Large files:',
    'Files over 100 MB may take several seconds to decode. MP3s at 320kbps load fastest.',
    C_BG_YELLOW, C_YELLOW)

# ══════════════════════════════════════════════
# 6. PLAYBACK CONTROLS
# ══════════════════════════════════════════════
add_heading(doc, '6. Playback Controls', level=1)
add_feature_table(doc,
    ['Action', 'How To'],
    [
        ['Play / Pause',      'Click the ▶ button or press Space (Deck A) / Enter (Deck B)'],
        ['Seek forward 5 sec', 'Press → (Deck A)'],
        ['Seek back 5 sec',   'Press ← (Deck A)'],
        ['Jump to position',  'Click anywhere on the waveform or timeline'],
        ['Set Cue Point',     'Pause at desired position, then click CUE'],
        ['Jump to Cue Point', 'Click CUE while playing — snaps back to cue point'],
    ],
    col_widths=[6, 10]
)

add_heading(doc, 'Hot Cues', level=2)
add_para(doc, 'Hot cues let you mark up to 8 specific points in a track and jump to them instantly — perfect for verse, chorus, and drop markers.', color=C_MID, size=10)
steps_hc = [
    ('Play or navigate to the spot you want to mark', 'For example, the start of the chorus.'),
    ('Click an empty Hot Cue button (1–8)',           'The button changes color. The time position is saved.'),
    ('Click the colored button at any time to jump',  'Instantly jumps to that saved position. If playing, continues from there.'),
]
for i, (t, b) in enumerate(steps_hc, 1): add_step(doc, i, t, b)
add_callout(doc, '💡', 'Pro Tip:',
    'Set Hot Cue 1 at the intro, Hot Cue 2 at the first drop, Hot Cue 3 at the breakdown. Jump between sections instantly during a live set.',
    C_BG_PURPLE, C_PURPLE)

add_heading(doc, 'Loop Controls', level=2)
add_feature_table(doc,
    ['Action', 'How To'],
    [
        ['Start loop',   'Click LOOP during playback — loop activates from current beat'],
        ['Exit loop',    'Click LOOP again — returns to normal playback'],
        ['Loop length',  'Auto-calculated from the detected BPM (4 beats)'],
    ],
    col_widths=[5, 11]
)

page_break(doc)

# ══════════════════════════════════════════════
# 7. BPM & KEY DETECTION
# ══════════════════════════════════════════════
add_heading(doc, '7. BPM & Key Detection', level=1)
add_heading(doc, 'Automatic BPM Detection', level=2)
add_para(doc,
    'When you load a track, ProMix analyzes the audio waveform using onset energy detection — it finds sudden increases '
    'in loudness (beats) and calculates the average interval between them to determine the tempo. BPM is always clamped '
    'to a musical range (70–200 BPM), so double-time and half-time tracks are handled correctly.',
    color=C_MID, size=10)

add_heading(doc, 'Correcting BPM with Tap Tempo', level=2)
steps_tap = [
    ('Play the track',         'Listen to the beat.'),
    ('Click TAP on the beat',  'Click at least 4 times. The BPM updates using the average interval.'),
    ('Tap after 2+ seconds to reset', 'A 2-second gap clears previous taps and starts fresh.'),
]
for i, (t, b) in enumerate(steps_tap, 1): add_step(doc, i, t, b)

add_heading(doc, 'Automatic Key Detection', level=2)
add_para(doc,
    'ProMix uses a pitch class profile analysis (Krumhansl–Schmuckler model) to detect the musical key. '
    'The detected key is displayed in Camelot Wheel notation (e.g., 8A = A minor, 8B = C major).',
    color=C_MID, size=10)

# ══════════════════════════════════════════════
# 8. CAMELOT WHEEL
# ══════════════════════════════════════════════
add_heading(doc, '8. Camelot Wheel — Harmonic Mixing', level=1)
add_para(doc,
    'Click the ⟳ Camelot Wheel button in the header to open the interactive harmonic mixing guide. '
    'This wheel shows all 24 musical keys arranged so that harmonically compatible keys are adjacent.',
    color=C_MID, size=10)
add_feature_table(doc,
    ['Ring / Position', 'Meaning'],
    [
        ['Outer ring (B keys)',          'Major keys (e.g., 8B = C major, 9B = G major)'],
        ['Inner ring (A keys)',           'Minor keys (e.g., 8A = A minor, 9A = E minor)'],
        ['Same number (e.g., 8A + 8B)',  'Relative major/minor — very compatible (parallel keys)'],
        ['Adjacent number (e.g., 8A+7A)','Neighboring key — sounds natural and smooth'],
        ['±2 positions',                 'Can work — slightly more harmonic tension'],
        ['Opposite side of wheel',       'Clashing keys — avoid for smooth transitions'],
    ],
    col_widths=[6, 10]
)
add_callout(doc, 'ℹ️', 'Camelot Notation:',
    'Numbers 1–12 represent key families. "A" = minor keys, "B" = major keys. Adjacent numbers are harmonically compatible.',
    C_BG_BLUE, C_BLUE)

# ══════════════════════════════════════════════
# 9. EQ & MIXING
# ══════════════════════════════════════════════
add_heading(doc, '9. EQ & Mixing', level=1)
add_feature_table(doc,
    ['Band', 'Frequency', 'Controls'],
    [
        ['HI (High shelf)', '10,000 Hz+',    'Cymbals, hi-hats, air, sibilance'],
        ['MID (Peaking)',    '~1,000 Hz',     'Vocals, guitars, presence'],
        ['LO (Low shelf)',   '200 Hz and below', 'Kick drum, bass, sub frequencies'],
    ],
    col_widths=[4, 4, 8]
)
add_heading(doc, 'Using the Knobs', level=2)
for txt in [
    'Drag UP on a knob to boost that frequency range (+12 dB max)',
    'Drag DOWN to cut that frequency range (–12 dB max)',
    'Scroll the mouse wheel on a knob to adjust it precisely',
    'Knob at center position = flat (no boost or cut)',
]:
    add_bullet(doc, txt)
doc.add_paragraph()
add_callout(doc, '💡', 'Classic EQ Transition Technique:',
    'Before bringing in Deck B, cut the lows (LO knob down) on Deck B. As you crossfade to B, gradually bring the lows back up. This prevents bass clash between the two tracks.',
    C_BG_PURPLE, C_PURPLE)

# ══════════════════════════════════════════════
# 10. CROSSFADER
# ══════════════════════════════════════════════
add_heading(doc, '10. Crossfader', level=1)
add_feature_table(doc,
    ['Mode', 'Best For', 'Behavior'],
    [
        ['Linear', 'General mixing',     'Smooth, proportional blend across the full range'],
        ['Scratch', 'Hip-hop, turntablism', 'Quick cut — audio switches sharply near the edges'],
        ['Smooth',  'Deep house, ambient', 'Cosine curve — gradual fade that sounds natural'],
    ],
    col_widths=[4, 4, 8]
)

page_break(doc)

# ══════════════════════════════════════════════
# 11. EFFECTS
# ══════════════════════════════════════════════
add_heading(doc, '11. Effects', level=1)
add_callout(doc, '⚠️', 'FX Deck Routing:',
    'Effects are applied to one deck at a time. Use the A / B buttons in the Effects panel to choose which deck receives the effects.',
    C_BG_YELLOW, C_YELLOW)

for fx_name, fx_desc, knobs in [
    ('REVERB',  'Adds room/hall ambience, making the sound feel like it\'s in a large space.',
     [('Wet', 'How much reverb is mixed in (0 = dry, max = very wet)'), ('Size', 'Size of the virtual room (larger = longer reverb tail)')]),
    ('DELAY',   'Creates repeating echoes of the audio.',
     [('Time', 'Delay time in seconds (0–1.5 sec)'), ('Feedback', 'How many times the echo repeats (0 = once, max = infinite)')]),
    ('FILTER',  'A bandpass filter that lets through only a narrow range of frequencies — great for buildups.',
     [('Freq', 'Center frequency (sweep from 80 Hz to 16,000 Hz)'), ('Q', 'Width of the filter (higher = narrower, more resonant)')]),
    ('FLANGER', 'A classic modulation effect creating a swirling, jet-engine-like sound.',
     [('Rate', 'Speed of the flanger oscillation'), ('Depth', 'Intensity of the effect')]),
    ('ECHO',    'A single-repeat echo effect for dramatic rhythmic accents.',
     [('Beat Div', 'Echo timing relative to the beat')]),
]:
    add_heading(doc, fx_name, level=2)
    add_para(doc, fx_desc, color=C_MID, size=10)
    add_feature_table(doc, ['Knob', 'Function'], knobs, col_widths=[4, 12])

# ══════════════════════════════════════════════
# 12. YOUTUBE SEARCH & MP3 DOWNLOAD
# ══════════════════════════════════════════════
add_heading(doc, '12. YouTube Search & MP3 Download', level=1)
add_heading(doc, 'Getting a Free YouTube API Key', level=2)
steps_api = [
    ('Go to Google Cloud Console', 'Visit console.cloud.google.com and sign in with your Google account.'),
    ('Create a new project',       'Click the project dropdown → New Project → name it "DJ Mixer" → Create.'),
    ('Enable YouTube Data API v3', 'Go to APIs & Services → Library → search "YouTube Data API v3" → Enable.'),
    ('Create credentials',         'Go to APIs & Services → Credentials → Create Credentials → API Key. Copy the key.'),
    ('Paste it in ProMix',         'In the ProMix header bar, paste your key into the API Key field. Searches are now live.'),
]
for i, (t, b) in enumerate(steps_api, 1): add_step(doc, i, t, b)
add_callout(doc, 'ℹ️', 'Free quota:',
    'Google gives you 10,000 search units per day for free — roughly 100 searches. No credit card required.',
    C_BG_BLUE, C_BLUE)

add_heading(doc, 'Downloading MP3s', level=2)
steps_dl = [
    ('Find a track via YouTube Search', 'Search for the song in the search bar.'),
    ('Click ⬇ MP3',                    'The app requests a download link from cobalt.tools (free, no account needed).'),
    ('The MP3 file downloads',          'Your browser saves it to your Downloads folder.'),
    ('Load the file into ProMix',       'Use 📂 Load to load your downloaded MP3 to a deck.'),
]
for i, (t, b) in enumerate(steps_dl, 1): add_step(doc, i, t, b)
add_callout(doc, '⚠️', 'Important:',
    'Only download music you have the right to use. Respect copyright law. ProMix DJ Studio is intended for personal use, private events, and licensed DJ performances.',
    C_BG_YELLOW, C_YELLOW)

page_break(doc)

# ══════════════════════════════════════════════
# 13. RECORDING
# ══════════════════════════════════════════════
add_heading(doc, '13. Recording Your Mix', level=1)
steps_rec = [
    ('Before you start mixing, click ● REC', 'The button turns red and pulses. Everything from this point is captured.'),
    ('Perform your mix',                      'Load tracks, play, crossfade, apply effects — mix as normal.'),
    ('Click ⏹ STOP REC when finished',        'The button returns to normal. Your mix is stored in memory.'),
    ('Click ⬇ Download Mix',                  'A file named ProMix_[timestamp].webm downloads to your computer.'),
]
for i, (t, b) in enumerate(steps_rec, 1): add_step(doc, i, t, b)

add_heading(doc, 'Converting to MP3', level=2)
for txt in [
    'VLC Media Player — Media → Convert/Save → Add file → choose MP3 output',
    'Audacity — File → Import → export as MP3',
    'Online converter — cloudconvert.com (free)',
    'FFmpeg (command line): ffmpeg -i ProMix_123.webm output.mp3',
]:
    add_bullet(doc, txt)
doc.add_paragraph()

# ══════════════════════════════════════════════
# 14. AUTO MIX
# ══════════════════════════════════════════════
add_heading(doc, '14. Auto Mix', level=1)
add_para(doc,
    'The ⚡ Auto Mix feature automates the transition between Deck A and Deck B — '
    'it syncs the BPM and smoothly moves the crossfader over about 4 seconds.',
    color=C_MID, size=10)
steps_am = [
    ('Load tracks on both decks',          'Make sure both tracks are loaded and BPM is detected.'),
    ('Start playing Deck A',               'Let the track play to the point where you want to mix out.'),
    ('Click ⚡ Auto Mix',                   'The app: (a) syncs Deck B\'s tempo to Deck A, (b) starts Deck B, (c) smoothly fades crossfader from A to B.'),
]
for i, (t, b) in enumerate(steps_am, 1): add_step(doc, i, t, b)

# ══════════════════════════════════════════════
# 15. STEM SEPARATION
# ══════════════════════════════════════════════
add_heading(doc, '15. Stem Separation', level=1)
add_para(doc,
    'Stems are the individual components of a song: Vocals, Drums, Bass, and Melody. '
    'By separating them, you can create acapellas, instrumentals, or completely remix a track.',
    color=C_MID, size=10)
steps_stem = [
    ('Click ⚗ Separate Stems (AI)',   'This opens Fadr.com in a new tab.'),
    ('Upload your track to Fadr',      'Drag and drop your audio file onto the Fadr interface.'),
    ('Wait for AI processing',         'Under 60 seconds. Fadr splits into: Vocals, Drums, Bass, Melody (and up to 16 components on Pro).'),
    ('Download the individual stems',  'Download the acapella (vocals only), instrumental, or any stem.'),
    ('Load stems back into ProMix',    'Load an acapella on Deck A and an instrumental on Deck B — you\'ve made a mashup!'),
]
for i, (t, b) in enumerate(steps_stem, 1): add_step(doc, i, t, b)

# ══════════════════════════════════════════════
# 16. SYNC & PITCH
# ══════════════════════════════════════════════
add_heading(doc, '16. Sync & Pitch', level=1)
add_feature_table(doc,
    ['Control', 'Function'],
    [
        ['SYNC Button',      'Click on a deck to match its BPM to the other deck automatically'],
        ['Pitch Slider',     'Adjusts playback speed from –8% to +8%. Right = faster, left = slower.'],
        ['Key Shift –/+',    'Shifts pitch by one semitone per click. Changes key display and playback pitch.'],
    ],
    col_widths=[5, 11]
)

page_break(doc)

# ══════════════════════════════════════════════
# 17. KEYBOARD SHORTCUTS
# ══════════════════════════════════════════════
add_heading(doc, '17. Keyboard Shortcuts', level=1)
add_para(doc, 'Shortcuts do not fire when the cursor is inside a text input field.', color=C_MID, size=10)
add_kb_table(doc, [
    ('Space',  'Play / Pause Deck A'),
    ('Enter',  'Play / Pause Deck B'),
    ('←',      'Seek Deck A back 5 seconds'),
    ('→',      'Seek Deck A forward 5 seconds'),
    ('Q',      'Set/Jump Cue on Deck A'),
    ('W',      'Set/Jump Cue on Deck B'),
    ('S',      'Sync Deck B to Deck A\'s BPM'),
    ('A',      'Toggle Mix Recording on/off'),
])

page_break(doc)

# ══════════════════════════════════════════════
# TUTORIALS
# ══════════════════════════════════════════════
add_heading(doc, '18. Tutorial 1 — Your First Mix', level=1)
add_para(doc, 'Goal: Load two tracks and perform a manual crossfader mix. Time: ~5 minutes.', italic=True, color=C_MID, size=10)
steps_t1 = [
    ('Open dj-mixer.html in Chrome',         'Double-click the file. The app loads instantly.'),
    ('Load a track on Deck A',               'Click 📂 Load. Choose an upbeat song with a steady kick drum.'),
    ('Load a different track on Deck B',     'Choose something with similar energy and tempo.'),
    ('Move crossfader all the way left (A)', 'Deck A is now the only thing heard.'),
    ('Press Space to play Deck A',           'Listen and enjoy.'),
    ('Click SYNC on Deck B',                 'Matches Deck B\'s tempo to Deck A automatically.'),
    ('Position Deck B at the start',         'Click the far left of Deck B\'s waveform.'),
    ('Start Deck B at a phrase boundary',    'When Deck A reaches a chorus or phrase end, press Enter to start Deck B.'),
    ('Slowly slide the crossfader to the right', 'Over 8–16 beats, fade from A to B. You\'ve performed your first mix!'),
    ('Pause Deck A',                          'Load your next track on Deck A and repeat.'),
]
for i, (t, b) in enumerate(steps_t1, 1): add_step(doc, i, t, b)

add_heading(doc, '19. Tutorial 2 — DJ Party Set with Recording', level=1)
add_para(doc, 'Goal: Record a 3-track mix with effects. Time: ~15 minutes.', italic=True, color=C_MID, size=10)
steps_t2 = [
    ('Prepare your tracklist',              'Choose 3 songs ~120–130 BPM. Check Camelot keys for harmonic compatibility.'),
    ('Click ● REC to start recording',     'The button pulses red. Everything is being captured.'),
    ('Load Track 1 on Deck A and play it', 'Press Space. Let it run for a minute or two.'),
    ('Load Track 2 on Deck B',             'Sync it. Position Deck B at the intro or drop.'),
    ('Apply a build-up effect',            'Enable FILTER on Deck A. Sweep the Freq knob upward to signal the incoming transition.'),
    ('Cut the bass on Deck A (LO knob down)', 'Drag Deck A\'s LO knob all the way counter-clockwise. Removes bass clash.'),
    ('Start Deck B and move crossfader to B', 'Slowly push crossfader right over 8 bars. Bring Deck B\'s LO up as you go.'),
    ('Load Track 3 on Deck A and repeat',  'Deck A is free. Load Track 3, sync it, repeat the second transition.'),
    ('Click ⏹ STOP REC and ⬇ Download Mix', 'Your 3-track mix is saved as a WebM audio file.'),
]
for i, (t, b) in enumerate(steps_t2, 1): add_step(doc, i, t, b)

page_break(doc)

add_heading(doc, '20. Tutorial 3 — YouTube to Mix', level=1)
add_para(doc, 'Goal: Search YouTube, download two tracks as MP3, and mix them. Requires API key.', italic=True, color=C_MID, size=10)
steps_t3 = [
    ('Enter your YouTube API key',      'Paste it into the header API key field.'),
    ('Search for your first song',      'Type in the search bar: "Calvin Harris Summer". Press Enter.'),
    ('Click ⬇ MP3 on the best result',  'Wait for the download. Save the file.'),
    ('Search for your second song',     'E.g., "Avicii Wake Me Up". Download as MP3.'),
    ('Load both MP3s into the decks',   '📂 Load → Track 1 to Deck A, Track 2 to Deck B.'),
    ('Check harmonic compatibility',    'Note both Camelot keys. Open Camelot Wheel to verify compatibility. Use Key Shift to adjust if needed.'),
    ('Mix them',                        'Click ● REC, perform your mix, click ⬇ Download Mix to save.'),
]
for i, (t, b) in enumerate(steps_t3, 1): add_step(doc, i, t, b)

add_heading(doc, '21. Tutorial 4 — Harmonic Mixing', level=1)
add_para(doc, 'Goal: Build a harmonically perfect set using only compatible keys.', italic=True, color=C_MID, size=10)
steps_t4 = [
    ('Load your first track on Deck A',         'Note the KEY display. Example: 8B (C major).'),
    ('Open the Camelot Wheel',                  'Find position 8B. Compatible keys: 8A (A minor), 7B (F major), 9B (G major).'),
    ('Find a track in a compatible key',        'Load candidates on Deck B and check their detected key.'),
    ('If close but not exact, use Key Shift',   'If Deck B shows 9A instead of 8A, click – once. Now both are compatible.'),
    ('Mix in Deck B harmonically',              'The transition will sound musically natural — no clashing.'),
    ('Plan the rest of your set',               'From 8A, compatible next keys are 8B, 7A, 9A. Move around the wheel for a harmonic journey.'),
]
for i, (t, b) in enumerate(steps_t4, 1): add_step(doc, i, t, b)

page_break(doc)

# ══════════════════════════════════════════════
# 22. FAQ
# ══════════════════════════════════════════════
add_heading(doc, '22. FAQ', level=1)
faqs = [
    ('No sound is coming out',
     'Browsers require a user interaction before playing audio. Click anywhere on the page first. '
     'Also check system volume and that the browser tab is not muted (right-click tab → Unmute).'),
    ('BPM is half or double the actual tempo',
     'Common with some genres. Use the TAP button to correct it — tap along to the beat 4–8 times.'),
    ('YouTube search returns an error',
     'Your API key may be invalid, expired, or over daily quota. Check in Google Cloud Console. '
     'Free quota is 10,000 units (~100 searches) per day.'),
    ('MP3 download does not start',
     'cobalt.tools may be temporarily busy. Try again in a moment. Alternatively, click ▶ YT to open on YouTube.'),
    ('Can I use this app offline?',
     'Yes — all core mixing features work offline. You only need internet for YouTube search and MP3 download.'),
    ('Can I save my mix settings or hot cues?',
     'Currently, settings are not saved between sessions. Hot cues and loaded tracks reset when you refresh the page. Save your recording before closing.'),
    ('The file does not open on my computer',
     'Right-click dj-mixer.html → Open with → choose Google Chrome, Firefox, or Microsoft Edge.'),
    ('Can I use this on my phone or tablet?',
     'Yes — it runs in mobile browsers. Landscape mode on a tablet is recommended for the full layout.'),
]
for q, a in faqs:
    add_heading(doc, q, level=2)
    add_para(doc, a, color=C_MID, size=10)

# ══════════════════════════════════════════════
# 23. TROUBLESHOOTING
# ══════════════════════════════════════════════
add_heading(doc, '23. Troubleshooting', level=1)
add_feature_table(doc,
    ['Issue', 'Solution'],
    [
        ['App is silent',                 'Click anywhere on the page first. Check system volume. Ensure tab is not muted.'],
        ['Waveform does not appear',      'Wait 2–3 seconds. If still missing, try converting the file to MP3 first.'],
        ['BPM says "--"',                 'Use TAP tempo to set it manually.'],
        ['Key says "--"',                 'Key detection needs audio data. Click in the middle of the track and reload.'],
        ['Reverb/Delay has no effect',    'Ensure the FX toggle is active (purple) AND the correct deck (A or B) is selected.'],
        ['YouTube search is blank',       'Confirm API key is correctly entered. Try disabling VPN or ad-blocker.'],
        ['Download Mix is empty/tiny',    'You must click ● REC before starting your mix. Re-record if needed.'],
        ['Tracks are out of sync',        'The BPM of one track may be wrong (double-time). Correct with TAP, then SYNC again.'],
        ['App is slow on long tracks',    'Very long WAV files may be slow. Use MP3 format for best performance.'],
    ],
    col_widths=[6, 10]
)

# Footer
doc.add_paragraph()
line2 = doc.add_paragraph()
line2.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(line2, '─' * 55, color=C_LAVENDER, size=10)
footer_p = doc.add_paragraph()
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(footer_p, 'ProMix DJ Studio  ·  User Manual v1.0\n', bold=True, color=C_PURPLE, size=10)
add_run(footer_p, 'Built on Web Audio API  ·  MP3 download via cobalt.tools  ·  Stem separation via Fadr\n', color=C_LIGHT, size=9)
add_run(footer_p, 'Open source  ·  No installation required  ·  Works in any modern browser', color=C_LIGHT, size=9)

# ─────────────────────────────────────────────
# SAVE DOCX
# ─────────────────────────────────────────────
docx_path = '/home/user/joz-repository/ProMix_DJ_Studio_User_Manual.docx'
doc.save(docx_path)
print(f'DOCX saved: {docx_path}')

# ─────────────────────────────────────────────
# GENERATE PDF via ReportLab
# ─────────────────────────────────────────────
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib.colors import Color, HexColor, white, black
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, PageBreak, HRFlowable, KeepTogether)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus.flowables import Flowable
from reportlab.lib import colors

pdf_path = '/home/user/joz-repository/ProMix_DJ_Studio_User_Manual.pdf'

# Colours
COL_PURPLE  = HexColor('#7C3AED')
COL_INDIGO  = HexColor('#4F46E5')
COL_LAVEND  = HexColor('#A78BFA')
COL_BLUE    = HexColor('#60A5FA')
COL_GREEN   = HexColor('#34D399')
COL_YELLOW  = HexColor('#FBBF24')
COL_RED     = HexColor('#F87171')
COL_DARK    = HexColor('#1E1B4B')
COL_MID     = HexColor('#4B5563')
COL_LIGHT   = HexColor('#6B7280')
COL_BG_PUR  = HexColor('#EDE9FE')
COL_BG_BLU  = HexColor('#DBEAFE')
COL_BG_GRN  = HexColor('#D1FAE5')
COL_BG_YEL  = HexColor('#FEF3C7')
COL_SIDEBAR = HexColor('#1A1A2E')
COL_ROW_ALT = HexColor('#F3F0FF')
COL_THEAD   = HexColor('#1A1A2E')

doc_pdf = SimpleDocTemplate(
    pdf_path, pagesize=A4,
    topMargin=2*cm, bottomMargin=2*cm,
    leftMargin=2.5*cm, rightMargin=2.5*cm,
    title='ProMix DJ Studio User Manual',
    author='ProMix DJ Studio',
)

styles = getSampleStyleSheet()
S = {
    'h1':  ParagraphStyle('h1',  parent=styles['Heading1'], fontName='Helvetica-Bold',
                           fontSize=22, textColor=COL_PURPLE, spaceAfter=8, spaceBefore=18),
    'h2':  ParagraphStyle('h2',  parent=styles['Heading2'], fontName='Helvetica-Bold',
                           fontSize=14, textColor=COL_DARK, spaceAfter=6, spaceBefore=14),
    'h3':  ParagraphStyle('h3',  parent=styles['Heading3'], fontName='Helvetica-Bold',
                           fontSize=11, textColor=COL_INDIGO, spaceAfter=4, spaceBefore=10),
    'body': ParagraphStyle('body', parent=styles['Normal'],  fontName='Helvetica',
                            fontSize=10, textColor=COL_MID, spaceAfter=6, leading=15),
    'bold': ParagraphStyle('bold', parent=styles['Normal'],  fontName='Helvetica-Bold',
                            fontSize=10, textColor=COL_DARK, spaceAfter=4),
    'small': ParagraphStyle('small', parent=styles['Normal'], fontName='Helvetica',
                              fontSize=8.5, textColor=COL_LIGHT, spaceAfter=4),
    'mono': ParagraphStyle('mono', parent=styles['Normal'],  fontName='Courier',
                             fontSize=8, textColor=COL_DARK, spaceAfter=4, leading=12),
    'center': ParagraphStyle('center', parent=styles['Normal'], fontName='Helvetica',
                               fontSize=10, textColor=COL_MID, alignment=TA_CENTER, spaceAfter=6),
    'toc':  ParagraphStyle('toc', parent=styles['Normal'],   fontName='Helvetica',
                             fontSize=10, textColor=COL_MID, spaceAfter=3, leftIndent=12),
    'step_title': ParagraphStyle('step_title', parent=styles['Normal'], fontName='Helvetica-Bold',
                                  fontSize=10.5, textColor=COL_DARK, spaceAfter=2),
    'step_body':  ParagraphStyle('step_body',  parent=styles['Normal'], fontName='Helvetica',
                                  fontSize=9.5, textColor=COL_MID, spaceAfter=4, leftIndent=8),
    'bullet': ParagraphStyle('bullet', parent=styles['Normal'], fontName='Helvetica',
                               fontSize=10, textColor=COL_MID, spaceAfter=3, leftIndent=16, bulletIndent=6),
    'italic': ParagraphStyle('italic', parent=styles['Normal'], fontName='Helvetica-Oblique',
                               fontSize=10, textColor=COL_MID, spaceAfter=6),
    'cover_title': ParagraphStyle('cover_title', fontName='Helvetica-Bold', fontSize=40,
                                   textColor=COL_PURPLE, alignment=TA_CENTER, spaceAfter=12),
    'cover_sub':   ParagraphStyle('cover_sub',   fontName='Helvetica-Oblique', fontSize=16,
                                   textColor=COL_MID, alignment=TA_CENTER, spaceAfter=20),
    'cover_small': ParagraphStyle('cover_small', fontName='Helvetica', fontSize=11,
                                   textColor=COL_LAVEND, alignment=TA_CENTER, spaceAfter=6),
}

def tbl_style(has_header=True, stripe=True):
    cmds = [
        ('FONTNAME',  (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE',  (0,0), (-1,-1), 9),
        ('TEXTCOLOR', (0,0), (-1,-1), COL_MID),
        ('ALIGN',     (0,0), (-1,-1), 'LEFT'),
        ('VALIGN',    (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING',  (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 0.5, HexColor('#E5E7EB')),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [white, COL_ROW_ALT]),
    ]
    if has_header:
        cmds += [
            ('BACKGROUND',  (0,0), (-1,0), COL_THEAD),
            ('TEXTCOLOR',   (0,0), (-1,0), white),
            ('FONTNAME',    (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',    (0,0), (-1,0), 9),
        ]
    if stripe:
        cmds.append(('FONTNAME', (0,1), (0,-1), 'Helvetica-Bold'))
        cmds.append(('TEXTCOLOR', (0,1), (0,-1), COL_DARK))
    return TableStyle(cmds)

def callout_table(icon, label, body, bg, border):
    data = [[Paragraph(icon, S['body']),
             Paragraph(f'<b>{label}</b>  {body}', S['body'])]]
    t = Table(data, colWidths=[1*cm, None])
    t.setStyle(TableStyle([
        ('BACKGROUND',   (0,0), (-1,-1), bg),
        ('LEFTPADDING',  (0,0), (0,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING',   (0,0), (-1,-1), 8),
        ('BOTTOMPADDING',(0,0), (-1,-1), 8),
        ('LINECOLOR',    (0,0), (0,-1), border),
        ('LINEBEFORE',   (0,0), (0,-1), 4, border),
        ('VALIGN',       (0,0), (-1,-1), 'TOP'),
    ]))
    return t

def step_table(num, title, body):
    data = [[Paragraph(f'<font color="white"><b>{num}</b></font>', S['center']),
             [Paragraph(f'<b>{title}</b>', S['step_title']),
              Paragraph(body, S['step_body'])]]]
    t = Table(data, colWidths=[1.0*cm, None])
    t.setStyle(TableStyle([
        ('BACKGROUND',  (0,0), (0,-1), COL_PURPLE),
        ('VALIGN',      (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING',  (0,0), (-1,-1), 6),
        ('BOTTOMPADDING',(0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING',(0,0), (-1,-1), 6),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [COL_BG_PUR]),
    ]))
    return t

story = []
W = A4[0] - 5*cm  # usable width

# ── COVER ──
story.append(Spacer(1, 3*cm))
story.append(Paragraph('ProMix DJ Studio', S['cover_title']))
story.append(Paragraph('The Ultimate Browser-Based Party Mixer', S['cover_sub']))
story.append(HRFlowable(width='80%', thickness=2, color=COL_LAVEND, spaceAfter=20))
story.append(Paragraph('USER MANUAL  ·  Version 1.0', S['cover_small']))
story.append(Spacer(1, 1*cm))
feat_line = '  ●  '.join(['Dual Decks','BPM & Key Detection','Camelot Wheel',
                           'YouTube Search','MP3 Download','Mix Recording',
                           '3-Band EQ','5 Effects','Auto Mix','Stem Separation'])
story.append(Paragraph(feat_line, S['small']))
story.append(PageBreak())

# ── TOC ──
story.append(Paragraph('Table of Contents', S['h1']))
for num, title in toc_items:
    story.append(Paragraph(f'<font color="#7C3AED"><b>{num}</b></font>  {title}', S['toc']))
story.append(PageBreak())

def section(num_title, paragraphs):
    story.append(Paragraph(num_title, S['h1']))
    for p in paragraphs:
        story.append(p)

def sub(title):
    story.append(Paragraph(title, S['h2']))

def p(text, style='body'):
    return Paragraph(text, S[style])

def data_table(headers, rows, col_widths=None):
    data = [headers] + rows
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(tbl_style())
    story.append(t)
    story.append(Spacer(1, 6))

def callout(icon, label, body, bg, border):
    story.append(callout_table(icon, label, body, bg, border))
    story.append(Spacer(1, 6))

def steps(items):
    for i, (title, body) in enumerate(items, 1):
        story.append(step_table(i, title, body))
        story.append(Spacer(1, 4))

def bullet(text):
    story.append(Paragraph(f'• {text}', S['bullet']))

# ── 1. OVERVIEW ──
section('1. Overview', [
    p('ProMix DJ Studio is a fully self-contained DJ mixing application delivered as a single HTML file. '
      'Open it in any modern web browser — no installation, no app store, no subscription. '
      'Everything runs locally using the Web Audio API.'),
    p('It was designed by analyzing DJ.Studio, Algoriddim djay Pro, VirtualDJ, and Fadr — '
      'combining the best features of all four into one free, portable tool.'),
])
sub('Feature Summary')
data_table(['Feature', 'Description'], [
    ['Dual Decks',         'Two independent decks (A and B) with animated vinyl turntables'],
    ['Waveform Display',   'Real-time waveform with click-to-seek and playhead position'],
    ['BPM Detection',      'Automatic beats-per-minute analysis on every loaded track'],
    ['Key Detection',      'Musical key detection mapped to Camelot Wheel notation'],
    ['Camelot Wheel',      'Interactive harmonic mixing guide for compatible key transitions'],
    ['3-Band EQ',          'Hi / Mid / Lo equalizer per deck with rotary knobs'],
    ['Crossfader',         'Linear, Scratch, and Smooth curve modes'],
    ['5 Effects',          'Reverb, Delay, Filter, Flanger, Echo — all with parameter knobs'],
    ['Hot Cues',           '8 hot cue points per deck — set and jump with one click'],
    ['Loop Control',       '4-beat auto-loop with one button'],
    ['Pitch / Sync',       '±8% pitch slider, key shift (±semitones), one-click BPM sync'],
    ['Auto Mix',           'One-click BPM-synced crossfader transition'],
    ['YouTube Search',     'Search by song/artist, load results directly to a deck'],
    ['MP3 Download',       'Download any YouTube track as MP3 via cobalt.tools'],
    ['Mix Recording',      'Record the master output and download as an audio file'],
    ['Stems Panel',        'Vocal/Drum/Bass/Melody controls + AI stem separation via Fadr'],
], col_widths=[4.5*cm, 12*cm])

# ── 2. SYSTEM REQUIREMENTS ──
section('2. System Requirements', [])
data_table(['Item', 'Requirement'], [
    ['Browser',       'Chrome 90+, Firefox 88+, Edge 90+, Safari 15+'],
    ['OS',            'Windows 10/11, macOS 11+, Linux'],
    ['RAM',           '4 GB min, 8 GB recommended'],
    ['Audio Files',   'MP3, WAV, FLAC, OGG, AAC, M4A'],
    ['Internet',      'Only for YouTube search and MP3 download'],
    ['API Key',       'Optional — enables YouTube in-app search (free)'],
], col_widths=[4.5*cm, 12*cm])
callout('💡', 'Best Experience:', 'Use Google Chrome for the best Web Audio API performance.',
        COL_BG_PUR, COL_PURPLE)

# ── 3. QUICK START ──
section('3. Quick Start Guide', [p('Get mixing in under 2 minutes:')])
steps([
    ('Open the file',       'Download dj-mixer.html from GitHub. Double-click to open in your browser.'),
    ('Load Track A',        'Click 📂 Load on Deck A. Select any MP3 or audio file. Waveform appears and BPM/key are auto-detected.'),
    ('Load Track B',        'Click 📂 Load on Deck B. Load a second track.'),
    ('Play Deck A',         'Press the green ▶ Play button or press Space on your keyboard.'),
    ('Sync and mix',        'Press SYNC on Deck B to match tempo. Press Play on Deck B. Slide the crossfader to blend.'),
    ('Record your mix',     'Click ● REC before mixing. Click ⬇ Download Mix to save when done.'),
])
callout('✅', 'You\'re mixing!', 'That\'s all it takes. Read on to explore all professional features.',
        COL_BG_GRN, COL_GREEN)
story.append(PageBreak())

# ── 4. INTERFACE ──
section('4. Interface Layout', [
    p('The interface is organized in a professional DJ layout. Below is a schematic overview:'),
    Paragraph(
        '┌─ Header Bar: Logo · API Key · REC · Download · Camelot · Auto Mix ─┐\n'
        '├─ Search Bar + Results Panel ──────────────────────────────────────┤\n'
        '├─ DECK A ──────────┬─── MIXER CENTER ────┬─── DECK B ─────────────┤\n'
        '│  Vinyl · Waveform │  EQ A/B · VU Meters │  Vinyl · Waveform      │\n'
        '│  BPM · Key · Time │  Vol Faders         │  BPM · Key · Time      │\n'
        '│  Play/Cue/Sync    │  Crossfader         │  Play/Cue/Sync         │\n'
        '│  Hot Cues 1-8     │  Master Vol         │  Hot Cues 1-8          │\n'
        '├───────────────────┴─────────────────────┴────────────────────────┤\n'
        '│  Effects: REVERB · DELAY · FILTER · FLANGER · ECHO               │\n'
        '│  Stems: VOCALS · DRUMS · BASS · MELODY · AI Separate             │\n'
        '│  Timeline: Deck A overview + Deck B overview (click to seek)      │\n'
        '└────────────────────────────────────────────────────────────────────┘',
        S['mono']
    ),
])
sub('Deck Controls Reference')
data_table(['Control', 'Function'], [
    ['📂 Load',     'Open file browser to load an audio file'],
    ['Vinyl Canvas','Animated spinning record'],
    ['Waveform',    'Click anywhere to jump to that position'],
    ['BPM',         'Auto-detected tempo (use TAP to correct)'],
    ['KEY',         'Camelot notation key (e.g. 8A = A minor)'],
    ['▶ Play',      'Start or pause playback'],
    ['CUE',         'Set cue point (paused) or jump to it (playing)'],
    ['SYNC',        'Match BPM to the other deck'],
    ['LOOP',        'Activate a 4-beat loop from current position'],
    ['Hot Cues 1-8','Set and jump to up to 8 saved positions'],
    ['PITCH slider','Adjust playback speed ±8%'],
    ['Key Shift –/+','Shift pitch by semitones'],
], col_widths=[4.5*cm, 12*cm])

# ── 5. LOADING TRACKS ──
section('5. Loading Tracks', [])
steps([
    ('Click 📂 Load', 'The file browser opens. Navigate to your music folder.'),
    ('Select your audio file', 'Supported: MP3, WAV, FLAC, OGG, AAC, M4A.'),
    ('Wait for analysis', 'The waveform appears and BPM/key are detected. Takes 1–3 seconds.'),
    ('Check the readout', 'If BPM looks wrong, use TAP tempo to correct it.'),
])
callout('⚠️', 'Large files:', 'Files over 100 MB may take several seconds. MP3s at 320kbps load fastest.',
        COL_BG_YEL, COL_YELLOW)

# ── 6. PLAYBACK ──
section('6. Playback Controls, Hot Cues & Loops', [])
data_table(['Action', 'How To'], [
    ['Play / Pause',     'Click ▶ or press Space (Deck A) / Enter (Deck B)'],
    ['Seek forward 5s',  'Press → (Deck A)'],
    ['Seek back 5s',     'Press ← (Deck A)'],
    ['Jump to position', 'Click anywhere on waveform or timeline'],
    ['Set Cue Point',    'Pause at desired position, click CUE'],
    ['Jump to Cue',      'Click CUE while playing'],
], col_widths=[5*cm, 11.5*cm])
sub('Hot Cues')
steps([
    ('Navigate to the spot to mark',        'For example, the start of the chorus.'),
    ('Click an empty Hot Cue button (1–8)', 'The button changes color. Position is saved.'),
    ('Click the colored button to jump',    'Instantly jumps to that saved position at any time.'),
])
callout('💡', 'Pro Tip:', 'Set Hot Cue 1 at intro, Cue 2 at first drop, Cue 3 at breakdown for fast navigation.',
        COL_BG_PUR, COL_PURPLE)
sub('Loop Controls')
data_table(['Action', 'How To'], [
    ['Start loop', 'Click LOOP during playback — activates 4-beat loop from current beat'],
    ['Exit loop',  'Click LOOP again — returns to normal playback'],
], col_widths=[5*cm, 11.5*cm])
story.append(PageBreak())

# ── 7. BPM & KEY ──
section('7. BPM & Key Detection', [
    p('ProMix analyzes audio using onset energy detection to determine tempo. '
      'BPM is clamped to 70–200 BPM to handle half-time and double-time tracks correctly.'),
])
sub('Correcting BPM with Tap Tempo')
steps([
    ('Play the track',               'Listen to the beat.'),
    ('Click TAP on the beat',        'Click 4+ times. BPM updates after each tap.'),
    ('Tap after 2+ seconds to reset','A 2-second gap clears previous taps.'),
])
sub('Key Detection')
story.append(p('ProMix uses a pitch class profile (Krumhansl–Schmuckler model) to detect key. '
               'Result is shown in Camelot Wheel notation (e.g., 8A = A minor, 8B = C major).'))
callout('ℹ️', 'Camelot Notation:', 'Numbers 1–12 represent key families. A = minor, B = major. Adjacent numbers are compatible.',
        COL_BG_BLU, COL_BLUE)

# ── 8. CAMELOT ──
section('8. Camelot Wheel', [
    p('Open the Camelot Wheel from the header to see all 24 musical keys arranged by harmonic compatibility.'),
])
data_table(['Position', 'Meaning'], [
    ['Same number (e.g., 8A+8B)',   'Relative major/minor — very compatible'],
    ['Adjacent number (e.g., 8A+7A)', 'Neighboring key — smooth and natural'],
    ['±2 positions',                'Slightly more tension but can work'],
    ['Opposite side of wheel',      'Clashing keys — avoid for smooth mixes'],
    ['Outer ring (B keys)',         'Major keys'],
    ['Inner ring (A keys)',         'Minor keys'],
], col_widths=[5*cm, 11.5*cm])
steps([
    ('Note the key of your current track', 'Example: Deck A shows 8A (A minor).'),
    ('Open the Camelot Wheel',             'Find position 8A.'),
    ('Choose a compatible key for Deck B', 'Compatible: 8B (C major), 7A (D minor), 9A (E minor).'),
    ('Mix in harmonically',                'The transition sounds musically natural.'),
])

# ── 9. EQ ──
section('9. EQ & Mixing', [])
data_table(['Band', 'Frequency', 'Controls'], [
    ['HI (High shelf)', '10,000 Hz+',    'Cymbals, hi-hats, air'],
    ['MID (Peaking)',   '~1,000 Hz',     'Vocals, guitars, presence'],
    ['LO (Low shelf)',  '200 Hz and below', 'Kick drum, bass, sub'],
], col_widths=[4*cm, 4*cm, 8.5*cm])
story.append(p('Drag knobs up to boost (+12 dB max), down to cut (–12 dB max). '
               'Scroll the mouse wheel for precise control. Center = flat.'))
callout('💡', 'Classic EQ Technique:',
        'Cut the LO on Deck B before mixing in. As you crossfade to B, gradually bring the lows back up. Prevents bass clash.',
        COL_BG_PUR, COL_PURPLE)

# ── 10. CROSSFADER ──
section('10. Crossfader', [p('Far left = only Deck A. Far right = only Deck B. Center = equal blend.')])
data_table(['Mode', 'Best For', 'Behavior'], [
    ['Linear', 'General mixing',     'Smooth, proportional blend'],
    ['Scratch', 'Hip-hop, turntablism', 'Sharp cut near edges'],
    ['Smooth',  'Deep house, ambient', 'Cosine curve, very gradual fade'],
], col_widths=[3.5*cm, 5*cm, 8*cm])
story.append(PageBreak())

# ── 11. EFFECTS ──
section('11. Effects', [])
callout('⚠️', 'FX Deck Routing:', 'Use the A/B buttons in the FX panel to select which deck receives the effects.',
        COL_BG_YEL, COL_YELLOW)
for name, desc, knobs in [
    ('REVERB',  'Adds room/hall ambience.',
     [['Wet','How much reverb is mixed in'],['Size','Size of the virtual room']]),
    ('DELAY',   'Creates repeating echoes.',
     [['Time','Delay time 0–1.5 sec'],['Feedback','How many repeats (0=once, max=infinite)']]),
    ('FILTER',  'Bandpass filter — great for buildups.',
     [['Freq','Center frequency 80 Hz–16 kHz'],['Q','Filter width (higher = narrower)']]),
    ('FLANGER', 'Swirling modulation effect.',
     [['Rate','Speed of oscillation'],['Depth','Intensity']]),
    ('ECHO',    'Single-repeat accent echo.',
     [['Beat Div','Echo timing relative to beat']]),
]:
    sub(name)
    story.append(p(desc))
    data_table(['Knob', 'Function'], knobs, col_widths=[4*cm, 12.5*cm])

# ── 12. YOUTUBE ──
section('12. YouTube Search & MP3 Download', [])
sub('Getting a Free YouTube API Key')
steps([
    ('Go to Google Cloud Console',  'Visit console.cloud.google.com and sign in with your Google account.'),
    ('Create a new project',        'Click project dropdown → New Project → name it "DJ Mixer" → Create.'),
    ('Enable YouTube Data API v3',  'APIs & Services → Library → search "YouTube Data API v3" → Enable.'),
    ('Create credentials',          'APIs & Services → Credentials → Create Credentials → API Key. Copy the key.'),
    ('Paste it in ProMix',          'Paste into the API Key field in the header. Searches are now live.'),
])
callout('ℹ️', 'Free quota:', '10,000 search units per day — roughly 100 searches. No credit card required.',
        COL_BG_BLU, COL_BLUE)
sub('Downloading MP3s')
steps([
    ('Search for a track',     'Type in the search bar and press Enter.'),
    ('Click ⬇ MP3',             'The app requests a download from cobalt.tools (free, no account).'),
    ('The file downloads',     'Saved to your Downloads folder.'),
    ('Load into ProMix',       'Use 📂 Load to load your downloaded MP3.'),
])
callout('⚠️', 'Important:', 'Only download music you have the right to use. Respect copyright law.',
        COL_BG_YEL, COL_YELLOW)
story.append(PageBreak())

# ── 13. RECORDING ──
section('13. Recording Your Mix', [])
steps([
    ('Click ● REC before mixing',      'Button turns red and pulses. Recording starts.'),
    ('Perform your mix',               'Everything you do is captured.'),
    ('Click ⏹ STOP REC when finished', 'Mix is stored in memory.'),
    ('Click ⬇ Download Mix',           'File downloads as ProMix_[timestamp].webm.'),
])
sub('Converting to MP3')
for txt in [
    'VLC Media Player — Media → Convert/Save → Add file → choose MP3',
    'Audacity — File → Import → export as MP3',
    'Online: cloudconvert.com (free)',
    'FFmpeg: ffmpeg -i ProMix_123.webm output.mp3',
]:
    bullet(txt)
story.append(Spacer(1, 8))

# ── 14. AUTO MIX ──
section('14. Auto Mix', [
    p('⚡ Auto Mix automates the transition: syncs BPM and smoothly moves the crossfader over ~4 seconds.')
])
steps([
    ('Load tracks on both decks', 'Make sure both tracks are loaded and BPM is detected.'),
    ('Start playing Deck A',      'Let it reach the point where you want to mix out.'),
    ('Click ⚡ Auto Mix',          'App syncs Deck B to Deck A\'s BPM, starts Deck B, then fades the crossfader.'),
])

# ── 15. STEMS ──
section('15. Stem Separation', [
    p('Stems are the individual components of a song: Vocals, Drums, Bass, and Melody. '
      'Separate them to create acapellas, instrumentals, or mashups.')
])
steps([
    ('Click ⚗ Separate Stems (AI)',  'Opens Fadr.com in a new tab.'),
    ('Upload your track to Fadr',    'Drag and drop your audio file.'),
    ('Wait for AI processing',       'Under 60 seconds. Splits into up to 16 stems.'),
    ('Download stems',               'Download acapella, instrumental, or any individual stem.'),
    ('Load stems into ProMix',       'Load acapella on Deck A + instrumental on Deck B = instant mashup!'),
])

# ── 16. SYNC & PITCH ──
section('16. Sync & Pitch', [])
data_table(['Control', 'Function'], [
    ['SYNC Button',   'Matches this deck\'s BPM to the other deck'],
    ['Pitch Slider',  'Adjusts speed –8% to +8%'],
    ['Key Shift –/+', 'Shifts pitch one semitone per click'],
], col_widths=[4.5*cm, 12*cm])
story.append(PageBreak())

# ── 17. KEYBOARD SHORTCUTS ──
section('17. Keyboard Shortcuts', [p('Shortcuts do not fire when cursor is in a text input.')])
data = [['Key', 'Action']] + [
    ['Space', 'Play / Pause Deck A'],
    ['Enter', 'Play / Pause Deck B'],
    ['←',     'Seek Deck A back 5 seconds'],
    ['→',     'Seek Deck A forward 5 seconds'],
    ['Q',     'Set/Jump Cue on Deck A'],
    ['W',     'Set/Jump Cue on Deck B'],
    ['S',     'Sync Deck B to Deck A\'s BPM'],
    ['A',     'Toggle Mix Recording'],
]
kb_t = Table(data, colWidths=[3*cm, 13.5*cm], repeatRows=1)
kb_t.setStyle(TableStyle([
    ('BACKGROUND',  (0,0), (-1,0), COL_THEAD),
    ('TEXTCOLOR',   (0,0), (-1,0), white),
    ('FONTNAME',    (0,0), (-1,0), 'Helvetica-Bold'),
    ('FONTNAME',    (0,1), (0,-1), 'Courier-Bold'),
    ('TEXTCOLOR',   (0,1), (0,-1), COL_PURPLE),
    ('BACKGROUND',  (0,1), (0,-1), COL_BG_PUR),
    ('ROWBACKGROUNDS', (1,1), (-1,-1), [white, COL_ROW_ALT]),
    ('GRID', (0,0), (-1,-1), 0.5, HexColor('#E5E7EB')),
    ('ALIGN', (0,0), (0,-1), 'CENTER'),
    ('FONTSIZE', (0,0), (-1,-1), 10),
    ('TOPPADDING', (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING', (0,0), (-1,-1), 8),
]))
story.append(kb_t)
story.append(PageBreak())

# ── TUTORIALS ──
section('18. Tutorial 1 — Your First Mix', [p('Goal: Load two tracks and perform a manual crossfader mix. Time: ~5 minutes.', 'italic')])
steps([
    ('Open dj-mixer.html in Chrome',          'Double-click the file. Loads instantly.'),
    ('Load a track on Deck A',                'Click 📂 Load. Choose an upbeat song with a steady kick drum.'),
    ('Load a different track on Deck B',      'Choose something with similar energy and tempo.'),
    ('Move crossfader all the way left (A)',  'Deck A is the only thing heard.'),
    ('Press Space to play Deck A',            'Listen and enjoy.'),
    ('Click SYNC on Deck B',                  'Matches Deck B\'s tempo to Deck A.'),
    ('Position Deck B at the start',          'Click far left of Deck B\'s waveform.'),
    ('Start Deck B at a phrase boundary',     'When Deck A reaches a chorus, press Enter.'),
    ('Slowly slide crossfader to the right',  'Over 8–16 beats, fade from A to B. First mix done!'),
    ('Pause Deck A',                          'Load your next track and repeat.'),
])

section('19. Tutorial 2 — Party Set with Recording', [p('Goal: Record a 3-track mix with effects. Time: ~15 minutes.', 'italic')])
steps([
    ('Prepare your tracklist',              'Choose 3 songs ~120–130 BPM. Check Camelot keys for compatibility.'),
    ('Click ● REC to start recording',     'Button pulses red.'),
    ('Load Track 1 on Deck A and play it', 'Press Space. Let it run for a minute.'),
    ('Load Track 2 on Deck B',             'Sync it. Position at intro or drop.'),
    ('Apply a build-up effect',            'Enable FILTER on Deck A. Sweep Freq knob upward.'),
    ('Cut the bass on Deck A (LO knob)',   'Drag Deck A\'s LO knob counter-clockwise.'),
    ('Start Deck B and move crossfader',   'Slowly push crossfader right over 8 bars.'),
    ('Load Track 3 on Deck A and repeat',  'Deck A is free. Repeat the transition process.'),
    ('STOP REC and Download Mix',          'Your 3-track mix is saved.'),
])
story.append(PageBreak())

section('20. Tutorial 3 — YouTube to Mix', [p('Goal: Search YouTube, download two MP3s, and mix them. Requires API key.', 'italic')])
steps([
    ('Enter your YouTube API key',     'Paste into the header API key field.'),
    ('Search for your first song',     'E.g., "Calvin Harris Summer". Press Enter.'),
    ('Click ⬇ MP3 and save',           'Download the file to your computer.'),
    ('Search for your second song',    'E.g., "Avicii Wake Me Up". Download as MP3.'),
    ('Load both MP3s into the decks',  '📂 Load → Track 1 to Deck A, Track 2 to Deck B.'),
    ('Check harmonic compatibility',   'Open the Camelot Wheel. Use Key Shift if needed.'),
    ('Record and download your mix',   'Click ● REC, mix, then ⬇ Download Mix.'),
])

section('21. Tutorial 4 — Harmonic Mixing', [p('Goal: Build a harmonically perfect set using the Camelot Wheel.', 'italic')])
steps([
    ('Load your first track on Deck A',       'Note the KEY display. Example: 8B (C major).'),
    ('Open the Camelot Wheel',                'Find 8B. Compatible: 8A (A minor), 7B (F major), 9B (G major).'),
    ('Find a track in a compatible key',      'Load candidates on Deck B, check detected key.'),
    ('Use Key Shift if close but not exact',  'If Deck B shows 9A instead of 8A, click – once.'),
    ('Mix in Deck B',                         'The transition sounds musically natural.'),
    ('Plan ahead around the wheel',           'From 8A, next options are 8B, 7A, 9A — take the audience on a journey.'),
])
story.append(PageBreak())

# ── FAQ ──
section('22. FAQ', [])
for q, a in [
    ('No sound is coming out',
     'Browsers require a user click before playing audio. Click anywhere on the page first. '
     'Check system volume and that the browser tab is not muted.'),
    ('BPM is half or double the actual tempo',
     'Use the TAP button to correct it manually — tap 4–8 times on the beat.'),
    ('YouTube search returns an error',
     'API key may be invalid or over quota (10,000 units/day free). Check Google Cloud Console.'),
    ('MP3 download does not start',
     'cobalt.tools may be temporarily busy. Try again, or click ▶ YT and use another download tool.'),
    ('Can I use this offline?',
     'Yes — all core mixing features work offline. Internet only needed for YouTube search/download.'),
    ('Can I save hot cues between sessions?',
     'Not currently. Hot cues and settings reset when you refresh the page.'),
    ('The file does not open on my computer',
     'Right-click dj-mixer.html → Open with → choose Chrome, Firefox, or Edge.'),
]:
    story.append(Paragraph(q, S['h3']))
    story.append(p(a))

# ── TROUBLESHOOTING ──
section('23. Troubleshooting', [])
data_table(['Issue', 'Solution'], [
    ['App is silent',              'Click anywhere on page. Check volume. Unmute browser tab.'],
    ['Waveform does not appear',   'Wait 2–3 sec. Try converting to MP3 if it persists.'],
    ['BPM says "--"',              'Use TAP tempo to set it manually.'],
    ['Key says "--"',              'Click the middle of the waveform then reload.'],
    ['Effects have no effect',     'Check FX toggle is active (purple) AND correct deck is selected.'],
    ['YouTube search is blank',    'Verify API key. Try disabling VPN or ad-blocker.'],
    ['Download Mix is tiny',       'You must click ● REC before starting your mix.'],
    ['Tracks out of sync',         'Correct BPM with TAP, then click SYNC again.'],
    ['App slow on long tracks',    'Use MP3 format. Very long WAV files may be slow.'],
], col_widths=[5.5*cm, 11*cm])

# Footer
story.append(Spacer(1, 1*cm))
story.append(HRFlowable(width='100%', thickness=1, color=COL_LAVEND))
story.append(Spacer(1, 6))
story.append(Paragraph(
    '<b>ProMix DJ Studio</b> — User Manual v1.0<br/>'
    'Built on Web Audio API  ·  MP3 download via cobalt.tools  ·  Stem separation via Fadr<br/>'
    'Open source  ·  No installation required  ·  Works in any modern browser',
    S['center']
))

doc_pdf.build(story)
print(f'PDF saved: {pdf_path}')
print('Done!')
